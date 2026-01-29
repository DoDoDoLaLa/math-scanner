# app.py
import io
import os
import re
import tempfile
from typing import List, Tuple, Optional

import streamlit as st
from PIL import Image

from google import genai
from google.genai import types

import pypandoc
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


# -----------------------------
# Gemini OCR prompt (Chinese + LaTeX in-place)
# -----------------------------
OCR_PROMPT_ZH = """你是一个用于学术文档的 OCR 引擎。

要求：
1) 把图片中所有可见文字逐字转写（保持中文，不要翻译，不要改写）。
2) 所有数学表达式必须转为 LaTeX，并保持在原本位置（行内/独立公式都要正确）。
3) 只输出 Markdown（不要输出解释）。
   - 行内公式必须用 $...$
   - 独立居多行公式必须用 $$...$$（单独成段）
4) 保持原有阅读顺序、换行、项目符号、标题层级。
5) 看不清的内容用 [UNK]，不要猜。

注意：
- 不要给每个英文字母都强行加 $，只在确实是数学符号/变量时才用。
- 如果原图有公式编号（例如 (1)(2) 或 \\tag{1}），请保留编号信息。
只输出 Markdown。
"""


# -----------------------------
# Pandoc availability helper
# -----------------------------
def ensure_pandoc_available() -> Tuple[bool, str]:
    """
    Returns (ok, message).
    Tries to detect pandoc; if missing, tries to download via pypandoc.
    """
    try:
        path = pypandoc.get_pandoc_path()
        if path and os.path.exists(path):
            return True, f"pandoc 已可用：{path}"
    except Exception:
        pass

    try:
        pypandoc.download_pandoc()
        path = pypandoc.get_pandoc_path()
        if path and os.path.exists(path):
            return True, f"已自动下载 pandoc：{path}"
        return False, "已尝试下载 pandoc，但仍不可用。"
    except Exception as e:
        return False, f"未检测到 pandoc，自动下载失败：{e}"


# -----------------------------
# Gemini OCR (cached)
# -----------------------------
@st.cache_data(show_spinner=False)
def gemini_ocr_one(image_bytes: bytes, mime_type: str, model_id: str) -> str:
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("缺少 GEMINI_API_KEY：请在 Streamlit Cloud 的 Secrets 中配置。")

    client = genai.Client(api_key=api_key)

    resp = client.models.generate_content(
        model=model_id,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            OCR_PROMPT_ZH,
        ],
        config=types.GenerateContentConfig(
            temperature=0.0,
            top_p=1.0,
            top_k=1,
            max_output_tokens=8192,
        ),
    )
    return (resp.text or "").strip()


# -----------------------------
# Preprocess Markdown for better conversions
# -----------------------------
def normalize_math_delimiters(md: str) -> str:
    md = md.replace(r"\(", "$").replace(r"\)", "$")
    md = md.replace(r"\[", "$$").replace(r"\]", "$$")
    return md


TAG_RE = re.compile(r"\\tag\{([^}]+)\}")


def tag_to_text_in_equation(md: str) -> str:
    """
    pandoc 对 \\tag 兼容性不稳定：
    把 \\tag{1} 变成 \\qquad (1) 放在公式末尾。
    仅当你勾选“更稳”时才使用；不勾选则保留 \\tag 原样进入代码环境。
    """
    def repl(m):
        t = m.group(1).strip()
        if not (t.startswith("(") and t.endswith(")")):
            t = f"({t})"
        return rf"\qquad {t}"

    return TAG_RE.sub(repl, md)


# -----------------------------
# Image processing: downscale + vertical tiling
# -----------------------------
def downscale_image(img: Image.Image, max_side: int) -> Image.Image:
    img = img.convert("RGB")
    w, h = img.size
    scale = min(1.0, max_side / max(w, h))
    if scale < 1.0:
        img = img.resize((int(w * scale), int(h * scale)))
    return img


def pil_to_png_bytes(img: Image.Image) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def split_long_image_vertical(img: Image.Image, tile_h: int, overlap: int) -> List[Image.Image]:
    """
    将长图按竖向切片。每片高度 tile_h，片与片之间重叠 overlap。
    """
    w, h = img.size
    if h <= tile_h:
        return [img]

    overlap = max(0, min(overlap, tile_h // 2))
    step = tile_h - overlap
    tiles: List[Image.Image] = []

    y0 = 0
    while y0 < h:
        y1 = min(y0 + tile_h, h)
        tile = img.crop((0, y0, w, y1))
        tiles.append(tile)
        if y1 >= h:
            break
        y0 = y0 + step

    return tiles


def _norm_line(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def dedup_overlap_by_lines(prev_md: str, next_md: str, max_check_lines: int = 8) -> str:
    """
    重叠切片拼接去重：若 prev 末尾若干行 == next 开头若干行（忽略空白差异），删除 next 重复开头。
    """
    prev_lines_raw = prev_md.splitlines()
    next_lines_raw = next_md.splitlines()

    prev_lines = [_norm_line(x) for x in prev_lines_raw if _norm_line(x)]
    next_lines = [_norm_line(x) for x in next_lines_raw if _norm_line(x)]

    if not prev_lines or not next_lines:
        return next_md

    k = min(max_check_lines, len(prev_lines), len(next_lines))
    for m in range(k, 1, -1):
        if prev_lines[-m:] == next_lines[:m]:
            new_lines = []
            removed = 0
            for line in next_lines_raw:
                if removed < m and _norm_line(line):
                    removed += 1
                    continue
                new_lines.append(line)
            return "\n".join(new_lines).lstrip("\n")

    return next_md


# -----------------------------
# Convert $$...$$ to LaTeX equation environment inside ```latex code fence
# -----------------------------
DISPLAY_MATH_RE = re.compile(r"\$\$(.+?)\$\$", re.DOTALL)


def split_by_fenced_code(md: str) -> List[Tuple[str, str]]:
    """
    Split markdown into [(type, text)], type in {"text","code"}.
    Fenced code blocks are preserved as "code".
    """
    lines = md.splitlines(True)
    out: List[Tuple[str, str]] = []
    buf: List[str] = []
    in_code = False

    for ln in lines:
        if ln.strip().startswith("```"):
            if buf:
                out.append(("code" if in_code else "text", "".join(buf)))
                buf = []
            buf.append(ln)
            in_code = not in_code
            if not in_code:
                out.append(("code", "".join(buf)))
                buf = []
        else:
            buf.append(ln)

    if buf:
        out.append(("code" if in_code else "text", "".join(buf)))
    return out


def display_math_to_equation_env_code(md: str) -> str:
    """
    Convert display math $$...$$ into:
    ```latex
    \\begin{equation}
    ...
    \\end{equation}
    ```
    Does NOT touch existing fenced code blocks.
    """
    parts = split_by_fenced_code(md)
    new_parts: List[str] = []

    for typ, txt in parts:
        if typ == "code":
            new_parts.append(txt)
            continue

        def repl(m):
            inner = m.group(1)
            inner = inner.strip("\n").strip()
            return (
                "\n\n```latex\n"
                "\\begin{equation}\n"
                f"{inner}\n"
                "\\end{equation}\n"
                "```\n\n"
            )

        converted = DISPLAY_MATH_RE.sub(repl, txt)
        new_parts.append(converted)

    return "".join(new_parts)


# -----------------------------
# DOCX export helpers
# -----------------------------
def set_doc_style(doc: Document):
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")


def markdown_to_docx_editable(md_text: str) -> bytes:
    """
    Version A: editable equations via pandoc (LaTeX -> OMML) if pandoc exists.
    """
    ok, _ = ensure_pandoc_available()
    if not ok:
        raise RuntimeError("pandoc 不可用，无法生成“可编辑公式”的 docx。")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        out_path = f.name
    try:
        pypandoc.convert_text(
            md_text,
            to="docx",
            format="markdown+tex_math_dollars+raw_tex",
            outputfile=out_path,
            extra_args=["--wrap=none"],
        )
        with open(out_path, "rb") as r:
            return r.read()
    finally:
        try:
            os.remove(out_path)
        except Exception:
            pass


def docx_plain_latex(md_text: str) -> bytes:
    """
    Version B: keep LaTeX as plain text in docx (no math objects).
    """
    doc = Document()
    set_doc_style(doc)

    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for line in lines:
        p = doc.add_paragraph(line)
        if line.strip().startswith("$$") or line.strip().endswith("$$"):
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


def docx_code_from_fenced_markdown(md_text: str) -> bytes:
    """
    Turn fenced code blocks into monospace paragraphs in docx.
    - Lines inside ```...``` become Consolas monospace.
    - Fence markers are not printed.
    """
    doc = Document()
    set_doc_style(doc)

    md = md_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = md.split("\n")

    in_code = False
    for line in lines:
        stripped = line.strip()

        if stripped.startswith("```"):
            in_code = not in_code
            doc.add_paragraph("")  # spacing
            continue

        if not in_code and stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            title = stripped[level:].strip()
            p = doc.add_paragraph()
            run = p.add_run(title)
            run.bold = True
            run.font.size = Pt(16 if level == 1 else 14 if level == 2 else 12)
            continue

        if in_code:
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.name = "Consolas"
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Consolas")
            run.font.size = Pt(10.5)
        else:
            doc.add_paragraph(line)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# -----------------------------
# DOCX -> Markdown (pandoc) for Word upload feature
# -----------------------------
def docx_bytes_to_markdown_via_pandoc(docx_bytes: bytes) -> Tuple[Optional[str], str]:
    ok, msg = ensure_pandoc_available()
    if not ok:
        return None, f"无法使用 pandoc 将 docx 转为 Markdown：{msg}"

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        in_path = f.name
        f.write(docx_bytes)

    try:
        md = pypandoc.convert_file(
            in_path,
            to="markdown",
            format="docx",
            extra_args=["--wrap=none"],
        )
        md = (md or "").strip()
        return md, "已使用 pandoc 将 Word 转为 Markdown（包含公式的数学标记）。"
    except Exception as e:
        return None, f"pandoc 转换失败：{e}"
    finally:
        try:
            os.remove(in_path)
        except Exception:
            pass


def docx_bytes_to_plaintext_fallback(docx_bytes: bytes) -> str:
    """
    Fallback: extract only paragraph text using python-docx.
    NOTE: OMML 公式对象通常不会出现在 para.text 中，因此此模式可能拿不到公式。
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        p = f.name
        f.write(docx_bytes)
    try:
        d = Document(p)
        paras = [para.text for para in d.paragraphs]
        return "\n".join(paras).strip()
    finally:
        try:
            os.remove(p)
        except Exception:
            pass


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="OCR → Word（三版本：含 equation 代码版） + Word公式→equation代码替换", layout="wide")
st.title("多图 OCR（Gemini）→ Word（可编辑公式 / LaTeX纯文本 / equation代码版） + Word 公式替换为 equation LaTeX 代码")

tab1, tab2 = st.tabs(["① 图片OCR → 三种 Word 版本", "② 上传 Word：公式替换为 equation 代码"])


# =============================
# TAB 1: Image OCR
# =============================
with tab1:
    col1, col2 = st.columns([1, 1])

    with col1:
        files = st.file_uploader(
            "上传图片（可多选：png/jpg/webp）",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
            key="img_uploader",
        )
        model_id = st.selectbox(
            "模型（Flash 更快，Pro 更稳）",
            ["gemini-3-flash-preview", "gemini-3-pro-preview"],
            index=0,
            key="model_select",
        )

    with col2:
        max_side = st.slider("最大边长（下采样提速）", 800, 3500, 2000, 50, key="max_side")
        keep_tag = st.checkbox("把 \\tag{n} 变成公式末尾编号（更稳；不勾选则保留\\tag原样进入代码）", value=True, key="keep_tag")

        st.markdown("**长图切片（解决超长图片只识别上半部分）**")
        enable_tiling = st.checkbox("对超长图片自动切片", value=True, key="enable_tiling")
        tile_h = st.slider("单片高度（像素，按下采样后的尺寸）", 700, 2000, 1200, 50, key="tile_h")
        overlap = st.slider("相邻切片重叠（像素）", 0, 400, 120, 10, key="overlap")

    if files:
        with st.expander("预览上传的图片", expanded=True):
            for f in files:
                img = Image.open(io.BytesIO(f.getvalue()))
                st.image(img, caption=f.name, use_container_width=True)

        if st.button("开始识别并生成（三种版本）", type="primary", key="run_ocr"):
            all_md: List[str] = []

            with st.spinner("正在逐张识别（支持长图切片）..."):
                for img_idx, f in enumerate(files, start=1):
                    img = Image.open(io.BytesIO(f.getvalue()))
                    img = downscale_image(img, max_side=max_side)

                    tiles = [img]
                    if enable_tiling and img.size[1] > tile_h:
                        tiles = split_long_image_vertical(img, tile_h=tile_h, overlap=overlap)

                    page_parts: List[str] = []
                    for tile in tiles:
                        tile_bytes = pil_to_png_bytes(tile)
                        md_part = gemini_ocr_one(tile_bytes, "image/png", model_id)

                        md_part = normalize_math_delimiters(md_part)
                        if keep_tag:
                            md_part = tag_to_text_in_equation(md_part)

                        if page_parts:
                            md_part = dedup_overlap_by_lines(page_parts[-1], md_part, max_check_lines=8)

                        page_parts.append(md_part)

                    page_md = "\n\n".join([p for p in page_parts if p.strip()])
                    all_md.append(f"## 第 {img_idx} 页\n\n{page_md}\n")

            merged_md = "\n\n\\newpage\n\n".join(all_md)

            st.subheader("合并后的识别结果（Markdown）")
            st.code(merged_md, language="markdown")

            # --- Version A: editable equations (pandoc) ---
            docx_a = None
            a_msg = ""
            with st.spinner("生成版本A：可编辑公式（pandoc）..."):
                try:
                    docx_a = markdown_to_docx_editable(merged_md)
                    a_msg = "版本A（可编辑公式）生成成功。"
                except Exception as e:
                    docx_a = None
                    a_msg = f"版本A 生成失败（不影响版本B/C）：{e}"

            # --- Version B: LaTeX plain text ---
            with st.spinner("生成版本B：LaTeX 原样（纯文本）..."):
                docx_b = docx_plain_latex(merged_md)

            # --- Version C: equation environment code ---
            with st.spinner("生成版本C：LaTeX equation 环境代码（可复制）..."):
                md_equation_code = display_math_to_equation_env_code(merged_md)
                docx_c = docx_code_from_fenced_markdown(md_equation_code)

            st.success("三种版本已生成（A 若失败不影响 B/C）。")
            st.info(a_msg)

            st.download_button(
                "下载 result_editable.docx（版本A：可编辑公式）",
                data=(docx_a or b""),
                file_name="result_editable.docx",
                disabled=(docx_a is None),
            )
            st.download_button(
                "下载 result_plain_latex.docx（版本B：LaTeX纯文本）",
                data=docx_b,
                file_name="result_plain_latex.docx",
            )
            st.download_button(
                "下载 result_equation_code.docx（版本C：equation代码版）",
                data=docx_c,
                file_name="result_equation_code.docx",
            )

            st.download_button("下载 result.md（原始Markdown）", data=merged_md.encode("utf-8"), file_name="result.md")
            st.download_button(
                "下载 result_equation_code.md（公式→equation代码）",
                data=md_equation_code.encode("utf-8"),
                file_name="result_equation_code.md",
            )
    else:
        st.caption("请先上传图片。")


# =============================
# TAB 2: Word -> Replace formulas with equation LaTeX code
# =============================
with tab2:
    st.markdown("### 上传 Word（.docx）并把公式替换为 LaTeX equation 代码")
    st.caption("优先使用 pandoc 将 Word 公式（OMML）转为 Markdown 数学，再统一把 $$...$$ 转为 equation 环境代码。")

    word_file = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="word_uploader")
    keep_tag2 = st.checkbox("把 \\tag{n} 变成末尾编号（更稳；不勾选则保留\\tag原样进入代码）", value=True, key="keep_tag2")

    if word_file and st.button("开始转换：Word公式 → equation代码并导出", type="primary", key="run_word"):
        docx_bytes_in = word_file.getvalue()

        with st.spinner("将 Word 转为 Markdown（优先 pandoc）..."):
            md_from_docx, msg = docx_bytes_to_markdown_via_pandoc(docx_bytes_in)
            st.info(msg)

        if md_from_docx is None:
            with st.spinner("pandoc 不可用：退化为仅提取段落文本（Word 公式对象可能无法提取）..."):
                md_from_docx = docx_bytes_to_plaintext_fallback(docx_bytes_in)

        md_from_docx = normalize_math_delimiters(md_from_docx)
        if keep_tag2:
            md_from_docx = tag_to_text_in_equation(md_from_docx)

        md_equation_code = display_math_to_equation_env_code(md_from_docx)

        st.subheader("转换后的 Markdown（公式已变成 equation 环境代码）")
        st.code(md_equation_code, language="markdown")

        with st.spinner("生成新的 Word（公式为 equation LaTeX 代码段）..."):
            out_docx = docx_code_from_fenced_markdown(md_equation_code)

        st.success("已生成新 Word（公式替换为 equation LaTeX 代码）。")
        st.download_button("下载 word_equation_code.docx", data=out_docx, file_name="word_equation_code.docx")
        st.download_button("下载 word_equation_code.md", data=md_equation_code.encode("utf-8"), file_name="word_equation_code.md")
