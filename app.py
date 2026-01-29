import io
import os
import re
import tempfile
from typing import List

import streamlit as st
from PIL import Image

from google import genai
from google.genai import types

import pypandoc
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


# =============================
# 1) Gemini OCR prompt (Chinese + LaTeX in-place)
# =============================
OCR_PROMPT_ZH = """你是一个用于学术文档的 OCR 引擎。

要求：
1) 把图片中所有可见文字逐字转写（保持中文，不要翻译，不要改写）。
2) 所有数学表达式必须转为 LaTeX，并保持在原本位置（行内/独立公式都要正确）。
3) 只输出 Markdown（不要输出解释）。
   - 行内公式必须用 $...$
   - 独立居中公式必须用 $$...$$（单独成段）
4) 保持原有阅读顺序、换行、项目符号、标题层级。
5) 看不清的内容用 [UNK]，不要猜。

注意：
- 不要给每个英文字母都强行加 $，只在确实是数学符号/变量时才用。
- 如果原图有公式编号（例如 (1)(2) 或 \\tag{1}），请保留编号信息。
只输出 Markdown。
"""


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


# =============================
# 2) Markdown preprocess
# =============================
def normalize_math_delimiters(md: str) -> str:
    # \( \) -> $ $
    md = md.replace(r"\(", "$").replace(r"\)", "$")
    # \[ \] -> $$ $$
    md = md.replace(r"\[", "$$").replace(r"\]", "$$")
    return md


TAG_RE = re.compile(r"\\tag\{([^}]+)\}")


def tag_to_text_in_equation(md: str) -> str:
    """把 \\tag{1} 变成 \\qquad (1)，避免 pandoc 对 \\tag 不稳定。"""
    def repl(m):
        t = m.group(1).strip()
        if not (t.startswith("(") and t.endswith(")")):
            t = f"({t})"
        return rf"\qquad {t}"
    return TAG_RE.sub(repl, md)


# =============================
# 3) Long image tiling (vertical)
# =============================
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
    w, h = img.size
    if h <= tile_h:
        return [img]
    overlap = max(0, min(overlap, tile_h // 2))
    step = tile_h - overlap
    tiles: List[Image.Image] = []

    y0 = 0
    while y0 < h:
        y1 = min(y0 + tile_h, h)
        tiles.append(img.crop((0, y0, w, y1)))
        if y1 >= h:
            break
        y0 += step
    return tiles


def _norm_line(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def dedup_overlap_by_lines(prev_md: str, next_md: str, max_check_lines: int = 8) -> str:
    """处理切片重叠导致的重复行：prev尾部 == next头部，则删掉next重复头部。"""
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


# =============================
# 4) Export DOCX
#    mode A: editable equations via pandoc (LaTeX -> OMML)
#    mode B: LaTeX plain text
# =============================
def markdown_to_docx_editable(md_text: str) -> bytes:
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
        return open(out_path, "rb").read()
    finally:
        try:
            os.remove(out_path)
        except Exception:
            pass


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


def docx_plain_latex(md_text: str) -> bytes:
    doc = Document()
    set_doc_style(doc)
    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for line in lines:
        if line.strip().startswith("$$") or line.strip().endswith("$$"):
            p = doc.add_paragraph(line)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            doc.add_paragraph(line)
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# =============================
# 5) NEW: Convert math -> LaTeX code format (for 3rd version)
#    - display $$...$$ -> fenced code block ```latex ... ```
#    - inline $...$ -> inline code `$...$`
# =============================
INLINE_MATH_RE = re.compile(r"(?<!\$)\$([^$\n]+)\$(?!\$)")  # match single $...$

def md_math_to_latex_code(md_text: str) -> str:
    """
    把数学都变成“代码格式”：
      - $$ ... $$ 段 -> ```latex\n$$\n...\n$$\n```
      - $ ... $ 行内 -> `$...$`
    注意：如果你本来就有代码块，我们尽量不改动其内部。
    """
    lines = md_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    out_lines: List[str] = []
    in_fence = False
    i = 0

    while i < len(lines):
        line = lines[i]

        # code fence toggle
        if line.strip().startswith("```"):
            in_fence = not in_fence
            out_lines.append(line)
            i += 1
            continue

        if in_fence:
            out_lines.append(line)
            i += 1
            continue

        # handle display math blocks
        if line.strip().startswith("$$"):
            # single-line $$...$$
            if line.strip().endswith("$$") and len(line.strip()) > 4:
                content = line.strip()
                out_lines.append("```latex")
                out_lines.append(content)
                out_lines.append("```")
                i += 1
                continue

            # multi-line $$ block
            buf = []
            # remove leading $$
            buf.append(line.strip()[2:].strip())
            i += 1
            while i < len(lines):
                if lines[i].strip().endswith("$$"):
                    tail = lines[i].strip()[:-2].strip()
                    if tail:
                        buf.append(tail)
                    break
                buf.append(lines[i])
                i += 1
            i += 1  # skip ending $$ line

            out_lines.append("```latex")
            out_lines.append("$$")
            for b in buf:
                out_lines.append(b)
            out_lines.append("$$")
            out_lines.append("```")
            continue

        # inline math -> inline code
        def repl(m):
            inner = m.group(0)  # includes $...$
            return f"`{inner}`"

        out_lines.append(INLINE_MATH_RE.sub(repl, line))
        i += 1

    return "\n".join(out_lines)


def markdown_to_docx_plain(md_text: str) -> bytes:
    """把普通 markdown（不渲染数学）转 docx，适用于“公式已变成代码块”的第三版本。"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        out_path = f.name
    try:
        pypandoc.convert_text(
            md_text,
            to="docx",
            format="markdown+fenced_code+tables",
            outputfile=out_path,
            extra_args=["--wrap=none"],
        )
        return open(out_path, "rb").read()
    finally:
        try:
            os.remove(out_path)
        except Exception:
            pass


# =============================
# 6) NEW: DOCX upload -> markdown -> math->latex-code -> new docx
# =============================
@st.cache_data(show_spinner=False)
def docx_bytes_to_markdown(docx_bytes: bytes) -> str:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        f.write(docx_bytes)
        in_path = f.name
    try:
        md = pypandoc.convert_file(
            in_path,
            to="markdown",
            format="docx",
            extra_args=["--wrap=none"],
        )
        return (md or "").strip()
    finally:
        try:
            os.remove(in_path)
        except Exception:
            pass


# =============================
# UI
# =============================
st.set_page_config(page_title="OCR + 长图切片 + 3版本导出 + DOCX公式转LaTeX代码", layout="wide")
st.title("Gemini OCR（多图/长图切片）→ Word 导出（含第三版 LaTeX 代码） + 上传 Word 转 LaTeX 代码")

tabs = st.tabs(["① 图片 OCR → 导出", "② 上传 Word(.docx) → 公式转 LaTeX 代码"])

# ---------- Tab 1: Images ----------
with tabs[0]:
    col1, col2 = st.columns([1, 1])
    with col1:
        files = st.file_uploader(
            "上传图片（可多选：png/jpg/webp）",
            type=["png", "jpg", "jpeg", "webp"],
            accept_multiple_files=True,
        )
        model_id = st.selectbox(
            "模型（Flash 更快，Pro 更稳）",
            ["gemini-3-flash-preview", "gemini-3-pro-preview"],
            index=0,
        )

    with col2:
        export_mode = st.radio(
            "导出 Word 方式（前两版）",
            ["可编辑公式（推荐）", "LaTeX 原样（纯文本）"],
            index=0,
        )
        max_side = st.slider("最大边长（下采样提速）", 800, 3500, 2000, 50)
        keep_tag = st.checkbox("把 \\tag{n} 变成公式末尾编号（更稳）", value=True)

        st.markdown("**长图切片（解决超长图片只识别上半部分）**")
        enable_tiling = st.checkbox("对超长图片自动切片", value=True)
        tile_h = st.slider("单片高度（像素，按下采样后的尺寸）", 700, 2200, 1200, 50)
        overlap = st.slider("相邻切片重叠（像素）", 0, 500, 120, 10)

    if not files:
        st.stop()

    with st.expander("预览上传的图片", expanded=True):
        for f in files:
            st.image(Image.open(f), caption=f.name, use_container_width=True)

    if st.button("开始识别并生成（含第三版）", type="primary"):
        all_md: List[str] = []

        with st.spinner("正在逐张识别（支持长图切片）..."):
            for img_idx, f in enumerate(files, start=1):
                img = Image.open(f)
                img = downscale_image(img, max_side=max_side)

                tiles = [img]
                if enable_tiling and img.size[1] > tile_h:
                    tiles = split_long_image_vertical(img, tile_h=tile_h, overlap=overlap)

                page_parts: List[str] = []
                for tile in tiles:
                    md_part = gemini_ocr_one(pil_to_png_bytes(tile), "image/png", model_id)
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

        # --- Version 1/2: docx ---
        with st.spinner("生成前两版 Word..."):
            if export_mode == "可编辑公式（推荐）":
                try:
                    docx_bytes_v1 = markdown_to_docx_editable(merged_md)
                except Exception as e:
                    st.warning(f"可编辑公式转换失败，已自动降级为纯文本：{e}")
                    docx_bytes_v1 = docx_plain_latex(merged_md)
            else:
                docx_bytes_v1 = docx_plain_latex(merged_md)

            docx_bytes_v2 = docx_plain_latex(merged_md)  # 第二版固定：LaTeX原样（纯文本）

        # --- Version 3: latex code ---
        latex_code_md = md_math_to_latex_code(merged_md)
        with st.spinner("生成第三版（公式→LaTeX代码格式）..."):
            docx_bytes_v3 = markdown_to_docx_plain(latex_code_md)

        st.success("已生成 3 个版本。")
        st.download_button("下载 V1：result.docx（按你选择的方式）", data=docx_bytes_v1, file_name="result.docx")
        st.download_button("下载 V2：result_plain_latex.docx（LaTeX 原样纯文本）", data=docx_bytes_v2, file_name="result_plain_latex.docx")
        st.download_button("下载 V3：result_latex_code.docx（公式变 LaTeX 代码）", data=docx_bytes_v3, file_name="result_latex_code.docx")

        st.download_button("下载 result.md", data=merged_md.encode("utf-8"), file_name="result.md")
        st.download_button("下载 result_latex_code.md", data=latex_code_md.encode("utf-8"), file_name="result_latex_code.md")


# ---------- Tab 2: DOCX Upload ----------
with tabs[1]:
    st.markdown("上传你的 Word(.docx)。如果其中有 **Word 原生公式** 或文本形式的 `$$...$$`，会统一提取为 LaTeX 并替换成 **LaTeX 代码格式**，输出一个新的 Word。")

    docx_up = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], accept_multiple_files=False)
    keep_tag2 = st.checkbox("把 \\tag{n} 变成公式末尾编号（更稳）", value=True, key="keep_tag2")

    if docx_up and st.button("把 Word 里的公式替换成 LaTeX 代码格式", type="primary"):
        with st.spinner("读取 Word → 提取为 Markdown（含 LaTeX）..."):
            md_from_docx = docx_bytes_to_markdown(docx_up.read())
            md_from_docx = normalize_math_delimiters(md_from_docx)
            if keep_tag2:
                md_from_docx = tag_to_text_in_equation(md_from_docx)

        st.subheader("提取出的 Markdown（含 LaTeX）")
        st.code(md_from_docx, language="markdown")

        latex_code_md2 = md_math_to_latex_code(md_from_docx)
        with st.spinner("生成“公式为 LaTeX 代码”的新 Word..."):
            new_docx = markdown_to_docx_plain(latex_code_md2)

        st.success("已生成新文档（公式已替换成 LaTeX 代码格式）。")
        st.download_button("下载 new_latex_code.docx", data=new_docx, file_name="new_latex_code.docx")
        st.download_button("下载 extracted.md", data=md_from_docx.encode("utf-8"), file_name="extracted.md")
        st.download_button("下载 extracted_latex_code.md", data=latex_code_md2.encode("utf-8"), file_name="extracted_latex_code.md")
