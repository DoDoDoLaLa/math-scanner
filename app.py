import io
import os
import re
import tempfile
from typing import List, Tuple

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
# Preprocess Markdown for better docx conversion
# -----------------------------
def normalize_math_delimiters(md: str) -> str:
    md = md.replace(r"\(", "$").replace(r"\)", "$")
    md = md.replace(r"\[", "$$").replace(r"\]", "$$")
    return md


TAG_RE = re.compile(r"\\tag\{([^}]+)\}")


def tag_to_text_in_equation(md: str) -> str:
    """
    pandoc 对 \\tag 兼容性不稳定。
    这里把 \\tag{1} 变成 \\qquad (1) 放在公式末尾，保证编号能显示。
    """
    def repl(m):
        t = m.group(1).strip()
        if not (t.startswith("(") and t.endswith(")")):
            t = f"({t})"
        return rf"\qquad {t}"

    return TAG_RE.sub(repl, md)


# -----------------------------
# Image processing: downscale + optional vertical tiling for long images
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
    将长图按竖向切片。每片高度 tile_h，片与片之间重叠 overlap（避免切到行/公式）。
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
    对“重叠切片”拼接时的重复内容做轻量去重：
    如果 prev 末尾若干行 与 next 开头若干行完全一致（忽略空白差异），就删掉 next 的重复开头。
    """
    prev_lines_raw = prev_md.splitlines()
    next_lines_raw = next_md.splitlines()

    prev_lines = [_norm_line(x) for x in prev_lines_raw if _norm_line(x)]
    next_lines = [_norm_line(x) for x in next_lines_raw if _norm_line(x)]

    if not prev_lines or not next_lines:
        return next_md

    k = min(max_check_lines, len(prev_lines), len(next_lines))
    # 从长到短匹配
    for m in range(k, 1, -1):
        if prev_lines[-m:] == next_lines[:m]:
            # 删除 next_md 中对应的前 m 个“非空行”
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
# DOCX mode A: Editable equations via pandoc (LaTeX -> OMML)
# -----------------------------
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
        with open(out_path, "rb") as r:
            return r.read()
    finally:
        try:
            os.remove(out_path)
        except Exception:
            pass


# -----------------------------
# DOCX mode B: Keep LaTeX as plain text in docx (no images)
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


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="多图 OCR → Word(可编辑公式) + 长图切片", layout="wide")
st.title("多图 OCR（Gemini）→ Word 文档（公式可编辑 / 或 LaTeX 原样）")

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
        "导出 Word 方式",
        ["可编辑公式（推荐）", "LaTeX 原样（纯文本）"],
        index=0,
    )
    max_side = st.slider("最大边长（下采样提速）", 800, 3500, 2000, 50)
    keep_tag = st.checkbox("把 \\tag{n} 变成公式末尾编号（更稳）", value=True)

    # ✅ 新增：长图切片控制（不影响其他功能）
    st.markdown("**长图切片（解决超长图片只识别上半部分）**")
    enable_tiling = st.checkbox("对超长图片自动切片", value=True)
    tile_h = st.slider("单片高度（像素，按下采样后的尺寸）", 700, 2000, 1200, 50)
    overlap = st.slider("相邻切片重叠（像素）", 0, 400, 120, 10)

if not files:
    st.stop()

with st.expander("预览上传的图片", expanded=True):
    for f in files:
        st.image(Image.open(f), caption=f.name, use_container_width=True)

if st.button("开始识别并生成 Word", type="primary"):
    all_md: List[str] = []

    with st.spinner("正在逐张识别（支持长图切片）..."):
        for img_idx, f in enumerate(files, start=1):
            img = Image.open(f)
            img = downscale_image(img, max_side=max_side)

            # ✅ 新增：对长图切片
            tiles = [img]
            if enable_tiling:
                # 超过阈值才切（避免短图也切）
                if img.size[1] > tile_h:
                    tiles = split_long_image_vertical(img, tile_h=tile_h, overlap=overlap)

            page_parts: List[str] = []

            for part_idx, tile in enumerate(tiles, start=1):
                tile_bytes = pil_to_png_bytes(tile)
                md_part = gemini_ocr_one(tile_bytes, "image/png", model_id)

                md_part = normalize_math_delimiters(md_part)
                if keep_tag:
                    md_part = tag_to_text_in_equation(md_part)

                # 轻量去重：处理切片重叠造成的重复段落
                if page_parts:
                    md_part = dedup_overlap_by_lines(page_parts[-1], md_part, max_check_lines=8)

                page_parts.append(md_part)

            # 合并该“第 img_idx 张图片”的所有切片 OCR 结果
            page_md = "\n\n".join([p for p in page_parts if p.strip()])

            # 保持你原来的分页与顺序
            all_md.append(f"## 第 {img_idx} 页\n\n{page_md}\n")

    merged_md = "\n\n\\newpage\n\n".join(all_md)

    st.subheader("合并后的识别结果（Markdown）")
    st.code(merged_md, language="markdown")

    with st.spinner("生成 Word 文档..."):
        if export_mode == "可编辑公式（推荐）":
            try:
                docx_bytes = markdown_to_docx_editable(merged_md)
            except Exception as e:
                st.warning(
                    "可编辑公式转换失败（pandoc 环境问题）。已自动改为 LaTeX 原样（纯文本）导出。\n"
                    f"错误信息：{e}"
                )
                docx_bytes = docx_plain_latex(merged_md)
        else:
            docx_bytes = docx_plain_latex(merged_md)

    st.success("已生成 Word 文档。")
    st.download_button("下载 result.docx", data=docx_bytes, file_name="result.docx")
    st.download_button("下载 result.md", data=merged_md.encode("utf-8"), file_name="result.md")
