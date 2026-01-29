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


def guess_mime(filename: str) -> str:
    fn = (filename or "").lower()
    if fn.endswith(".png"):
        return "image/png"
    if fn.endswith(".webp"):
        return "image/webp"
    return "image/jpeg"


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
# - normalize \(...\), \[...\] to $...$, $$...$$
# - optional: convert \tag{1} to right-side text if needed
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


def downscale_image_to_png_bytes(img: Image.Image, max_side: int) -> bytes:
    img = img.convert("RGB")
    w, h = img.size
    scale = min(1.0, max_side / max(w, h))
    if scale < 1.0:
        img = img.resize((int(w * scale), int(h * scale)))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


# -----------------------------
# DOCX mode A: Editable equations via pandoc (LaTeX -> OMML)
# -----------------------------
def markdown_to_docx_editable(md_text: str) -> bytes:
    # pandoc wants a file output
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as f:
        out_path = f.name

    try:
        # markdown+tex_math_dollars tells pandoc to treat $...$/$$...$$ as math
        pypandoc.convert_text(
            md_text,
            to="docx",
            format="markdown+tex_math_dollars+raw_tex",
            outputfile=out_path,
            extra_args=[
                "--wrap=none",
            ],
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

    # very simple line-based writing: preserve text and $$ blocks as is
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
st.set_page_config(page_title="多图 OCR → Word(可编辑公式)", layout="wide")
st.title("多图 OCR（Gemini）→ Word 文档（公式可编辑/或 LaTeX 原样）")

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
        help="可编辑公式模式会用 pandoc 把 LaTeX 转成 Word 原生公式；纯文本模式只保留 $$...$$。",
    )
    max_side = st.slider("最大边长（下采样提速）", 800, 3500, 2000, 50)
    keep_tag = st.checkbox("把 \\tag{n} 变成公式末尾编号（更稳）", value=True)

if not files:
    st.stop()

# Preview images
with st.expander("预览上传的图片", expanded=True):
    for f in files:
        st.image(Image.open(f), caption=f.name, use_container_width=True)

if st.button("开始识别并生成 Word", type="primary"):
    all_md: List[str] = []
    with st.spinner("正在逐张识别（多图）..."):
        for idx, f in enumerate(files, start=1):
            img = Image.open(f)
            img_bytes = downscale_image_to_png_bytes(img, max_side=max_side)
            md = gemini_ocr_one(img_bytes, "image/png", model_id)
            md = normalize_math_delimiters(md)
            if keep_tag:
                md = tag_to_text_in_equation(md)

            # add a page title to keep order
            all_md.append(f"## 第 {idx} 页\n\n{md}\n")

    # Add pagebreak between images for docx
    merged_md = "\n\n\\newpage\n\n".join(all_md)

    st.subheader("合并后的识别结果（Markdown）")
    st.code(merged_md, language="markdown")

    with st.spinner("生成 Word 文档..."):
        if export_mode == "可编辑公式（推荐）":
            try:
                docx_bytes = markdown_to_docx_editable(merged_md)
            except Exception as e:
                st.warning(
                    "可编辑公式转换失败（通常是 pandoc 环境问题）。已自动改为 LaTeX 原样（纯文本）导出。\n"
                    f"错误信息：{e}"
                )
                docx_bytes = docx_plain_latex(merged_md)
        else:
            docx_bytes = docx_plain_latex(merged_md)

    st.success("已生成 Word 文档。")
    st.download_button("下载 result.docx", data=docx_bytes, file_name="result.docx")
    st.download_button("下载 result.md", data=merged_md.encode("utf-8"), file_name="result.md")
