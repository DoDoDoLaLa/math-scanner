import io
import re
import hashlib
from dataclasses import dataclass
from typing import List, Tuple, Optional

import streamlit as st
from PIL import Image

from google import genai
from google.genai import types

# DOCX
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

# LaTeX render -> PNG
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# -------------------------
# 1) Gemini OCR（中文 + LaTeX）
# -------------------------
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


def guess_mime(filename: str) -> str:
    fn = (filename or "").lower()
    if fn.endswith(".png"):
        return "image/png"
    if fn.endswith(".webp"):
        return "image/webp"
    return "image/jpeg"


@st.cache_data(show_spinner=False)
def gemini_ocr_markdown(image_bytes: bytes, mime_type: str, model_id: str) -> str:
    api_key = st.secrets.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError("缺少 GEMINI_API_KEY：请在 Streamlit Cloud 的 Secrets 中配置。")

    client = genai.Client(api_key=api_key)

    # 官方示例：types.Part.from_bytes + client.models.generate_content :contentReference[oaicite:3]{index=3}
    resp = client.models.generate_content(
        model=model_id,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            OCR_PROMPT_ZH,
        ],
        # OCR 场景建议 temperature=0
        config=types.GenerateContentConfig(
            temperature=0.0,
            top_p=1.0,
            top_k=1,
            max_output_tokens=8192,
        ),
    )
    return (resp.text or "").strip()


# -------------------------
# 2) Markdown 解析（保留原位：段落/标题/列表/行内公式/独立公式）
# -------------------------
@dataclass
class Block:
    kind: str  # "heading" | "ul" | "ol" | "para" | "display_math"
    text: str
    level: int = 0


def split_markdown_blocks(md: str) -> List[Block]:
    lines = md.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blocks: List[Block] = []

    i = 0
    while i < len(lines):
        line = lines[i]

        # skip leading blank
        if line.strip() == "":
            i += 1
            continue

        # display math block: $$ ... $$ (may span multiple lines)
        if line.strip().startswith("$$"):
            buf = []
            # if same line contains both start/end
            if line.strip().endswith("$$") and len(line.strip()) > 4:
                content = line.strip()[2:-2].strip()
                blocks.append(Block("display_math", content))
                i += 1
                continue

            # start multi-line
            start = line
            # remove leading $$
            buf.append(start.strip()[2:].strip())
            i += 1
            while i < len(lines):
                if lines[i].strip().endswith("$$"):
                    tail = lines[i].strip()[:-2].strip()
                    if tail:
                        buf.append(tail)
                    break
                buf.append(lines[i])
                i += 1
            blocks.append(Block("display_math", "\n".join(buf).strip()))
            i += 1
            continue

        # heading
        m = re.match(r"^(#{1,6})\s+(.*)$", line)
        if m:
            level = len(m.group(1))
            blocks.append(Block("heading", m.group(2).strip(), level=level))
            i += 1
            continue

        # ul / ol item
        m_ul = re.match(r"^[-*]\s+(.*)$", line)
        m_ol = re.match(r"^\d+\.\s+(.*)$", line)
        if m_ul:
            blocks.append(Block("ul", m_ul.group(1).strip()))
            i += 1
            continue
        if m_ol:
            blocks.append(Block("ol", m_ol.group(1).strip()))
            i += 1
            continue

        # paragraph: merge consecutive non-blank lines until next special block
        buf = [line]
        i += 1
        while i < len(lines):
            peek = lines[i]
            if peek.strip() == "":
                break
            if peek.strip().startswith("$$"):
                break
            if re.match(r"^(#{1,6})\s+", peek):
                break
            if re.match(r"^[-*]\s+", peek) or re.match(r"^\d+\.\s+", peek):
                break
            buf.append(peek)
            i += 1
        blocks.append(Block("para", "\n".join(buf).strip()))
        i += 1

    return blocks


INLINE_MATH_RE = re.compile(r"(\$[^$\n]+\$)")  # 简化版：覆盖大多数常见行内公式


def split_inline_math(text: str) -> List[Tuple[str, str]]:
    """
    return list of (kind, content)
    kind: "text" | "math"
    """
    parts: List[Tuple[str, str]] = []
    pos = 0
    for m in INLINE_MATH_RE.finditer(text):
        if m.start() > pos:
            parts.append(("text", text[pos:m.start()]))
        math = m.group(1)[1:-1]  # remove $ $
        parts.append(("math", math))
        pos = m.end()
    if pos < len(text):
        parts.append(("text", text[pos:]))
    return parts


def extract_tag(latex: str) -> Tuple[str, Optional[str]]:
    """
    支持 \tag{1} 或 \tag{(1)}。返回 (latex_without_tag, tag_text)
    """
    m = re.search(r"\\tag\{([^}]+)\}", latex)
    if not m:
        return latex, None
    tag = m.group(1).strip()
    latex2 = (latex[:m.start()] + latex[m.end():]).strip()
    # normalize number style
    if not (tag.startswith("(") and tag.endswith(")")):
        tag = f"({tag})"
    return latex2, tag


# -------------------------
# 3) LaTeX 渲染为 PNG（用于插入 Word）
# -------------------------
def render_latex_png(latex: str, display: bool, font_size: int = 16, dpi: int = 220) -> io.BytesIO:
    """
    用 matplotlib mathtext 渲染（不依赖系统 LaTeX）。
    display=True 时用 \displaystyle 增大公式风格。
    """
    latex = latex.strip()
    if display:
        expr = r"$\displaystyle " + latex + r"$"
    else:
        expr = r"$" + latex + r"$"

    fig = plt.figure(figsize=(0.01, 0.01), dpi=dpi)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    t = ax.text(0, 0, expr, fontsize=font_size)

    # draw to get bbox
    fig.canvas.draw()
    bbox = t.get_window_extent(renderer=fig.canvas.get_renderer()).expanded(1.10, 1.25)

    # set figure size to bbox
    w_in = bbox.width / dpi
    h_in = bbox.height / dpi
    fig.set_size_inches(w_in, h_in)

    # reposition text
    ax.cla()
    ax.axis("off")
    ax.text(0, 0, expr, fontsize=font_size)

    out = io.BytesIO()
    fig.savefig(out, format="png", dpi=dpi, transparent=True, bbox_inches="tight", pad_inches=0.02)
    plt.close(fig)
    out.seek(0)
    return out


# -------------------------
# 4) 生成 DOCX（公式“渲染后插入原位”）
# -------------------------
def set_doc_style(doc: Document):
    # Page margins
    sec = doc.sections[0]
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)

    # Normal style: 中文宋体 + 英文 Times
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    r = style._element.rPr.rFonts
    r.set(qn("w:eastAsia"), "宋体")

def add_text_with_inline_math(paragraph, text: str):
    # 保留换行：把段内 \n 拆成多个 run / 新段落
    lines = text.split("\n")
    for li, line in enumerate(lines):
        chunks = split_inline_math(line)
        for kind, content in chunks:
            if kind == "text":
                paragraph.add_run(content)
            else:
                # inline math -> render png and insert as inline picture
                try:
                    img_stream = render_latex_png(content, display=False, font_size=13)
                    run = paragraph.add_run()
                    run.add_picture(img_stream, height=Pt(14))  # 行内高度
                except Exception:
                    paragraph.add_run(f"${content}$")  # fallback
        if li != len(lines) - 1:
            paragraph.add_run("\n")


def add_display_equation(doc: Document, latex: str):
    latex, tag = extract_tag(latex)

    # 用两列表格实现“居中公式 + 右侧编号”
    if tag:
        table = doc.add_table(rows=1, cols=2)
        table.style = "Table Normal"
        left = table.cell(0, 0)
        right = table.cell(0, 1)

        # left: centered equation image
        p_left = left.paragraphs[0]
        p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            img_stream = render_latex_png(latex, display=True, font_size=16)
            run = p_left.add_run()
            run.add_picture(img_stream, height=Pt(28))
        except Exception:
            p_left.add_run(f"$$\n{latex}\n$$")

        # right: equation number
        p_right = right.paragraphs[0]
        p_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_right.add_run(tag)

        # spacing after table
        doc.add_paragraph("")
    else:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        try:
            img_stream = render_latex_png(latex, display=True, font_size=16)
            run = p.add_run()
            run.add_picture(img_stream, height=Pt(28))
        except Exception:
            p.add_run(f"$$\n{latex}\n$$")


def build_docx_from_markdown(md_text: str) -> bytes:
    doc = Document()
    set_doc_style(doc)

    blocks = split_markdown_blocks(md_text)

    for b in blocks:
        if b.kind == "heading":
            # level 1~3 用 heading，其他用加粗段落
            lvl = min(max(b.level, 1), 3)
            h = doc.add_heading(b.text, level=lvl)
            # 中文字体更美观
            for r in h.runs:
                r.font.name = "Times New Roman"
                r._element.rPr.rFonts.set(qn("w:eastAsia"), "微软雅黑")
        elif b.kind == "ul":
            p = doc.add_paragraph(style="List Bullet")
            add_text_with_inline_math(p, b.text)
        elif b.kind == "ol":
            p = doc.add_paragraph(style="List Number")
            add_text_with_inline_math(p, b.text)
        elif b.kind == "display_math":
            add_display_equation(doc, b.text)
        else:  # para
            p = doc.add_paragraph()
            add_text_with_inline_math(p, b.text)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# -------------------------
# 5) Streamlit UI（中文更美观）
# -------------------------
st.set_page_config(page_title="图片 OCR → Word(含公式渲染)", layout="wide")
st.title("图片 OCR（Gemini）→ Word 文档（公式渲染后原位插入）")

st.markdown(
    "上传图片后，系统会识别**中文正文 + 数学公式**，公式用 LaTeX 表达并**渲染后插入到 Word(.docx)** 的原本位置（行内/独立公式）。"
)

col1, col2 = st.columns([1, 1])
with col1:
    uploaded = st.file_uploader("上传图片（png/jpg/webp）", type=["png", "jpg", "jpeg", "webp"])
    model_id = st.selectbox(
        "选择模型（Flash 更快，Pro 更稳）",
        options=["gemini-3-flash-preview", "gemini-3-pro-preview"],
        index=0,
    )
with col2:
    st.info("提示：图片太大时可下采样，速度会明显提升；公式渲染会在导出 Word 时完成。")

if not uploaded:
    st.stop()

img = Image.open(uploaded).convert("RGB")
st.image(img, caption="已上传图片", use_container_width=True)

max_side = st.slider("最大边长（超过则下采样）", 800, 3500, 2000, 50)
w, h = img.size
scale = min(1.0, max_side / max(w, h))
if scale < 1.0:
    img = img.resize((int(w * scale), int(h * scale)))

buf = io.BytesIO()
img.save(buf, format="PNG")
image_bytes = buf.getvalue()
mime_type = "image/png"

if st.button("开始识别", type="primary"):
    with st.spinner("Gemini 识别中..."):
        md_out = gemini_ocr_markdown(image_bytes, mime_type, model_id)

    st.subheader("识别结果（Markdown + LaTeX）")
    st.code(md_out, language="markdown")

    with st.spinner("生成 Word(.docx)：渲染公式并原位插入..."):
        docx_bytes = build_docx_from_markdown(md_out)

    st.success("已生成 Word 文档（含渲染后的公式）。")
    st.download_button("下载 result.docx", data=docx_bytes, file_name="result.docx")
    st.download_button("下载 result.md", data=md_out.encode("utf-8"), file_name="result.md")