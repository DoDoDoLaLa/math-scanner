import hashlib
import re
from typing import Tuple

import streamlit as st
from PIL import Image
from google import genai
from google.genai import types

import markdown as md
import bleach


# ---------- LaTeX delimiter normalization ----------
INLINE_OPEN = r"\("
INLINE_CLOSE = r"\)"
BLOCK_OPEN = r"\["
BLOCK_CLOSE = r"\]"

def normalize_latex_delimiters(s: str) -> str:
    # Convert \( \) -> $ $
    s = s.replace(INLINE_OPEN, "$").replace(INLINE_CLOSE, "$")
    # Convert \[ \] -> $$ $$
    s = s.replace(BLOCK_OPEN, "$$").replace(BLOCK_CLOSE, "$$")
    return s


# ---------- Markdown -> Safe HTML (keep $...$ for MathJax) ----------
ALLOWED_TAGS = [
    "p","br","hr","em","strong","ul","ol","li","code","pre","blockquote",
    "h1","h2","h3","h4","h5","h6","table","thead","tbody","tr","th","td",
    "span","div"
]
ALLOWED_ATTRS = {
    "*": ["class", "style"],
    "th": ["colspan", "rowspan"],
    "td": ["colspan", "rowspan"],
}

def markdown_to_safe_html(markdown_text: str) -> str:
    html = md.markdown(markdown_text, extensions=["tables", "fenced_code"])
    clean = bleach.clean(html, tags=ALLOWED_TAGS, attributes=ALLOWED_ATTRS, strip=True)
    return clean


def wrap_with_mathjax(html_body: str) -> str:
    # MathJax config: use $...$ and $$...$$
    return f"""
<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<script>
window.MathJax = {{
  tex: {{
    inlineMath: [['$', '$']],
    displayMath: [['$$', '$$']],
    processEscapes: true
  }},
  options: {{
    skipHtmlTags: ['script','noscript','style','textarea','pre','code']
  }}
}};
</script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
<style>
body {{ font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; padding: 16px; }}
table {{ border-collapse: collapse; width: 100%; }}
th, td {{ border: 1px solid #ddd; padding: 6px 8px; }}
pre {{ background: #f6f8fa; padding: 10px; overflow-x: auto; }}
code {{ background: #f6f8fa; padding: 2px 4px; }}
</style>
</head>
<body>
{html_body}
</body>
</html>
"""


# ---------- Gemini OCR ----------
OCR_PROMPT = """You are an OCR engine for academic documents.

Task:
1) Transcribe ALL visible text exactly (do not paraphrase, do not translate).
2) Convert ALL mathematical expressions into LaTeX and keep them in the original position.
3) Output ONLY Markdown.
   - Inline math MUST be wrapped with $...$
   - Display/centered standalone equations MUST be wrapped with $$...$$ on their own lines
4) Preserve reading order, line breaks, bullet lists, and headings if present.
5) If a token is unreadable, write [UNK] (do NOT guess).

Return ONLY the Markdown.
"""

@st.cache_data(show_spinner=False)
def gemini_ocr_markdown(image_bytes: bytes, mime_type: str, model_id: str) -> str:
    api_key = st.secrets.get("GEMINI_API_KEY", None)
    if not api_key:
        raise RuntimeError("Missing GEMINI_API_KEY in Streamlit secrets.")

    client = genai.Client(api_key=api_key)

    resp = client.models.generate_content(
        model=model_id,
        contents=[
            types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
            OCR_PROMPT,
        ],
        config=types.GenerateContentConfig(
            temperature=0.0,
            top_p=1.0,
            top_k=1,
            max_output_tokens=8192,
        ),
    )
    text = resp.text or ""
    return normalize_latex_delimiters(text)


def guess_mime(filename: str) -> str:
    fn = filename.lower()
    if fn.endswith(".png"):
        return "image/png"
    if fn.endswith(".webp"):
        return "image/webp"
    return "image/jpeg"


# ---------- Streamlit UI ----------
st.set_page_config(page_title="Gemini OCR → LaTeX", layout="wide")
st.title("Gemini OCR: Image → Text + LaTeX (in-place)")

st.markdown(
    "Upload an image. The app returns **Markdown** with formulas as **LaTeX** using `$...$` / `$$...$$`, "
    "and renders it **in the original position** via MathJax."
)

colA, colB = st.columns([1, 1])

with colA:
    uploaded = st.file_uploader("Upload image (png/jpg/webp)", type=["png", "jpg", "jpeg", "webp"])
    model_id = st.selectbox(
        "Gemini model",
        options=["gemini-3-flash-preview", "gemini-3-pro-preview"],
        index=0,
        help="Flash is faster; Pro is typically more accurate for messy scans."
    )

with colB:
    st.info("Tip: If inline math doesn’t render in Streamlit markdown, this app uses MathJax HTML rendering to keep formulas in-place.")

if not uploaded:
    st.stop()

img = Image.open(uploaded).convert("RGB")
st.image(img, caption="Uploaded image", use_container_width=True)

# Optional downscale for speed
max_side = st.slider("Max side length (downscale if larger)", 800, 3000, 1800, 50)
w, h = img.size
scale = min(1.0, max_side / max(w, h))
if scale < 1.0:
    img = img.resize((int(w * scale), int(h * scale)))

# Encode to bytes (PNG)
import io
buf = io.BytesIO()
img.save(buf, format="PNG")
image_bytes = buf.getvalue()
mime_type = "image/png"

if st.button("Run OCR", type="primary"):
    with st.spinner("Calling Gemini ..."):
        md_out = gemini_ocr_markdown(image_bytes, mime_type, model_id)

    tab1, tab2 = st.tabs(["Rendered (MathJax, in-place)", "Raw Markdown"])

    with tab1:
        html_body = markdown_to_safe_html(md_out)
        full_html = wrap_with_mathjax(html_body)
        st.components.v1.html(full_html, height=800, scrolling=True)

    with tab2:
        st.code(md_out, language="markdown")

    st.download_button("Download Markdown", md_out.encode("utf-8"), file_name="result.md")
    st.download_button("Download HTML", full_html.encode("utf-8"), file_name="result.html")
