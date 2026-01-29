# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import io
import time
import json
import math
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Tuple, Optional

import numpy as np
import streamlit as st
from PIL import Image

# Gemini
from google import genai
from google.genai import types
from google.genai import errors as genai_errors

# Pandoc
import pypandoc
def ensure_pandoc():
    try:
        _ = pypandoc.get_pandoc_path()
    except OSError:
        # 若环境里没找到 pandoc，就尝试下载（某些云环境可能被限制）
        pypandoc.download_pandoc()

ensure_pandoc()


# -----------------------------
# UI basics
# -----------------------------
st.set_page_config(page_title="Gemini OCR → Word (LaTeX)", layout="wide")

TITLE = "Gemini OCR（多图/长图切片）→ Word 导出（含第三版 LaTeX 代码） + Word 转 LaTeX 代码"
st.title(TITLE)

# -----------------------------
# Prompts
# -----------------------------
OCR_PROMPT_ZH = r"""
你是一个严谨的中文学术 OCR 转写器。请从图片中提取内容并输出 **Markdown**，要求：
1) 文字保持中文、语句通顺、稍微美观（适当分段、标题加粗/##），但不要乱改原意。
2) 公式必须转为 LaTeX：
   - 行间公式用 $$ ... $$（保留 \tag{n} 若图片里有编号）
   - 行内公式用 $ ... $
3) 保持原本的顺序与相对位置：段落/公式出现的位置尽量与图片一致。
4) 不要输出“我看到了什么/解释/推导”，只输出最终 Markdown 正文。
5) 如果某一块是表格，请用 Markdown 表格输出（不要用大量对齐空格）。
6) 若图片里有页眉页脚/无关按钮元素，忽略。
输出：仅 Markdown，不要 JSON，不要代码块包裹全文。
""".strip()

FORMULA_ONLY_PROMPT = r"""
你会收到一张图片，这张图片大概率是“公式截图/公式图片”。
任务：如果它确实是数学公式/符号表达式，请只输出 **LaTeX 代码**（不要解释，不要前后文字，不要 $$ 包裹）。
如果不是公式（例如普通插图、地图、照片），请输出空字符串。
""".strip()

# -----------------------------
# Helpers: image preprocess
# -----------------------------
def load_pil(uploaded_file) -> Image.Image:
    img = Image.open(uploaded_file).convert("RGB")
    return img

def downscale_max_side(img: Image.Image, max_side: int) -> Image.Image:
    w, h = img.size
    m = max(w, h)
    if m <= max_side:
        return img
    scale = max_side / float(m)
    nw, nh = int(w * scale), int(h * scale)
    return img.resize((nw, nh), Image.LANCZOS)

def slice_long_image(img: Image.Image, tile_h: int, overlap: int) -> List[Tuple[Image.Image, Tuple[int,int,int,int]]]:
    """Vertical slicing with overlap. Returns list of (tile_img, (left, top, right, bottom))."""
    w, h = img.size
    if h <= tile_h:
        return [(img, (0, 0, w, h))]
    tiles = []
    y = 0
    while y < h:
        top = y
        bottom = min(y + tile_h, h)
        tile = img.crop((0, top, w, bottom))
        tiles.append((tile, (0, top, w, bottom)))
        if bottom >= h:
            break
        y = bottom - overlap
    return tiles

def pil_to_jpeg_bytes(img: Image.Image, quality: int = 85) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()

# -----------------------------
# Helpers: Markdown postprocess / dedupe
# -----------------------------
def normalize_md(md: str) -> str:
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return md

def dedupe_adjacent_chunks(chunks: List[str]) -> str:
    """
    Simple overlap-dedupe: remove identical tail/head blocks due to overlapped tiles.
    """
    chunks = [normalize_md(c) for c in chunks if c and c.strip()]
    if not chunks:
        return ""
    out = [chunks[0]]
    for cur in chunks[1:]:
        prev = out[-1]
        # Compare last N lines of prev with first N lines of cur
        prev_lines = prev.splitlines()
        cur_lines = cur.splitlines()
        max_n = min(30, len(prev_lines), len(cur_lines))
        cut = 0
        for n in range(max_n, 5, -1):
            if prev_lines[-n:] == cur_lines[:n]:
                cut = n
                break
        if cut > 0:
            cur = "\n".join(cur_lines[cut:]).strip()
        if cur:
            out.append(cur)
    return normalize_md("\n\n".join(out))

# -----------------------------
# Helpers: math -> code version (third download)
# -----------------------------
INLINE_MATH = re.compile(r"(?<!\\)\$(.+?)(?<!\\)\$")
DISPLAY_MATH = re.compile(r"(?s)(?<!\\)\$\$(.+?)(?<!\\)\$\$")

def md_math_to_code(md: str) -> str:
    """
    Convert math rendering into LaTeX CODE blocks/text:
      $$ ... $$  -> ```latex\n$$ ... $$\n```
      $ ... $    -> ` $ ... $ `
    """
    def repl_display(m):
        body = m.group(1).strip()
        return "\n```latex\n$$ " + body + " $$\n```\n"
    md2 = DISPLAY_MATH.sub(repl_display, md)

    def repl_inline(m):
        body = m.group(1).strip()
        return "`$ " + body + " $`"
    md2 = INLINE_MATH.sub(repl_inline, md2)
    return normalize_md(md2)

# -----------------------------
# Pandoc conversions
# -----------------------------
def md_to_docx(md_text: str, out_path: Path) -> Path:
    md_text = normalize_md(md_text) + "\n"
    # Ensure pandoc available
    _ = pypandoc.get_pandoc_path()
    # Use markdown with tex math dollars
    fmt = "markdown+tex_math_dollars+tex_math_single_backslash"
    pypandoc.convert_text(md_text, to="docx", format=fmt, outputfile=str(out_path))
    return out_path

def docx_to_md(docx_path: Path, media_dir: Path) -> str:
    media_dir.mkdir(parents=True, exist_ok=True)
    fmt = "docx"
    md = pypandoc.convert_file(
        str(docx_path),
        to="markdown",
        format=fmt,
        extra_args=[f"--extract-media={str(media_dir)}"]
    )
    return normalize_md(md)

# -----------------------------
# Gemini client + safe call
# -----------------------------
@dataclass
class GeminiResult:
    text: str
    status_code: Optional[int] = None
    error_message: Optional[str] = None
    raw_error: Optional[str] = None

def get_gemini_client() -> genai.Client:
    api_key = None
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
    api_key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")

    if not api_key:
        st.stop()

    # Force Gemini Developer API (NOT Vertex AI)
    # If your environment accidentally sets Vertex-related vars, this avoids 401 "API keys not supported..."
    client = genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(api_version="v1beta"),
    )
    return client

def safe_generate_content(
    client: genai.Client,
    model_id: str,
    contents,
    config: Optional[types.GenerateContentConfig] = None,
    retries: int = 3,
    sleep_base: float = 1.2
) -> GeminiResult:
    last_err = None
    for attempt in range(retries):
        try:
            resp = client.models.generate_content(
                model=model_id,
                contents=contents,
                config=config,
            )
            return GeminiResult(text=(resp.text or "").strip())
        except genai_errors.ClientError as e:
            # Key point: do NOT crash streamlit; expose status_code & message safely
            status = getattr(e, "status_code", None)
            msg = str(e)
            return GeminiResult(text="", status_code=status, error_message=msg, raw_error=repr(e))
        except genai_errors.ServerError as e:
            # 5xx -> retry
            last_err = e
            time.sleep(sleep_base * (2 ** attempt))
        except Exception as e:
            last_err = e
            time.sleep(sleep_base * (2 ** attempt))

    # If we reached here, retries exhausted
    msg = f"{type(last_err).__name__}: {last_err}"
    return GeminiResult(text="", status_code=None, error_message=msg, raw_error=repr(last_err))

def gemini_ocr_one_image(
    client: genai.Client,
    model_id: str,
    image_bytes: bytes,
    mime_type: str,
    prompt: str,
    max_output_tokens: int = 4096,
) -> GeminiResult:
    # If inline bytes too large, use Files API automatically (20MB request limit) :contentReference[oaicite:4]{index=4}
    part = None
    if len(image_bytes) > 18_000_000:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as f:
            f.write(image_bytes)
            tmp = f.name
        up = client.files.upload(file=tmp)
        os.unlink(tmp)
        contents = [up, prompt]
    else:
        part = types.Part.from_bytes(data=image_bytes, mime_type=mime_type)
        contents = [part, prompt]  # prompt after image (best practice) :contentReference[oaicite:5]{index=5}

    cfg = types.GenerateContentConfig(
        # 对 Gemini 3：官方建议 temperature 保持默认 1.0（避免循环/异常行为） :contentReference[oaicite:6]{index=6}
        temperature=1.0,
        max_output_tokens=max_output_tokens,
    )
    return safe_generate_content(client, model_id, contents, cfg)

# -----------------------------
# Render preview with MathJax (in-place)
# -----------------------------
MATHJAX_HTML = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<script>
window.MathJax = {
  tex: {
    inlineMath: [['$', '$'], ['\\\\(', '\\\\)']],
    displayMath: [['$$','$$'], ['\\\\[','\\\\]']],
    processEscapes: true,
    tags: 'ams'
  },
  options: { skipHtmlTags: ['script','noscript','style','textarea','pre','code'] }
};
</script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
<style>
body { font-family: -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,"Noto Sans SC","PingFang SC","Microsoft YaHei",sans-serif; line-height: 1.55; padding: 6px 10px; }
h1,h2,h3 { margin: 0.6em 0 0.3em; }
pre, code { background: #f6f6f6; padding: 2px 4px; border-radius: 4px; }
table { border-collapse: collapse; }
td,th { border: 1px solid #ddd; padding: 6px 8px; }
</style>
</head>
<body>
<div id="content">{{CONTENT}}</div>
<script>
(function(){
  // simple markdown-ish conversion for preview (keep it light)
  const el = document.getElementById("content");
  let t = el.textContent;
  // Convert headings
  t = t.replace(/^### (.*)$/gm, "<h3>$1</h3>");
  t = t.replace(/^## (.*)$/gm, "<h2>$1</h2>");
  t = t.replace(/^# (.*)$/gm, "<h1>$1</h1>");
  // bold
  t = t.replace(/\\*\\*(.+?)\\*\\*/g, "<b>$1</b>");
  // line breaks
  t = t.replace(/\\n/g, "<br/>");
  el.innerHTML = t;
})();
</script>
</body>
</html>
"""

def render_mathjax(md: str):
    import html
    safe = html.escape(md)
    html_doc = MATHJAX_HTML.replace("{{CONTENT}}", safe)
    st.components.v1.html(html_doc, height=600, scrolling=True)

# -----------------------------
# UI: Model select + debug panel
# -----------------------------
with st.sidebar:
    st.subheader("Gemini 设置")
    model_id = st.selectbox(
        "Model ID",
        [
            "gemini-3-flash-preview",
            "gemini-flash-latest",
            "gemini-2.5-flash",
            "gemini-2.5-pro",
        ],
        index=0
    )
    max_side = st.slider("Max side length (downscale if larger)", 800, 3200, 2000, 100)
    tile_h = st.slider("Long image tile height", 800, 2600, 1600, 100)
    overlap = st.slider("Tile overlap (px)", 0, 400, 160, 10)
    jpeg_q = st.slider("JPEG quality (smaller=faster)", 50, 95, 85, 1)
    out_tokens = st.slider("Max output tokens", 1024, 8192, 4096, 256)

    st.divider()
    st.caption("如果你遇到红框报错：下面的页面会显示 status_code 和 message，按错误码处理即可。")

tabs = st.tabs(["① 图片 OCR → Word 导出", "② 上传 Word(docx) → 公式转 LaTeX 代码并替换输出"])

# -----------------------------
# TAB 1: Image OCR -> Word
# -----------------------------
with tabs[0]:
    st.subheader("① 图片 OCR（支持多图 + 长图自动切片）")
    imgs = st.file_uploader("上传 1 张或多张图片（png/jpg/webp）", type=["png","jpg","jpeg","webp"], accept_multiple_files=True)

    run = st.button("Run OCR", type="primary", disabled=(not imgs))
    if run and imgs:
        client = get_gemini_client()

        all_md_parts = []
        debug_rows = []

        for idx, f in enumerate(imgs, start=1):
            img = load_pil(f)
            img = downscale_max_side(img, max_side=max_side)
            tiles = slice_long_image(img, tile_h=tile_h, overlap=overlap)

            st.write(f"**Image {idx}:** {img.size[0]}×{img.size[1]}  → tiles: {len(tiles)}")

            tile_mds = []
            for ti, (tile, box) in enumerate(tiles, start=1):
                b = pil_to_jpeg_bytes(tile, quality=jpeg_q)
                res = gemini_ocr_one_image(
                    client=client,
                    model_id=model_id,
                    image_bytes=b,
                    mime_type="image/jpeg",
                    prompt=OCR_PROMPT_ZH,
                    max_output_tokens=out_tokens,
                )

                debug_rows.append({
                    "img": idx,
                    "tile": ti,
                    "box": box,
                    "bytes": len(b),
                    "status_code": res.status_code,
                    "error": (res.error_message[:200] if res.error_message else ""),
                })

                if res.error_message:
                    st.error(f"Gemini error on Image {idx} Tile {ti}: status={res.status_code} message={res.error_message}")
                    # 官方建议：按 status code 对照处理（FAILED_PRECONDITION/429/INVALID_ARGUMENT 等） :contentReference[oaicite:7]{index=7}
                    st.stop()

                tile_mds.append(res.text)

            merged = dedupe_adjacent_chunks(tile_mds)
            all_md_parts.append(merged)

        final_md = normalize_md("\n\n---\n\n".join(all_md_parts))
        st.success("OCR 完成")

        colA, colB = st.columns([1,1], gap="large")
        with colA:
            st.markdown("### Rendered (MathJax, in-place)")
            render_mathjax(final_md)

        with colB:
            st.markdown("### Raw Markdown")
            st.code(final_md, language="markdown")

        # Export files (3 versions)
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            md_path = td / "ocr.md"
            md_path.write_text(final_md, encoding="utf-8")

            docx_rendered = td / "OCR_Rendered.docx"
            md_to_docx(final_md, docx_rendered)

            latex_code_md = md_math_to_code(final_md)
            docx_latex_code = td / "OCR_LaTeX_Code.docx"
            md_to_docx(latex_code_md, docx_latex_code)

            st.download_button("下载版本1：Rendered Word (.docx)", data=docx_rendered.read_bytes(), file_name=docx_rendered.name)
            st.download_button("下载版本2：Raw Markdown (.md)", data=md_path.read_bytes(), file_name=md_path.name)
            st.download_button("下载版本3：LaTeX 代码版 Word (.docx)", data=docx_latex_code.read_bytes(), file_name=docx_latex_code.name)

        with st.expander("Debug（如果你再遇到 ClientError，先看这里）", expanded=False):
            st.json(debug_rows)

# -----------------------------
# TAB 2: Word -> LaTeX code
# -----------------------------
with tabs[1]:
    st.subheader("② 上传 Word(.docx) → 把公式转成 LaTeX 代码形式并替换输出")
    st.caption("说明：docx 内原生公式（OMML）Pandoc 会转成 $...$/$$...$$；若有“公式截图图片”，会尝试用 Gemini 识别为 LaTeX 后替换。")

    docx_file = st.file_uploader("上传 Word (.docx)", type=["docx"])
    run2 = st.button("Convert Word → LaTeX Code Word", type="primary", disabled=(not docx_file))

    if run2 and docx_file:
        client = get_gemini_client()

        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            src = td / "input.docx"
            src.write_bytes(docx_file.read())

            media_dir = td / "media"
            md = docx_to_md(src, media_dir=media_dir)

            # 1) turn docx math to code style (wrap existing $/$$)
            md2 = md_math_to_code(md)

            # 2) replace formula-like images in markdown
            # markdown image pattern: ![](media/image.png)
            img_pat = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
            replacements = {}
            for m in img_pat.finditer(md2):
                rel = m.group(1)
                # pandoc usually uses something like media/image1.png
                p = (td / rel).resolve()
                if not p.exists():
                    continue

                try:
                    im = Image.open(p).convert("RGB")
                    w, h = im.size
                except Exception:
                    continue

                # heuristic: small & wide images more likely formulas
                looks_formula = (h <= 280 and w >= 240) or (h <= 220 and w/h >= 2.0)
                if not looks_formula:
                    continue

                b = pil_to_jpeg_bytes(im, quality=90)
                res = gemini_ocr_one_image(
                    client=client,
                    model_id=model_id,
                    image_bytes=b,
                    mime_type="image/jpeg",
                    prompt=FORMULA_ONLY_PROMPT,
                    max_output_tokens=1024,
                )
                if res.error_message:
                    st.error(f"Gemini error while converting formula image: status={res.status_code} message={res.error_message}")
                    st.stop()

                latex = (res.text or "").strip()
                if not latex:
                    continue

                # Replace image markdown with LaTeX code block (you asked for code format)
                rep = "\n```latex\n$$ " + latex + " $$\n```\n"
                replacements[m.group(0)] = rep

            for k, v in replacements.items():
                md2 = md2.replace(k, v)

            md2 = normalize_md(md2)

            out_docx = td / "Word_LaTeX_Code.docx"
            md_to_docx(md2, out_docx)

            out_md = td / "Word_LaTeX_Code.md"
            out_md.write_text(md2, encoding="utf-8")

            st.success("转换完成")
            st.download_button("下载：LaTeX 代码版 Word (.docx)", data=out_docx.read_bytes(), file_name=out_docx.name)
            st.download_button("下载：中间产物 Markdown (.md)", data=out_md.read_bytes(), file_name=out_md.name)

            st.markdown("### 预览（Raw Markdown）")
            st.code(md2, language="markdown")
