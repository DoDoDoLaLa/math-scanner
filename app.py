# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
import re
import json
import time
import base64
import tempfile
import concurrent.futures
from functools import lru_cache
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Callable, Iterable

import streamlit as st
from PIL import Image, ImageFilter, ImageOps

from openai import OpenAI

import httpx

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import pypandoc

# LaTeX-in-text -> Word equation (OMML) support (optional)
try:
    from lxml import etree  # type: ignore
    import latex2mathml.converter  # type: ignore
    HAVE_LATEX_OMML = True
except Exception:
    HAVE_LATEX_OMML = False

# PDF rendering (optional but recommended)
try:
    import fitz  # PyMuPDF
    HAVE_PYMUPDF = True
except Exception:
    HAVE_PYMUPDF = False


# ============================================================
# 0) App UI
# ============================================================
st.set_page_config(page_title="Ark OCR / DOCX Translate / LaTeX Export", layout="wide")
st.title("学术 OCR & Word→LaTeX 工具（Ark EP 已接入）")

st.caption(
    "面向论文/讲义/教材：PDF/图片OCR→Markdown/Word；Word（含可编辑公式）→LaTeX/Markdown；Word 原排版内翻译（best-effort）。"
)

def ensure_pandoc():
    try:
        _ = pypandoc.get_pandoc_path()
    except OSError:
        try:
            pypandoc.download_pandoc()
        except Exception:
            pass

ensure_pandoc()


# ============================================================
# 1) Prompts
# ============================================================
OCR_PROMPT_ZH = r"""
你是一个严谨的中文学术 OCR 转写器。请从图片中提取内容并输出 Markdown，要求：
1) 把图片中所有可见文字逐字转写（保持中文，不要翻译，不要改写）。
2) 数学公式必须转为 LaTeX，并尽量保持原位：
   - 行内用 $...$
   - 行间用 $$...$$
   - 如果有公式编号（如 (6)(7) 或 \tag{6}），保留编号信息。
3) 保持原本顺序，表格用 Markdown 表格输出。
4) 只输出 Markdown 正文，不要解释。
5) 保持原有阅读顺序、换行、项目符号、标题层级。
""".strip()

TRANSLATE_PROMPT_TEMPLATE = r"""
你是一个专业学术翻译引擎。请把以下内容从 __SRC_LANG__ 翻译到 __DST_LANG__。

严格要求：
- 输入是 JSON，包含 items 列表，每个 item 有 id 和 segments（字符串列表）。
- 输出必须是 JSON，结构必须为：{"items":[{"id":"...","segments":[...]}, ...]}
- segments 的数量必须与输入完全一致；每个 segments[i] 对应翻译输入 segments[i]。
- 保留所有占位符不变：例如 __MATH_0__、__KEEP_12__、{{ }} 这种标记必须原样输出，不能翻译、不能改大小写、不能删。
- LaTeX 代码、公式环境（如 \begin{equation}...\end{equation} 或 $...$）不得改动。
- 不要添加多余字段、不要输出解释、不要 Markdown。
""".strip()


# ============================================================
# 2) Ark config / client / retry / timeout
# ============================================================
@dataclass
class ArkResult:
    text: str = ""
    error_message: Optional[str] = None


class TranslateCallError(RuntimeError):
    """Raised when a translate call fails or returns invalid JSON schema."""
    def __init__(self, message: str, raw: Optional[str] = None):
        super().__init__(message)
        self.raw = raw

DEFAULT_ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_ARK_MODEL = "ep-20260203141749-992fx"
DEFAULT_TRANSLATE_MODEL = "ep-20260203211039-n9v2f"  # dedicated translate endpoint
DEFAULT_MML2OMML_XSL_ENV = "MML2OMML_XSL"  # path to MML2OMML.XSL
DEFAULT_OCR_TIMEOUT_S = 120
DEFAULT_TRANSLATE_TIMEOUT_S = 180

def get_api_key() -> str:
    k = None
    try:
        if "ARK_API_KEY" in st.secrets:
            k = st.secrets["ARK_API_KEY"]
    except Exception:
        pass
    k = k or os.environ.get("ARK_API_KEY")
    if not k:
        st.error("缺少 ARK_API_KEY。请在 .streamlit/secrets.toml 或环境变量中配置。")
        st.stop()
    return k

def get_ark_base_url() -> str:
    v = None
    try:
        v = st.secrets.get("ARK_BASE_URL", None)
    except Exception:
        v = None
    return v or os.environ.get("ARK_BASE_URL") or DEFAULT_ARK_BASE_URL

def get_default_model() -> str:
    v = None
    try:
        v = st.secrets.get("ARK_MODEL", None)
    except Exception:
        v = None
    return v or os.environ.get("ARK_MODEL") or os.environ.get("ARK_ENDPOINT_ID") or DEFAULT_ARK_MODEL

def get_timeout_defaults() -> Tuple[int, int]:
    ocr_t = None
    tr_t = None
    try:
        ocr_t = st.secrets.get("ARK_OCR_TIMEOUT_S", None)
        tr_t = st.secrets.get("ARK_TRANSLATE_TIMEOUT_S", None)
    except Exception:
        pass
    ocr_t = ocr_t or os.environ.get("ARK_OCR_TIMEOUT_S") or DEFAULT_OCR_TIMEOUT_S
    tr_t = tr_t or os.environ.get("ARK_TRANSLATE_TIMEOUT_S") or DEFAULT_TRANSLATE_TIMEOUT_S
    return int(ocr_t), int(tr_t)

@st.cache_resource(show_spinner=False)
def get_ark_client_cached(api_key: str, base_url: str, default_timeout_s: int) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=base_url, timeout=default_timeout_s)

def get_ark_client(default_timeout_s: int) -> OpenAI:
    return get_ark_client_cached(get_api_key(), get_ark_base_url(), default_timeout_s)

def _parse_retry_delay_seconds(msg: str) -> Optional[float]:
    m = re.search(r"retry in ([0-9.]+)s", msg, flags=re.IGNORECASE)
    if m:
        return float(m.group(1))
    m = re.search(r'"retryDelay"\s*:\s*"(\d+)s"', msg)
    if m:
        return float(m.group(1))
    return None

def safe_chat_completions(
    client: OpenAI,
    model: str,
    messages: List[dict],
    temperature: float,
    max_tokens: int,
    timeout_s: int,
    retries: int = 6,
    total_timeout_s: Optional[int] = None,
    on_attempt: Optional[Callable[[int, int, str], None]] = None,
    response_format: Optional[dict] = None,
) -> ArkResult:
    """
    - timeout_s: 单次请求超时（传给 OpenAI SDK / httpx）
    - total_timeout_s: 整体“墙钟”超时（包含重试等待）。用于避免卡太久看起来像“死了”。
    - on_attempt: (attempt_idx, retries, phase_msg) 回调，用于在 Streamlit UI 中展示当前在做什么。
    """
    t0 = time.time()
    last_err: Any = None

    for attempt in range(1, retries + 1):
        if total_timeout_s is not None and (time.time() - t0) > float(total_timeout_s):
            return ArkResult(text="", error_message=f"overall timeout after {total_timeout_s}s (last_err={last_err})")

        if on_attempt:
            on_attempt(attempt, retries, "request")

        try:
            kw = {}
            if response_format is not None:
                kw["response_format"] = response_format

            rsp = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                **kw,
            )
            text = rsp.choices[0].message.content or ""
            return ArkResult(text=text, error_message=None)

        except Exception as e:
            last_err = e
            msg = str(e)

            # Try parse retry delay from error message, else exponential backoff.
            delay = _parse_retry_delay_seconds(msg)
            if delay is None:
                delay = min(8.0, 0.8 * (2 ** (attempt - 1)))

            if on_attempt:
                on_attempt(attempt, retries, f"backoff {delay:.1f}s")

            # Retry on typical transient failures
            if any(x in msg.lower() for x in ["timeout", "timed out", "rate limit", "429", "overloaded", "try again", "internal"]):
                time.sleep(delay)
                continue

            return ArkResult(text="", error_message=f"{type(e).__name__}: {e}")

    return ArkResult(text="", error_message=f"failed after retries: {last_err}")


# ============================================================
# 3) Image utils (tiling / enhance)
# ============================================================
def to_base64_jpeg(img: Image.Image, quality: int = 85) -> str:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def enhance_for_ocr(img: Image.Image) -> Image.Image:
    img = img.convert("RGB")
    img = ImageOps.exif_transpose(img)
    img = ImageOps.autocontrast(img)
    img = img.filter(ImageFilter.SHARPEN)
    return img

def tile_image(img: Image.Image, tile_h: int, overlap: int) -> List[Image.Image]:
    w, h = img.size
    if h <= tile_h:
        return [img]
    tiles = []
    y = 0
    while y < h:
        y2 = min(h, y + tile_h)
        crop = img.crop((0, y, w, y2))
        tiles.append(crop)
        if y2 >= h:
            break
        y = y2 - overlap
    return tiles


# ============================================================
# 4) OCR
# ============================================================
def normalize_md(md: str) -> str:
    md = (md or "").replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip()

def dedupe_tail_head(prev: str, cur: str) -> str:
    prev = prev.strip()
    cur = cur.strip()
    if not prev or not cur:
        return cur
    max_k = min(300, len(prev), len(cur))
    best = 0
    for k in range(10, max_k + 1):
        if prev[-k:] == cur[:k]:
            best = k
    return cur[best:].lstrip("\n")

def ocr_images_to_markdown(
    client: OpenAI,
    images: List[Image.Image],
    model: str,
    max_side: int,
    tile_h: int,
    overlap: int,
    jpeg_q: int,
    out_tokens: int,
    timeout_s: int,
) -> str:
    chunks: List[str] = []
    for idx, img in enumerate(images, start=1):
        img = enhance_for_ocr(img)
        # resize
        w, h = img.size
        scale = 1.0
        m = max(w, h)
        if m > max_side:
            scale = max_side / float(m)
        if scale != 1.0:
            img = img.resize((int(w * scale), int(h * scale)))

        tiles = tile_image(img, tile_h=tile_h, overlap=overlap)
        st.info(f"第 {idx}/{len(images)} 页：分块 {len(tiles)}")

        for ti, tile in enumerate(tiles, start=1):
            b64 = to_base64_jpeg(tile, quality=jpeg_q)
            messages = [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": OCR_PROMPT_ZH},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    ],
                }
            ]
            res = safe_chat_completions(
                client=client,
                model=model,
                messages=messages,
                temperature=0.2,
                max_tokens=out_tokens,
                timeout_s=timeout_s,
                retries=6,
            )
            if res.error_message:
                st.error(f"OCR 调用失败：\n\n{res.error_message}")
                st.stop()

            md = normalize_md(res.text)
            if chunks:
                md = dedupe_tail_head(chunks[-1], md)
            chunks.append(md)

    return normalize_md("\n\n".join([c for c in chunks if c.strip()]))


# ============================================================
# 4.5) PDF -> images (PyMuPDF)
# ============================================================
def pdf_bytes_to_images(pdf_bytes: bytes, dpi: int = 300, max_pages: Optional[int] = None) -> List[Image.Image]:
    if not HAVE_PYMUPDF:
        raise RuntimeError("缺少依赖 pymupdf。请先 pip install pymupdf")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs: List[Image.Image] = []
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    n = doc.page_count if max_pages is None else min(doc.page_count, max_pages)
    for i in range(n):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
        imgs.append(img)
    return imgs


# ============================================================
# 5) LaTeX code style for OCR result
# ============================================================
DISPLAY_MATH_RE = re.compile(r"(?s)\$\$(.+?)\$\$")
INLINE_MATH_RE = re.compile(r"(?<!\$)\$([^$\n]+)\$(?!\$)")

def display_to_equation_env(md: str) -> str:
    def repl(m):
        body = m.group(1).strip()
        return "\\begin{equation}\n" + body + "\n\\end{equation}"
    return DISPLAY_MATH_RE.sub(repl, md)

def inline_to_paren(md: str) -> str:
    def repl(m):
        body = m.group(1).strip()
        return "\\(" + body + "\\)"
    return INLINE_MATH_RE.sub(repl, md)

def md_to_latex_code_style(md: str) -> str:
    md2 = normalize_md(md)
    md2 = display_to_equation_env(md2)
    md2 = inline_to_paren(md2)
    eq_env = re.compile(r"(?s)(\\begin\{equation\}.*?\\end\{equation\})")
    md2 = eq_env.sub(lambda m: "\n```latex\n" + m.group(1).strip() + "\n```\n", md2)
    return normalize_md(md2)


# ============================================================
# 6) DOCX translate (preserve layout, best-effort) + equation replacement
# ============================================================
MATH_TOKEN_RE = re.compile(
    r"(\$\$.*?\$\$|\$[^$\n]+\$|\\begin\{equation\}.*?\\end\{equation\}|\\\(.+?\\\))",
    re.DOTALL
)

def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    mapping: Dict[str, str] = {}
    k = 0
    def repl(m):
        nonlocal k
        token = f"__MATH_{k}__"
        mapping[token] = m.group(0)
        k += 1
        return token
    out = MATH_TOKEN_RE.sub(repl, text)
    return out, mapping

def restore_tokens(text: str, mapping: Dict[str, str]) -> str:
    for tok, val in mapping.items():
        text = text.replace(tok, val)
    return text

def iter_table_paragraphs(table: Table) -> List[Paragraph]:
    out: List[Paragraph] = []
    for row in table.rows:
        for cell in row.cells:
            out.extend(cell.paragraphs)
            for t2 in cell.tables:
                out.extend(iter_table_paragraphs(t2))
    return out

def iter_part_paragraphs(part) -> List[Paragraph]:
    """part: doc, header, footer"""
    ps: List[Paragraph] = []
    try:
        ps.extend(part.paragraphs)
    except Exception:
        pass
    try:
        for table in part.tables:
            ps.extend(iter_table_paragraphs(table))
    except Exception:
        pass
    return ps

def iter_all_paragraphs_extended(doc: Document) -> List[Paragraph]:
    """
    ✅ 覆盖：正文 paragraphs + tables + 每个 section 的 header/footer（含表格）
    仍不覆盖：textbox/shape、批注、脚注尾注（python-docx 限制）
    """
    ps: List[Paragraph] = []
    ps.extend(iter_part_paragraphs(doc))

    # headers/footers
    try:
        for sec in doc.sections:
            ps.extend(iter_part_paragraphs(sec.header))
            ps.extend(iter_part_paragraphs(sec.footer))
            # first/even page headers/footers（若存在）
            if hasattr(sec, "first_page_header"):
                ps.extend(iter_part_paragraphs(sec.first_page_header))
            if hasattr(sec, "first_page_footer"):
                ps.extend(iter_part_paragraphs(sec.first_page_footer))
            if hasattr(sec, "even_page_header"):
                ps.extend(iter_part_paragraphs(sec.even_page_header))
            if hasattr(sec, "even_page_footer"):
                ps.extend(iter_part_paragraphs(sec.even_page_footer))
    except Exception:
        pass
    return ps

def extract_math_from_docx_with_pandoc(docx_bytes: bytes) -> List[Tuple[str, str]]:
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        f = td / "in.docx"
        f.write_bytes(docx_bytes)
        md = pypandoc.convert_file(str(f), to="markdown", format="docx", extra_args=["--wrap=none"])
        md = md.replace("\r\n", "\n").replace("\r", "\n")

    matches: List[Tuple[int, int, str, str]] = []
    for m in re.finditer(r"(?s)\$\$(.+?)\$\$", md):
        matches.append((m.start(), m.end(), "display", m.group(1).strip()))
    for m in re.finditer(r"(?<!\$)\$([^$\n]+)\$(?!\$)", md):
        matches.append((m.start(), m.end(), "inline", m.group(1).strip()))
    matches.sort(key=lambda x: x[0])
    return [(k, b) for _, _, k, b in matches]

def insert_text_run_at_paragraph_child(p_elm, idx: int, text: str):
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    if text.startswith(" ") or text.endswith(" "):
        t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    p_elm.insert(idx, r)

def replace_omml_with_latex_code(doc: Document, math_seq: List[Tuple[str, str]], use_equation_env: bool = False) -> int:
    all_ps = iter_all_paragraphs_extended(doc)
    seq_idx = 0
    replaced = 0

    for p in all_ps:
        p_elm = p._p
        nodes = p_elm.xpath(".//*[local-name()='oMath' or local-name()='oMathPara']")
        for node in nodes:
            if seq_idx >= len(math_seq):
                return replaced
            kind, body = math_seq[seq_idx]
            seq_idx += 1

            if use_equation_env:
                latex_text = (f"\\begin{{equation}} {body} \\end{{equation}}") if kind == "display" else (f"\\({body}\\)")
            else:
                # ✅ Pandoc 可直接渲染（tex_math_dollars）：inline=$...$，display=$$...$$
                latex_text = (f"$$\n{body}\n$$") if kind == "display" else (f"${body}$")

            parent = node.getparent()
            if parent is None:
                continue

            try:
                idx_in_parent = list(parent).index(node)
            except Exception:
                idx_in_parent = None

            if parent is not p_elm or idx_in_parent is None:
                insert_text_run_at_paragraph_child(p_elm, len(p_elm), latex_text)
            else:
                insert_text_run_at_paragraph_child(parent, idx_in_parent, latex_text)

            try:
                parent.remove(node)
            except Exception:
                pass

            replaced += 1

    return replaced


# ---------------------------
# LaTeX ($...$ / $$...$$ / \(..\) / \[..]) -> Word equation (OMML)
# ---------------------------
LATEX_INLINE_RE = re.compile(
    r"(\$\$.*?\$\$|\$[^$\n]+\$|\\\(.+?\\\)|\\\[.+?\\\])",
    re.DOTALL,
)

def _strip_latex_delims(s: str) -> Tuple[str, bool]:
    s = s.strip()
    if s.startswith("$$") and s.endswith("$$"):
        return s[2:-2].strip(), True
    if s.startswith("$") and s.endswith("$"):
        return s[1:-1].strip(), False
    if s.startswith(r"\(") and s.endswith(r"\)"):
        return s[2:-2].strip(), False
    if s.startswith(r"\[") and s.endswith(r"\]"):
        return s[2:-2].strip(), True
    return s, False

def _guess_office_mml2omml_paths() -> List[str]:
    # Common installs; user can override via sidebar/env.
    return [
        r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\Office16\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\Office16\MML2OMML.XSL",
    ]

@lru_cache(maxsize=1)
def _load_mml2omml_xslt(xsl_path: str):
    if not HAVE_LATEX_OMML:
        raise RuntimeError("Missing deps: lxml / latex2mathml")
    if not xsl_path:
        # try guess
        for p in _guess_office_mml2omml_paths():
            if os.path.exists(p):
                xsl_path = p
                break
    if not xsl_path or not os.path.exists(xsl_path):
        raise RuntimeError(
            "MML2OMML.XSL not found. Provide its path in sidebar (or env MML2OMML_XSL)."
        )
    xslt_doc = etree.parse(xsl_path)
    return etree.XSLT(xslt_doc)

def latex_to_omml_element(latex_expr: str, xsl_path: str):
    """Return an lxml element (OMML) for insertion into docx XML."""
    if not HAVE_LATEX_OMML:
        raise RuntimeError("This feature requires: pip install lxml latex2mathml")
    body, is_block = _strip_latex_delims(latex_expr)
    # latex2mathml returns <math xmlns="http://www.w3.org/1998/Math/MathML">...</math>
    mml = latex2mathml.converter.convert(body)
    # Ensure no newlines (Word is sensitive per some implementations)
    mml = re.sub(r"\s+", " ", mml).strip()
    mml_root = etree.fromstring(mml.encode("utf-8"))
    xslt = _load_mml2omml_xslt(xsl_path)
    omml_tree = xslt(mml_root)
    omml_root = omml_tree.getroot()
    # Word accepts either m:oMath or m:oMathPara; keep as produced by stylesheet.
    return omml_root, is_block

def _clone_rPr(src_r):
    # clone run properties to keep styling for newly inserted text runs
    rpr = src_r.find(qn("w:rPr"))
    if rpr is None:
        return None
    return deepcopy(rpr)

def convert_inline_latex_to_omml_in_doc(doc: Document, xsl_path: str, log: Optional[Callable[[str], None]] = None) -> int:
    """Best-effort: convert LaTeX math delimited with $...$ etc into OMML equations inside a docx.

    Limitations:
    - Only detects math fully contained inside a single run.
    - Keeps paragraph structure; tries to preserve run formatting for surrounding text.
    """
    if not HAVE_LATEX_OMML:
        raise RuntimeError("Missing deps: lxml / latex2mathml")
    converted = 0
    for p in doc.paragraphs:
        # iterate by underlying XML runs to allow insertion
        r_elems = list(p._p.findall(qn("w:r")))
        for r in r_elems:
            t = r.find(qn("w:t"))
            if t is None or not t.text:
                continue
            text = t.text
            # find first match (one per run; iterate after mutation)
            m = LATEX_INLINE_RE.search(text)
            if not m:
                continue
            before = text[:m.start()]
            expr = m.group(0)
            after = text[m.end():]

            try:
                omml_elem, _is_block = latex_to_omml_element(expr, xsl_path)
            except Exception as e:
                if log:
                    log(f"公式转换失败（跳过）：{type(e).__name__}: {e} | expr={expr[:80]}")
                continue

            # Replace current run text with "before"
            t.text = before

            # Insert OMML element right after this run
            # Note: OMML uses its own namespace (m:). etree element is fine to append.
            r.addnext(omml_elem)

            # Insert "after" as a new run, trying to keep same rPr
            if after:
                new_r = OxmlElement("w:r")
                rpr_clone = _clone_rPr(r)
                if rpr_clone is not None:
                    new_r.append(rpr_clone)
                new_t = OxmlElement("w:t")
                new_t.text = after
                new_r.append(new_t)
                omml_elem.addnext(new_r)

            converted += 1
    return converted


def chunk_items_for_api(items: List[dict], max_chars: int = 12000) -> List[List[dict]]:
    batches: List[List[dict]] = []
    cur: List[dict] = []
    cur_len = 0
    for it in items:
        s = json.dumps(it, ensure_ascii=False)
        if cur and cur_len + len(s) > max_chars:
            batches.append(cur)
            cur = []
            cur_len = 0
        cur.append(it)
        cur_len += len(s)
    if cur:
        batches.append(cur)
    return batches

def extract_json_object(text: str) -> dict:
    if not text:
        raise ValueError("empty response")
    s = text.strip()

    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", s, flags=re.DOTALL | re.IGNORECASE)
    if fence:
        s = fence.group(1).strip()

    try:
        return json.loads(s)
    except Exception:
        pass

    start = s.find("{")
    if start == -1:
        raise ValueError("no '{' found in response")

    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(s)):
        ch = s[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    candidate = s[start:i + 1]
                    return json.loads(candidate)

    raise ValueError("no complete JSON object found (unbalanced braces)")


def _ark_extract_output_text(resp_obj: dict) -> str:
    """Extract assistant text from an Ark/OpenAI-style Responses API payload.

    Ark's /api/v3/responses is designed to be compatible with the OpenAI Responses schema:
    resp_obj.output -> list[message] -> content -> list[{type: "output_text", text: "..."}]
    (Some gateways may also return ChatCompletions-style `choices`.)
    """
    texts: List[str] = []

    if isinstance(resp_obj, dict):
        out = resp_obj.get("output")
        if isinstance(out, list):
            for msg in out:
                if not isinstance(msg, dict):
                    continue
                content = msg.get("content")
                if not isinstance(content, list):
                    continue
                for part in content:
                    if not isinstance(part, dict):
                        continue
                    if part.get("type") == "output_text" and isinstance(part.get("text"), str):
                        texts.append(part["text"])

        # Fallback: ChatCompletions-like shape
        if not texts:
            choices = resp_obj.get("choices")
            if isinstance(choices, list):
                for ch in choices:
                    try:
                        t = ch["message"]["content"]
                        if isinstance(t, str):
                            texts.append(t)
                    except Exception:
                        pass

        # Fallback: sometimes gateways add a root `text`
        if not texts and isinstance(resp_obj.get("text"), str):
            texts.append(resp_obj["text"])

    return "".join(texts).strip()


def ark_translate_text_via_responses(
    api_key: str,
    base_url: str,
    text: str,
    src_lang: str,
    dst_lang: str,
    timeout_s: int,
    model: str = DEFAULT_TRANSLATE_MODEL,
    retries: int = 6,
    on_attempt: Optional[Callable[[int, int, str], None]] = None,
) -> str:
    """Call Ark Responses API translation model using translation_options.

    According to Ark's translation model usage, each input_text can include:
      translation_options: {source_language?, target_language}.

    - If src_lang == "Auto", we omit source_language (model auto-detect).
    - dst_lang must be a language code (e.g. en, zh, ja ...).
    """
    url = base_url.rstrip("/") + "/responses"
    src_code: Optional[str] = None if src_lang == "Auto" else src_lang

    trans_opts: Dict[str, str] = {"target_language": dst_lang}
    if src_code:
        trans_opts["source_language"] = src_code

    payload = {
        "model": model,
        "input": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "input_text",
                        "text": text,
                        "translation_options": trans_opts,
                    }
                ],
            }
        ],
    }

    t0 = time.time()
    last_err: Optional[str] = None

    for attempt in range(1, retries + 1):
        if on_attempt:
            on_attempt(attempt, retries, "send")

        try:
            with httpx.Client(timeout=timeout_s) as hc:
                r = hc.post(
                    url,
                    headers={
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": "application/json",
                    },
                    json=payload,
                )

            if r.status_code >= 400:
                # Try parse structured error
                try:
                    err_obj = r.json()
                except Exception:
                    err_obj = None

                msg = None
                if isinstance(err_obj, dict):
                    # common shapes
                    msg = (
                        err_obj.get("error", {}) if isinstance(err_obj.get("error"), dict) else None
                    )
                    if isinstance(msg, dict):
                        msg = msg.get("message") or msg.get("msg")
                    if not msg:
                        msg = err_obj.get("message") or err_obj.get("msg")

                msg = msg or f"HTTP {r.status_code}: {r.text[:400]}"
                last_err = msg

                # Retry on rate limit / transient errors
                if r.status_code in (408, 409, 429) or (500 <= r.status_code < 600):
                    delay = _parse_retry_delay_seconds(msg) or min(8.0, 0.8 * (2 ** (attempt - 1)))
                    if on_attempt:
                        on_attempt(attempt, retries, f"backoff {delay:.1f}s")
                    time.sleep(delay)
                    continue

                raise TranslateCallError(f"translate http error: {msg}")

            obj = r.json()
            out_text = _ark_extract_output_text(obj)
            if not out_text:
                raise TranslateCallError("empty translate response text", raw=json.dumps(obj, ensure_ascii=False)[:800])
            return out_text

        except (httpx.TimeoutException, httpx.TransportError) as e:
            last_err = f"{type(e).__name__}: {e}"
            delay = min(8.0, 0.8 * (2 ** (attempt - 1)))
            if on_attempt:
                on_attempt(attempt, retries, f"network backoff {delay:.1f}s")
            time.sleep(delay)
            continue

        # overall wall-clock safeguard (avoid endless retries)
        if (time.time() - t0) > (timeout_s + 120):
            break

    raise TranslateCallError(f"translate failed after retries: {last_err}")


def doubao_translate_items(
    client: OpenAI,
    model: str,
    items: List[dict],
    src_lang: str,
    dst_lang: str,
    timeout_s: int,
    on_attempt: Optional[Callable[[int, int, str], None]] = None,
    debug_sink: Optional[Callable[[str], None]] = None,
) -> Dict[str, List[str]]:
    """Translate a batch of items via the dedicated translation model (Responses API + translation_options).

    ⚠️ IMPORTANT CHANGE (per requirement):
    - When "translate" is enabled, we ALWAYS use DEFAULT_TRANSLATE_MODEL (ep-20260203211039-n9v2f).
    - We do NOT use any other model for translation.
    - We do NOT rely on JSON-structured output from the model; we translate each segment directly.

    Returns:
        dict[item_id] = translated_segments (same length as input segments)
    """
    api_key = get_api_key()
    base_url = get_ark_base_url()
    translate_model = DEFAULT_TRANSLATE_MODEL

    # Flatten tasks
    tasks: List[Tuple[str, int, str]] = []
    for it in items:
        _id = it.get("id")
        segs = it.get("segments")
        if not isinstance(_id, str) or not isinstance(segs, list):
            continue
        for j, seg in enumerate(segs):
            if not isinstance(seg, str):
                seg = ""
            tasks.append((_id, j, seg))

    # Pre-allocate output buffers
    out: Dict[str, List[Optional[str]]] = {}
    for it in items:
        _id = it.get("id")
        segs = it.get("segments")
        if isinstance(_id, str) and isinstance(segs, list):
            out[_id] = [None] * len(segs)

    # Translate in parallel (best-effort). Keep worker count conservative.
    max_workers = min(8, max(2, (os.cpu_count() or 4)))
    errors: List[str] = []

    def _do_one(t: Tuple[str, int, str]) -> Tuple[str, int, str]:
        _id, j, seg = t
        if not seg.strip():
            return _id, j, seg  # keep whitespace
        translated = ark_translate_text_via_responses(
            api_key=api_key,
            base_url=base_url,
            text=seg,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            model=translate_model,
            retries=6,
            on_attempt=on_attempt,
        )
        if debug_sink:
            # only store a short snippet to avoid huge logs
            debug_sink(translated[:400])
        return _id, j, translated

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(_do_one, t) for t in tasks]
        for fut in concurrent.futures.as_completed(futs):
            try:
                _id, j, translated = fut.result()
                if _id in out and 0 <= j < len(out[_id]):
                    out[_id][j] = translated
            except Exception as e:
                errors.append(f"{type(e).__name__}: {e}")

    if errors:
        raise TranslateCallError("; ".join(errors[:3]) + (" ..." if len(errors) > 3 else ""))

    # Finalize: convert Optional[str] -> str, and validate completeness
    final: Dict[str, List[str]] = {}
    for _id, segs in out.items():
        if any(s is None for s in segs):
            raise TranslateCallError(f"missing translated segments for {_id}")
        final[_id] = [s or "" for s in segs]

    return final


def estimate_auto_batch_chars(items: List[dict], target_batches: int, clamp_min: int = 4000, clamp_max: int = 20000) -> int:
    """
    自动分批大小：
    - 先估 total_json_chars
    - 再除以 target_batches 得到 max_batch_chars
    - clamp 到 [clamp_min, clamp_max]
    """
    if not items:
        return 12000
    total = 0
    for it in items:
        total += len(json.dumps(it, ensure_ascii=False))
    est = int(total / max(1, target_batches))
    return int(max(clamp_min, min(clamp_max, est)))


def translate_items_adaptive(
    client: OpenAI,
    model: str,
    items: List[dict],
    src_lang: str,
    dst_lang: str,
    timeout_s: int,
    on_attempt: Optional[Callable[[int, int, str], None]] = None,
    debug_sink: Optional[Callable[[str], None]] = None,
     max_depth: int = 3,
    target_batches: int = 6,
) -> Dict[str, List[str]]:
    """
    自适应翻译调度器（保留接口不动，内部仍强制走 dedicated translation model）。

    设计目标：
    - 不动你其他功能；
    - 翻译时全部使用 DEFAULT_TRANSLATE_MODEL（ep-20260203211039-n9v2f）；
    - 如果某次翻译失败，按“items 批次”递归二分，直到 depth 用尽；
    - 最终返回 dict[id] = segments[]（长度一致）。

    注意：实际翻译粒度已在 doubao_translate_items() 内部做到了 segment-level 并发，
    这里的分批更多是为了“失败时缩小范围 + 更稳”。
    """
    if not items:
        return {}

    # 首先尝试一次性翻译（最省时）
    try:
        return doubao_translate_items(
            client=client,
            model=model,  # ignored for translation (kept for compatibility)
            items=items,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            on_attempt=on_attempt,
            debug_sink=debug_sink,
        )
    except Exception as e:
        if max_depth <= 0 or len(items) == 1:
            raise

        # 失败则二分递归（更稳）
        mid = max(1, len(items) // 2)
        left = items[:mid]
        right = items[mid:]

        left_res = translate_items_adaptive(
            client=client,
            model=model,
            items=left,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            on_attempt=on_attempt,
            debug_sink=debug_sink,
            max_depth=max_depth - 1,
            target_batches=target_batches,
        )
        right_res = translate_items_adaptive(
            client=client,
            model=model,
            items=right,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            on_attempt=on_attempt,
            debug_sink=debug_sink,
            max_depth=max_depth - 1,
            target_batches=target_batches,
        )
        left_res.update(right_res)
        return left_res


# ============================================================
# 7) DOCX translate pipeline (extract->translate->apply)
# ============================================================
def paragraph_get_text(p: Paragraph) -> str:
    # python-docx 里 p.text 会合并 runs，但会丢失某些 field；这里仍用 p.text 作为 best-effort
    return p.text or ""

def paragraph_set_text_preserve_style(p: Paragraph, new_text: str):
    """
    best-effort：保留段落级样式，但 runs 样式不完全保留（除非做更复杂的 run-diff）。
    这里不动你现有结构，采用“清空 runs -> 写一个 run”策略。
    """
    # 清空原 runs
    for r in list(p.runs):
        try:
            r._r.getparent().remove(r._r)
        except Exception:
            pass
    # 写入新 run
    run = p.add_run(new_text)
    # 尽量保留段落 style：p.style 不动即可

def build_items_from_docx(doc: Document) -> Tuple[List[dict], Dict[str, List[Paragraph]]]:
    """
    将 docx 中所有段落（含表格、页眉页脚）抽取成 items:
      items = [{"id": "...", "segments":[seg]}]
    同时返回 id -> Paragraph list 映射，用于回写。
    """
    ps = iter_all_paragraphs_extended(doc)
    id2ps: Dict[str, List[Paragraph]] = {}
    items: List[dict] = []

    for i, p in enumerate(ps):
        txt = paragraph_get_text(p)
        # 跳过空段落
        if not txt or not txt.strip():
            continue

        pid = f"p_{i}"
        id2ps[pid] = [p]
        items.append({"id": pid, "segments": [txt]})

    return items, id2ps

def apply_translation_to_docx(
    doc: Document,
    translated: Dict[str, List[str]],
    id2ps: Dict[str, List[Paragraph]],
    protect_math_enabled: bool = True,
):
    """
    回写翻译结果到 doc（best-effort）
    - protect_math_enabled：对段落文本中的公式/环境先做 __MATH_k__ 保护，翻译后再恢复
    """
    for pid, ps in id2ps.items():
        if pid not in translated:
            continue
        segs = translated[pid]
        if not segs:
            continue
        new_text = segs[0] if isinstance(segs[0], str) else ""

        for p in ps:
            paragraph_set_text_preserve_style(p, new_text)

def translate_docx_bytes(
    docx_bytes: bytes,
    src_lang: str,
    dst_lang: str,
    translate_timeout_s: int,
    on_progress: Optional[Callable[[str], None]] = None,
) -> bytes:
    """
    Word 内翻译（保留排版 best-effort）：
    - 抽取段落文本
    - 对每段进行公式占位保护（避免翻译破坏 $...$ / $$...$$ / \(...\) / equation env）
    - 使用翻译模型（Responses API）翻译
    - 恢复公式占位符
    - 回写 docx
    """
    doc = Document(io.BytesIO(docx_bytes))

    items, id2ps = build_items_from_docx(doc)

    # 1) protect math tokens per segment
    protected_items: List[dict] = []
    token_maps: Dict[str, Dict[str, str]] = {}

    for it in items:
        pid = it["id"]
        seg = it["segments"][0]
        protected, mapping = protect_math(seg)
        token_maps[pid] = mapping
        protected_items.append({"id": pid, "segments": [protected]})

    client = get_ark_client(default_timeout_s=translate_timeout_s)

    def _attempt_cb(a: int, n: int, phase: str):
        if on_progress:
            on_progress(f"Translate attempt {a}/{n}: {phase}")

    translated_protected = translate_items_adaptive(
        client=client,
        model=get_default_model(),  # ignored for translation, kept for compatibility
        items=protected_items,
        src_lang=src_lang,
        dst_lang=dst_lang,
        timeout_s=translate_timeout_s,
        on_attempt=_attempt_cb,
        debug_sink=None,
        max_depth=3,
    )

    # 2) restore math tokens
    translated_restored: Dict[str, List[str]] = {}
    for pid, segs in translated_protected.items():
        mapping = token_maps.get(pid, {})
        restored = restore_tokens(segs[0], mapping) if segs else ""
        translated_restored[pid] = [restored]

    # 3) apply back
    apply_translation_to_docx(doc, translated_restored, id2ps, protect_math_enabled=True)

    out_buf = io.BytesIO()
    doc.save(out_buf)
    return out_buf.getvalue()


# ============================================================
# 8) DOCX -> Markdown/LaTeX (Pandoc) + OMML -> $...$ / $$...$$ injection
# ============================================================
def docx_to_markdown_bytes(docx_bytes: bytes) -> str:
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        f = td / "in.docx"
        f.write_bytes(docx_bytes)
        md = pypandoc.convert_file(
            str(f),
            to="markdown",
            format="docx",
            extra_args=["--wrap=none"],
        )
    return normalize_md(md)

def docx_to_latex_bytes(docx_bytes: bytes) -> str:
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        f = td / "in.docx"
        f.write_bytes(docx_bytes)
        tex = pypandoc.convert_file(
            str(f),
            to="latex",
            format="docx",
            extra_args=["--wrap=none"],
        )
    return tex.replace("\r\n", "\n").replace("\r", "\n").strip()

def export_docx_equations_to_pandoc_math_docx(docx_bytes: bytes, use_equation_env: bool = False) -> bytes:
    """
    核心：把 Word OMML 公式替换成 Pandoc 可渲染的数学标记（默认 $...$ / $$...$$）
    - 先用 pandoc 把 docx 转 md 提取公式序列（math_seq）
    - 再在原 docx XML 里按顺序把 oMath/oMathPara 替换成 latex 文本
    """
    math_seq = extract_math_from_docx_with_pandoc(docx_bytes)
    doc = Document(io.BytesIO(docx_bytes))
    n = replace_omml_with_latex_code(doc, math_seq, use_equation_env=use_equation_env)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ============================================================
# 9) Streamlit UI
# ============================================================
with st.sidebar:
    st.header("配置")
    st.write("Ark Base URL:", get_ark_base_url())
    st.write("默认（非翻译）模型:", get_default_model())
    st.write("翻译固定模型:", DEFAULT_TRANSLATE_MODEL)

    ocr_timeout_s, tr_timeout_s = get_timeout_defaults()
    st.subheader("超时")
    ocr_timeout_s = st.number_input("OCR 单次超时(秒)", min_value=30, max_value=600, value=int(ocr_timeout_s), step=10)
    tr_timeout_s = st.number_input("翻译单次超时(秒)", min_value=30, max_value=600, value=int(tr_timeout_s), step=10)

    st.subheader("OCR 参数")
    max_side = st.number_input("最大边长缩放", min_value=800, max_value=4000, value=2200, step=100)
    tile_h = st.number_input("分块高度(px)", min_value=800, max_value=4000, value=1800, step=100)
    overlap = st.number_input("分块重叠(px)", min_value=0, max_value=800, value=120, step=20)
    jpeg_q = st.number_input("JPEG质量", min_value=50, max_value=95, value=85, step=1)
    ocr_out_tokens = st.number_input("OCR 输出 tokens", min_value=256, max_value=8192, value=3000, step=256)

    st.subheader("公式转换")
    use_equation_env = st.checkbox("Word公式导出为 \\begin{equation}...（否则 $$...$$）", value=False)

    if HAVE_LATEX_OMML:
        st.subheader("LaTeX->Word 公式（可选）")
        xsl_path = st.text_input("MML2OMML.XSL 路径(可留空自动猜)", value=os.environ.get(DEFAULT_MML2OMML_XSL_ENV, ""))
    else:
        xsl_path = ""


tab1, tab2, tab3, tab4 = st.tabs(["OCR（PDF/图片）→Markdown", "DOCX 内翻译", "DOCX 公式→Pandoc数学标记", "LaTeX($..$)→Word公式(可选)"])


# ----------------------------
# Tab 1: OCR
# ----------------------------
with tab1:
    st.subheader("OCR（保持中文，不翻译）")
    uploaded = st.file_uploader("上传 PDF 或图片", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

    if st.button("开始 OCR", type="primary", disabled=not uploaded):
        api_key = get_api_key()
        client = get_ark_client(default_timeout_s=int(ocr_timeout_s))
        model = get_default_model()  # ✅ OCR 仍用默认模型，不动

        all_images: List[Image.Image] = []
        for f in uploaded:
            data = f.read()
            if f.name.lower().endswith(".pdf"):
                if not HAVE_PYMUPDF:
                    st.error("缺少 pymupdf，无法渲染 PDF。请安装：pip install pymupdf")
                    st.stop()
                imgs = pdf_bytes_to_images(data, dpi=300, max_pages=None)
                all_images.extend(imgs)
            else:
                img = Image.open(io.BytesIO(data)).convert("RGB")
                all_images.append(img)

        md = ocr_images_to_markdown(
            client=client,
            images=all_images,
            model=model,
            max_side=int(max_side),
            tile_h=int(tile_h),
            overlap=int(overlap),
            jpeg_q=int(jpeg_q),
            out_tokens=int(ocr_out_tokens),
            timeout_s=int(ocr_timeout_s),
        )
        st.success("OCR 完成")
        st.text_area("OCR Markdown 输出", value=md, height=400)

        st.download_button("下载 Markdown", data=md.encode("utf-8"), file_name="ocr.md", mime="text/markdown")


# ----------------------------
# Tab 2: DOCX translate
# ----------------------------
with tab2:
    st.subheader("DOCX 内翻译（翻译模型固定）")
    docx_up = st.file_uploader("上传 DOCX", type=["docx"], key="docx_translate")

    colA, colB = st.columns(2)
    with colA:
        src_lang = st.selectbox("源语言（可 Auto）", ["Auto", "zh", "zh-Hant", "en", "ja", "ko", "de", "fr", "es", "it", "pt", "ru", "th", "vi", "ar", "cs", "da"], index=0)
    with colB:
        dst_lang = st.selectbox("目标语言", ["zh", "zh-Hant", "en", "ja", "ko", "de", "fr", "es", "it", "pt", "ru", "th", "vi", "ar", "cs", "da"], index=2)

    if st.button("开始翻译并导出 DOCX", type="primary", disabled=not docx_up):
        raw = docx_up.read()

        prog = st.empty()
        def _log(s: str):
            prog.info(s)

        try:
            out_bytes = translate_docx_bytes(
                docx_bytes=raw,
                src_lang=src_lang,
                dst_lang=dst_lang,
                translate_timeout_s=int(tr_timeout_s),
                on_progress=_log,
            )
        except Exception as e:
            st.error(f"翻译失败：{type(e).__name__}: {e}")
            st.stop()

        st.success("翻译完成")
        st.download_button(
            "下载 翻译后的 DOCX",
            data=out_bytes,
            file_name="translated.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )


# ----------------------------
# Tab 3: DOCX equation -> pandoc math delimiters
# ----------------------------
with tab3:
    st.subheader("把 Word OMML 公式替换为 Pandoc 可渲染数学标记（$...$ / $$...$$）")
    docx_eq = st.file_uploader("上传 DOCX（含可编辑公式）", type=["docx"], key="docx_eq")

    if st.button("转换并导出 DOCX（公式变为 $ 标记）", type="primary", disabled=not docx_eq):
        raw = docx_eq.read()
        try:
            out = export_docx_equations_to_pandoc_math_docx(raw, use_equation_env=bool(use_equation_env))
        except Exception as e:
            st.error(f"公式转换失败：{type(e).__name__}: {e}")
            st.stop()

        st.success("转换完成")
        st.download_button(
            "下载 DOCX（Pandoc 数学标记）",
            data=out,
            file_name="equations_as_pandoc_math.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

        # 也附带导出 markdown 方便你快速验证 pandoc 渲染是否正常
        try:
            md = docx_to_markdown_bytes(out)
            st.text_area("可选：转换后 Markdown（pandoc 导出）", value=md, height=260)
            st.download_button("下载 Markdown", data=md.encode("utf-8"), file_name="equations.md", mime="text/markdown")
        except Exception:
            pass


# ----------------------------
# Tab 4: LaTeX($..$) -> Word OMML (optional)
# ----------------------------
with tab4:
    st.subheader("把文本中的 $...$ / $$...$$ 转成 Word 可编辑公式（可选功能）")
    if not HAVE_LATEX_OMML:
        st.warning("此功能需要安装：pip install lxml latex2mathml，并提供 MML2OMML.XSL（Office 自带）。")
    else:
        docx_l2w = st.file_uploader("上传 DOCX（里面包含 $...$ 公式文本）", type=["docx"], key="docx_l2w")
        if st.button("转换并导出 DOCX（$ 公式 -> OMML）", type="primary", disabled=not docx_l2w):
            raw = docx_l2w.read()
            doc = Document(io.BytesIO(raw))
            logbox = st.empty()

            def _log(s: str):
                logbox.info(s)

            try:
                n = convert_inline_latex_to_omml_in_doc(doc, xsl_path=xsl_path, log=_log)
            except Exception as e:
                st.error(f"转换失败：{type(e).__name__}: {e}")
                st.stop()

            out = io.BytesIO()
            doc.save(out)
            st.success(f"完成：转换 {n} 处公式")
            st.download_button(
                "下载 DOCX（含可编辑公式）",
                data=out.getvalue(),
                file_name="latex_to_omml.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

