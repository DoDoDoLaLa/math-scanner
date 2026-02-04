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
            try:
                on_attempt(attempt, retries, "request")
            except Exception:
                pass

        try:
            # 重要：某些环境里 http 超时/流控可能导致请求“看起来卡死”。
            # 这里用线程包一层硬超时：即使底层库未按预期抛超时，我们也能继续重试并给出可见反馈。
            def _do_req():
                kwargs = dict(
                    model=model,
                    messages=messages,
                    temperature=temperature,
                    max_tokens=max_tokens,
                    timeout=timeout_s,
                )
                if response_format is not None:
                    kwargs["response_format"] = response_format
                try:
                    return client.chat.completions.create(**kwargs)
                except TypeError:
                    # 某些 OpenAI 兼容实现不支持 response_format；回退到普通请求
                    kwargs.pop("response_format", None)
                    return client.chat.completions.create(**kwargs)

            with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
                fut = ex.submit(_do_req)
                resp = fut.result(timeout=float(timeout_s) + 5.0)
            txt = ""
            if resp and resp.choices and resp.choices[0].message and resp.choices[0].message.content:
                txt = resp.choices[0].message.content
            return ArkResult(text=(txt or "").strip())

        except concurrent.futures.TimeoutError:
            msg = f"TimeoutError: request exceeded {timeout_s}s"
            last_err = msg
            wait_s = min(2 ** (attempt - 1), 20)
            if on_attempt:
                try:
                    on_attempt(attempt, retries, f"request timeout, sleep {wait_s:.1f}s")
                except Exception:
                    pass
            time.sleep(wait_s)
            continue

        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            last_err = msg

            # 429 / 限流：按服务端提示或指数退避，但仍受 total_timeout_s 约束
            if ("429" in msg) or ("rate limit" in msg.lower()) or ("RESOURCE_EXHAUSTED" in msg):
                wait_s = _parse_retry_delay_seconds(msg) or min(2 ** (attempt - 1), 60)
                if on_attempt:
                    try:
                        on_attempt(attempt, retries, f"rate-limited, sleep {wait_s:.1f}s")
                    except Exception:
                        pass
                time.sleep(min(wait_s + 0.3, 90.0))
                continue

            # 其它错误：短暂退避
            wait_s = min(2 ** (attempt - 1), 20)
            if on_attempt:
                try:
                    on_attempt(attempt, retries, f"error, sleep {wait_s:.1f}s")
                except Exception:
                    pass
            time.sleep(wait_s)

    return ArkResult(text="", error_message=f"retry exhausted: {last_err}")


# ============================================================
# 3) OCR: image analysis + auto recommend
# ============================================================
def analyze_image_density(img: Image.Image) -> Dict[str, float]:
    g = ImageOps.grayscale(img)
    g_small = g.copy()
    g_small.thumbnail((900, 900))
    w, h = g_small.size

    edges = g_small.filter(ImageFilter.FIND_EDGES)
    eb = edges.point(lambda p: 255 if p > 40 else 0)
    edge_px = sum(1 for p in eb.getdata() if p > 0)
    edge_density = edge_px / float(w * h + 1e-9)

    db = g_small.point(lambda p: 1 if p < 160 else 0)
    dark_px = sum(db.getdata())
    dark_ratio = dark_px / float(w * h + 1e-9)

    return {"edge_density": edge_density, "dark_ratio": dark_ratio, "w": float(img.size[0]), "h": float(img.size[1])}

def recommend_ocr_params(img: Image.Image, mode: str = "Balanced") -> Dict[str, int]:
    m = analyze_image_density(img)
    h = int(m["h"])
    edge = m["edge_density"]
    dark = m["dark_ratio"]
    complexity = 0.6 * edge + 0.4 * dark

    if mode == "Fast":
        max_side, jpeg_q, out_tokens = 1600, 80, 3072
        base_tile, base_overlap = 2000, 120
    elif mode == "Accurate":
        max_side, jpeg_q, out_tokens = 2800, 90, 6144
        base_tile, base_overlap = 1400, 220
    else:
        max_side, jpeg_q, out_tokens = 2200, 85, 4096
        base_tile, base_overlap = 1600, 160

    if complexity > 0.18:
        max_side = min(3200, max_side + 400)
        out_tokens = min(8192, out_tokens + 1024)
        tile_h = max(1000, base_tile - 300)
        overlap = min(320, base_overlap + 80)
        jpeg_q = min(95, jpeg_q + 5)
    elif complexity < 0.10:
        tile_h = min(2400, base_tile + 300)
        overlap = max(80, base_overlap - 40)
    else:
        tile_h = base_tile
        overlap = base_overlap

    if h >= 4000 and complexity > 0.14:
        tile_h = max(900, tile_h - 200)
        out_tokens = min(8192, out_tokens + 512)

    max_side = int(max(800, min(3200, max_side)))
    tile_h = int(max(800, min(2600, tile_h)))
    overlap = int(max(0, min(400, overlap)))
    jpeg_q = int(max(50, min(95, jpeg_q)))
    out_tokens = int(max(1024, min(8192, out_tokens)))

    return {"max_side": max_side, "tile_h": tile_h, "overlap": overlap, "jpeg_q": jpeg_q, "out_tokens": out_tokens}


# ============================================================
# 4) OCR helpers
# ============================================================
def downscale(img: Image.Image, max_side: int) -> Image.Image:
    img = img.convert("RGB")
    w, h = img.size
    s = min(1.0, max_side / float(max(w, h)))
    if s < 1.0:
        img = img.resize((int(w * s), int(h * s)), Image.LANCZOS)
    return img

def pil_to_jpeg_bytes(img: Image.Image, quality: int = 85) -> bytes:
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality, optimize=True)
    return buf.getvalue()

def slice_long(img: Image.Image, tile_h: int, overlap: int) -> List[Image.Image]:
    w, h = img.size
    if h <= tile_h:
        return [img]
    overlap = max(0, min(overlap, tile_h // 2))
    step = tile_h - overlap
    out = []
    y = 0
    while y < h:
        y2 = min(y + tile_h, h)
        out.append(img.crop((0, y, w, y2)))
        if y2 >= h:
            break
        y = y + step
    return out

def normalize_md(md: str) -> str:
    md = md.replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return md

def dedupe_tail_head(prev: str, cur: str, max_lines: int = 18) -> str:
    pl = [x.strip() for x in prev.splitlines() if x.strip()]
    cl = [x.strip() for x in cur.splitlines() if x.strip()]
    if not pl or not cl:
        return cur
    k = min(max_lines, len(pl), len(cl))
    for n in range(k, 3, -1):
        if pl[-n:] == cl[:n]:
            raw = cur.splitlines()
            removed = 0
            keep = []
            for line in raw:
                if removed < n and line.strip():
                    removed += 1
                    continue
                keep.append(line)
            return "\n".join(keep).lstrip("\n")
    return cur

def _jpeg_bytes_to_data_url(jpeg_bytes: bytes) -> str:
    b64 = base64.b64encode(jpeg_bytes).decode("utf-8")
    return f"data:image/jpeg;base64,{b64}"

def ocr_image_to_markdown(
    client: OpenAI,
    model: str,
    img: Image.Image,
    max_side: int,
    tile_h: int,
    overlap: int,
    jpeg_q: int,
    out_tokens: int,
    timeout_s: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> str:
    img = downscale(img, max_side=max_side)
    tiles = slice_long(img, tile_h=tile_h, overlap=overlap)

    chunks: List[str] = []
    total = len(tiles)

    for idx, timg in enumerate(tiles, start=1):
        if progress_cb:
            progress_cb(idx, total)

        b = pil_to_jpeg_bytes(timg, quality=jpeg_q)
        data_url = _jpeg_bytes_to_data_url(b)

        messages = [{
            "role": "user",
            "content": [
                {"type": "text", "text": OCR_PROMPT_ZH},
                {"type": "image_url", "image_url": {"url": data_url}},
            ],
        }]

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

def replace_omml_with_latex_code(doc: Document, math_seq: List[Tuple[str, str]], use_equation_env: bool = True) -> int:
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

    Ark /api/v3/responses follows the OpenAI Responses schema:
      resp.output -> list[message] -> content -> list[{type: "output_text", text: "..."}]
    Some gateways may also return ChatCompletions-style `choices`.
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
                    if isinstance(part, dict) and part.get("type") == "output_text" and isinstance(part.get("text"), str):
                        texts.append(part["text"])

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

    Docs: /api/v3/responses supports input_text.translation_options.source_language/target_language.
    Response: output.content.type=output_text, output.content.text holds the text.
    """
    url = base_url.rstrip("/") + "/responses"

    # Ark translation expects language codes; we accept friendly labels from UI and map to codes.
    lang_map = {
        "Auto": None,
        "Chinese": "zh",
        "English": "en",
        "Japanese": "ja",
        "Korean": "ko",
        "Spanish": "es",
        "French": "fr",
        "German": "de",
        "Portuguese": "pt",
        "Russian": "ru",
        "Thai": "th",
        "Vietnamese": "vi",
        "Arabic": "ar",
        "Czech": "cs",
        "Danish": "da",
    }

    src_code = lang_map.get(src_lang, src_lang)
    dst_code = lang_map.get(dst_lang, dst_lang) or dst_lang
    if dst_code in (None, "", "Auto"):
        raise TranslateCallError("target language must be specified (not Auto)")

    trans_opts: Dict[str, str] = {"target_language": dst_code}
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
                try:
                    err_obj = r.json()
                except Exception:
                    err_obj = None
                msg = None
                if isinstance(err_obj, dict):
                    err = err_obj.get("error")
                    if isinstance(err, dict):
                        msg = err.get("message") or err.get("msg")
                    msg = msg or err_obj.get("message") or err_obj.get("msg")
                msg = msg or f"HTTP {r.status_code}: {r.text[:400]}"
                last_err = msg

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
                raise TranslateCallError("empty translate response text")
            return out_text

        except (httpx.TimeoutException, httpx.TransportError) as e:
            last_err = f"{type(e).__name__}: {e}"
            delay = min(8.0, 0.8 * (2 ** (attempt - 1)))
            if on_attempt:
                on_attempt(attempt, retries, f"network backoff {delay:.1f}s")
            time.sleep(delay)

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
    """Translate a batch of items via the dedicated translation model (Responses API).

    IMPORTANT (per requirement):
    - When translation is enabled, we ALWAYS use DEFAULT_TRANSLATE_MODEL = ep-20260203211039-n9v2f.
    - We do NOT use other models for translation and do NOT ask the model to output JSON.
    - We translate each segment directly so $...$/__MATH_x__ placeholders are preserved by protect_math().
    """
    api_key = get_api_key()
    base_url = get_ark_base_url()
    translate_model = DEFAULT_TRANSLATE_MODEL

    # Flatten tasks
    tasks: List[Tuple[str, int, str]] = []
    out_buf: Dict[str, List[Optional[str]]] = {}

    for it in items:
        _id = it.get("id")
        segs = it.get("segments")
        if not isinstance(_id, str) or not isinstance(segs, list):
            continue
        out_buf[_id] = [None] * len(segs)
        for j, seg in enumerate(segs):
            tasks.append((_id, j, seg if isinstance(seg, str) else ""))

    # Concurrency: keep conservative to avoid rate-limit bursts.
    max_workers = min(4, max(2, (os.cpu_count() or 4) // 2))
    errors: List[str] = []

    def _do_one(t: Tuple[str, int, str]) -> Tuple[str, int, str]:
        _id, j, seg = t
        if not seg.strip():
            return _id, j, seg

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
            debug_sink(translated[:400])
        return _id, j, translated

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(_do_one, t) for t in tasks]
        for fut in concurrent.futures.as_completed(futs):
            try:
                _id, j, translated = fut.result()
                if _id in out_buf and 0 <= j < len(out_buf[_id]):
                    out_buf[_id][j] = translated
            except Exception as e:
                errors.append(f"{type(e).__name__}: {e}")

    # Retry failed segments sequentially once (often fixes transient 429).
    if any(v is None for segs in out_buf.values() for v in segs):
        for _id, segs in list(out_buf.items()):
            for j, val in enumerate(segs):
                if val is not None:
                    continue
                seg = next((s for (iid, jj, s) in tasks if iid == _id and jj == j), "")
                try:
                    out_buf[_id][j] = ark_translate_text_via_responses(
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
                except Exception as e:
                    errors.append(f"{type(e).__name__}: {e}")

    if errors and any(v is None for segs in out_buf.values() for v in segs):
        raise TranslateCallError("; ".join(errors[:3]) + (" ..." if len(errors) > 3 else ""))

    final: Dict[str, List[str]] = {}
    for _id, segs in out_buf.items():
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
    max_depth: int = 10,
) -> Dict[str, List[str]]:
    """Translate with automatic batch splitting on timeout/invalid JSON.

    This is the main fix for “small text works, large text times out / hangs”.
    Strategy:
    - Try translating the batch as-is.
    - If it fails (timeout, rate-limit, invalid JSON), split the batch into halves and retry.
    - Merge results. Stops splitting when batch has 1 item or max_depth reached.
    """
    if not items:
        return {}
    try:
        return doubao_translate_items(
            client=client,
            model=model,
            items=items,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            on_attempt=on_attempt,
            debug_sink=debug_sink,
        )
    except TranslateCallError as e:
        # Only split if we can
        if len(items) <= 1 or max_depth <= 0:
            raise
        mid = len(items) // 2
        left = translate_items_adaptive(client, model, items[:mid], src_lang, dst_lang, timeout_s,
                                       on_attempt=on_attempt, debug_sink=debug_sink, max_depth=max_depth - 1)
        right = translate_items_adaptive(client, model, items[mid:], src_lang, dst_lang, timeout_s,
                                        on_attempt=on_attempt, debug_sink=debug_sink, max_depth=max_depth - 1)
        left.update(right)
        return left

def translate_docx_in_place(
    doc: Document,
    client: OpenAI,
    model: str,
    src_lang: str,
    dst_lang: str,
    max_batch_chars: int,
    timeout_s: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
    diagnostic_cb: Optional[Callable[[Dict[str, Any]], None]] = None,
    attempt_msg_cb: Optional[Callable[[str], None]] = None,
):
    all_ps = iter_all_paragraphs_extended(doc)

    items: List[dict] = []
    para_refs: Dict[str, Tuple[List[Any], List[Dict[str, str]]]] = {}
    pid = 0

    total_runs = 0
    total_visible_chars = 0

    for p in all_ps:
        runs = [r for r in p.runs if r.text is not None and r.text != ""]
        if not runs:
            continue

        segs = []
        maps = []
        for r in runs:
            total_runs += 1
            total_visible_chars += len(r.text or "")
            protected, mp = protect_math(r.text)
            segs.append(protected)
            maps.append(mp)

        if not "".join(segs).strip():
            continue

        pid += 1
        item_id = f"p{pid}"
        items.append({"id": item_id, "segments": segs})
        para_refs[item_id] = (runs, maps)

    if diagnostic_cb is not None:
        diagnostic_cb({
            "paragraphs_scanned": len(all_ps),
            "items_built": len(items),
            "runs_counted": total_runs,
            "visible_chars_counted": total_visible_chars,
        })

    if not items:
        # 关键：明确提示为什么“没有翻译”
        st.warning(
            "没有抓到可翻译的正文文本（items=0）。\n\n"
            "常见原因：\n"
            "1) 内容在文本框/形状（python-docx 默认无法读取）；\n"
            "2) 内容主要在批注/脚注/尾注；\n"
            "3) 内容是嵌入对象或图片。\n\n"
            "建议：\n"
            "- 若是文本框：优先用 Tab③ Pandoc 导出，或将文本框内容复制到正文段落后再翻译；\n"
            "- 或改走 PDF/图片 OCR 路线。"
        )
        return

    batches = chunk_items_for_api(items, max_chars=max_batch_chars)
    total = len(batches)

    for bi, batch in enumerate(batches, start=1):
        if progress_cb:
            progress_cb(bi, total)

        if attempt_msg_cb:
            try:
                approx_chars = sum(len(json.dumps(it, ensure_ascii=False)) for it in batch)
            except Exception:
                approx_chars = -1
            attempt_msg_cb(
                f"批次 {bi}/{total}：发送翻译请求（items={len(batch)}，approx_chars={approx_chars}）"
            )

        translated_map = translate_items_adaptive(
            client=client,
            model=model,
            items=batch,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            on_attempt=(lambda a, n, ph: attempt_msg_cb(f"批次 {bi}/{total}：请求尝试 {a}/{n} · {ph}") if attempt_msg_cb else None),
            debug_sink=(lambda raw: attempt_msg_cb(f"批次 {bi}/{total}：已收到原始输出（len={len(raw)})") if attempt_msg_cb else None),
        )

        if attempt_msg_cb:
            attempt_msg_cb(f"批次 {bi}/{total}：收到响应，开始解析/写回…")

        for it in batch:
            item_id = it["id"]
            if item_id not in translated_map:
                continue

            runs, maps = para_refs[item_id]
            out_segs = translated_map[item_id]

            if len(out_segs) != len(runs):
                whole = " ".join(out_segs)
                whole = restore_tokens(whole, {k: v for mp in maps for k, v in mp.items()})
                runs[0].text = whole
                for r in runs[1:]:
                    r.text = ""
                continue

            for r, seg, mp in zip(runs, out_segs, maps):
                r.text = restore_tokens(seg, mp)

        if attempt_msg_cb:
            attempt_msg_cb(f"批次 {bi}/{total}：写回完成")


# ============================================================
# 7) Export helpers
# ============================================================
def doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def pandoc_md_to_docx(md: str) -> bytes:
    md = normalize_md(md) + "\n"
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.docx"
        fmt = "markdown+fenced_code_blocks+tex_math_dollars"
        pypandoc.convert_text(md, to="docx", format=fmt, outputfile=str(out))
        return out.read_bytes()

# ============================================================
# 7.7) AI-assisted "DOCX text -> OCR-style math -> OMML" (Tab 3 enhancement)
# ============================================================
AI_MATH_CLEAN_PROMPT_ZH = """你是论文排版助理。请把我给你的 Markdown 文本“尽量保持原文不变”，只对其中的数学公式做纠错与统一格式，使其更像“图片 OCR 输出的学术 LaTeX”风格，便于后续用 Pandoc 解析成 Word 可编辑公式（OMML）。

严格要求：
1) 仅处理数学公式：位于 $...$ 或 $$...$$ 或 \\( ... \\) 或 \\[ ... \\] 或 \\begin{equation}...\\end{equation} 内的内容。普通文本不要改写措辞。
2) 统一分隔符：
   - 行内公式统一成 $...$
   - 行间公式统一成 $$...$$（注意：$$...$$ 内部不要出现空行；如果有换行，保持为单个换行即可）
3) 纠错策略（尽量保守，宁可不改也不要瞎改）：
   - 把明显的“文本型下标/上标”统一为直立体：例如 P_seed -> P_{\\mathrm{seed}}；t_k^+ 中的 + 保留；类似 \\Delta_h、\\rho_h 这类命令保持不变。
   - 保留 \\alpha \\beta 等 LaTeX 命令；不要把 \\\\alpha 之类的转义搞坏。
   - 不要引入新的宏包命令；不要新增 \\label/\\ref 等。
4) 不要输出解释，不要输出代码块围栏，不要输出 JSON。只输出修正后的 Markdown 纯文本。
"""

def docx_to_pandoc_markdown_for_math(docx_bytes: bytes, wrap_none: bool = True) -> str:
    """DOCX -> Markdown(+tex_math_dollars) suitable for later math parsing."""
    if not docx_bytes:
        return ""
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        in_path = td / "in.docx"
        in_path.write_bytes(docx_bytes)
        extra_args = ["--wrap=none"] if wrap_none else []
        md = pypandoc.convert_file(
            str(in_path),
            to="markdown+tex_math_dollars",
            format="docx",
            extra_args=extra_args,
        )
    md = (md or "").replace("\r\n", "\n").replace("\r", "\n")

    # Undo common escaping patterns that break TeX math parsing
    md = md.replace(r"\\$", "$")
    md = md.replace("\\\\", "\\")  # DOCX writer often escapes backslashes
    md = normalize_pandoc_math(md, display_style="dollars", inline_style="dollars")
    md = _sanitize_tex_math_for_pandoc(md)
    return normalize_md(md)

def chunk_text_for_ai(text: str, max_chars: int = 8000) -> List[str]:
    """Chunk long markdown by paragraph boundaries."""
    if not text:
        return []
    paras = text.split("\n\n")
    chunks = []
    cur = []
    cur_len = 0
    for p in paras:
        p2 = p.strip("\n")
        if not p2:
            continue
        add = p2 + "\n\n"
        if cur and cur_len + len(add) > max_chars:
            chunks.append("".join(cur).strip() + "\n")
            cur = []
            cur_len = 0
        cur.append(add)
        cur_len += len(add)
    if cur:
        chunks.append("".join(cur).strip() + "\n")
    return chunks

def ai_clean_markdown_math_like_ocr(
    client: OpenAI,
    model: str,
    md: str,
    *,
    max_batch_chars: int,
    timeout_s: int,
    out_tokens: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> str:
    """Use LLM to normalize math to OCR-like $...$/$$...$$ style (best-effort)."""
    if not md.strip():
        return md
    parts = chunk_text_for_ai(md, max_chars=max_batch_chars)
    out_parts: List[str] = []
    total = len(parts)
    for i, part in enumerate(parts, start=1):
        if progress_cb:
            progress_cb(i, total)

        messages = [{
            "role": "user",
            "content": AI_MATH_CLEAN_PROMPT_ZH.strip() + "\n\n---\n\n" + part
        }]

        res = safe_chat_completions(
            client=client,
            model=model,
            messages=messages,
            temperature=0.0,
            max_tokens=out_tokens,
            timeout_s=timeout_s,
            retries=6,
        )
        if res.error_message:
            raise RuntimeError(res.error_message)

        cleaned = normalize_md(res.text)
        if out_parts:
            cleaned = dedupe_tail_head(out_parts[-1], cleaned)
        out_parts.append(cleaned)

    merged = normalize_md("\n\n".join([p for p in out_parts if p.strip()]))
    # Final safety: ensure pandoc-friendly
    merged = normalize_pandoc_math(merged, display_style="dollars", inline_style="dollars")
    merged = _sanitize_tex_math_for_pandoc(merged)
    return normalize_md(merged)

def docx_ai_roundtrip_make_equations_editable(
    docx_bytes: bytes,
    *,
    client: OpenAI,
    model: str,
    max_batch_chars: int,
    timeout_s: int,
    out_tokens: int,
) -> Tuple[bytes, str]:
    """DOCX -> pandoc md -> AI math clean -> docx(OMML). Returns (docx_bytes, cleaned_markdown)."""
    if not docx_bytes:
        return docx_bytes, ""
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        in_path = td / "in.docx"
        in_path.write_bytes(docx_bytes)

        md = docx_to_pandoc_markdown_for_math(docx_bytes, wrap_none=True)

        # AI refinement with Streamlit progress
        prog_bar = st.progress(0.0)
        status = st.empty()
        def _cb(i, total):
            prog_bar.progress(min(1.0, i / max(total, 1)))
            status.info(f"AI 校正中：{i}/{total}")
        md2 = ai_clean_markdown_math_like_ocr(
            client=client,
            model=model,
            md=md,
            max_batch_chars=int(max_batch_chars),
            timeout_s=int(timeout_s),
            out_tokens=int(out_tokens),
            progress_cb=_cb,
        )
        status.empty()
        prog_bar.empty()

        out_path = td / "out.docx"
        pypandoc.convert_text(
            md2,
            to="docx",
            format="markdown+tex_math_dollars",
            outputfile=str(out_path),
            extra_args=["--reference-doc", str(in_path)],
        )
        return out_path.read_bytes(), md2

# ============================================================
# 7.5) Pandoc math normalization helpers (for thesis-friendly editing)
# ============================================================
PANDOC_DISPLAY_MATH_BRACKET_RE = re.compile(r"(?s)\\\[(.+?)\\\]")  # \[ ... \]
PANDOC_DISPLAY_MATH_DOLLAR_RE = re.compile(r"(?s)\$\$(.+?)\$\$")    # $$ ... $$
PANDOC_INLINE_MATH_PAREN_RE = re.compile(r"(?s)\\\((.+?)\\\)")    # \( ... \)
PANDOC_INLINE_MATH_DOLLAR_RE = re.compile(r"(?s)(?<!\$)\$([^$\n]+)\$(?!\$)")  # $ ... $

def normalize_pandoc_math(
    text: str,
    *,
    display_style: str = "equation",   # equation | equation* | bracket | dollars
    inline_style: str = "paren",       # paren | dollars
) -> str:
    """Normalize math delimiters produced by pandoc for easier paper editing.

    - Pandoc LaTeX writer commonly emits inline math as \(...\) and display math as \[...\].
    - Some sources may already contain $$...$$. We normalize both.
    - display_style:
        - equation  : \begin{equation} ... \end{equation} (numbered)
        - equation* : \begin{equation*} ... \end{equation*} (unnumbered)
        - bracket   : \[ ... \]
        - dollars   : $$ ... $$
    - inline_style:
        - paren     : \( ... \)
        - dollars   : $ ... $
    """
    if not text:
        return text

    t = text.replace("\r\n", "\n").replace("\r", "\n")

    def _to_equation(body: str, starred: bool) -> str:
        env = "equation*" if starred else "equation"
        body = body.strip()
        return f"\\begin{{{env}}}\n{body}\n\\end{{{env}}}"

    def _display_repl(m):
        body = (m.group(1) or "").strip()
        if display_style == "equation":
            return _to_equation(body, starred=False)
        if display_style == "equation*":
            return _to_equation(body, starred=True)
        if display_style == "dollars":
            return f"$$\n{body}\n$$"
        # bracket
        return f"\\[\n{body}\n\\]"

    # Normalize display math first (both \[\] and $$ $$)
    t = PANDOC_DISPLAY_MATH_BRACKET_RE.sub(_display_repl, t)
    t = PANDOC_DISPLAY_MATH_DOLLAR_RE.sub(_display_repl, t)

    def _inline_repl(m):
        body = (m.group(1) or "").strip()
        if inline_style == "dollars":
            return f"${body}$"
        return f"\\({body}\\)"

    # Normalize inline math (both \(\) and $ $)
    t = PANDOC_INLINE_MATH_PAREN_RE.sub(_inline_repl, t)
    t = PANDOC_INLINE_MATH_DOLLAR_RE.sub(lambda m: _inline_repl(m), t)

    return t


# ============================================================
# 8) UI: sidebar
# ============================================================
ocr_timeout_default, translate_timeout_default = get_timeout_defaults()

def _init_state():
    defaults = {
        "ocr_model_id": get_default_model(),
        "translate_model_id": os.environ.get("ARK_TRANSLATE_MODEL") or DEFAULT_TRANSLATE_MODEL,
        "mml2omml_xsl": os.environ.get(DEFAULT_MML2OMML_XSL_ENV, ""),
        "max_side": 2200,
        "tile_h": 1600,
        "overlap": 160,
        "jpeg_q": 85,
        "out_tokens": 4096,
        "ocr_timeout_s": int(ocr_timeout_default),
        "translate_timeout_s": int(translate_timeout_default),
        "preset_mode": "Balanced",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()

# ✅ 关键修复：在任何 widget 创建之前应用 pending 推荐参数
if "__pending_rec__" in st.session_state:
    rec = st.session_state.pop("__pending_rec__")
    for k in ["max_side", "tile_h", "overlap", "jpeg_q", "out_tokens"]:
        if k in rec:
            st.session_state[k] = int(rec[k])
    st.session_state["__pending_rec_msg__"] = f"已回填推荐参数：{rec}"

with st.sidebar:
    st.subheader("Ark 配置")
    st.text_input("OCR model（默认 EP 已填）", key="ocr_model_id")
    st.text_input("Translate model（翻译专用 EP）", key="translate_model_id")
    st.text_input("MML2OMML.XSL 路径（可选：仅用于把 $...$ 转 Word 公式）", key="mml2omml_xsl")
    st.caption(f"Base URL: {get_ark_base_url()}")

    st.divider()
    st.subheader("OCR 参数与自动推荐")

    st.selectbox("预设", ["Fast", "Balanced", "Accurate"], key="preset_mode",
                 help="Fast：更快更省；Accurate：更清晰更稳但更慢；Balanced：折中。")

    st.slider("Max side（最长边像素）", 800, 3200, key="max_side", step=100,
              help="大：更清晰更准但更慢更贵；小：更快更省但小字/公式易错。")
    st.slider("Tile height（切片高度）", 800, 2600, key="tile_h", step=100,
              help="大：切片少更快但易截断；小：切片多更稳但更慢更贵。")
    st.slider("Overlap（切片重叠）", 0, 400, key="overlap", step=10,
              help="大：边界漏字更少但可能重复；小：更快但边界更易漏。")
    st.slider("JPEG quality（压缩质量）", 50, 95, key="jpeg_q", step=1,
              help="高：细节更清晰更准但更慢；低：更快但公式/小字更易糊。")
    st.slider("OCR max tokens（输出上限）", 1024, 8192, key="out_tokens", step=256,
              help="输出太短会截断漏字。截断优先：减 tile_h；其次：增 tokens。")

    st.number_input("OCR 请求超时（秒）", min_value=30, max_value=600, key="ocr_timeout_s", step=10,
                    help="大图/网络慢可调到 180-300。")
    st.number_input("翻译请求超时（秒）", min_value=30, max_value=900, key="translate_timeout_s", step=10,
                    help="10页以上建议 180-300；更稳可更高。")

    with st.expander("参数速查（给不懂的人）", expanded=False):
        st.markdown(
            """
- **输出经常断在半页**：先把 **Tile height** 调小（1600→1200），再把 **OCR max tokens** 调大（4096→6144）。
- **速度太慢/切片太多**：把 **Tile height** 调大（1600→2000），再把 **Max side** 略降（2200→1800）。
- **小字/公式错多**：把 **Max side** 调大（2200→2800+），或 **JPEG** 85→90/95。
- **翻译大文档更稳**：把 **max_batch_chars** 调小（12000→8000/6000），同时把超时调高（180→300）。
"""
        )


# ============================================================
# 9) Tabs
# ============================================================
tabs = st.tabs([
    "① PDF/图片 OCR → 导出",
    "② Word(.docx) → 保排版翻译/就地替换公式（best-effort）",
    "③ Word(.docx) → LaTeX/Markdown 直接导出（推荐）",
])

# ---------------------------
# Tab 1: OCR (PDF + images)
# ---------------------------
with tabs[0]:
    st.subheader("PDF/图片 OCR（带自动参数推荐 + 进度条）")
    st.write("建议：先上传 1 页代表性页面 → 点“自动推荐参数” → 再批量处理。")

    if "__pending_rec_msg__" in st.session_state:
        st.success(st.session_state.pop("__pending_rec_msg__"))

    files = st.file_uploader(
        "上传 PDF 或图片（可多选，PDF 将按页渲染）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True
    )

    col_pdf = st.columns([1, 1, 2])
    with col_pdf[0]:
        pdf_dpi = st.selectbox("PDF 渲染 DPI", [200, 300, 400], index=1)
    with col_pdf[1]:
        pdf_max_pages = st.number_input("PDF 最多页数（0=不限制）", min_value=0, max_value=500, value=0, step=1)
    with col_pdf[2]:
        if not HAVE_PYMUPDF:
            st.warning("未检测到 pymupdf，PDF 上传不可用。请 pip install pymupdf")

    # 先把所有输入转换成 images list
    images: List[Image.Image] = []
    page_meta: List[str] = []  # for page labeling

    if files:
        for f in files:
            name = getattr(f, "name", "upload")
            ext = (name.split(".")[-1] or "").lower()
            if ext == "pdf":
                if not HAVE_PYMUPDF:
                    st.error("你上传了 PDF，但环境缺少 pymupdf。请安装后重试：pip install pymupdf")
                    st.stop()
                pdf_bytes = f.read()
                try:
                    imgs = pdf_bytes_to_images(
                        pdf_bytes,
                        dpi=int(pdf_dpi),
                        max_pages=None if int(pdf_max_pages) == 0 else int(pdf_max_pages),
                    )
                except Exception as e:
                    st.error(f"PDF 渲染失败：{e}")
                    st.stop()
                for i, im in enumerate(imgs, start=1):
                    images.append(im)
                    page_meta.append(f"{name} - p{i}")
            else:
                try:
                    im = Image.open(f)
                    images.append(im)
                    page_meta.append(name)
                except Exception as e:
                    st.warning(f"无法读取图片 {name}：{e}")

    cols = st.columns([1, 1, 2])
    with cols[0]:
        auto_btn = st.button("🪄 自动推荐参数（基于第1页）", disabled=not images)
    with cols[1]:
        preset_btn = st.button("🎛️ 应用预设（Fast/Balanced/Accurate）", disabled=not images)
    with cols[2]:
        st.caption("自动推荐会根据图片高度与文字/边缘密度估计调整 max_side/tile/overlap/tokens。")

    if images and auto_btn:
        rec = recommend_ocr_params(images[0], mode=st.session_state["preset_mode"])
        st.session_state["__pending_rec__"] = rec
        st.rerun()

    if images and preset_btn:
        rec = recommend_ocr_params(images[0], mode=st.session_state["preset_mode"])
        st.session_state["__pending_rec__"] = rec
        st.rerun()

    join_lines = st.checkbox("可选：合并断行（适合 PDF 每行强制换行）", value=False)

    def merge_hard_wraps(md: str) -> str:
        parts = md.split("\n\n")
        out = []
        for p in parts:
            lines = [x.strip() for x in p.splitlines()]
            if any(l.startswith(("-", "*", "|", "#")) for l in lines):
                out.append("\n".join(p.splitlines()))
            else:
                out.append(" ".join([l for l in lines if l != ""]).strip())
        return "\n\n".join(out).strip()

    if st.button("开始 OCR 并导出", type="primary", disabled=not images):
        if not st.session_state["ocr_model_id"].strip():
            st.error("请先填写 model（ep-xxxx 或模型ID）。")
            st.stop()

        client = get_ark_client(default_timeout_s=int(st.session_state["ocr_timeout_s"]))

        pages = []
        total_pages = len(images)
        page_bar = st.progress(0, text="准备 OCR…")

        for pi, img in enumerate(images, start=1):
            tile_bar = st.progress(0, text=f"OCR 第 {pi}/{total_pages} 页：准备切片…")

            def _tile_cb(cur: int, total: int):
                tile_bar.progress(int(cur / total * 100), text=f"OCR 第 {pi}/{total_pages} 页：切片 {cur}/{total}")

            md = ocr_image_to_markdown(
                client=client,
                model=st.session_state["ocr_model_id"].strip(),
                img=img,
                max_side=int(st.session_state["max_side"]),
                tile_h=int(st.session_state["tile_h"]),
                overlap=int(st.session_state["overlap"]),
                jpeg_q=int(st.session_state["jpeg_q"]),
                out_tokens=int(st.session_state["out_tokens"]),
                timeout_s=int(st.session_state["ocr_timeout_s"]),
                progress_cb=_tile_cb,
            )

            if join_lines:
                md = merge_hard_wraps(md)

            label = page_meta[pi - 1] if pi - 1 < len(page_meta) else f"Page {pi}"
            pages.append(f"## 第 {pi} 页（{label}）\n\n{md}")
            page_bar.progress(int(pi / total_pages * 100), text=f"已完成 {pi}/{total_pages} 页")
            tile_bar.empty()

        merged_md = normalize_md("\n\n---\n\n".join(pages))

        st.success("OCR 完成")
        st.markdown("### 预览（Markdown）")
        st.code(merged_md, language="markdown")

        v1_docx = pandoc_md_to_docx(merged_md)
        v2_md_bytes = merged_md.encode("utf-8")

        v3_md = md_to_latex_code_style(merged_md)
        v3_docx = pandoc_md_to_docx(v3_md)

        st.download_button("下载 V1：Rendered.docx（pandoc渲染公式）", data=v1_docx, file_name="OCR_Rendered.docx")
        st.download_button("下载 V2：Result.md（原始 Markdown）", data=v2_md_bytes, file_name="OCR_Result.md")
        st.download_button("下载 V3：LaTeX_equation_code.docx（公式为 LaTeX 代码块）", data=v3_docx,
                           file_name="OCR_LaTeX_equation_code.docx")
        st.download_button("下载 V3：LaTeX_equation_code.md", data=v3_md.encode("utf-8"),
                           file_name="OCR_LaTeX_equation_code.md")


# ---------------------------
# Tab 2: DOCX translate + best-effort equation replace
# ---------------------------
with tabs[1]:
    st.subheader("Word(.docx) → 保排版翻译 / 公式就地替换（best-effort）")
    st.warning(
        "说明：Word 可编辑公式（OMML）要“可靠”转 LaTeX，推荐使用 Tab③ 的 Pandoc 直接导出。\n"
        "本 Tab 的“就地替换公式”为 best-effort，可能出现错位/不全。"
    )

    docx_file = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_inplace")

    colA, colB, colC = st.columns([1, 1, 1])
    with colA:
        do_equation_replace = st.checkbox("把 Word 原生公式（OMML）替换为 LaTeX 代码（best-effort）", value=False)
        do_latex_to_omml = st.checkbox("检测 $...$ / $$...$$ 并转换为 Word 公式（OMML）", value=False)
        do_translate = st.checkbox("翻译文档（尽量保持原排版）", value=False)

    with colB:
        src_lang = st.selectbox("源语言", ["Auto", "Chinese", "English", "Japanese", "Korean", "Spanish"], index=0)
        dst_lang = st.selectbox("目标语言", ["Chinese", "English", "Japanese", "Korean", "Spanish"], index=1)

    with colC:
        auto_batch = st.checkbox("自动分批（推荐）", value=True)
        target_batches = st.number_input("目标批次数（自动分批用）", min_value=1, max_value=50, value=8, step=1)

    if not auto_batch:
        max_batch_chars = st.slider("翻译分批大小（max_batch_chars）", 4000, 20000, 12000, 500)
    else:
        max_batch_chars = 12000  # will be computed after parsing doc

    show_translate_diagnostics = st.checkbox("显示翻译抓取诊断（建议打开排错）", value=True)

    if st.button("开始处理并导出（new.docx）", type="primary", disabled=not docx_file):
        if not st.session_state["ocr_model_id"].strip():
            st.error("请先在侧边栏配置 OCR model（用于 OCR/PDF，不影响翻译）。")
            st.stop()

        # 翻译不依赖侧边栏输入的 translate_model_id（会强制使用 DEFAULT_TRANSLATE_MODEL），
        # 但这里仍保留输入框，便于你看到当前配置。
        client = get_ark_client(default_timeout_s=int(st.session_state["translate_timeout_s"]))

        doc_bytes = docx_file.read()
        doc = Document(io.BytesIO(doc_bytes))

        # 0) 将文本中的 LaTeX 公式 ($...$ / $$...$$ / \(\) / \[\]) 转成 Word 原生公式（可选）
        if do_latex_to_omml:
            if not HAVE_LATEX_OMML:
                st.error("该功能需要额外依赖：lxml + latex2mathml（见 requirements 更新）。")
                st.stop()

            xsl_cfg = (st.session_state.get("mml2omml_xsl", "") or "").strip()
            xsl_env = (os.environ.get(DEFAULT_MML2OMML_XSL_ENV, "") or "").strip()
            xsl_path = xsl_cfg or xsl_env

            if not xsl_path:
                st.warning("未提供 MML2OMML.XSL 路径：已跳过 LaTeX→OMML（不影响翻译/公式导出）。")
            else:
                try:
                    converted = convert_inline_latex_to_omml_in_doc(
                        doc,
                        xsl_path=xsl_path,
                        log=lambda s: st.info(s),
                    )
                    st.success(f"LaTeX→OMML 转换完成：{converted} 处")
                except Exception as e:
                    st.warning(f"LaTeX→OMML 转换失败（已跳过，不影响后续）：{type(e).__name__}: {e}")

        # 1) 把 Word 原生公式（OMML）替换为 Pandoc 可渲染的 LaTeX（$...$ / $$...$$）
        if do_equation_replace:
            with st.spinner("提取公式序列（pandoc）..."):
                math_seq = extract_math_from_docx_with_pandoc(doc_bytes)
            with st.spinner("替换 Word 原生公式为 LaTeX 代码（best-effort）..."):
                replaced_count = replace_omml_with_latex_code(doc, math_seq, use_equation_env=False)
            st.info(f"已替换公式数量（best-effort）：{replaced_count}")

        # 2) 翻译（保持原排版：仅替换 runs 的文本；图片/表格结构不动）
        if do_translate:
            batch_bar = st.progress(0, text="准备翻译…")
            diag_box = st.empty()
            attempt_box = st.empty()

            def _batch_cb(cur: int, total: int):
                batch_bar.progress(int(cur / total * 100), text=f"翻译批次 {cur}/{total}")

            diag_payload_holder: Dict[str, Any] = {}

            def _diag_cb(payload: Dict[str, Any]):
                diag_payload_holder.update(payload)

            # 预扫描一次 items 用于“自动分批”估算
            all_ps = iter_all_paragraphs_extended(doc)
            items_preview: List[dict] = []
            pid = 0
            for p in all_ps:
                runs = [r for r in p.runs if r.text is not None and r.text != ""]
                if not runs:
                    continue
                segs = []
                for r in runs:
                    protected, _ = protect_math(r.text)
                    segs.append(protected)
                if not "".join(segs).strip():
                    continue
                pid += 1
                items_preview.append({"id": f"p{pid}", "segments": segs})

            if auto_batch:
                max_batch_chars_local = estimate_auto_batch_chars(items_preview, target_batches=int(target_batches))
                st.info(f"自动分批：items={len(items_preview)}，估算 max_batch_chars={max_batch_chars_local}")
            else:
                max_batch_chars_local = int(max_batch_chars)

            if show_translate_diagnostics:
                diag_box.info(
                    "诊断将在开始后显示：扫描段落数 / 可翻译 items 数 / runs 数 / 字符数。\n"
                    "若 items=0，多半是内容在文本框/形状里（python-docx 读不到）。"
                )

            with st.spinner("翻译中（分批提交，保持图片/公式位置不动）..."):
                translate_docx_in_place(
                    doc=doc,
                    client=client,
                    model=st.session_state.get("translate_model_id", "").strip(),
                    src_lang=src_lang,
                    dst_lang=dst_lang,
                    max_batch_chars=int(max_batch_chars_local),
                    timeout_s=int(st.session_state["translate_timeout_s"]),
                    progress_cb=_batch_cb,
                    diagnostic_cb=_diag_cb,
                    attempt_msg_cb=(lambda msg: attempt_box.info(msg)),
                )

            if show_translate_diagnostics and diag_payload_holder:
                diag_box.success(
                    f"翻译抓取诊断：\n"
                    f"- paragraphs_scanned = {diag_payload_holder.get('paragraphs_scanned')}\n"
                    f"- items_built       = {diag_payload_holder.get('items_built')}\n"
                    f"- runs_counted      = {diag_payload_holder.get('runs_counted')}\n"
                    f"- visible_chars     = {diag_payload_holder.get('visible_chars_counted')}\n"
                )

        out_docx_bytes = doc_to_bytes(doc)
        st.success("处理完成")
        st.download_button("下载 new.docx", data=out_docx_bytes, file_name="new.docx")



# ---------------------------
# Tab 3: DOCX -> LaTeX/Markdown export (recommended)
# ---------------------------
with tabs[2]:
    st.subheader("Word(.docx) → LaTeX / Markdown 直接导出（推荐：可编辑公式最稳）")
    st.write("这个模式不追求保留 Word 排版，而追求“学术 LaTeX 输出正确性”，尤其适合大量可编辑公式。")

    docx_file2 = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_export")

    out_format = st.selectbox("导出格式", ["latex (.tex)", "markdown (.md)"], index=0)
    wrap_none = st.checkbox("wrap=none（不自动换行）", value=True)

    # ✅ 新增：把 pandoc 默认的 \[...\]/\(...\) 统一成更“论文友好”的格式
    st.markdown("**公式输出格式（可选）**")
    col_m1, col_m2 = st.columns([1, 1])
    with col_m1:
        display_style = st.selectbox(
            "行间公式（display math）",
            ["equation（带编号）", "equation*（不编号）", "bracket（\\[...\\]）", "dollars（$$...$$）"],
            index=0,
            help="pandoc 默认常用 \\[...\\]。论文通常更喜欢 equation/equation*，便于统一编号与引用。",
        )
    with col_m2:
        inline_style = st.selectbox(
            "行内公式（inline math）",
            ["paren（\\(...\\)）", "dollars（$...$）"],
            index=0,
            help="pandoc 默认常用 \\(...\\)。如果你后续要在 Markdown 里编辑，可能更喜欢 $...$。",
        )

    # 映射到 normalize_pandoc_math 的参数值
    display_style_key = display_style.split("（", 1)[0].strip()
    inline_style_key = inline_style.split("（", 1)[0].strip()

    if st.button("导出", type="primary", disabled=not docx_file2):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            in_path = td / "in.docx"
            # ⚠️ 不要用 .read()（会消耗 UploadedFile 的指针，导致后续同一文件无法复用）
            in_path.write_bytes(docx_file2.getvalue())

            extra_args = []
            if wrap_none:
                extra_args += ["--wrap=none"]

            if out_format.startswith("latex"):
                # docx -> latex
                out_text = pypandoc.convert_file(
                    str(in_path),
                    to="latex",
                    format="docx",
                    extra_args=extra_args,
                )
                # 统一公式风格（\[\] / $$ $$ / \(\) / $ $）
                out_text = normalize_pandoc_math(
                    out_text,
                    display_style=display_style_key,
                    inline_style=("dollars" if inline_style_key.startswith("dollars") else "paren"),
                )

                st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="latex")
                st.download_button("下载 .tex", data=out_text.encode("utf-8"), file_name="export.tex")
            else:
                # docx -> markdown
                # 关键：显式打开 tex_math_dollars 扩展，保证 $...$ / $$...$$ 被当作数学处理
                # Pandoc 支持通过 format name +EXTENSION 来启用扩展（见官方手册）。
                out_text = pypandoc.convert_file(
                    str(in_path),
                    to="markdown+tex_math_dollars",
                    format="docx",
                    extra_args=extra_args,
                )

                # Markdown 输出也可按同一套规则统一（通常建议保留 dollars 方便后续编辑）
                out_text = normalize_pandoc_math(
                    out_text,
                    display_style=display_style_key,
                    inline_style=("dollars" if inline_style_key.startswith("dollars") else "paren"),
                )

                st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="markdown")
                st.download_button("下载 .md", data=out_text.encode("utf-8"), file_name="export.md")

    # ------------------------------------------------------------
    # ③-补充功能：把文本公式 $...$ / $$...$$ 转回可编辑 Word 公式（OMML）
    # ------------------------------------------------------------
    st.divider()
    st.subheader("把文档中的 .../... 文本公式转为可编辑 Word 公式（OMML）")
    st.caption(
        "适用场景：你在 Word 里手打了 $...$（或 $$...$$）占位公式，"
        "希望一键转成可编辑的原生公式对象（OMML）。实现方式：docx → pandoc markdown → docx（用原文档作为 reference-doc 尽量保留样式）。"
    )

    if st.button("生成：EditableEquations.docx", disabled=not docx_file2, key="btn_make_omml"):
        try:
            raw_bytes = docx_file2.getvalue()
            out_docx = docx_roundtrip_make_equations_editable(raw_bytes)
            st.success("已生成，可下载。")
            st.download_button(
                "下载 EditableEquations.docx",
                data=out_docx,
                file_name="EditableEquations.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_make_omml",
            )
        except Exception as e:
            st.error(f"生成失败：{type(e).__name__}: {e}")
st.markdown("---")
st.subheader("AI 增强：DOCX 文本 → 类似图片 OCR 的公式风格 → 再转 OMML / Pandoc")
st.caption("思路：先用 Pandoc 把 docx 里的公式变成 $...$/$$...$$ 文本，再交给模型做“只改公式不改正文”的校正，最后再用 Pandoc 写回 Word（OMML）。对少量顽固公式通常更稳。")

enable_ai = st.checkbox("开启 AI 公式校正（更慢但更准）", value=False, key="ai_math_clean_enable")
ai_model = st.text_input("AI 校正使用的 model（默认用 OCR model）", value=st.session_state.get("ocr_model_id",""), key="ai_math_clean_model")
max_batch_chars_ai = st.number_input("AI 每次输入最大字符数（自动分段）", min_value=2000, max_value=20000, value=8000, step=500, key="ai_math_clean_max_chars")
out_tokens_ai = st.number_input("AI 输出 tokens 上限", min_value=512, max_value=16000, value=4096, step=256, key="ai_math_clean_out_tokens")

docx_file4 = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_omml_ai_roundtrip")
col_ai1, col_ai2 = st.columns([1, 1])
with col_ai1:
    if st.button("生成：AI_EditableEquations.docx", key="btn_omml_ai_roundtrip", disabled=not docx_file4):
        try:
            if not enable_ai:
                # 如果没开 AI，就直接走原本的 roundtrip
                out_docx_bytes = docx_roundtrip_make_equations_editable(docx_file4.read())
                st.download_button(
                    "下载 AI_EditableEquations.docx",
                    data=out_docx_bytes,
                    file_name="AI_EditableEquations.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_omml_ai_roundtrip_noai",
                )
                st.success("已生成：AI_EditableEquations.docx（未开启 AI，走原 roundtrip）")
            else:
                client2 = OpenAI(
                    api_key=os.environ.get("ARK_API_KEY", "").strip() or st.secrets.get("ARK_API_KEY", ""),
                    base_url=(os.environ.get("ARK_BASE_URL") or st.secrets.get("ARK_BASE_URL","https://ark.cn-beijing.volces.com/api/v3")).strip(),
                )
                out_docx_bytes, cleaned_md = docx_ai_roundtrip_make_equations_editable(
                    docx_file4.read(),
                    client=client2,
                    model=(ai_model or st.session_state.get("ocr_model_id","")).strip(),
                    max_batch_chars=int(max_batch_chars_ai),
                    timeout_s=int(st.session_state["ocr_timeout_s"]),
                    out_tokens=int(out_tokens_ai),
                )
                st.download_button(
                    "下载 AI_EditableEquations.docx",
                    data=out_docx_bytes,
                    file_name="AI_EditableEquations.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="dl_omml_ai_roundtrip",
                )
                with st.expander("查看 AI 校正后的 Markdown（可用于排查）", expanded=False):
                    st.code(cleaned_md[:8000] + ("\n...\n" if len(cleaned_md) > 8000 else ""), language="markdown")
                st.success("已生成：AI_EditableEquations.docx（AI 校正 + OMML）")
        except Exception as e:
            st.error(f"生成失败：{type(e).__name__}: {e}")

with col_ai2:
    st.write("提示：如果你只想拿到“图片 OCR 风格”的 Markdown，可以先导出 markdown（上面导出功能），再手动粘贴进这里的 AI 校正。")
