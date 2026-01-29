# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
import re
import json
import time
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import streamlit as st
from PIL import Image

# Gemini
from google import genai
from google.genai import types
from google.genai import errors as genai_errors

# DOCX
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import _Cell, Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Pandoc (for extracting equations from DOCX robustly)
import pypandoc


# =========================
# 0) Basic settings
# =========================
st.set_page_config(page_title="OCR + LaTeX(equation) + DOCX Translate", layout="wide")
st.title("Gemini OCR → Word（含第三版 LaTeX equation 代码） + 上传 Word 可选翻译/公式转 LaTeX 代码")

# ---- IMPORTANT: ensure pandoc exists ----
def ensure_pandoc():
    try:
        _ = pypandoc.get_pandoc_path()
    except OSError:
        # In some cloud env this may fail; if so, pypandoc-binary usually already provides it.
        # But keep the fallback.
        try:
            pypandoc.download_pandoc()
        except Exception:
            pass

ensure_pandoc()

def xpath_omml(el, names: list[str]):
    cond = " or ".join([f"local-name()='{n}'" for n in names])
    return el.xpath(f".//*[{cond}]")


# =========================
# 1) Prompts
# =========================
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



# =========================
# 2) Gemini client and safe call
# =========================
@dataclass
class GeminiResult:
    text: str = ""
    status_code: Optional[int] = None
    error_message: Optional[str] = None


def get_api_key() -> str:
    k = None
    if "GEMINI_API_KEY" in st.secrets:
        k = st.secrets["GEMINI_API_KEY"]
    k = k or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not k:
        st.error("缺少 GEMINI_API_KEY。请在 Streamlit Cloud → Secrets 中配置。")
        st.stop()
    return k


def get_gemini_client() -> genai.Client:
    # Force developer API
    return genai.Client(api_key=get_api_key(), http_options=types.HttpOptions(api_version="v1beta"))


def _parse_retry_delay_seconds(msg: str) -> Optional[float]:
    # Examples:
    # "Please retry in 37.161915127s."
    m = re.search(r"retry in ([0-9.]+)s", msg)
    if m:
        return float(m.group(1))
    # Sometimes includes JSON-like '"retryDelay":"37s"'
    m = re.search(r'"retryDelay"\s*:\s*"(\d+)s"', msg)
    if m:
        return float(m.group(1))
    return None

def _is_free_tier_daily_quota(msg: str) -> bool:
    # daily quota strings often include:
    # GenerateRequestsPerDayPerProjectPerModel-FreeTier
    return ("PerDayPerProjectPerModel-FreeTier" in msg
            or "generate_content_free_tier_requests" in msg)

def safe_generate(
    client: genai.Client,
    model: str,
    contents,
    config: Optional[types.GenerateContentConfig] = None,
    retries: int = 6,
) -> GeminiResult:
    last_err = None
    for attempt in range(retries):
        try:
            resp = client.models.generate_content(model=model, contents=contents, config=config)
            return GeminiResult(text=(resp.text or "").strip())
        except genai_errors.ClientError as e:
            msg = str(e)
            code = getattr(e, "status_code", None)

            # 429: quota / rate limit
            if code == 429 or "RESOURCE_EXHAUSTED" in msg or "Quota exceeded" in msg:
                # 如果是 free tier 的“每日配额已用完”，等多久都没用
                if _is_free_tier_daily_quota(msg):
                    return GeminiResult(text="", status_code=429, error_message=msg)

                # 否则按 retryDelay 等待并重试
                wait_s = _parse_retry_delay_seconds(msg)
                if wait_s is None:
                    wait_s = min(2 ** attempt, 60)
                time.sleep(min(wait_s + 0.3, 90.0))
                last_err = (code, msg)
                continue

            return GeminiResult(text="", status_code=code, error_message=msg)

        except genai_errors.ServerError as e:
            last_err = ("5xx", str(e))
            time.sleep(min(2 ** attempt, 30))
            continue
        except Exception as e:
            last_err = ("unknown", f"{type(e).__name__}: {e}")
            time.sleep(min(2 ** attempt, 30))
            continue

    return GeminiResult(text="", status_code=None, error_message=f"retry exhausted: {last_err}")



# =========================
# 3) OCR helpers (image → markdown)
# =========================
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
            # remove first n non-empty lines in cur (raw)
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


def ocr_image_to_markdown(client: genai.Client, model: str, img: Image.Image, max_side: int, tile_h: int, overlap: int, jpeg_q: int, out_tokens: int) -> str:
    img = downscale(img, max_side=max_side)
    tiles = slice_long(img, tile_h=tile_h, overlap=overlap)

    chunks: List[str] = []
    for i, timg in enumerate(tiles, start=1):
        b = pil_to_jpeg_bytes(timg, quality=jpeg_q)

        # If > 18MB use Files API (keeps under 20MB inline limit)
        if len(b) > 18_000_000:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as f:
                f.write(b)
                tmp = f.name
            up = client.files.upload(file=tmp)
            os.unlink(tmp)
            contents = [up, OCR_PROMPT_ZH]
        else:
            part = types.Part.from_bytes(data=b, mime_type="image/jpeg")
            contents = [part, OCR_PROMPT_ZH]

        cfg = types.GenerateContentConfig(temperature=1.0, max_output_tokens=out_tokens)
        res = safe_generate(client, model, contents, cfg)
        if res.error_message:
            raise RuntimeError(f"Gemini OCR error: status={res.status_code} msg={res.error_message}")

        md = normalize_md(res.text)
        if chunks:
            md = dedupe_tail_head(chunks[-1], md)
        chunks.append(md)

    return normalize_md("\n\n".join([c for c in chunks if c.strip()]))


# =========================
# 4) LaTeX code format: use \begin{equation}...\end{equation}
# =========================
DISPLAY_MATH_RE = re.compile(r"(?s)\$\$(.+?)\$\$")
INLINE_MATH_RE = re.compile(r"(?<!\$)\$([^$\n]+)\$(?!\$)")

def display_to_equation_env(md: str) -> str:
    """
    Convert $$ ... $$ blocks into \begin{equation} ... \end{equation}
    Keep \tag if present inside.
    """
    def repl(m):
        body = m.group(1).strip()
        return "\\begin{equation}\n" + body + "\n\\end{equation}"
    return DISPLAY_MATH_RE.sub(repl, md)

def inline_to_paren(md: str) -> str:
    """
    Convert inline $...$ into \( ... \) for code style.
    """
    def repl(m):
        body = m.group(1).strip()
        return "\\(" + body + "\\)"
    return INLINE_MATH_RE.sub(repl, md)

def md_to_latex_code_style(md: str) -> str:
    """
    Third-version code-style output:
      - display -> equation env
      - inline  -> \( ... \)
      - wrap equation env blocks into ```latex code fences for "代码格式"
    """
    md2 = normalize_md(md)
    md2 = display_to_equation_env(md2)
    md2 = inline_to_paren(md2)

    # Wrap equation env blocks into fenced code blocks so Word shows "code"
    # We'll wrap each equation environment occurrence.
    eq_env = re.compile(r"(?s)(\\begin\{equation\}.*?\\end\{equation\})")
    md2 = eq_env.sub(lambda m: "\n```latex\n" + m.group(1).strip() + "\n```\n", md2)
    return normalize_md(md2)


# =========================
# 5) DOCX translate (preserve layout) + optional equation replacement
# =========================
MATH_TOKEN_RE = re.compile(r"(\$\$.*?\$\$|\$[^$\n]+\$|\\begin\{equation\}.*?\\end\{equation\}|\\\(.+?\\\))", re.DOTALL)

def protect_math(text: str) -> Tuple[str, Dict[str, str]]:
    """
    Replace math fragments with tokens __MATH_k__ so translator won't modify them.
    """
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

def iter_all_paragraphs(doc: Document) -> List[Paragraph]:
    ps: List[Paragraph] = []
    for p in doc.paragraphs:
        ps.append(p)
    for table in doc.tables:
        ps.extend(iter_table_paragraphs(table))
    return ps

def iter_table_paragraphs(table: Table) -> List[Paragraph]:
    out: List[Paragraph] = []
    for row in table.rows:
        for cell in row.cells:
            out.extend(cell.paragraphs)
            for t2 in cell.tables:
                out.extend(iter_table_paragraphs(t2))
    return out

# ---- Extract OMML count and replace with LaTeX code (best-effort) ----
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"m": M_NS, "w": W_NS}

def extract_math_from_docx_with_pandoc(docx_bytes: bytes) -> List[Tuple[str, str]]:
    """
    Use pandoc: docx -> markdown, then scan math in order.
    Return list of ("display"|"inline", latex_body_with_delims_removed).
    """
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        f = td / "in.docx"
        f.write_bytes(docx_bytes)

        md = pypandoc.convert_file(str(f), to="markdown", format="docx", extra_args=["--wrap=none"])
        md = md.replace("\r\n", "\n").replace("\r", "\n")

    # scan in order (both $$ and $)
    matches: List[Tuple[int, int, str, str]] = []  # (start,end,kind,body)
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
    """
    Replace OMML math nodes with LaTeX code text in-place.
    math_seq order is used to map. Returns replaced count.
    """
    # Collect all paragraph elements including in tables
    all_ps = iter_all_paragraphs(doc)
    seq_idx = 0
    replaced = 0

    for p in all_ps:
        p_elm = p._p  # lxml element
        # math nodes can be m:oMath or m:oMathPara
        nodes = p_elm.xpath(".//*[local-name()='oMath' or local-name()='oMathPara']")
        # We will replace each node in the paragraph in document order.
        for node in nodes:
            if seq_idx >= len(math_seq):
                return replaced

            kind, body = math_seq[seq_idx]
            seq_idx += 1

            if use_equation_env:
                if kind == "display":
                    latex_text = "\\begin{equation} " + body + " \\end{equation}"
                else:
                    latex_text = "\\(" + body + "\\)"
            else:
                # fallback
                latex_text = "$$ " + body + " $$" if kind == "display" else "$ " + body + " $"

            # Insert a run at node position then remove node
            parent = node.getparent()
            if parent is None:
                continue
            try:
                idx_in_parent = list(parent).index(node)
            except Exception:
                idx_in_parent = None

            # If math node is deep, best effort: insert into paragraph root
            if parent is not p_elm or idx_in_parent is None:
                # insert at end of paragraph
                insert_text_run_at_paragraph_child(p_elm, len(p_elm), latex_text)
            else:
                insert_text_run_at_paragraph_child(parent, idx_in_parent, latex_text)

            # remove original math node
            try:
                parent.remove(node)
            except Exception:
                pass

            replaced += 1

    return replaced


# ---- Translation batching (JSON in/out) ----
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
    """
    Robust JSON extraction:
    1) If model returns ```json ... ```, strip fences
    2) Try direct json.loads
    3) Fallback: locate the first JSON object and parse with bracket matching
    """
    if not text:
        raise ValueError("empty response")

    s = text.strip()

    # 1) Strip fenced code blocks
    # e.g. ```json\n{...}\n```
    fence = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", s, flags=re.DOTALL | re.IGNORECASE)
    if fence:
        s = fence.group(1).strip()

    # 2) Try direct parse first
    try:
        return json.loads(s)
    except Exception:
        pass

    # 3) Bracket-matching to find first valid JSON object
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
                    candidate = s[start:i+1]
                    try:
                        return json.loads(candidate)
                    except Exception as e:
                        raise ValueError(f"json parse failed after bracket match: {e}\nCandidate:\n{candidate[:800]}")

    raise ValueError("no complete JSON object found (unbalanced braces)")


def gemini_translate_items(
    client: genai.Client,
    model: str,
    items: List[dict],
    src_lang: str,
    dst_lang: str,
) -> Dict[str, List[str]]:
    prompt = (TRANSLATE_PROMPT_TEMPLATE
              .replace("__SRC_LANG__", src_lang)
              .replace("__DST_LANG__", dst_lang))

    payload = {"items": items}

    cfg = types.GenerateContentConfig(
    temperature=0.2,
    max_output_tokens=8192,
    response_mime_type="application/json",
)

    res = safe_generate(client, model, [prompt, json.dumps(payload, ensure_ascii=False)], cfg)

    if res.error_message:
        st.error(f"Gemini translate error\nstatus={res.status_code}\n\n{res.error_message}")

        if res.status_code == 429 and ("PerDayPerProjectPerModel-FreeTier" in res.error_message
                                       or "generate_content_free_tier_requests" in res.error_message):
            st.warning(
                "这是 Free Tier 的“每日请求数”配额用完了（不是临时限流）。\n"
                "解决：1) 换模型（每个模型单独计数） 2) 换项目/API Key 3) 开通 Billing 4) 等到配额刷新。"
            )
        elif res.status_code == 429:
            st.info("这是临时限流/配额，程序会自动等待重试；若仍失败，请降低切片数量或增大翻译批次以减少请求次数。")

        try:
    obj = extract_json_object(res.text)
except Exception as e:
    st.error(f"JSON parse failed: {e}")
    st.markdown("#### Raw model output (first 2000 chars)")
    st.code((res.text or "")[:2000], language="text")
    st.stop()


    obj = extract_json_object(res.text)
    out: Dict[str, List[str]] = {}
    for it in obj.get("items", []):
        out[it["id"]] = it["segments"]
    return out


def translate_docx_in_place(
    doc: Document,
    client: genai.Client,
    model: str,
    src_lang: str,
    dst_lang: str,
    max_batch_chars: int = 12000,
):
    """
    Translate all paragraph runs and table-cell runs while preserving formatting.
    Strategy:
      - For each paragraph, collect run texts (only runs with text).
      - Protect math tokens.
      - Batch multiple paragraphs to reduce API calls.
      - Apply translated segments back to same runs (keeps bold/italic/fonts).
    """
    all_ps = iter_all_paragraphs(doc)

    # Build items
    items = []
    para_refs = {}  # id -> (runs, per-run math_maps)
    pid = 0

    for p in all_ps:
        runs = [r for r in p.runs if r.text is not None and r.text != ""]
        if not runs:
            continue

        segs = []
        maps = []
        for r in runs:
            protected, mp = protect_math(r.text)
            segs.append(protected)
            maps.append(mp)

        # skip if all segments are just tokens/whitespace
        joined = "".join(segs).strip()
        if not joined:
            continue

        pid += 1
        item_id = f"p{pid}"
        items.append({"id": item_id, "segments": segs})
        para_refs[item_id] = (runs, maps)

    batches = chunk_items_for_api(items, max_chars=max_batch_chars)

    # Translate in batches
    for bi, batch in enumerate(batches, start=1):
        translated_map = gemini_translate_items(client, model, batch, src_lang=src_lang, dst_lang=dst_lang)

        # Apply back
        for it in batch:
            item_id = it["id"]
            if item_id not in translated_map:
                continue
            runs, maps = para_refs[item_id]
            out_segs = translated_map[item_id]

            # Ensure same length
            if len(out_segs) != len(runs):
                # fallback: if mismatch, translate whole paragraph as one string into first run, clear others
                whole = " ".join(out_segs)
                whole = restore_tokens(whole, {k:v for mp in maps for k,v in mp.items()})
                runs[0].text = whole
                for r in runs[1:]:
                    r.text = ""
                continue

            for r, seg, mp in zip(runs, out_segs, maps):
                r.text = restore_tokens(seg, mp)


# =========================
# 6) Export helpers
# =========================
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


# =========================
# 7) UI
# =========================
tabs = st.tabs(["① 图片 OCR → 导出", "② 上传 Word(.docx) → LaTeX 代码/可选翻译"])

with st.sidebar:
    st.subheader("通用设置")
    model_id = st.selectbox("Gemini model", ["gemini-3-flash-preview", "gemini-2.5-pro", "gemini-2.5-flash"], index=0)
    st.divider()
    st.subheader("OCR 设置")
    max_side = st.slider("Max side", 800, 3200, 2000, 100)
    tile_h = st.slider("Tile height", 800, 2600, 1600, 100)
    overlap = st.slider("Overlap", 0, 400, 160, 10)
    jpeg_q = st.slider("JPEG quality", 50, 95, 85, 1)
    out_tokens = st.slider("OCR max tokens", 1024, 8192, 4096, 256)

with tabs[0]:
    st.subheader("① 图片 OCR（多图 + 长图切片）")
    imgs = st.file_uploader("上传图片（可多选）", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

    if st.button("开始 OCR 并导出", type="primary", disabled=not imgs):
        client = get_gemini_client()

        pages = []
        for i, f in enumerate(imgs, start=1):
            img = Image.open(f)
            md = ocr_image_to_markdown(client, model_id, img, max_side, tile_h, overlap, jpeg_q, out_tokens)
            pages.append(f"## 第 {i} 页\n\n{md}")

        merged_md = normalize_md("\n\n---\n\n".join(pages))

        st.success("OCR 完成")
        st.markdown("### 预览（Markdown）")
        st.code(merged_md, language="markdown")

        # ---- Export 3 versions ----
        # V1: Rendered docx (pandoc will render $$...$$ to Word equations)
        v1_docx = pandoc_md_to_docx(merged_md)

        # V2: Raw markdown
        v2_md_bytes = merged_md.encode("utf-8")

        # V3: LaTeX code style: display -> equation env + fenced code blocks
        v3_md = md_to_latex_code_style(merged_md)
        v3_docx = pandoc_md_to_docx(v3_md)

        st.download_button("下载 V1：Rendered.docx（可编辑公式）", data=v1_docx, file_name="OCR_Rendered.docx")
        st.download_button("下载 V2：Result.md（原始 Markdown）", data=v2_md_bytes, file_name="OCR_Result.md")
        st.download_button("下载 V3：LaTeX_equation_code.docx（公式为 equation 代码格式）", data=v3_docx,
                           file_name="OCR_LaTeX_equation_code.docx")
        st.download_button("下载 V3：LaTeX_equation_code.md", data=v3_md.encode("utf-8"),
                           file_name="OCR_LaTeX_equation_code.md")

    with tabs[1]:
        st.subheader("② 上传 Word(.docx) → LaTeX 代码 / 可选翻译（保持原排版）")

        docx_file = st.file_uploader("上传 Word 文档（.docx）", type=["docx"])

        colA, colB = st.columns([1, 1])
        with colA:
            do_equation_replace = st.checkbox("把 Word 原生公式（OMML）替换为 LaTeX 代码（equation/\\(\\)）", value=True)
            do_translate = st.checkbox("翻译文档（保留图片/公式位置不动，仅替换文本内容）", value=False)

        with colB:
            src_lang = st.selectbox("源语言", ["Auto", "Chinese", "English", "Japanese", "Korean", "Spanish"], index=0)
            dst_lang = st.selectbox("目标语言", ["Chinese", "English", "Japanese", "Korean", "Spanish"], index=1)
            max_batch_chars = st.slider("翻译分批大小（文档很长就调小）", 4000, 20000, 12000, 500)

        if st.button("开始处理并导出", type="primary", disabled=not docx_file):
            client = get_gemini_client()

            # Load docx
            doc_bytes = docx_file.read()
            doc = Document(io.BytesIO(doc_bytes))

            # (A) replace OMML equations with LaTeX code text
            replaced_count = 0
            if do_equation_replace:
                with st.spinner("提取公式序列（pandoc）..."):
                    math_seq = extract_math_from_docx_with_pandoc(doc_bytes)

                with st.spinner("替换 Word 原生公式为 LaTeX 代码..."):
                    replaced_count = replace_omml_with_latex_code(doc, math_seq, use_equation_env=True)

                st.info(f"已替换公式数量（best-effort）：{replaced_count}")

            # (B) translate text in-place (preserve layout by keeping runs)
            if do_translate:
                with st.spinner("翻译中（分批提交，保持图片/公式位置不动）..."):
                    # If user chooses Auto, pass "Auto" as src_lang; prompt will still work
                    translate_docx_in_place(
                        doc=doc,
                        client=client,
                        model=model_id,
                        src_lang=src_lang,
                        dst_lang=dst_lang,
                        max_batch_chars=max_batch_chars,
                    )

            # Output docx
            out_docx_bytes = doc_to_bytes(doc)

            st.success("处理完成")
            st.download_button("下载 new.docx", data=out_docx_bytes, file_name="new.docx")

            # Optional: also export an intermediate md (useful for debugging)
            with st.expander("导出调试用 Markdown（可选）", expanded=False):
                try:
                    with tempfile.TemporaryDirectory() as td:
                        td = Path(td)
                        p = td / "out.docx"
                        p.write_bytes(out_docx_bytes)
                        md_dbg = pypandoc.convert_file(str(p), to="markdown", format="docx", extra_args=["--wrap=none"])
                    md_dbg = normalize_md(md_dbg)
                    st.code(md_dbg, language="markdown")
                    st.download_button("下载 debug.md", data=md_dbg.encode("utf-8"), file_name="debug.md")
                except Exception as e:
                    st.warning(f"导出 Markdown 失败：{e}")
