# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
import re
import json
import time
import base64
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Callable

import streamlit as st
from PIL import Image, ImageFilter, ImageOps

from openai import OpenAI

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import pypandoc

# PDF rendering (optional)
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
st.caption("面向论文/讲义/教材：PDF/图片OCR→Markdown/Word；Word→翻译（best-effort + 强校验）；Word→LaTeX/Markdown。")


def ensure_pandoc() -> bool:
    """尽量保证 pandoc 可用；不可用则不阻塞启动。"""
    try:
        _ = pypandoc.get_pandoc_path()
        return True
    except OSError:
        try:
            pypandoc.download_pandoc()
            _ = pypandoc.get_pandoc_path()
            return True
        except Exception:
            return False

PANDOC_OK = ensure_pandoc()
if not PANDOC_OK:
    st.warning("Pandoc 不可用：Tab③ 导出 / Tab① Markdown→docx 渲染可能受影响。")


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

严格要求（必须遵守）：
- 输入是 JSON，包含 items 列表，每个 item 有 id 和 segments（字符串列表）。
- 输出必须是 JSON，结构必须为：{"items":[{"id":"...","segments":[...]}, ...]}
- 输出 JSON 必须可被标准 json.loads 解析。
- segments 的数量必须与输入完全一致；每个 segments[i] 对应翻译输入 segments[i]。
- 保留所有占位符不变：例如 __MATH_0__、__KEEP_12__、{{ }} 这种标记必须原样输出，不能翻译、不能改大小写、不能删。
- LaTeX 代码、公式环境（如 \begin{equation}...\end{equation} 或 $...$）不得改动。
- 只输出 JSON，不要输出解释，不要输出 Markdown，不要输出多余文本。
""".strip()

# 更强的二次重试提示：明确“禁止保留原文”
TRANSLATE_PROMPT_STRONG_TEMPLATE = r"""
你是一个严格的学术翻译引擎。你必须把文本翻译成 __DST_LANG__，禁止保留原文语言（除非是专有名词/缩写）。

输入是 JSON，输出必须是 JSON：
{"items":[{"id":"...","segments":[...]}, ...]}

硬性要求：
- 输出必须可被 json.loads 解析，且只能输出一个 JSON 对象。
- segments 数量必须与输入完全一致。
- 占位符（__MATH_0__、__KEEP_12__、{{}}）必须原样保留。
- LaTeX/公式不得改动。
- 若输入是中文，输出必须明显是 __DST_LANG__（例如英文），不能原样照抄。
""".strip()


# ============================================================
# 2) Ark config / client / retry / timeout
# ============================================================
@dataclass
class ArkResult:
    text: str = ""
    error_message: Optional[str] = None

DEFAULT_ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_ARK_MODEL = "ep-20260203141749-992fx"
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
) -> ArkResult:
    last_err: Any = None
    for attempt in range(retries):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=messages,
                temperature=temperature,
                max_tokens=max_tokens,
                timeout=timeout_s,
            )
            txt = ""
            if resp and resp.choices and resp.choices[0].message and resp.choices[0].message.content:
                txt = resp.choices[0].message.content
            return ArkResult(text=(txt or "").strip())
        except Exception as e:
            msg = f"{type(e).__name__}: {e}"
            last_err = msg

            if ("429" in msg) or ("rate limit" in msg.lower()) or ("RESOURCE_EXHAUSTED" in msg):
                wait_s = _parse_retry_delay_seconds(msg) or min(2 ** attempt, 60)
                time.sleep(min(wait_s + 0.3, 90.0))
                continue

            time.sleep(min(2 ** attempt, 30))

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
# 5) LaTeX style helpers
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
# 6) DOCX translate
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
    ps: List[Paragraph] = []
    ps.extend(iter_part_paragraphs(doc))
    try:
        for sec in doc.sections:
            ps.extend(iter_part_paragraphs(sec.header))
            ps.extend(iter_part_paragraphs(sec.footer))
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

def _schema_validate_translation(input_items: List[dict], obj: dict) -> Dict[str, List[str]]:
    """严格校验输出 JSON 是否满足 schema，并返回 id->segments。"""
    if not isinstance(obj, dict):
        raise ValueError("output is not a JSON object")
    items_out = obj.get("items", None)
    if not isinstance(items_out, list):
        raise ValueError("JSON schema error: items must be a list")

    input_map = {it["id"]: it["segments"] for it in input_items}
    out: Dict[str, List[str]] = {}

    for it in items_out:
        if not isinstance(it, dict):
            raise ValueError("items element is not an object")
        if "id" not in it or "segments" not in it:
            raise ValueError("items element missing id/segments")
        _id = it["id"]
        segs = it["segments"]
        if _id not in input_map:
            # 允许模型输出多余 id 也行，但我们只处理输入 id
            continue
        if not isinstance(segs, list) or not all(isinstance(x, str) for x in segs):
            raise ValueError("segments must be a list of strings")
        if len(segs) != len(input_map[_id]):
            raise ValueError(f"segments length mismatch for {_id}: {len(segs)} != {len(input_map[_id])}")
        out[_id] = segs

    # 确保每个输入 id 都有输出
    missing = [k for k in input_map.keys() if k not in out]
    if missing:
        raise ValueError(f"missing translated items: {missing[:5]} ... total {len(missing)}")

    return out

def _looks_untranslated(orig: str, trans: str) -> bool:
    """粗略判断：译文是否几乎等于原文（忽略空白）。"""
    o = re.sub(r"\s+", "", orig or "")
    t = re.sub(r"\s+", "", trans or "")
    if not o:
        return False
    if o == t:
        return True
    # 过高重合也判为可疑（保守）
    if len(o) > 40:
        # overlap ratio by common prefix length
        common = 0
        for a, b in zip(o, t):
            if a == b:
                common += 1
            else:
                break
        if common / max(1, len(o)) > 0.6:
            return True
    return False

def doubao_translate_items(
    client: OpenAI,
    model: str,
    items: List[dict],
    src_lang: str,
    dst_lang: str,
    timeout_s: int,
    debug_raw_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, List[str]]:
    src = "auto-detect" if src_lang == "Auto" else src_lang
    prompt = TRANSLATE_PROMPT_TEMPLATE.replace("__SRC_LANG__", src).replace("__DST_LANG__", dst_lang)

    payload = {"items": items}
    messages = [{"role": "user", "content": prompt + "\n\n" + json.dumps(payload, ensure_ascii=False)}]

    # 翻译建议 temperature=0 更稳
    res = safe_chat_completions(
        client=client,
        model=model,
        messages=messages,
        temperature=0.0,
        max_tokens=8192,
        timeout_s=timeout_s,
        retries=6,
    )
    if res.error_message:
        st.error(f"翻译调用失败：\n\n{res.error_message}")
        st.stop()

    raw1 = res.text
    if debug_raw_cb:
        debug_raw_cb(raw1)

    # 解析 + schema 校验
    try:
        obj = extract_json_object(raw1)
        out = _schema_validate_translation(items, obj)
    except Exception as e:
        # 失败就强提示重试一次
        strong_prompt = TRANSLATE_PROMPT_STRONG_TEMPLATE.replace("__DST_LANG__", dst_lang)
        messages2 = [{"role": "user", "content": strong_prompt + "\n\n" + json.dumps(payload, ensure_ascii=False)}]
        res2 = safe_chat_completions(
            client=client,
            model=model,
            messages=messages2,
            temperature=0.0,
            max_tokens=8192,
            timeout_s=timeout_s,
            retries=3,
        )
        if res2.error_message:
            st.error(f"翻译重试失败：\n\n{res2.error_message}")
            st.stop()
        raw2 = res2.text
        if debug_raw_cb:
            debug_raw_cb("\n\n--- RETRY ---\n\n" + raw2)

        obj2 = extract_json_object(raw2)
        out = _schema_validate_translation(items, obj2)

    # “翻译有效性”检测：如果大量段落完全没变化，则再强制重试一次
    suspicious = 0
    for it in items[: min(20, len(items))]:
        _id = it["id"]
        orig = it["segments"][0] if it.get("segments") else ""
        trans = out[_id][0] if _id in out else ""
        if _looks_untranslated(orig, trans):
            suspicious += 1
    if len(items) >= 5 and suspicious >= max(3, len(items[:20]) // 2):
        # 大概率模型在“照抄”，再强制重试（最强约束）
        strong_prompt = TRANSLATE_PROMPT_STRONG_TEMPLATE.replace("__DST_LANG__", dst_lang)
        messages3 = [{"role": "user", "content": strong_prompt + "\n\n" + json.dumps(payload, ensure_ascii=False)}]
        res3 = safe_chat_completions(
            client=client,
            model=model,
            messages=messages3,
            temperature=0.0,
            max_tokens=8192,
            timeout_s=timeout_s,
            retries=3,
        )
        if not res3.error_message and res3.text:
            try:
                obj3 = extract_json_object(res3.text)
                out3 = _schema_validate_translation(items, obj3)
                out = out3
            except Exception:
                pass

    return out

# Textbox support
_W_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

def iter_textbox_paragraph_elements(doc: Document):
    try:
        root = doc.part.element
        return root.xpath(".//w:txbxContent//w:p", namespaces=_W_NS)
    except Exception:
        return []

def get_wt_text(p_elm) -> str:
    ts = p_elm.xpath(".//w:t", namespaces=_W_NS)
    return "".join([(t.text or "") for t in ts])

def set_wt_text(p_elm, text: str):
    ts = p_elm.xpath(".//w:t", namespaces=_W_NS)
    if not ts:
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        r.append(t)
        p_elm.append(r)
        return
    ts[0].text = text
    for t in ts[1:]:
        t.text = ""

def build_translate_items_paragraph_level(doc: Document):
    items = []
    writers = {}
    pid = 0

    for p in iter_all_paragraphs_extended(doc):
        full = "".join([r.text or "" for r in p.runs])
        if not full or not full.strip():
            continue
        protected, mp = protect_math(full)
        if not protected.strip():
            continue
        pid += 1
        item_id = f"p{pid}"
        items.append({"id": item_id, "segments": [protected]})

        def _make_writer(paragraph: Paragraph, mapping: Dict[str, str]):
            def _w(out_text: str):
                out_text = restore_tokens(out_text, mapping)
                if paragraph.runs:
                    paragraph.runs[0].text = out_text
                    for r in paragraph.runs[1:]:
                        r.text = ""
                else:
                    paragraph.add_run(out_text)
            return _w
        writers[item_id] = _make_writer(p, mp)

    for p_elm in iter_textbox_paragraph_elements(doc):
        full = get_wt_text(p_elm)
        if not full or not full.strip():
            continue
        protected, mp = protect_math(full)
        if not protected.strip():
            continue
        pid += 1
        item_id = f"tb{pid}"
        items.append({"id": item_id, "segments": [protected]})

        def _make_writer_xml(par_elm, mapping: Dict[str, str]):
            def _w(out_text: str):
                out_text = restore_tokens(out_text, mapping)
                set_wt_text(par_elm, out_text)
            return _w
        writers[item_id] = _make_writer_xml(p_elm, mp)

    return items, writers

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
    sample_cb: Optional[Callable[[List[Tuple[str, str]]], None]] = None,
    raw_cb: Optional[Callable[[str], None]] = None,
):
    items, writers = build_translate_items_paragraph_level(doc)

    total_chars = sum(len(it["segments"][0]) for it in items if it.get("segments"))
    contains_tb = any(it["id"].startswith("tb") for it in items)

    if diagnostic_cb is not None:
        diagnostic_cb({
            "items_built": len(items),
            "estimated_chars": total_chars,
            "contains_textbox_items": contains_tb,
        })

    if not items:
        st.error("没有抓到任何可翻译文本（items=0）。扫描件请用 Tab① OCR。")
        return

    batches = chunk_items_for_api(items, max_chars=max_batch_chars)
    total = len(batches)

    # 收集样本对照
    before_after: List[Tuple[str, str]] = []

    for bi, batch in enumerate(batches, start=1):
        if progress_cb:
            progress_cb(bi, total)

        translated_map = doubao_translate_items(
            client=client,
            model=model,
            items=batch,
            src_lang=src_lang,
            dst_lang=dst_lang,
            timeout_s=timeout_s,
            debug_raw_cb=raw_cb,
        )

        for it in batch:
            item_id = it["id"]
            segs = translated_map.get(item_id)
            if not segs:
                continue
            out_text = segs[0]
            w = writers.get(item_id)
            if w:
                # 样本：只记录前几条
                if len(before_after) < 5:
                    before_after.append((it["segments"][0], out_text))
                w(out_text)

    if sample_cb is not None and before_after:
        sample_cb(before_after)


# ============================================================
# 7) Export helpers
# ============================================================
def doc_to_bytes(doc: Document) -> bytes:
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

def pandoc_md_to_docx(md: str) -> bytes:
    if not PANDOC_OK:
        raise RuntimeError("Pandoc 不可用，无法导出 docx。")
    md = normalize_md(md) + "\n"
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.docx"
        fmt = "markdown+fenced_code_blocks+tex_math_dollars"
        pypandoc.convert_text(md, to="docx", format=fmt, outputfile=str(out))
        return out.read_bytes()


# ============================================================
# 8) UI: sidebar
# ============================================================
ocr_timeout_default, translate_timeout_default = get_timeout_defaults()

def _init_state():
    defaults = {
        "model_id": get_default_model(),
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

if "__pending_rec__" in st.session_state:
    rec = st.session_state.pop("__pending_rec__")
    for k in ["max_side", "tile_h", "overlap", "jpeg_q", "out_tokens"]:
        if k in rec:
            st.session_state[k] = int(rec[k])
    st.session_state["__pending_rec_msg__"] = f"已回填推荐参数：{rec}"

with st.sidebar:
    st.subheader("Ark 配置")
    st.text_input("model（默认 EP 已填）", key="model_id")
    st.caption(f"Base URL: {get_ark_base_url()}")

    st.divider()
    st.subheader("OCR 参数与自动推荐")

    st.selectbox("预设", ["Fast", "Balanced", "Accurate"], key="preset_mode")

    st.slider("Max side", 800, 3200, key="max_side", step=100)
    st.slider("Tile height", 800, 2600, key="tile_h", step=100)
    st.slider("Overlap", 0, 400, key="overlap", step=10)
    st.slider("JPEG quality", 50, 95, key="jpeg_q", step=1)
    st.slider("OCR max tokens", 1024, 8192, key="out_tokens", step=256)

    st.number_input("OCR 超时（秒）", min_value=30, max_value=600, key="ocr_timeout_s", step=10)
    st.number_input("翻译超时（秒）", min_value=30, max_value=900, key="translate_timeout_s", step=10)


# ============================================================
# 9) Tabs
# ============================================================
tabs = st.tabs([
    "① PDF/图片 OCR → 导出",
    "② Word(.docx) → 翻译/公式替换（强校验）",
    "③ Word(.docx) → LaTeX/Markdown 导出（推荐）",
])


# ---------------------------
# Tab 1: OCR
# ---------------------------
with tabs[0]:
    st.subheader("PDF/图片 OCR（带自动参数推荐 + 进度条）")
    if "__pending_rec_msg__" in st.session_state:
        st.success(st.session_state.pop("__pending_rec_msg__"))

    files = st.file_uploader(
        "上传 PDF 或图片（可多选，PDF 将按页渲染）",
        type=["pdf", "png", "jpg", "jpeg", "webp"],
        accept_multiple_files=True
    )

    col_pdf = st.columns([1, 1, 2])
    with col_pdf[0]:
        pdf_dpi = st.selectbox("PDF DPI", [200, 300, 400], index=1)
    with col_pdf[1]:
        pdf_max_pages = st.number_input("PDF 最多页数（0=不限制）", min_value=0, max_value=500, value=0, step=1)
    with col_pdf[2]:
        if not HAVE_PYMUPDF:
            st.warning("未检测到 pymupdf，PDF 上传不可用（pip install pymupdf）。")

    images: List[Image.Image] = []
    page_meta: List[str] = []

    if files:
        for f in files:
            name = getattr(f, "name", "upload")
            ext = (name.split(".")[-1] or "").lower()
            if ext == "pdf":
                if not HAVE_PYMUPDF:
                    st.error("你上传了 PDF，但环境缺少 pymupdf。请安装后重试：pip install pymupdf")
                    st.stop()
                pdf_bytes = f.read()
                imgs = pdf_bytes_to_images(
                    pdf_bytes,
                    dpi=int(pdf_dpi),
                    max_pages=None if int(pdf_max_pages) == 0 else int(pdf_max_pages),
                )
                for i, im in enumerate(imgs, start=1):
                    images.append(im)
                    page_meta.append(f"{name} - p{i}")
            else:
                im = Image.open(f)
                images.append(im)
                page_meta.append(name)

    cols = st.columns([1, 1, 2])
    with cols[0]:
        auto_btn = st.button("🪄 自动推荐参数（基于第1页）", disabled=not images)
    with cols[1]:
        preset_btn = st.button("🎛️ 应用预设", disabled=not images)
    with cols[2]:
        st.caption("推荐会根据文字密度/边缘密度估计参数。")

    if images and auto_btn:
        rec = recommend_ocr_params(images[0], mode=st.session_state["preset_mode"])
        st.session_state["__pending_rec__"] = rec
        st.rerun()

    if images and preset_btn:
        rec = recommend_ocr_params(images[0], mode=st.session_state["preset_mode"])
        st.session_state["__pending_rec__"] = rec
        st.rerun()

    join_lines = st.checkbox("合并断行（适合 PDF 强制换行）", value=False)

    def merge_hard_wraps(md: str) -> str:
        parts = md.split("\n\n")
        out = []
        for p in parts:
            lines = [x.strip() for x in p.splitlines()]
            if any(l.startswith(("-", "*", "|", "#")) for l in lines):
                out.append("\n".join(p.splitlines()))
            else:
                out.append(" ".join([l for l in lines if l]).strip())
        return "\n\n".join(out).strip()

    if st.button("开始 OCR 并导出", type="primary", disabled=not images):
        if not st.session_state["model_id"].strip():
            st.error("请先填写 model（ep-xxxx 或模型ID）。")
            st.stop()

        client = get_ark_client(default_timeout_s=int(st.session_state["ocr_timeout_s"]))

        pages = []
        total_pages = len(images)
        page_bar = st.progress(0, text="准备 OCR…")

        for pi, img in enumerate(images, start=1):
            tile_bar = st.progress(0, text=f"OCR {pi}/{total_pages}：准备切片…")

            def _tile_cb(cur: int, total: int):
                tile_bar.progress(int(cur / total * 100), text=f"OCR {pi}/{total_pages}：切片 {cur}/{total}")

            md = ocr_image_to_markdown(
                client=client,
                model=st.session_state["model_id"].strip(),
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
            page_bar.progress(int(pi / total_pages * 100), text=f"已完成 {pi}/{total_pages}")
            tile_bar.empty()

        merged_md = normalize_md("\n\n---\n\n".join(pages))

        st.success("OCR 完成")
        st.code(merged_md, language="markdown")

        if PANDOC_OK:
            v1_docx = pandoc_md_to_docx(merged_md)
            st.download_button("下载 Rendered.docx", data=v1_docx, file_name="OCR_Rendered.docx")

        st.download_button("下载 Result.md", data=merged_md.encode("utf-8"), file_name="OCR_Result.md")


# ---------------------------
# Tab 2: DOCX translate
# ---------------------------
with tabs[1]:
    st.subheader("Word(.docx) → 翻译（强校验 + 自动重试）")
    st.info("如果导出仍然是原文，本页面会显示“原文-译文对照样本”和“是否判定为未翻译”。")

    docx_file = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_inplace")

    colA, colB, colC = st.columns([1, 1, 1])
    with colA:
        do_translate = st.checkbox("翻译文档（保持尽量原排版）", value=True)
    with colB:
        src_lang = st.selectbox("源语言", ["Auto", "Chinese", "English", "Japanese", "Korean", "Spanish"], index=0)
        dst_lang = st.selectbox("目标语言", ["English", "Chinese", "Japanese", "Korean", "Spanish"], index=0)
    with colC:
        max_batch_chars = st.slider("max_batch_chars", 4000, 20000, 12000, 500)

    show_debug = st.checkbox("显示调试信息（推荐开）", value=True)

    if st.button("开始翻译并导出（new.docx）", type="primary", disabled=not docx_file):
        if src_lang == dst_lang and src_lang != "Auto":
            st.warning("源语言和目标语言相同，翻译可能看起来“没变化”。")

        if not st.session_state["model_id"].strip():
            st.error("请先填写 model（ep-xxxx 或模型ID）。")
            st.stop()

        client = get_ark_client(default_timeout_s=int(st.session_state["translate_timeout_s"]))
        doc_bytes = docx_file.read()
        doc = Document(io.BytesIO(doc_bytes))

        diag_box = st.empty()
        sample_box = st.empty()
        raw_box = st.empty()

        diag_payload_holder: Dict[str, Any] = {}

        def _diag_cb(payload: Dict[str, Any]):
            diag_payload_holder.update(payload)
            if show_debug:
                diag_box.info(
                    f"诊断：items_built={payload.get('items_built')}  "
                    f"estimated_chars={payload.get('estimated_chars')}  "
                    f"contains_textbox_items={payload.get('contains_textbox_items')}"
                )

        def _sample_cb(pairs: List[Tuple[str, str]]):
            if show_debug and pairs:
                md_lines = []
                for i, (o, t) in enumerate(pairs, start=1):
                    md_lines.append(f"**样本 {i} 原文：** {o[:200]}")
                    md_lines.append(f"**样本 {i} 译文：** {t[:200]}")
                    md_lines.append("---")
                sample_box.markdown("\n\n".join(md_lines))

                # 判定是否“明显没翻译”
                bad = 0
                for o, t in pairs:
                    if _looks_untranslated(o, t):
                        bad += 1
                if bad >= max(2, len(pairs)//2):
                    sample_box.warning("检测到样本译文与原文高度一致：模型可能在照抄（已做自动重试，但仍可能无效）。建议更换翻译模型/EP。")

        raw_snippets: List[str] = []
        def _raw_cb(raw: str):
            if show_debug:
                raw_snippets.append(raw[:1200])

        if do_translate:
            batch_bar = st.progress(0, text="准备翻译…")

            def _batch_cb(cur: int, total: int):
                batch_bar.progress(int(cur / total * 100), text=f"翻译批次 {cur}/{total}")

            with st.spinner("翻译中…"):
                translate_docx_in_place(
                    doc=doc,
                    client=client,
                    model=st.session_state["model_id"].strip(),
                    src_lang=src_lang,
                    dst_lang=dst_lang,
                    max_batch_chars=int(max_batch_chars),
                    timeout_s=int(st.session_state["translate_timeout_s"]),
                    progress_cb=_batch_cb,
                    diagnostic_cb=_diag_cb,
                    sample_cb=_sample_cb,
                    raw_cb=_raw_cb,
                )

            if show_debug and raw_snippets:
                raw_box.code("\n\n====\n\n".join(raw_snippets[:2]), language="text")

        out_docx_bytes = doc_to_bytes(doc)
        st.success("完成")
        st.download_button("下载 new.docx", data=out_docx_bytes, file_name="new.docx")


# ---------------------------
# Tab 3: DOCX -> LaTeX/Markdown export
# ---------------------------
with tabs[2]:
    st.subheader("Word(.docx) → LaTeX / Markdown 直接导出（推荐）")
    docx_file2 = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_export")

    out_format = st.selectbox("导出格式", ["latex (.tex)", "markdown (.md)"], index=0)
    wrap_none = st.checkbox("wrap=none（不自动换行）", value=True)

    if st.button("导出", type="primary", disabled=not docx_file2):
        if not PANDOC_OK:
            st.error("Pandoc 不可用，无法导出。")
            st.stop()

        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            in_path = td / "in.docx"
            in_path.write_bytes(docx_file2.read())

            extra_args = []
            if wrap_none:
                extra_args += ["--wrap=none"]

            if out_format.startswith("latex"):
                out_text = pypandoc.convert_file(str(in_path), to="latex", format="docx", extra_args=extra_args)
                st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="latex")
                st.download_button("下载 .tex", data=out_text.encode("utf-8"), file_name="export.tex")
            else:
                out_text = pypandoc.convert_file(str(in_path), to="markdown", format="docx", extra_args=extra_args)
                st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="markdown")
                st.download_button("下载 .md", data=out_text.encode("utf-8"), file_name="export.md")
