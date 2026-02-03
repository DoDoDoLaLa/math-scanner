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


# ============================================================
# 0) App UI
# ============================================================
st.set_page_config(page_title="Ark OCR / DOCX Translate / LaTeX Export", layout="wide")
st.title("学术 OCR & Word→LaTeX 工具（Ark EP 已接入）")

st.caption(
    "面向论文/讲义/教材：图片OCR→Markdown/Word；Word（含可编辑公式）→LaTeX/Markdown；Word 原排版内翻译（best-effort）。"
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

DEFAULT_ARK_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
DEFAULT_ARK_MODEL = "ep-20260203141749-992fx"   # ✅ 你的 EP
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
    # ✅ 全局复用 + cache（Key/BaseURL/timeout 改变会自动重建）
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
                timeout=timeout_s,  # ✅ 请求级超时
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
    """
    粗略估计“文字密度/边缘密度”：用于推荐切片和分辨率参数。
    - edge_density: 边缘像素比例（论文/表格/公式通常更高）
    - dark_ratio: 暗像素比例（文字多通常更高）
    """
    g = ImageOps.grayscale(img)
    # 统一缩小到可控尺寸做分析（不影响 OCR 用原图）
    g_small = g.copy()
    g_small.thumbnail((900, 900))
    w, h = g_small.size

    # 边缘
    edges = g_small.filter(ImageFilter.FIND_EDGES)
    # 将边缘图二值化
    eb = edges.point(lambda p: 255 if p > 40 else 0)
    edge_px = sum(1 for p in eb.getdata() if p > 0)
    edge_density = edge_px / float(w * h + 1e-9)

    # 暗像素（粗阈值）
    db = g_small.point(lambda p: 1 if p < 160 else 0)
    dark_px = sum(db.getdata())
    dark_ratio = dark_px / float(w * h + 1e-9)

    return {"edge_density": edge_density, "dark_ratio": dark_ratio, "w": float(img.size[0]), "h": float(img.size[1])}

def recommend_ocr_params(img: Image.Image, mode: str = "Balanced") -> Dict[str, int]:
    """
    mode: Fast / Balanced / Accurate
    返回：max_side, tile_h, overlap, jpeg_q, out_tokens
    """
    m = analyze_image_density(img)
    w, h = int(m["w"]), int(m["h"])
    edge = m["edge_density"]
    dark = m["dark_ratio"]

    # 复杂度评分：越大说明越“密排/公式表格多”
    complexity = 0.6 * edge + 0.4 * dark

    # 基础参数
    if mode == "Fast":
        max_side = 1600
        jpeg_q = 80
        out_tokens = 3072
        base_tile = 2000
        base_overlap = 120
    elif mode == "Accurate":
        max_side = 2800
        jpeg_q = 90
        out_tokens = 6144
        base_tile = 1400
        base_overlap = 220
    else:  # Balanced
        max_side = 2200
        jpeg_q = 85
        out_tokens = 4096
        base_tile = 1600
        base_overlap = 160

    # 根据复杂度调整
    # complexity 常见大致在 0.05~0.25 间（不同图差异很大，这里是启发式）
    if complexity > 0.18:         # 非常密排/表格多
        max_side = min(3200, max_side + 400)
        out_tokens = min(8192, out_tokens + 1024)
        tile_h = max(1000, base_tile - 300)
        overlap = min(320, base_overlap + 80)
        jpeg_q = min(95, jpeg_q + 5)
    elif complexity < 0.10:       # 较稀疏/字较大
        tile_h = min(2400, base_tile + 300)
        overlap = max(80, base_overlap - 40)
    else:
        tile_h = base_tile
        overlap = base_overlap

    # 根据高度决定切片策略
    # 高度很大 → 适当降低 tile 以减少单次输出截断风险（尤其复杂度高）
    if h >= 4000 and complexity > 0.14:
        tile_h = max(900, tile_h - 200)
        out_tokens = min(8192, out_tokens + 512)

    # clamp
    max_side = int(max(800, min(3200, max_side)))
    tile_h = int(max(800, min(2600, tile_h)))
    overlap = int(max(0, min(400, overlap)))
    jpeg_q = int(max(50, min(95, jpeg_q)))
    out_tokens = int(max(1024, min(8192, out_tokens)))

    return {
        "max_side": max_side,
        "tile_h": tile_h,
        "overlap": overlap,
        "jpeg_q": jpeg_q,
        "out_tokens": out_tokens,
    }


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

def iter_all_paragraphs(doc: Document) -> List[Paragraph]:
    ps: List[Paragraph] = []
    ps.extend(doc.paragraphs)
    for table in doc.tables:
        ps.extend(iter_table_paragraphs(table))
    return ps

def extract_math_from_docx_with_pandoc(docx_bytes: bytes) -> List[Tuple[str, str]]:
    """
    best-effort：用 pandoc 抽取 markdown 里的 $...$/$$...$$ 序列
    仍可能与 docx 内 oMath 遍历顺序不一致 —— 所以这里只用于“就地替换”模式。
    """
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
    all_ps = iter_all_paragraphs(doc)
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
                latex_text = (f"$$ {body} $$") if kind == "display" else (f"$ {body} $")

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

def doubao_translate_items(
    client: OpenAI,
    model: str,
    items: List[dict],
    src_lang: str,
    dst_lang: str,
    timeout_s: int,
) -> Dict[str, List[str]]:
    src = "auto-detect" if src_lang == "Auto" else src_lang
    prompt = (TRANSLATE_PROMPT_TEMPLATE.replace("__SRC_LANG__", src).replace("__DST_LANG__", dst_lang))

    payload = {"items": items}
    messages = [{"role": "user", "content": prompt + "\n\n" + json.dumps(payload, ensure_ascii=False)}]

    res = safe_chat_completions(
        client=client,
        model=model,
        messages=messages,
        temperature=0.2,
        max_tokens=8192,
        timeout_s=timeout_s,
        retries=6,
    )
    if res.error_message:
        st.error(f"翻译调用失败：\n\n{res.error_message}")
        st.stop()

    obj = extract_json_object(res.text)
    out: Dict[str, List[str]] = {}
    items_out = obj.get("items", [])
    if not isinstance(items_out, list):
        raise ValueError("JSON schema error: items must be a list")
    for it in items_out:
        out[it["id"]] = it["segments"]
    return out

def translate_docx_in_place(
    doc: Document,
    client: OpenAI,
    model: str,
    src_lang: str,
    dst_lang: str,
    max_batch_chars: int,
    timeout_s: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
):
    all_ps = iter_all_paragraphs(doc)

    items = []
    para_refs = {}
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

        if not "".join(segs).strip():
            continue

        pid += 1
        item_id = f"p{pid}"
        items.append({"id": item_id, "segments": segs})
        para_refs[item_id] = (runs, maps)

    batches = chunk_items_for_api(items, max_chars=max_batch_chars)
    total = len(batches)

    for bi, batch in enumerate(batches, start=1):
        if progress_cb:
            progress_cb(bi, total)

        translated_map = doubao_translate_items(
            client=client, model=model, items=batch,
            src_lang=src_lang, dst_lang=dst_lang,
            timeout_s=timeout_s,
        )

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

with st.sidebar:
    st.subheader("Ark 配置")
    st.text_input("model（默认 EP 已填）", key="model_id")
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
    "① 图片 OCR → 导出",
    "② Word(.docx) → 保排版翻译/就地替换公式（best-effort）",
    "③ Word(.docx) → LaTeX/Markdown 直接导出（推荐）",
])

# ---------------------------
# Tab 1: OCR
# ---------------------------
with tabs[0]:
    st.subheader("图片 OCR（带自动参数推荐 + 进度条）")
    st.write("建议：先上传 1 张代表性页面 → 点“自动推荐参数” → 再批量处理。")

    imgs = st.file_uploader("上传图片（可多选）", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

    cols = st.columns([1, 1, 2])
    with cols[0]:
        auto_btn = st.button("🪄 自动推荐参数（基于第1张图）", disabled=not imgs)
    with cols[1]:
        preset_btn = st.button("🎛️ 应用预设（Fast/Balanced/Accurate）", disabled=not imgs)
    with cols[2]:
        st.caption("自动推荐会根据图片高度与文字/边缘密度估计调整 max_side/tile/overlap/tokens。")

    if imgs and auto_btn:
        img0 = Image.open(imgs[0])
        rec = recommend_ocr_params(img0, mode=st.session_state["preset_mode"])
        st.session_state["max_side"] = rec["max_side"]
        st.session_state["tile_h"] = rec["tile_h"]
        st.session_state["overlap"] = rec["overlap"]
        st.session_state["jpeg_q"] = rec["jpeg_q"]
        st.session_state["out_tokens"] = rec["out_tokens"]
        st.success(f"已回填推荐参数：{rec}")

    if imgs and preset_btn:
        # 仅按预设，不基于图片分析
        rec = recommend_ocr_params(Image.open(imgs[0]), mode=st.session_state["preset_mode"])
        st.session_state["max_side"] = rec["max_side"]
        st.session_state["tile_h"] = rec["tile_h"]
        st.session_state["overlap"] = rec["overlap"]
        st.session_state["jpeg_q"] = rec["jpeg_q"]
        st.session_state["out_tokens"] = rec["out_tokens"]
        st.success(f"已应用预设参数：{rec}")

    # 可选：合并“PDF 换行”
    join_lines = st.checkbox("可选：合并断行（适合 PDF 每行强制换行的情况）", value=False,
                             help="会把同段落中间的单换行合并为一个空格；保留空行分段。")

    def merge_hard_wraps(md: str) -> str:
        # 简单规则：段内单换行合并，段间空行保留
        parts = md.split("\n\n")
        out = []
        for p in parts:
            lines = [x.strip() for x in p.splitlines()]
            # 不处理列表/表格/标题块（保守）
            if any(l.startswith(("-", "*", "|", "#")) for l in lines):
                out.append("\n".join(p.splitlines()))
            else:
                out.append(" ".join([l for l in lines if l != ""]).strip())
        return "\n\n".join(out).strip()

    if st.button("开始 OCR 并导出", type="primary", disabled=not imgs):
        if not st.session_state["model_id"].strip():
            st.error("请先填写 model（ep-xxxx 或模型ID）。")
            st.stop()

        client = get_ark_client(default_timeout_s=int(st.session_state["ocr_timeout_s"]))

        pages = []
        total_pages = len(imgs)
        page_bar = st.progress(0, text="准备 OCR…")

        for pi, f in enumerate(imgs, start=1):
            img = Image.open(f)
            tile_bar = st.progress(0, text=f"OCR 第 {pi}/{total_pages} 页：准备切片…")

            def _tile_cb(cur: int, total: int):
                tile_bar.progress(int(cur / total * 100), text=f"OCR 第 {pi}/{total_pages} 页：切片 {cur}/{total}")

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

            pages.append(f"## 第 {pi} 页\n\n{md}")
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
# Tab 2: DOCX in-place translate + best-effort equation replace
# ---------------------------
with tabs[1]:
    st.subheader("Word(.docx) → 保排版翻译/就地替换公式（best-effort）")
    st.warning(
        "说明：Word 可编辑公式（OMML）要“可靠”转 LaTeX，推荐使用 Tab③ 的 Pandoc 直接导出。"
        "本 Tab 的“就地替换公式”为 best-effort，可能出现错位/不全。"
    )

    docx_file = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_inplace")

    colA, colB = st.columns([1, 1])
    with colA:
        do_equation_replace = st.checkbox("把 Word 原生公式（OMML）替换为 LaTeX 代码（best-effort）", value=False)
        do_translate = st.checkbox("翻译文档（尽量保持原排版）", value=False)

    with colB:
        src_lang = st.selectbox("源语言", ["Auto", "Chinese", "English", "Japanese", "Korean", "Spanish"], index=0)
        dst_lang = st.selectbox("目标语言", ["Chinese", "English", "Japanese", "Korean", "Spanish"], index=1)
        max_batch_chars = st.slider("翻译分批大小（max_batch_chars）", 4000, 20000, 12000, 500)

    if st.button("开始处理并导出（new.docx）", type="primary", disabled=not docx_file):
        if not st.session_state["model_id"].strip():
            st.error("请先填写 model（ep-xxxx 或模型ID）。")
            st.stop()

        client = get_ark_client(default_timeout_s=int(st.session_state["translate_timeout_s"]))

        doc_bytes = docx_file.read()
        doc = Document(io.BytesIO(doc_bytes))

        if do_equation_replace:
            with st.spinner("提取公式序列（pandoc）..."):
                math_seq = extract_math_from_docx_with_pandoc(doc_bytes)

            with st.spinner("替换 Word 原生公式为 LaTeX 代码（best-effort）..."):
                replaced_count = replace_omml_with_latex_code(doc, math_seq, use_equation_env=True)

            st.info(f"已替换公式数量（best-effort）：{replaced_count}")

        if do_translate:
            batch_bar = st.progress(0, text="准备翻译…")

            def _batch_cb(cur: int, total: int):
                batch_bar.progress(int(cur / total * 100), text=f"翻译批次 {cur}/{total}")

            with st.spinner("翻译中（分批提交，保持图片/公式位置不动）..."):
                translate_docx_in_place(
                    doc=doc,
                    client=client,
                    model=st.session_state["model_id"].strip(),
                    src_lang=src_lang,
                    dst_lang=dst_lang,
                    max_batch_chars=int(max_batch_chars),
                    timeout_s=int(st.session_state["translate_timeout_s"]),
                    progress_cb=_batch_cb,
                )

        out_docx_bytes = doc_to_bytes(doc)
        st.success("处理完成")
        st.download_button("下载 new.docx", data=out_docx_bytes, file_name="new.docx")


# ---------------------------
# Tab 3: DOCX -> LaTeX/Markdown export (recommended)
# ---------------------------
with tabs[2]:
    st.subheader("Word(.docx) → LaTeX / Markdown 直接导出（推荐：可编辑公式最稳）")
    st.write(
        "这个模式不追求保留 Word 的排版，而是追求“学术 LaTeX 输出的正确性”，尤其适合含大量可编辑公式的论文/作业。"
    )

    docx_file2 = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_export")

    out_format = st.selectbox("导出格式", ["latex (.tex)", "markdown (.md)"], index=0)
    wrap_none = st.checkbox("wrap=none（不自动换行）", value=True)

    if st.button("导出", type="primary", disabled=not docx_file2):
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
                # markdown：尽量保留 tex_math_dollars
                out_text = pypandoc.convert_file(str(in_path), to="markdown", format="docx", extra_args=extra_args)
                st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="markdown")
                st.download_button("下载 .md", data=out_text.encode("utf-8"), file_name="export.md")
