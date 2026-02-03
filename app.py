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
st.caption("图片OCR→Markdown/Word；Word→LaTeX/Markdown；Word原排版内翻译（best-effort）。")


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
DEFAULT_ARK_MODEL = "ep-20260203141749-992fx"  # ✅ 你的 EP

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
# 3) OCR: auto recommend
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

    return {
        "max_side": int(max(800, min(3200, max_side))),
        "tile_h": int(max(800, min(2600, tile_h))),
        "overlap": int(max(0, min(400, overlap))),
        "jpeg_q": int(max(50, min(95, jpeg_q))),
        "out_tokens": int(max(1024, min(8192, out_tokens))),
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
# 6) DOCX translate + export helpers (略：保持你之前逻辑不动)
#     —— 为了让你能先跑通，这里保留 Tab③（Pandoc 导出）作为“公式最稳”方案
# ============================================================
def pandoc_convert_docx_to_text(docx_bytes: bytes, to: str, wrap_none: bool = True) -> str:
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        in_path = td / "in.docx"
        in_path.write_bytes(docx_bytes)
        extra_args = ["--wrap=none"] if wrap_none else []
        return pypandoc.convert_file(str(in_path), to=to, format="docx", extra_args=extra_args)

def pandoc_md_to_docx(md: str) -> bytes:
    md = normalize_md(md) + "\n"
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "out.docx"
        fmt = "markdown+fenced_code_blocks+tex_math_dollars"
        pypandoc.convert_text(md, to="docx", format=fmt, outputfile=str(out))
        return out.read_bytes()


# ============================================================
# 7) Session state init (关键：必须在 widgets 创建前完成)
# ============================================================
ocr_timeout_default, translate_timeout_default = get_timeout_defaults()

def init_state():
    defaults = dict(
        model_id=get_default_model(),
        preset_mode="Balanced",
        max_side=2200,
        tile_h=1600,
        overlap=160,
        jpeg_q=85,
        out_tokens=4096,
        ocr_timeout_s=int(ocr_timeout_default),
        translate_timeout_s=int(translate_timeout_default),
    )
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ============================================================
# 8) Callbacks (关键：用 on_click 修改 widget-key 对应 state，避免报错)
# ============================================================
def apply_recommendation_from_first_image():
    imgs = st.session_state.get("_uploaded_imgs_cache", None)
    if not imgs:
        return
    img0 = Image.open(imgs[0])
    rec = recommend_ocr_params(img0, mode=st.session_state["preset_mode"])
    st.session_state.update(rec)
    # update 中的 key 与 slider key 相同，回调阶段更新是允许的

def apply_preset_only():
    # 这里依然需要一张图来跑 recommend（因为你想“估计密度”），否则我们就按固定值填
    imgs = st.session_state.get("_uploaded_imgs_cache", None)
    if imgs:
        img0 = Image.open(imgs[0])
        rec = recommend_ocr_params(img0, mode=st.session_state["preset_mode"])
    else:
        # 没图时按纯预设给一套
        mode = st.session_state["preset_mode"]
        if mode == "Fast":
            rec = {"max_side": 1600, "tile_h": 2000, "overlap": 120, "jpeg_q": 80, "out_tokens": 3072}
        elif mode == "Accurate":
            rec = {"max_side": 2800, "tile_h": 1400, "overlap": 220, "jpeg_q": 90, "out_tokens": 6144}
        else:
            rec = {"max_side": 2200, "tile_h": 1600, "overlap": 160, "jpeg_q": 85, "out_tokens": 4096}
    st.session_state.update(rec)


# ============================================================
# 9) Sidebar UI
# ============================================================
with st.sidebar:
    st.subheader("Ark 配置")
    st.text_input("model（默认 EP 已填）", key="model_id")
    st.caption(f"Base URL: {get_ark_base_url()}")

    st.divider()
    st.subheader("OCR 参数与自动推荐")

    st.selectbox("预设模式", ["Fast", "Balanced", "Accurate"], key="preset_mode",
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

    st.number_input("OCR 请求超时（秒）", min_value=30, max_value=600, key="ocr_timeout_s", step=10)
    st.number_input("翻译请求超时（秒）", min_value=30, max_value=900, key="translate_timeout_s", step=10)

    with st.expander("参数速查", expanded=False):
        st.markdown(
            """
- **输出经常断在半页**：先把 **Tile height** 调小（1600→1200），再把 **OCR max tokens** 调大（4096→6144）。
- **速度太慢/切片太多**：把 **Tile height** 调大（1600→2000），再把 **Max side** 略降（2200→1800）。
- **小字/公式错多**：把 **Max side** 调大（2200→2800+），或 **JPEG** 85→90/95。
"""
        )


# ============================================================
# 10) Tabs
# ============================================================
tabs = st.tabs([
    "① 图片 OCR → 导出",
    "③ Word(.docx) → LaTeX/Markdown 直接导出（推荐）",
])

# ---------------------------
# Tab 1: OCR
# ---------------------------
with tabs[0]:
    st.subheader("图片 OCR（自动推荐参数 + 进度条）")
    st.write("建议：先上传 1 张代表性页面 → 点“自动推荐参数” → 再批量处理。")

    imgs = st.file_uploader("上传图片（可多选）", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)
    # ✅ 把 uploader 的结果临时放 session_state，供回调取用（避免闭包拿不到）
    st.session_state["_uploaded_imgs_cache"] = imgs

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        st.button("🪄 自动推荐参数（基于第1张图）", on_click=apply_recommendation_from_first_image, disabled=not imgs)
    with c2:
        st.button("🎛️ 应用预设", on_click=apply_preset_only)
    with c3:
        st.caption("已修复：按钮回填不再触发 Streamlit session_state 报错。")

    join_lines = st.checkbox("可选：合并断行（适合 PDF 每行强制换行的情况）", value=False)

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
# Tab 2: DOCX -> LaTeX/Markdown export (recommended)
# ---------------------------
with tabs[1]:
    st.subheader("Word(.docx) → LaTeX / Markdown 直接导出（推荐：可编辑公式最稳）")
    st.write("这个模式追求公式正确性，适合论文/作业（Word 的可编辑公式会被 Pandoc 稳定转为 LaTeX）。")

    docx_file2 = st.file_uploader("上传 Word 文档（.docx）", type=["docx"], key="docx_export")
    out_format = st.selectbox("导出格式", ["latex (.tex)", "markdown (.md)"], index=0)
    wrap_none = st.checkbox("wrap=none（不自动换行）", value=True)

    if st.button("导出", type="primary", disabled=not docx_file2):
        docx_bytes = docx_file2.read()

        if out_format.startswith("latex"):
            out_text = pandoc_convert_docx_to_text(docx_bytes, to="latex", wrap_none=wrap_none)
            st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="latex")
            st.download_button("下载 .tex", data=out_text.encode("utf-8"), file_name="export.tex")
        else:
            out_text = pandoc_convert_docx_to_text(docx_bytes, to="markdown", wrap_none=wrap_none)
            st.code(out_text[:4000] + ("\n...\n" if len(out_text) > 4000 else ""), language="markdown")
            st.download_button("下载 .md", data=out_text.encode("utf-8"), file_name="export.md")
