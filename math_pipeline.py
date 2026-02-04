# -*- coding: utf-8 -*-
"""
math_pipeline.py
集中放：DOCX<->Pandoc(Markdown/LaTeX)<->DOCX(OMML) 的“计算/转换”逻辑 + 统一的上传读取工具。

设计目标：
- app.py 尽量只做 UI（Streamlit 控件、布局、状态管理）。
- 这里提供纯函数（便于测试/缓存），不直接依赖 Streamlit。
"""
from __future__ import annotations

import io
import re
import hashlib
import tempfile
from pathlib import Path
from typing import Callable, List, Optional, Tuple

import pypandoc


# -----------------------------
# Common helpers
# -----------------------------
def ensure_pandoc() -> None:
    """Ensure pandoc exists. In Streamlit Cloud, pypandoc-binary is recommended."""
    try:
        _ = pypandoc.get_pandoc_path()
    except OSError:
        try:
            pypandoc.download_pandoc()
        except Exception:
            # If download fails (e.g. no internet), rely on system/pypandoc-binary.
            pass


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def read_uploaded_bytes(uploaded_file) -> bytes:
    """Robustly read bytes from Streamlit UploadedFile (or file-like).

    Prefer .getvalue() to avoid consuming the internal pointer of the uploaded stream.
    """
    if uploaded_file is None:
        return b""
    if hasattr(uploaded_file, "getvalue"):
        return uploaded_file.getvalue()
    if hasattr(uploaded_file, "read"):
        return uploaded_file.read()
    raise TypeError(f"Unsupported uploaded object: {type(uploaded_file)}")


def normalize_md(md: str) -> str:
    md = (md or "").replace("\r\n", "\n").replace("\r", "\n")
    md = re.sub(r"\n{3,}", "\n\n", md).strip()
    return md


# -----------------------------
# Pandoc math normalization / sanitization
# -----------------------------
PANDOC_DISPLAY_MATH_BRACKET_RE = re.compile(r"(?s)\\\[(.+?)\\\]")
PANDOC_DISPLAY_MATH_DOLLAR_RE = re.compile(r"(?s)\$\$(.+?)\$\$")
PANDOC_INLINE_MATH_PAREN_RE = re.compile(r"(?s)\\\((.+?)\\\)")
PANDOC_INLINE_MATH_DOLLAR_RE = re.compile(r"(?s)(?<!\$)\$([^$\n]+)\$(?!\$)")


def normalize_pandoc_math(
    text: str,
    *,
    display_style: str = "dollars",  # equation|equation*|bracket|dollars
    inline_style: str = "dollars",   # paren|dollars
) -> str:
    """Normalize delimiters so math is stable for later parsing and AI cleaning."""
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
        if display_style == "bracket":
            return f"\\[\n{body}\n\\]"
        # dollars
        return f"$$\n{body}\n$$"

    t = PANDOC_DISPLAY_MATH_BRACKET_RE.sub(_display_repl, t)
    t = PANDOC_DISPLAY_MATH_DOLLAR_RE.sub(_display_repl, t)

    def _inline_repl(m):
        body = (m.group(1) or "").strip()
        if inline_style == "paren":
            return f"\\({body}\\)"
        return f"${body}$"

    t = PANDOC_INLINE_MATH_PAREN_RE.sub(_inline_repl, t)
    t = PANDOC_INLINE_MATH_DOLLAR_RE.sub(lambda m: _inline_repl(m), t)
    return t


def sanitize_tex_math_for_pandoc(md: str) -> str:
    """Make TeX math more Pandoc-friendly:
    - Undo common escaping patterns (\$ -> $)
    - Ensure $$...$$ blocks contain NO blank lines (Pandoc can mis-parse)
    """
    if not md:
        return md

    t = md.replace("\r\n", "\n").replace("\r", "\n").replace("\u00A0", " ")
    t = t.replace(r"\\$", "$").replace(r"\$", "$")

    block_re = re.compile(r"(?s)\$\$(.+?)\$\$")
    inline_re = re.compile(r"(?s)(?<!\$)\$([^$\n]+)\$(?!\$)")

    def _clean_body(body: str, is_block: bool) -> str:
        b = (body or "").strip()
        if is_block:
            b = re.sub(r"\n\s*\n+", "\n", b)  # collapse blank lines
            b = re.sub(r"[ \t]+\n", "\n", b)
        else:
            b = b.replace("\n", " ")
            b = re.sub(r"\s{2,}", " ", b)
        return b.strip()

    def _fix_block(m):
        b = _clean_body(m.group(1), True)
        return "$$\n" + b + "\n$$"

    def _fix_inline(m):
        b = _clean_body(m.group(1), False)
        return "$" + b + "$"

    t = block_re.sub(_fix_block, t)
    t = inline_re.sub(_fix_inline, t)
    return t


# -----------------------------
# DOCX <-> Pandoc
# -----------------------------
def docx_to_pandoc_markdown_for_math(docx_bytes: bytes, *, wrap_none: bool = True) -> str:
    """DOCX -> Markdown(+tex_math_dollars) suitable for later math parsing."""
    if not docx_bytes:
        return ""

    ensure_pandoc()

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
    md = md.replace(r"\$", "$").replace(r"\\$", "$")

    # DOCX writer often escapes backslashes; be conservative:
    md = md.replace("\\\\", "\\")

    md = normalize_pandoc_math(md, display_style="dollars", inline_style="dollars")
    md = sanitize_tex_math_for_pandoc(md)
    return normalize_md(md)


def pandoc_markdown_to_docx(md: str, *, reference_docx_bytes: Optional[bytes] = None) -> bytes:
    """Markdown(+tex_math_dollars) -> DOCX. If reference_docx_bytes provided, preserve styles via --reference-doc."""
    ensure_pandoc()
    md = normalize_md(md) + "\n"

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        out_path = td / "out.docx"

        extra_args: List[str] = []
        if reference_docx_bytes:
            ref = td / "ref.docx"
            ref.write_bytes(reference_docx_bytes)
            extra_args += ["--reference-doc", str(ref)]

        pypandoc.convert_text(
            md,
            to="docx",
            format="markdown+tex_math_dollars",
            outputfile=str(out_path),
            extra_args=extra_args,
        )
        return out_path.read_bytes()




def pandoc_markdown_to_latex(md: str) -> str:
    """Markdown(+tex_math_dollars) -> LaTeX string (for export)."""
    ensure_pandoc()
    md = normalize_md(md) + "\n"
    return pypandoc.convert_text(md, to="latex", format="markdown+tex_math_dollars")


def docx_roundtrip_make_equations_editable(docx_bytes: bytes) -> Tuple[bytes, str]:
    """DOCX -> MD(tex_math_dollars) -> sanitize -> DOCX (Pandoc emits OMML). Returns (out_docx, md_used)."""
    md = docx_to_pandoc_markdown_for_math(docx_bytes, wrap_none=True)
    out = pandoc_markdown_to_docx(md, reference_docx_bytes=docx_bytes)
    return out, md


# -----------------------------
# AI-assisted cleaning (app passes in a callable)
# -----------------------------
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


def chunk_text_by_paragraph(md: str, *, max_chars: int) -> List[str]:
    if not md:
        return []
    paras = md.split("\n\n")
    chunks: List[str] = []
    cur: List[str] = []
    cur_len = 0
    for p in paras:
        p2 = p.strip("\n")
        if not p2:
            continue
        add = p2 + "\n\n"
        if cur and cur_len + len(add) > max_chars:
            chunks.append("".join(cur).strip() + "\n")
            cur, cur_len = [], 0
        cur.append(add)
        cur_len += len(add)
    if cur:
        chunks.append("".join(cur).strip() + "\n")
    return chunks


def ai_clean_markdown_math_like_ocr(
    md: str,
    *,
    call_llm: Callable[[str], str],
    max_batch_chars: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> str:
    """LLM cleans markdown math. call_llm(prompt_text)->response_text should be provided by app."""
    md = normalize_md(md)
    if not md:
        return md

    parts = chunk_text_by_paragraph(md, max_chars=int(max_batch_chars))
    out_parts: List[str] = []
    total = len(parts)

    for i, part in enumerate(parts, start=1):
        if progress_cb:
            progress_cb(i, total)
        prompt = AI_MATH_CLEAN_PROMPT_ZH.strip() + "\n\n---\n\n" + part
        cleaned = normalize_md(call_llm(prompt))
        out_parts.append(cleaned)

    merged = normalize_md("\n\n".join([p for p in out_parts if p.strip()]))
    merged = normalize_pandoc_math(merged, display_style="dollars", inline_style="dollars")
    merged = sanitize_tex_math_for_pandoc(merged)
    return normalize_md(merged)


def docx_ai_roundtrip_make_equations_editable(
    docx_bytes: bytes,
    *,
    call_llm: Callable[[str], str],
    max_batch_chars: int,
    progress_cb: Optional[Callable[[int, int], None]] = None,
) -> Tuple[bytes, str]:
    """DOCX -> MD -> AI clean -> DOCX(OMML). Returns (out_docx, cleaned_md)."""
    md = docx_to_pandoc_markdown_for_math(docx_bytes, wrap_none=True)
    md2 = ai_clean_markdown_math_like_ocr(
        md,
        call_llm=call_llm,
        max_batch_chars=int(max_batch_chars),
        progress_cb=progress_cb,
    )
    out = pandoc_markdown_to_docx(md2, reference_docx_bytes=docx_bytes)
    return out, md2