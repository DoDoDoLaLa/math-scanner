# Streamlit DOCX Math Tool

部署：直接把本目录推到 GitHub，然后在 https://share.streamlit.io/ 选择仓库 + app.py。

## 依赖
- requirements.txt 已包含：streamlit, openai, python-docx, pypandoc, pypandoc-binary 等。

## 关键改动
- 新增 `math_pipeline.py`：集中放 DOCX<->Pandoc<->DOCX(OMML) 以及 AI 校正逻辑
- `app.py`：主要负责 UI；并新增
  - 统一 `getvalue()` 读取：避免 UploadedFile.read() 导致二次读取为空
  - `st.cache_data` 缓存：Pandoc 转换 + AI 校正结果缓存，减少卡顿/空文件/重复付费

## 使用提示
- Tab3 的 “把.../... 文本公式转为可编辑 Word 公式(OMML)”：适合把 $...$ / $$...$$ 转成 Word 可编辑公式
- Tab3 的 “AI 增强”：
  - 开启后会先 DOCX->Markdown，再交给模型做“只校正公式不改正文”的清洗，再回写 DOCX
