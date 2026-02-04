# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/). Versions follow [SemVer](https://semver.org/).

## [0.1.0] - 初版 / First Release

Released: 2026‑02‑04

### Added / 新增

* **Math OCR**: Convert images (JPEG/PNG) and selected PDF pages into Markdown with LaTeX math using `$...$` for inline and `$$...$$` for block equations. The math can be exported back to Word (OMML) or LaTeX.
* **DOCX Translation**: Translate Word documents while preserving layout (best‑effort) and protect LaTeX equations from corruption. Support both direct translation and enhanced Pandoc‑based translation to produce editable equations in Word.
* **Export Formats**: Added buttons to export translated documents to Markdown and LaTeX. Pandoc is used to convert between formats and embed equations as OMML.
* **Glossary and QA**: Added glossary support for consistent terminology and a post‑translation quality check that retries segments if the output contains unwanted language or improper math delimiters.
* **PDF Page Range**: Users can select page ranges when processing PDFs to avoid unnecessary OCR on entire files.
* **Streamlit UI Enhancements**: Multiple tabs (OCR, Word translation, enhanced Pandoc translation) with configurable parameters and model settings.

### Changed / 更改

* N/A – initial release.

### Fixed / 修复

* N/A – initial release.
