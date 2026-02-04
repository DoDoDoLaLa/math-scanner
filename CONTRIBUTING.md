<!--
This file provides guidelines for contributors in both Chinese and English.
请参阅本文件的中文版，以了解如何参与贡献。
-->

# Contributing / 贡献指南

> 欢迎所有人参与改善 **math‑scanner** 项目！我们真诚感谢每一位贡献者的时间和建议。本指南列出了如何提问题、贡献代码以及提交拉取请求（PR）的步骤。

## 📌 Reporting Issues / 提交问题

请先确认您的问题尚未在 [Issues](https://github.com/DoDoDoLaLa/math-scanner/issues) 中被报告过。我们建议按照以下步骤提交问题：

1. 查看 [FAQ](./README.md#faq--troubleshooting) 部分，确认问题未被提及。
2. 使用仓库的搜索功能查找是否已经存在相同或相似的问题。
3. 如果问题未解决，请创建一个新的 issue，标题简明扼要并描述问题。
4. 在 issue 中提供以下信息：
   - 您使用的操作系统和浏览器版本。
   - 您使用的模型版本（如果更改过默认设置）。
   - 重现步骤：请提供足够的信息让我们能够复现问题（上传示例文件或截图）。

**English:**

1. Check the [FAQ](./README.md#faq--troubleshooting) to see if your issue is covered.
2. Search existing [Issues](https://github.com/DoDoDoLaLa/math-scanner/issues) to avoid duplicates.
3. If the problem has not been addressed, open a new issue with a clear title and description.
4. Include environment details (OS, browser), model configuration, and steps to reproduce. Attaching sample files or screenshots helps us diagnose quickly.

## 🛠️ Development Setup / 开发环境配置

要在本地开发并运行该项目，请遵循以下步骤：

```bash
git clone https://github.com/DoDoDoLaLa/math-scanner.git
cd math-scanner
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

需要注意的是，本项目依赖外部模型服务，请不要将任何真实的 `ARK_API_KEY` 或其他密钥提交到代码库。可以在 `.streamlit/secrets.toml` 或环境变量中配置 API 密钥。

**English:** To set up a local development environment:

```bash
git clone https://github.com/DoDoDoLaLa/math-scanner.git
cd math-scanner
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
streamlit run app.py
```

Configure your API key in `.streamlit/secrets.toml` or as an environment variable (e.g. `export ARK_API_KEY=YOUR_KEY`). Never commit secrets to the repository.

## ✍️ Coding Style & Commit Messages / 代码风格和提交信息

为了保持项目代码的整洁，请遵循以下指南：

* 使用 [PEP 8](https://www.python.org/dev/peps/pep-0008/) 作为 Python 代码风格参考。
* 在适当的位置添加注释，尤其是涉及提示词、模型配置和可能影响功能的部分。
* 每个提交信息应简洁明了，标题以动词开头，例如 `fix: 修复 OCR 图片旋转问题` 或 `feat: 添加词汇表支持`。
* 如果提交涉及重大改动或需要额外说明，请在提交正文中提供更多细节。

**English:**

* Follow [PEP 8](https://www.python.org/dev/peps/pep-0008/) for Python code style.
* Comment your code where it improves readability or explains model prompts/configuration.
* Craft clear commit messages. Use an imperative verb at the start, e.g. `fix: correct image rotation`, `feat: add glossary support`. Include details in the body if necessary.

## ➕ Submitting a Pull Request / 提交拉取请求（PR）

1. Fork this repository and create a new branch based on `main` (e.g. `feat/add-pdf-range`).
2. Commit your changes to this branch and push to your fork.
3. Open a pull request against `main` with a detailed description of your changes and the issue it addresses.
4. 请确保 CI 测试通过。如果有新的依赖，请更新 `requirements.txt` 并说明用途。
5. 请遵循仓库维护者的审查意见，必要时对 PR 进行修改。

**English:**

1. Fork the repository and create a feature branch (e.g. `feat/add-pdf-range`).
2. Commit and push your changes to your fork.
3. Open a pull request against `main` describing what your changes do and which issue they address.
4. Ensure tests (if any) pass. Update `requirements.txt` when introducing new dependencies and justify the addition.
5. Address review comments from maintainers. We may ask for changes before merging.

感谢您对 **math‑scanner** 的关注和贡献！
