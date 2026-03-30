# 📄 AI Document Generator

**Just say what you want — and get a real, beautiful Word, PowerPoint or Excel document instantly.**

This is a local, powerful version of OfficeGPT. Describe any document in natural language, and it automatically creates professional `.docx`, `.xlsx`, or `.pptx` files using real AI (powered by GPT or Grok).

---

## ✨ Features

- **Natural Language Interface** — Type anything like “Create a modern PowerPoint about the future of AI” or “Excel budget forecast for a startup”
- **Real LLM-Powered Content** — Uses OpenAI GPT or xAI Grok to generate high-quality, unique, context-aware content (not just templates)
- **Real Office Files** — Generates actual editable `.docx`, `.xlsx`, and `.pptx` files (using `python-docx`, `openpyxl`, `python-pptx`)
- **Smart Auto-Detection** — Automatically chooses Word, Excel, or PowerPoint based on your request
- **Multiple Styles** — Professional, Modern, Creative, Technical, Academic
- **Beautiful Web UI** — Clean, modern interface that opens automatically in your browser
- **Fallback Mode** — Works even without an API key using smart templates
- **Runs Locally** — Everything stays on your machine (no data sent to third parties except the LLM API you choose)

---

## 🚀 Quick Start

### 1. Installation

```bash
pip install openai python-docx openpyxl python-pptx
```

### 2. Set your API Key (recommended)

```bash
# For OpenAI (recommended for best quality)
export OPENAI_API_KEY=sk-...

# OR for Grok (xAI)
export XAI_API_KEY=xai-...
```

> You can also hard-code the key in the script if preferred.

### 3. Run the App

```bash
python officegpt.py
```

The app will automatically open `http://localhost:8000` in your browser.

---

## 📝 How to Use

Just type a natural request in the input box:

### Examples:

- `Create a PowerPoint about quantum computing in modern style`
- `Excel sheet with quarterly sales forecast for a tech startup`
- `Professional Word report on renewable energy`
- `Creative presentation on the future of work`
- `Technical document explaining machine learning algorithms`

The system will:
1. Detect the document type (Word / Excel / PowerPoint)
2. Choose the best style
3. Generate rich, relevant content using LLM
4. Create and download the real Office file instantly

---

## ⚙️ Configuration

You can easily change settings at the top of the script:

```python
USE_LLM = True                    # Set to False to force template mode
API_PROVIDER = "openai"           # or "grok"
```

Supported models:
- OpenAI: `gpt-4o-mini` (fast & cheap) or `gpt-4o`
- Grok: `grok-4` / `grok-beta`

---

## 🛠️ Tech Stack

- **Backend**: Python + `http.server`
- **LLM**: OpenAI GPT or xAI Grok
- **Documents**: `python-docx`, `openpyxl`, `python-pptx`
- **Frontend**: Clean HTML + Tailwind-like styling

---

## 📌 Notes

- First run may take a few seconds while the LLM generates content.
- All generated files are saved temporarily and automatically cleaned up after download.
- No internet required except for the LLM API call.

---

