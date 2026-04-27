# arXiv Flask App

English | [中文](README.zh-CN.md)

A Flask web application for academic paper workflows. It searches arXiv papers, previews candidate results, parses PDF content, and exports Word / TXT / Markdown files. The interface follows a simple “search -> preview -> export” flow for literature screening, paper organization, and local formatting experiments.

## Project Idea

This project combines paper search, result filtering, PDF retrieval, content parsing, and document export into a lightweight academic workspace. After a user enters keywords, the backend queries the arXiv API and returns candidate papers. The frontend displays titles, authors, years, abstracts, and PDF links. After a paper is selected, the system downloads the PDF, extracts text, images, and structural information with PyMuPDF, then exports editable documents through python-docx and related formatting logic.

The app also keeps a local LLM translation extension point. During Word export, it can call an Ollama-compatible API to process paper content in Chinese. The API endpoint and model are configurable through environment variables, so the app can run on a personal machine or connect to an existing local model service.

## Key Features

- Keyword search through arXiv, returning up to 50 candidate papers by default.
- Candidate list with title, authors, year, and abstract preview.
- Paper detail preview with original PDF access.
- Export support for Word, TXT, and Markdown.
- Word export with structured title, authors, abstract, body text, and images.
- Optional local translation through an Ollama-compatible API.
- Automatic download entry after export is completed.

## Screenshots

### 1. Home and Search Entry

![Home and Search Entry](docs/screenshots/466d39cca27a1abda8ddaf7db2e4b406.png)

### 2. Search Results

![Search Results](docs/screenshots/43f331cc46bcd673032a011a08f39ed5.png)

### 3. Candidate Papers and Detail Preview

![Candidate Papers and Detail Preview](docs/screenshots/5e64fd3810d7e7d848c077c9135c11f8.png)

### 4. Export Progress

![Export Progress](docs/screenshots/5bd535df714146039d0be7922f7466cc.png)

### 5. Word Export Output

![Word Export Output](docs/screenshots/bf450ea35621e0c17f3e7ddf2f67be87.png)

## Usage

1. Clone the repository.

```bash
git clone https://github.com/xinruliuresearch-maker/arxiv_flask_app.git
cd arxiv_flask_app
```

2. Create and activate a virtual environment.

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

macOS / Linux:

```bash
source .venv/bin/activate
```

3. Install dependencies.

```bash
pip install -r requirements.txt
```

4. Start the app.

```bash
python run_app.py
```

5. Open the app in your browser.

```text
http://127.0.0.1:5000
```

6. Enter keywords, choose an export format, and click the search button. Select a candidate paper to preview details, open the PDF, or export the selected paper.

## Optional Configuration

To enable local translation, start Ollama or another service compatible with `/api/generate`, then configure environment variables as needed.

```bash
export OLLAMA_API_URL="http://127.0.0.1:11434/api/generate"
export OLLAMA_MODEL="deepseek-r1:8b"
python run_app.py
```

Defaults:

- `OLLAMA_API_URL`: `http://127.0.0.1:11434/api/generate`
- `OLLAMA_MODEL`: `deepseek-r1:8b`

## Deployment

### Local Deployment

Best for personal use, paper organization, and feature testing.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python run_app.py
```

### Server Deployment

For production-like usage, use a WSGI server such as Gunicorn or Waitress instead of the Flask debug server.

Linux example:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install gunicorn
gunicorn -w 2 -b 0.0.0.0:5000 app:app
```

Windows example:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install waitress
waitress-serve --host=0.0.0.0 --port=5000 app:app
```

For public access, place Nginx / Caddy or another reverse proxy in front of the app, and configure HTTPS, upload/download limits, and access control.

## File Structure

```text
arxiv_flask_app/
├── app.py                    # Main Flask app: routes, arXiv search, PDF parsing, export logic
├── run_app.py                # Local startup entry
├── requirements.txt          # Python dependencies
├── README.md                 # Language selection entry
├── README.zh-CN.md           # Chinese documentation
├── README.en.md              # English documentation
├── static/
│   └── style.css             # Frontend styles
├── templates/
│   └── index.html            # Single-page workspace UI
└── docs/
    └── screenshots/          # Screenshots used by README
```

## Tech Stack

- Flask: Web server and routes
- requests / feedparser: arXiv API requests and Atom feed parsing
- PyMuPDF: PDF text, image, and layout extraction
- python-docx: Word document generation
- HTML / CSS / JavaScript: Frontend interaction
- Ollama-compatible API: Optional local translation support
