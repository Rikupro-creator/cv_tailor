# Professional CV Suite

An AI-powered Streamlit application that tailors your CV to a specific job description, generates a personalised cover letter, and produces professional Word documents — all with full hyperlink preservation.

---

## Features

- **CV Extraction** — Parses 15+ sections from PDF or Word uploads, including clickable hyperlinks
- **AI Tailoring** — Rewrites your CV using the CAR-L method (Challenge → Action → Result → Learning) to match the target job description
- **ATS Optimisation** — Embeds 15–20 keywords from the job description throughout the CV
- **Cover Letter Generation** — Writes a personalised, story-driven cover letter (380–430 words)
- **Hyperlink Preservation** — Extracts and re-embeds clickable links (LinkedIn, GitHub, portfolio, project URLs) in the output Word documents
- **Match Scoring** — Provides an honest 1–10 score of how well your CV matches the job
- **Gap Analysis** — Identifies missing qualifications and areas for improvement
- **Interview Talking Points** — Generates 5–6 specific stories to prepare for the interview
- **Bundle Download** — Download the tailored CV and cover letter individually or as a ZIP

---

## Supported AI Providers

| Provider | Setup |
|---|---|
| Google Gemini | Free tier available at [aistudio.google.com](https://aistudio.google.com) |
| OpenRouter | Access to GPT-4, Claude, Mistral, and more at [openrouter.ai](https://openrouter.ai) |

---

## Installation

### 1. Clone the repository

```bash
git clone https://github.com/Rikupro-creator/cv-suite.git
cd cv-suite
```

### 2. Create a virtual environment (recommended)

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the app

```bash
streamlit run cv_suite.py
```

The app will open at `http://localhost:8501`.

---

## Requirements

All dependencies are listed in `requirements.txt`:

| Package | Purpose |
|---|---|
| `streamlit` | Web application framework |
| `google-generativeai` | Google Gemini AI API |
| `openai` | OpenRouter API client (OpenAI-compatible) |
| `python-docx` | Read and write Word (.docx) files |
| `pdfplumber` | Extract text and word positions from PDFs |
| `PyMuPDF` | Extract hyperlink annotations from PDFs |
| `lxml` | Parse DOCX XML for hyperlink extraction |

---

## Usage

### Step 1 — Configure AI Provider

In the sidebar, select either **Gemini** or **OpenRouter**, paste your API key, and choose a model.

### Step 2 — Upload your CV

Upload a `.pdf` or `.docx` file. Word format gives the most reliable hyperlink extraction because links are stored in the document's relationship XML. PDF links work when they are proper clickable annotations (not just typed-out URLs).

### Step 3 — Paste the job description

Paste the full job posting, including responsibilities, requirements, and any company description. The more detail you provide, the better the tailoring.

### Step 4 — Generate

Click **Generate Tailored CV & Cover Letter**. The app runs four steps:

1. Extracts all data from your CV (text + hyperlinks)
2. Tailors the CV to the job description
3. Writes the cover letter
4. Builds the Word documents

### Step 5 — Review and download

Use the tabs to preview the tailored CV, edit the cover letter, review the ATS analysis, and inspect the raw extracted data. Download the CV, cover letter, or both as a ZIP bundle.

---

## How Hyperlinks Work

### Extraction

**Word (.docx):** The app iterates each paragraph's XML children, detects `<w:hyperlink>` nodes, and looks up the target URL from the document's `.rels` relationship table. This is the most reliable method.

**PDF:** The app uses `pdfplumber` to extract text and word bounding boxes, then uses `PyMuPDF` to read the PDF's link annotation rectangles and their URIs. Each link rectangle is matched to the nearest word by coordinate distance, and the text is annotated as `[word](url)`.

### Preservation through AI tailoring

Extracted hyperlinks are passed into the AI prompt as part of the structured JSON data. The prompt explicitly instructs the model to preserve all `link` fields without modification. After tailoring, links are written back into the output document.

### Output

Hyperlinks in the downloaded Word document are genuine clickable links (proper `w:hyperlink` elements with registered relationship IDs) — not just underlined blue text. They open in a browser when clicked inside Word.

---

## Project Structure

```
cv-suite/
├── cv_suite.py          # Main application (single-file)
├── requirements.txt     # Python dependencies
└── README.md            # This file
```

---

## Tips for Best Results

- **Use Word (.docx) format** for your CV upload — hyperlinks are extracted with 100% accuracy
- **Paste the full job description** — include the responsibilities, requirements, and company overview sections
- **Add metrics to your original CV** — the AI can only amplify numbers that already exist; it won't invent them
- **Check the Analysis tab** — the skill gap and interview talking points sections are useful for interview preparation
- **Edit the cover letter** — the editable text area in the Cover Letter tab lets you refine before downloading

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `ModuleNotFoundError: No module named 'fitz'` | Run `pip install PyMuPDF` |
| PDF links not extracted | Check that the PDF has proper link annotations (not just typed URLs). Open the PDF and try clicking a link — if it's not clickable in your PDF viewer, it cannot be extracted. |
| AI generation returns empty string | Check your API key is valid and has quota remaining |
| Word document won't open | Ensure `python-docx` and `lxml` are installed correctly |
| Cover letter is too long / too short | Use the Regenerate button in the Cover Letter tab for a fresh attempt |

---

## License

MIT License — free to use, modify, and distribute.
