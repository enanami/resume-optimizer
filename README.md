# Resume ATS Optimizer

A Streamlit app that analyzes your resume against a job description using Claude AI, identifies keyword gaps, and helps you produce an ATS-optimized Word document.

## Features

- **Resume upload** — supports PDF and DOCX
- **ATS score breakdown** — overall score plus keyword match, quantification, formatting, and relevance sub-scores
- **Keyword gap analysis** — missing technical skills, soft skills, tools, certifications, and domain knowledge grouped by category
- **Priority improvements** — ranked list of the highest-impact changes
- **Bullet rewrites** — side-by-side before/after rewrites with reasoning
- **Optimized professional summary** — tailored to the specific role
- **Inline resume editor** — apply suggestions directly in the browser
- **Word export** — download the edited resume as a `.docx` file

## Setup

### 1. Clone and install dependencies

```bash
git clone <repo-url>
cd resume-optimizer
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Add your Anthropic API key

```bash
cp .env.example .env
```

Edit `.env` and replace `your-api-key-here` with your key from [console.anthropic.com](https://console.anthropic.com).

### 3. Run the app

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Usage

1. Upload your resume (PDF or DOCX) using the file uploader.
2. Paste the full job description — include responsibilities, required qualifications, and preferred qualifications.
3. Click **Analyze Resume**.
4. Review the ATS score, keyword gaps, suggested rewrites, and optimized summary.
5. Edit your resume in the text editor at the bottom of the page.
6. Click **Download as Word (.docx)** to export the updated resume.

## Requirements

- Python 3.9+
- Anthropic API key
