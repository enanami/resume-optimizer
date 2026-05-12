# Resume ATS Optimizer

A Streamlit app that analyzes your resume against a job description using Claude AI, identifies keyword gaps, suggests ATS-optimized rewrites, and generates a tailored cover letter — all exportable as Word files.

## Features

### Resume Analysis
- **Resume upload** — supports PDF and DOCX (bold formatting preserved from DOCX)
- **ATS score breakdown** — overall score plus keyword match, quantification, formatting, and relevance sub-scores
- **Strengths & weaknesses** — what's working and what's hurting your chances for this specific role
- **Strategic advice** — actionable tips tailored to your background and the role

### Suggested Rewrites
- **Professional summary rewrite** — tailored to the role with high-priority keywords
- **Skills section update** — integrates missing keywords into your existing skills format
- **Bullet rewrites** — side-by-side editable before/after with priority ranking and reasoning
- **Apply All** — pushes all accepted rewrites into the editor with one click

### Keyword Gaps & Improvements
- **Keyword gap analysis** — missing technical skills, soft skills, tools, certifications, and domain knowledge
- **Priority improvements** — ranked list of the highest-impact changes

### Resume Preview & Export
- **Live preview** — formatted resume preview that updates as you edit (Arial font, right-aligned dates, bold text via `**word**`)
- **Edit Source tab** — make manual edits directly in the browser
- **Re-score** — run a fresh ATS score on your edited resume to measure improvement
- **Word export** — compact, ATS-safe `.docx` with proper formatting (one-page optimized)

### Cover Letter
- **AI-generated cover letter** — 3–4 paragraphs tailored to the job description, using achievements from your resume
- **Preview & edit tabs** — review and customize before downloading
- **Word export** — clean `.docx` in Arial 11pt

## Setup

### 1. Clone and install dependencies

```bash
git clone <repo-url>
cd resume-optimizer
python3 -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Add your Anthropic API key

```bash
cp .env.example .env
```

Edit `.env` and replace `your-api-key-here` with your key from [console.anthropic.com](https://console.anthropic.com).

```
ANTHROPIC_API_KEY=sk-ant-...
```

### 3. Run the app

```bash
streamlit run app.py
```

The app opens at `http://localhost:8501`.

## Usage

1. **Upload** your resume (PDF or DOCX).
2. **Paste** the full job description including responsibilities and qualifications.
3. Click **Analyze Resume**.
4. Review scores, strengths/weaknesses, and strategic advice.
5. Edit the **After** fields in Suggested Rewrites, then click **Apply All Rewrites to Editor**.
6. Check the live **Preview** tab — use `**word**` syntax to bold specific text.
7. Click **Score Edited Resume** to measure your improvement.
8. **Download** the optimized resume as a Word file.
9. Click **Generate Cover Letter** for a tailored cover letter, then download it.

## Requirements

- Python 3.9+
- Anthropic API key (Claude Sonnet)
