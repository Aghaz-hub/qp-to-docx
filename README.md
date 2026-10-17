# Question Paper → Word Document Converter

Upload **old scanned question papers** (images or PDF) and get a clean, editable **Word (.docx)** file.

Built for Manav Rachna University / similar university question-paper formats.

---

## Features

- Accepts **JPG / PNG / WEBP** images and **PDF** (single-page or multi-page)
- Runs **OCR** on every page
- Builds a structured Word document with:
  - University / School / Department header
  - Exam details table (Semester, Date, Course Code, Max Marks, Duration, etc.)
  - Full question text extracted from the scan
- Download ready `.docx` for proof-reading and reuse

---

## How to use

1. Take clear photos or scans of the question paper pages  
2. Upload images and/or a multi-page PDF  
3. Click **Convert to Word Document**  
4. Download the `.docx` and open in MS Word / Google Docs  
5. Proof-read (OCR is never 100% perfect on old scans)

---

## Deploy on Streamlit Community Cloud

1. Create a new GitHub repository  
2. Upload all files from this folder  
3. Go to [https://share.streamlit.io](https://share.streamlit.io) → **New app**  
4. Select the repo → main file `app.py` → Deploy  

`packages.txt` installs Tesseract OCR and Poppler so PDF + OCR work on Streamlit Cloud.

---

## Run locally

```bash
# System packages (Ubuntu/Debian)
sudo apt update && sudo apt install -y tesseract-ocr tesseract-ocr-eng poppler-utils

# Python packages
pip install -r requirements.txt

# Launch
streamlit run app.py
```

---

## Files

| File | Purpose |
|------|---------|
| `app.py` | Main Streamlit application |
| `requirements.txt` | Python dependencies |
| `packages.txt` | System packages for Streamlit Cloud |
| `README.md` | This file |
| `sample_QP.pdf` | Sample question paper (optional) |

---

**Note:** Always proof-read the generated Word file. Faded ink, handwriting, and mathematical symbols are the most common OCR error sources.
