# DOCX Language Detector

Detects non-English text in Word documents, highlights it in yellow, and tags it with the detected language and confidence score.

## Deploy to Render (free, 5 minutes)

1. Push this folder to a GitHub repository
2. Go to https://render.com and sign up (free)
3. Click **New → Web Service**
4. Connect your GitHub repo
5. Render auto-detects Python. Set:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Instance Type:** Free
6. Click **Deploy**
7. Your URL will be `https://your-app-name.onrender.com`

Share that URL with your colleagues — no install needed on their end.

## Run locally

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## How it works

- Extracts all text from the uploaded .docx (body paragraphs + table cells)
- Runs each segment through the Lingua language detector
- Skips: dates, IDs, codes, version numbers, short strings
- Highlights detected non-English runs in yellow
- Appends a [Language | confidence%] tag inline
- Adds a summary page at the end of the document
- Returns the modified .docx for download

## Notes

- No files are stored on the server
- Max file size: 20MB
- Supported: any .docx file
- Detection threshold: 55% confidence minimum
