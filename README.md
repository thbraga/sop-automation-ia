# sop-automation-ia

Automation of Standard Operating Procedures (SOPs) using GPT-4o and Google Sheets.

## 🚀 How it works

This project automates the standardization of procedural documents by:
1. Downloading .docx files from a Google Form response sheet.
2. Extracting and analyzing the text + embedded images.
3. Sending the content to GPT-4o for structured transformation.
4. Writing the structured result into the same Google Sheet.
5. Generating a final formatted .docx file with paragraphs, bullet lists and inserted images.

## 📂 Folder Structure

```
sop-automation-ia/
│
├── aut_pop_ia.py          # Main automation script
├── requirements.txt       # Python dependencies
├── .env.example           # Environment variable example
├── .gitignore             # Files to ignore in git
└── README.md              # Project overview
```

## ⚙️ Setup

1. Rename `.env.example` to `.env` and add your OpenAI key.
2. Ensure your Google service account has access to the response spreadsheet.
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Run the script:

```bash
python aut_pop_ia.py
```

---

Developed by Thainá Braga – 2025
