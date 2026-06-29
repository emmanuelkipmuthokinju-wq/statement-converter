# 📊 Customer Statement → Excel Converter

A free web app that converts customer statement PDFs to formatted Excel files.

## Supported Formats
- ✅ Savannah Cement statements
- ✅ National Cement statements

## How to Deploy (FREE on Streamlit Cloud)

### Step 1 — Put the files on GitHub
1. Go to https://github.com and create a free account
2. Click **New repository** → name it `statement-converter` → Public → Create
3. Upload these two files:
   - `app.py`
   - `requirements.txt`

### Step 2 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io and sign in with GitHub
2. Click **New app**
3. Select your `statement-converter` repository
4. Main file: `app.py`
5. Click **Deploy** — done! 🎉

Your app will be live at:
`https://your-username-statement-converter-app-xxxx.streamlit.app`

**It's completely free and shareable with anyone.**

## Run Locally (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```
