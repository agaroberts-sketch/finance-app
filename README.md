# Financial Sample Explorer

A first Streamlit dashboard for exploring the `Financial Sample.xlsx` workbook.

## Run locally

1. Install Python 3.10 or newer.
2. Open PowerShell in this folder.
3. Create and activate a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

4. Install dependencies:

   ```powershell
   python -m pip install -r requirements.txt
   ```

5. Start the app:

   ```powershell
   streamlit run app.py
   ```

The app automatically opens `Financial Sample.xlsx` from your Downloads folder. You can also upload another workbook from the sidebar.

## Enable AI chat

Set an OpenAI API key before starting Streamlit. The key is read from the environment and is never stored in the app source code.

```powershell
$env:OPENAI_API_KEY = "your-api-key"
streamlit run app.py
```

The AI chat uses worksheet metadata, a small sample of rows, and numeric summaries as context. Avoid using it with sensitive data unless your organization's AI data policy allows it.