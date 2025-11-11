
# 📊 JDE Financial App

This project contains a **Streamlit-based financial dashboard** to visualize and interact with budget and actuals (BvA) data from the JDE system. The app loads financial Excel reports and displays summary and detailed views for decision-making.

---

## 🚀 How to Run the App

1. **Install dependencies**

```bash
pip install -r requirements.txt
```

2. **Run the app**

```bash
streamlit run jde_app.py
```

3. **Navigate to**

Open your browser and go to: [http://localhost:8501](http://localhost:8501)

---

## 📁 Project Structure

```
JDE FINANCIAL/
├── .streamlit/                        # Streamlit configuration folder (theme, layout, etc.)
│
├── .venv-jde/                         # Virtual environment (do not track in Git)
│
├── 02-Aug FY26 QRA BvAs-copy.xlsx     # Excel source file for BvA data (official or clean copy)
├── 02-Aug FY26 QRA BvAs-copy - Copy2.xlsx
│                                      # Duplicate or working copy of the above file
│
├── FY 2026 - JDE Details QRA (YTD)... # Detailed financial breakdown (YTD) required by the app
│                                      # Possibly includes line item data by cost center or department
│
├── JDE Financial Report.docx          # Internal draft in Spanish with file structure & team context
│
├── jde_app.py                         # Main Streamlit application script
│                                      # Loads Excel data and renders dashboard
│
└── requirements.txt                   # Python dependencies for the app
```

---

## 📄 File Descriptions

| File | Description |
|------|-------------|
| `jde_app.py` | Main script to launch the Streamlit app. Loads and processes Excel inputs. |
| `requirements.txt` | List of required Python packages (Streamlit, pandas, openpyxl, etc.). |
| `02-Aug FY26 QRA BvAs-copy.xlsx` | Main Excel source for budget vs actuals (QRA). Likely the official version. |
| `02-Aug FY26 QRA BvAs-copy - Copy2.xlsx` | Backup or test version of the main BvA file. |
| `FY 2026 - JDE Details QRA (YTD)...xlsx` | Detailed YTD view of financials, possibly per business unit or account. |
| `JDE Financial Report.docx` | Draft document (Spanish) to understand file structure, usage needs, and team context. |

---

## ✅ Notes

- Ensure Excel files are kept up to date with the latest QRA inputs from the Finance team.
- You may want to rename or archive working copies (e.g., `Copy2`) to avoid confusion.
- Consider including logic in `jde_app.py` to validate and preprocess Excel sheets before rendering.
