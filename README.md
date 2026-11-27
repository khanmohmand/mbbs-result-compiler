# MBBS Result Compiler

A Streamlit web app that compiles **module-wise MBBS results** and exports a ready-to-print Excel workbook
with live formulas, a graphical dashboard and individual student reports.

## What it does

- Takes 1Best, Viva, SPE and CA marks for one module at a time (typed in or uploaded from Excel/CSV).
- Scales each component to its weightage: `marks obtained / max marks x weightage`
  (default MBBS weightage: 1Best 40, Viva 10, SPE 20, CA 30 = 100).
- Grand Total is rounded to a whole number. A student passes with the Pass Marks (default 55) or more.
- Shows how many marks are still required to pass, and the Pass / Fail status.
- Flags students whose marks in any component are below the **Minimum** (Min fills with 50% of Max and can be edited)
  and writes a professional improvement remark for them.
- Supports 1st Prof and 2nd Prof modules (module names are built in).

## Excel output (3 sheets)

| Sheet | Contents |
|---|---|
| **Result** | Roll No., 1Best, Viva, SPE with their adjusted marks, Adjusted Total, CA, Grand Total, Marks Required to Pass, Status. All calculations are live Excel formulas, so every figure can be rechecked. |
| **Dashboard** | KPI cards (students, appeared, absent, pass, fail, pass %, average, highest, lowest, needs improvement) and a Grand Total distribution chart. |
| **Student Reports** | Pick a Roll No. from the drop-down to see that student's graph, numbers and summary text. Below it, a table of all students with improvement remarks for those below the minimum (filterable). |

## Input format

Excel or CSV with these columns (names are matched loosely):

| Roll No. | 1Best | Viva | SPE | CA |
|---|---|---|---|---|

- Leave all exam marks empty to mark a student **Absent**.
- CA can be in the marks sheet, or in a separate CA sheet (Roll No., CA) uploaded in the app. The CA sheet fills CA where the marks sheet has none.
- A blank template can be downloaded from the app.

## Run locally

```bash
pip install -r requirements.txt
streamlit run mbbs_app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository and upload `mbbs_app.py`, `requirements.txt` and this `README.md`.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **Create app**, choose the repository and branch, and set **Main file path** to `mbbs_app.py`.
4. Click **Deploy**. The app gets a public URL; it redeploys automatically on every push to the repository.

If the repository also contains another app file (for example the BDS `app.py`), deploy it as a separate app
from the same repository by choosing its file as the main file path.

## Privacy note

Student marks are processed in memory only and are not stored by the app. If the repository or the app URL is public,
do not commit real student data files to GitHub. Keep the marks sheets on your own computer.
