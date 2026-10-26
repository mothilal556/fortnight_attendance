# FortNight Attendance Sheets Generator

A Flask application that reads a student Excel workbook and generates a formatted Excel workbook with one worksheet per subject, matching the uploaded workbook's tabular attendance layout.

The home page has two tabs: **Fortnight Attendance** and **Consolidated Report**. The consolidated tool is bundled under `consolidated/` and runs on the same Flask site at `/consolidated/`.

## Features
- Form fields: Degree, CollegeCode, BranchCode, Section, Regulation, Year, Semester.
- Upload `.xlsx` or `.xlsm` Excel file (maximum 20 MB).
- Validates the seven required column names on each non-empty worksheet.
- Reads multiple worksheets and creates one output worksheet per subject.
- Collects all non-empty HTNOs and distinct fortnight/date ranges from the upload, then creates every HTNO/fortnight combination on each subject worksheet.
- Preserves academic details, class counts, and fortnight dates when those columns are present; uses the form values for missing academic details.
- Removes duplicate student rows within the same subject and fortnight date range.
- Applies column headings, filters, `dd-mm-yyyy` date formatting, and print-friendly page settings.
- Processes Fortnight Attendance uploads in memory; no permanent upload is stored.
- Consolidates multi-sheet attendance or marks workbooks into a summary with totals and percentages in the **Consolidated Report** tab.

## Required Excel columns
`HTNO`, `SUBJECTCODE`, `SUBJECTNAME`, `SUBJECTTYPE`, `FORTNIGHT`, `FORTNIGHT_START_DATE`, `FORTNIGHT_END_DATE`

The column names are case-insensitive; leading/trailing spaces are removed. `From_date` and `To_date` are accepted as aliases for `FORTNIGHT_START_DATE` and `FORTNIGHT_END_DATE`. Academic details (`DEGREE`, `COLLEGECODE`, `BRANCHCODE`, `SECTION`, `REGULATION`, `YEAR`, and `SEMESTER`) plus `CLASSESCONDUCTED` and `CLASSESATTENDED` are preserved when present, but are not required. The Fortnight Attendance tab accepts `.xlsx`/`.xlsm`; the Consolidated Report tab accepts `.xlsx`/`.xls`.

## Run on Windows

1. Install Python 3.10 or newer.
2. Extract this folder and open Command Prompt/PowerShell inside it.
3. Create a virtual environment:

   ```powershell
   py -m venv .venv
   .\.venv\Scripts\activate
   ```

4. Install packages:

   ```powershell
   python -m pip install --upgrade pip
   pip install -r requirements.txt
   ```

5. Start the app:

   ```powershell
   python app.py
   ```

6. Open `http://127.0.0.1:5000` in your browser.

## Output behavior
The downloaded workbook contains one worksheet per subject. Each worksheet has the columns `DEGREE`, `COLLEGECODE`, `HTNO`, `BRANCHCODE`, `SECTION`, `REGULATION`, `YEAR`, `SEMESTER`, `SUBJECTCODE`, `SUBJECTNAME`, `SUBJECTTYPE`, `FORTNIGHT`, `CLASSESCONDUCTED`, `CLASSESATTENDED`, `FORTNIGHT_START_DATE`, and `FORTNIGHT_END_DATE`.

## Production notes
- Vercel deployment uses the Python runtime and Vercel Functions. When configured, uploads use a private Vercel Blob store; temporary report files are only kept during a single function invocation.
- Connect a **private** Vercel Blob store to the project and ensure Vercel provides `BLOB_READ_WRITE_TOKEN` for direct uploads up to 20 MB. Without Blob, the app remains available and accepts uploads up to 4 MB to stay below Vercel's request-body limit.
- Set `FLASK_SECRET_KEY` to a long, random secret in the Vercel project environment so Flask's signed flash-message cookies remain consistent across function instances.
- Import the project from its Git repository in Vercel and deploy from the repository root. Vercel installs Python dependencies from `requirements.txt` and the upload handler dependency from `package.json`.
- The main attendance workbook is returned immediately after processing. Consolidated reports are stored privately in Blob storage and removed when downloaded.
- The application has no authentication. Anyone can use the tools and request upload tokens, so the public Blob upload endpoint may incur storage or processing costs. Keep the upload size and Excel file-type restrictions in place, monitor usage, and use Vercel access controls if the app or student data should not be public.

### Deploy on Vercel
1. Push this project to a GitHub repository, then import that repository into Vercel with the project root as the root directory.
2. Optionally, in Vercel Storage, create a **private** Blob store and connect it to this project. Confirm `BLOB_READ_WRITE_TOKEN` is available to enable uploads up to 20 MB.
3. Add `FLASK_SECRET_KEY` for the Production environment (and Preview if needed). Generate a long random value.
4. Deploy or redeploy. Without Blob configured, the app still runs with the 4 MB upload limit.

To create a random Flask secret locally, run:

```powershell
python -c "import secrets; print(secrets.token_urlsafe(48))"
```
