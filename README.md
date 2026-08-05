# Tea with Nessie — Mail-Merge Invitation Automation

A Python mail-merge system built for a course exercise in **DGTR 3720 (Data Transformation)** at the University of Lethbridge. It generates personalized event invitations for 24 guests split across three venues, pulling guest data from an Excel workbook and merging it into a Word template.

## What it does

- Reads guest records from `TeaWithNessie_GuestData_NiyatThapa.xlsx` using `openpyxl`
- Validates that every required column is present and no required field is blank
- Opens a fresh copy of `TeaWithNessie_Invitation_Template_NiyatThapa.docx` for each guest using `python-docx`
- Replaces merge fields (e.g. `<<first_name>>`, `<<venue_name>>`, `<<tea_time>>`) with that guest's data
- Checks the finished letter for any unfilled `<<field>>` placeholders before saving
- Produces one personalized `.docx` invitation per guest, plus a summary report

## Result

24 invitations generated, 8 guests per venue across 3 venues, 0 unfilled placeholders. See `merge_summary.txt` for the full run output.

## Run it yourself

```bash
pip install python-docx openpyxl
python mail_merge.py
```

Generated invitations are written to an `output/` folder.

## Tech

Python · openpyxl · python-docx · regex-based field validation

---
Built by Niyat Thapa. AI tools (ChatGPT, Claude) were used to help debug and improve parts of the script; the plan, data structure, and final work were reviewed and are owned by the author.
