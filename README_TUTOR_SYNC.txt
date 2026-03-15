Tutor Verification Sync (Excel → Website)

What you edit:
- In Excel sheet TUTOR_REGISTER:
  Tutor Name | Subject Specialty | Highest Qualification | University/Institution | Status (Active/Inactive)
  (Optional but recommended): Verification Code | Verification Status

How to export tutors.json (Windows):
1) Put your Excel workbook in this same folder.
2) Double-click RUN_EXPORT_TUTORS.bat
   - It generates tutors.json
   - It also saves an Excel copy with codes: *_WITH_CODES.xlsx (first run only)

How to publish:
- Commit tutors.json to GitHub in the SAME folder as index.html
- Verify in browser:
  https://<username>.github.io/<repo>/tutors.json
