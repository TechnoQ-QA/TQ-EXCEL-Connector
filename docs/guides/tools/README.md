# Building the Word guides

The installation guide and the user guide are generated, so they can be rebuilt after a change.

1. **Screenshots** (Edge, headless, from the installed connector at https://localhost:3000). The account, user
   name and company are replaced with example values before each picture.
   `pip install --target <lib> playwright pymupdf`, then
   `python shots.py <images folder> install` and `python shots.py <images folder> user`.
2. **Word files**: in a folder with `npm install docx@9`, run
   `node install.js <images folder> "..\TQ Connector - Installation Guide.docx"` and
   `node user.js <images folder> "..\TQ Connector - User Guide.docx"`.
3. **Check**: `powershell -File render.ps1 -Docx <docx> -Pdf <pdf>` opens the file in Word (hidden), updates the
   table of contents and saves a PDF to look at.

To add screenshots of Excel itself (a report sheet, a dashboard, the ribbon, the right-click menu), save them as
PNG in the images folder and add `one('name', 'caption')` lines in `user.js`.
