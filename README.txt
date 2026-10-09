AM/NS SAFETY REGISTER — RENDER BACKEND + FRONTEND

WHAT THIS PACKAGE DOES
- Serves the existing dashboard frontend and backend from one Render web service.
- Reads the Confined Space and Gas Hazardous Area Excel registers.
- Generates data.json / gas_data.json from the Excel files and serves them through /api/data and /api/gas-data.
- Includes a protected upload page at /admin so an authorised admin can upload a replacement workbook and refresh the dashboard without IT or a GitHub commit.
- Stores the authoritative Excel files on Render's persistent disk at /var/data.

IMPORTANT RENDER REQUIREMENT
The included render.yaml uses a paid Starter web service and a 1 GB persistent disk. Render's free web services do not support persistent disks, so don't deploy this configuration as a free service if you need Excel uploads to survive restarts/redeploys.

DEPLOY STEPS
1. Extract this ZIP on your computer.
2. Create a NEW GitHub repository for this Render version, e.g. AMNSCS-Render.
3. Upload ALL extracted contents to the repository root. Do not upload the ZIP as a single file. Keep app.py, render.yaml, requirements.txt, index.html, tools/ and both Excel files at the root/folder structure shown.
4. In Render, choose New + > Blueprint and connect the GitHub repository. Render reads render.yaml and creates the web service and persistent disk.
5. In Render service Environment, set ADMIN_PASSWORD to a long unique password. Do not share it or put it in GitHub. Save changes/redeploy if asked.
6. Wait until the service is Live. Open https://YOUR-SERVICE.onrender.com/api/health. It should return status "ok" plus record counts.
7. Open https://YOUR-SERVICE.onrender.com/ for the dashboard.
8. To update data later, open https://YOUR-SERVICE.onrender.com/admin, enter ADMIN_PASSWORD, choose the updated workbook(s), and press Validate, Save & Refresh Dashboard. The dashboard reads the newly generated data on reload.

EXCEL RULES
- Confined Space workbook must be named/uploaded as .xlsx and contain the sheet "Confined Space Location List" with the same column layout as the supplied template.
- Gas workbook must contain the sheet "Hazard Assessment Matrix" with the same column headings as the supplied template.
- You can update either workbook independently. Select at least one file.
- Do not rename the required worksheets or change the column positions/headings without modifying the parser.
- The package's current confined-space workbook has 114 identified location records (not 116/121); this count is calculated from the supplied Excel. Gas workbook currently has 16 areas.

HOW IT WORKS
Browser -> Render Flask/Gunicorn backend -> Excel files on persistent disk -> openpyxl generators -> JSON API -> dashboard. Uploading Excel validates it, saves it to persistent storage, rebuilds the JSON data, then the browser reads the new API data. The source Excel itself is not publicly downloadable from the dashboard routes.

SECURITY / COMPANY DATA
Use only if your company permits this register to be hosted on Render. The dashboard URL is public unless you add access controls; the /admin upload page is protected by ADMIN_PASSWORD, but it is not a complete enterprise identity system. For production use, request company IT/security approval, use a strong password, and consider company SSO/network restrictions. Never upload confidential data to a personal cloud without approval.

TROUBLESHOOTING
- /api/health returns status error: check Render logs; verify both source workbooks and their required sheet names.
- Upload says incorrect password: reset ADMIN_PASSWORD in Render Environment.
- New data doesn't appear: open /api/health, then hard refresh the dashboard. The API checks source-file timestamps and rebuilds if needed.
- Render service is sleeping: first request on a free service may take time; this package's persistent disk requires a paid service plan.
