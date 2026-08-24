# GitHub + Apps Script Deployment (Hindi)

## Important
GitHub Pages par ye project direct run nahi hoga, kyunki `.gs` backend aur `google.script.run` sirf Google Apps Script HTML Service me chalte hain. GitHub ko **source control + automatic Apps Script deployment** ke liye use karein.

## One-time local setup
1. GitHub par private repository banayein, e.g. `moneywise-cfl-monitoring`.
2. Is project folder me CMD/PowerShell kholkar:
   ```bash
   git init
   git branch -M main
   git add .
   git commit -m "MoneyWise v3.9 final"
   git remote add origin https://github.com/YOUR-USER/moneywise-cfl-monitoring.git
   git push -u origin main
   ```
3. Apni working `.clasp.json` ko GitHub me commit **na** karein. `.gitignore` already isse block karta hai.

## GitHub Actions secrets
GitHub repo -> Settings -> Secrets and variables -> Actions -> New repository secret.

Create:
- `CLASPRC_JSON`: aapke Windows user profile me `~/.clasprc.json` ka poora content. Ye sensitive OAuth credential hai.
- `CLASP_JSON`: aapke project folder ki `.clasp.json` ka poora content.
- `APPS_SCRIPT_DEPLOYMENT_ID`: existing Web App deployment ID (optional but recommended). `clasp deployments` se dekh sakte hain.

## Automatic deployment
Ab `main` branch me `.gs`, `.html`, ya `appsscript.json` change push hote hi GitHub Action:
1. Apps Script par `clasp push --force` karega.
2. Naya immutable Apps Script version banayega.
3. Agar `APPS_SCRIPT_DEPLOYMENT_ID` secret diya hai to same `/exec` Web App ko new version par redeploy karega.

## Deployment ID kaise milega
Local project folder me:
```bash
clasp deployments
```
Web app deployment ke saamne jo deployment ID hai usko GitHub secret me paste karein.

## Security
- Repository ko private rakhein.
- `.clasprc.json` kabhi repository me commit na karein.
- GitHub Secrets me hi OAuth credential rakhein.
- GitHub deployment pipeline use karne ke baad Apps Script editor me direct code editing avoid karein; GitHub ko source of truth rakhein.
