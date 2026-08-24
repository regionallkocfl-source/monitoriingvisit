# MoneyWise CFL & Session Monitoring v3.9 FINAL

## Is version me kya final hai
- AM/CBO fast login -> direct data entry.
- CBO CFL access: `Cbo Master` District -> `CFL Master` same District ke CFLs.
- CFL Name/BCC Code searchable selector.
- CFL & Session forms always-open questions; Hindi/English switch.
- Visit Block / Adjacent Block dropdown = selected CFL ka Base Block + Adjacent Block 1 + Adjacent Block 2.
- Previous unresolved observations/action points carry-forward until verified completion.
- Dedicated Follow-up tab for AM/CBO/Admin.
- Admin Monitoring & Reminders + manual/daily reminder support.
- PDF A4-safe layout, centered blue headings, actual geotagged photos embedded.
- Separate Excel export.
- One consolidated admin email with PDF + Excel attachments.
- Consultant ID/Name-wise Drive folder structure.
- Login privacy: previous ID/PIN not retained; login UI has no PIN field.
- Mobile top navigation + Logout text button.

## v3.9 Admin CFL / Block rule
Admin -> CFL / Block:
1. CFL/BCC select karein.
2. BCC Code, Phase, Bank, State, District, Base Block auto-fetch aur read-only rahenge.
3. Sirf `Adjacent Block 1` aur `Adjacent Block 2` edit karein.
4. `Save Adjacent Blocks` dabayein.
5. Backend sirf selected CFL ke adjacent-block columns update karta hai; baaki master data overwrite nahi hota.
6. Saved blocks AM/CBO ke CFL Entry aur Session `Visit Block / Adjacent Block` dropdown me automatically aate hain.

## Apps Script update
Existing project folder me final files replace karne ke baad:
```bash
clasp push
```
Then Apps Script -> Deploy -> Manage deployments -> Edit -> New version -> Deploy.
Is v3.9 block change ke liye `setupMoneyWise()` compulsory nahi hai. Agar full package ko fresh project me install kar rahe hain to setup ek baar run karein.

## GitHub deployment
Detailed guide: `GITHUB_DEPLOY_HINDI.md`.
GitHub Pages par project direct host na karein. GitHub source/version control rakhega aur GitHub Actions `clasp` ke through Apps Script ko deploy karega.

## Android mobile app
`mobile-app` folder me Android Studio WebView wrapper source diya hai.
Build guide: `MOBILE_APP_BUILD_HINDI.md`.
Backend wahi Apps Script rahega, isliye web aur app dono same live data use karenge.
