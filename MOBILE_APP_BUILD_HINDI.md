# MoneyWise Android Mobile App Wrapper

Is folder me Android Studio source diya hai. Ye existing Google Apps Script Web App ko mobile app ke andar WebView me open karta hai aur camera/GPS/file chooser support deta hai.

## Important
- Ye source project hai, prebuilt APK nahi.
- Default Web App URL `app/src/main/res/values/strings.xml` me diya hai. Agar Web App deployment URL badle to wahi value update karein.
- Backend Google Apps Script par hi rahega; Android app usi live system ka mobile client hai.

## APK build steps
1. Android Studio install karein.
2. `mobile-app` folder ko **Open Project** karein.
3. Gradle sync complete hone dein.
4. Phone USB debugging se connect karke Run karein, ya emulator use karein.
5. Release APK ke liye: **Build -> Generate Signed App Bundle or APK -> APK**.
6. First launch par Camera aur Location permission Allow karein.

## Field test
- Consultant ID login
- CFL Name/BCC search
- Camera se multiple photos
- GPS capture
- Signature
- CFL/Session submit
- PDF/Excel open/download
- Follow-up tab
- Logout/Login privacy
