# Jo Office for Android

A small Android shell around the Jo Office web app. It opens **https://karthik1988ai-web.github.io/Jo-Office/**,
so every update to `index.html` reaches the phone without reinstalling. It gives the page the phone's
speech recognizer, text-to-speech, notifications and full screen (see `phone/.../NativeBridge.kt`;
the page side is the "Jo inside the Android app" script in `index.html`).

It is a separate app from **Jo** (`com.karthik.jo.office`, label "Jo Office"), so both can be installed.

## Get the APK
Every push that changes `android/` runs **Actions → Build APK** (about 5 minutes); it can also be started
by hand (Run workflow). Open the run and download **Jo-Office-apk**, unzip it to get `Jo-Office.apk`,
send it to the phone, tap it and allow "install unknown apps".

## Keep settings across updates (optional)
Android installs a new build over the old one only if both are signed with the same key. Create one
keystore (`keytool -genkeypair -keystore jo-office.keystore -alias jo-office -keyalg RSA -validity 10000`),
then add repository secrets **JO_KEYSTORE_BASE64** (`base64 -w0 jo-office.keystore`) and
**JO_KEYSTORE_PASSWORD**. Never commit the keystore.
