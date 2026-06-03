# Native Build — Signing & Icons Setup

One-time setup so the "Build Native Apps" workflow produces store-ready binaries.

## 1. Android signing (do this once)

The CI no longer generates a throwaway keystore each run (that breaks Google Play
updates, which require the *same* key every time). Instead it restores a persistent
keystore you create once and store as a secret.

**Generate the keystore** (on any machine with Java/`keytool`):

```
keytool -genkey -v -keystore framemedicine-release.jks \
  -keyalg RSA -keysize 2048 -validity 10000 \
  -alias framemedicine \
  -dname "CN=Tom, OU=Development, O=PinePeakMed LLC, L=Jacksonville, ST=FL, C=US"
```

Use a strong password and **save the `.jks` file and password somewhere safe and
permanent** (password manager). If you lose them, you can never update the app on
Play again.

**Base64-encode it** and copy the output:

- Windows PowerShell: `[Convert]::ToBase64String([IO.File]::ReadAllBytes("framemedicine-release.jks")) | Set-Clipboard`
- macOS/Linux: `base64 -w0 framemedicine-release.jks | pbcopy`

**Add two GitHub repo secrets** (Settings -> Secrets and variables -> Actions):

- `ANDROID_KEYSTORE_BASE64` = the base64 string above
- `ANDROID_KEYSTORE_PASSWORD` = the keystore password (already set — confirm it matches the password you used above)

## 2. App icon (do this once)

Add a square PNG at `native/resources/icon.png`, **1024x1024, no transparency,
no rounded corners** (the stores round it). Optionally add
`native/resources/splash.png` at 2732x2732 (logo centered on the #080808 background).

CI runs `@capacitor/assets` to generate every required icon/splash size for both
platforms. Without `icon.png`, builds fall back to the default Capacitor icon.

## 3. iOS (blocked until Apple enrollment)

The iOS job needs three more repo secrets, available once your Apple Developer
Program enrollment is complete:

- `IOS_TEAM_ID`
- `IOS_CODE_SIGN_IDENTITY`
- `IOS_PROVISIONING_PROFILE`

## 4. Run a build

Actions tab -> "Build Native Apps" -> Run workflow -> choose `android` (or `both`
once iOS secrets exist). The signed `.aab` is uploaded as a build artifact.
