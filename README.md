# Villa Expense Ledger app

A cash book for **Utilities, Visa, Tickets and Villa cash**, with Credit (money in),
Debit (money out) and Balance in AED. One app runs on iPhone, Android and Mac. Every
entry is saved as a row in a Google Sheet you own, so the sheet is also a readable record.

```
 iPhone / Mac (installed web app)  ─┐
 Android (APK)                      ─┼──►  Apps Script link (/exec)  ──►  Google Sheet "Entries" tab
 Any browser                        ─┘        checks the PIN
```

| Folder | What it is |
|---|---|
| `web/` | The app (one page plus icons, manifest and offline support) |
| `apps-script/Code.gs` | The small backend that reads and writes the Google Sheet |
| `android/` | Wrapper that turns `web/` into an Android APK (built by GitHub Actions) |

## 1. Connect the Google Sheet (one time, about 5 minutes)

1. Open the sheet **Villa Expense Ledger** in Google Drive. It already has an `Entries` tab with headers.
2. Click **Extensions → Apps Script**.
3. Delete what is in `Code.gs`, paste in everything from [`apps-script/Code.gs`](apps-script/Code.gs), and press **Save**.
4. At the top of the file, set the PINs:
   - `EDIT_PIN`: your PIN for adding and changing entries.
   - `VIEW_PIN`: a second PIN for your boss to look only. Leave `''` to turn it off.
5. Click **Deploy → New deployment**, choose the gear icon → **Web app**, then set:
   - *Execute as*: **Me**
   - *Who has access*: **Anyone**
6. Click **Deploy** and allow access when Google asks. You may need to click *Advanced → Go to project (unsafe)* the first time, because the script is your own and not published.
7. Copy the **Web app URL**. It ends in `/exec`. This is the app's *connection link*.

"Anyone" means the link answers requests from the app without a Google sign-in. Nothing
can be read or changed without a correct PIN. Keep the link and PINs private.

If you change `Code.gs` later, use **Deploy → Manage deployments → Edit → New version**,
so the same link keeps working.

## 2. Install the app

Open the app link, paste the connection link, enter your PIN and press **Connect**.
Each device remembers it.

**iPhone / iPad:** open the app link in **Safari** → Share button → **Add to Home Screen**.

**Mac:** open the app link in **Safari** → **File → Add to Dock** (macOS Sonoma or later).
In Chrome, use the install icon in the address bar.

**Android:** download `villa-ledger.apk` from
[Releases → Villa Ledger app (Android)](https://github.com/officepaddy1-rgb/villa-ledger/releases/tag/villa-ledger-latest), open it, and allow "Install unknown apps" when asked.
Or open the app link in Chrome → menu → **Add to Home screen**.

### App link

**https://officepaddy1-rgb.github.io/villa-ledger/**

The workflow `.github/workflows/web.yml` publishes `web/` with GitHub Pages whenever
`main` changes. If the link shows "404", turn Pages on once:
**Settings → Pages → Build and deployment → Source: GitHub Actions**, then re-run the
"Web app (GitHub Pages)" workflow from the Actions tab.

The Android APK doesn't need the link: the app is built into it. Each push to `main`
rebuilds it and replaces `villa-ledger.apk` on the
[Villa Ledger app (Android) release](https://github.com/officepaddy1-rgb/villa-ledger/releases/tag/villa-ledger-latest).

This repository is public, so it holds only the app's code. There are no entries,
PINs or connection links in it. Those live in your Google Sheet and on each device.

## Sharing with your boss

Give them the app link, the connection link and the **view PIN**. They see every entry
and the totals, but can't change anything. You can also share the Google Sheet itself
(**Share → Viewer**).

## Notes

- Without internet, the app shows the last copy it loaded. Saving needs a connection.
- The Android build is a debug-signed APK for installing directly, not a Play Store build.
- Currency is a display setting per device. Amounts are stored as plain numbers.
