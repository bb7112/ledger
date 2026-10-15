# ledger setup

Personal money tracker. Everything lives on your **personal** accounts only:
Google `brendaocana08@gmail.com` and GitHub `brenda-711`. Nothing First Chair.

## 1. Google Sheet + Drive (about 5 minutes, on your laptop)

1. Signed in as **brendaocana08@gmail.com**, go to sheets.new and name the sheet `ledger`.
2. Click **Extensions > Apps Script**.
3. Delete what's there and paste everything from `google-apps-script.gs`.
4. On the line `const SECRET = 'CHANGE-ME';`, replace `CHANGE-ME` with your own passcode (a few words only you know). Save.
5. Click **Deploy > New deployment**, then the gear icon > **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Click **Deploy**, then **Authorize access**. Pick your personal Gmail. On "Google hasn't verified this app", click **Advanced > Go to ledger (unsafe)**. It's safe: it's your own script, and it only uses your sheet and a `ledger receipts` folder in your Drive.
7. Copy the **Web app URL** (ends in `/exec`).

"Anyone" only means the URL can be reached. Without your passcode it refuses everything.

## 2. Put the app online (GitHub Pages, personal account)

1. On github.com, signed in as **brenda-711**, create a new repository called `ledger` (public is fine because none of your money data is in it, only the app).
2. Tell Claude when it exists. Claude pushes the code and turns on Pages.
3. The app is then at **https://brenda-711.github.io/ledger/**

## 3. On your iPhone

1. Open https://brenda-711.github.io/ledger/ in **Safari**.
2. Tap **Share > Add to Home Screen**.
3. Open **ledger** from your home screen, then tap the settings icon (top right of the dashboard).
4. Paste the web app URL from step 1.7, type your passcode, and tap **connect**.
5. Add your fixed bills, monthly budget and savings goal in the same settings screen.

## Changing the backend later

If you edit the Apps Script, use **Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy**.
That keeps the same URL so the app keeps working.
