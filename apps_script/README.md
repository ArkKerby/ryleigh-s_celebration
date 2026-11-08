# RSVP sheet setup

The Apps Script is prepared for the spreadsheet at:

`https://docs.google.com/spreadsheets/d/1xzzbHzWASeCMjvG9m6PuXlq5XVXgesqlxW6xRDQZmC8/edit`

## Create the tabs and godparent directory

1. Open the spreadsheet and choose **Extensions → Apps Script**.
2. Replace the editor contents with `Code.gs` from this folder. If the editor exposes `appsscript.json`, use the manifest in this folder too.
3. Save, select `setupRsvpSystem`, and click **Run**. Authorize spreadsheet access and email sending from the Google account that will own the RSVP service.
4. The script creates `RSVP Responses`, `Godparent Directory`, and `Summary` tabs. It seeds the directory with the Ninongs and Ninangs currently listed on the invitation. Keep that directory current if names change.

## Deploy the RSVP endpoint

1. In Apps Script, choose **Deploy → New deployment → Web app**.
2. Set **Execute as** to your account, and set access to **Anyone** if that option is available for your account.
3. Deploy and copy the URL ending in `/exec`.
4. Paste that URL into `rsvp-config.js` as the value of `window.RYLEIGH_RSVP_ENDPOINT`.

The website posts submissions to the endpoint in a hidden frame, so it can display the actual Apps Script response without exposing sheet data to the browser. The script appends each response before attempting the confirmation email. If email sending is unavailable, the response remains recorded and the page says the confirmation was not sent.

For an `Other guest`, the script matches at least two ordered name parts against the `Godparent Directory`, ignoring punctuation and letter case (so `John Mark Laguador` can match `John Mark T. Laguador`). A unique match is classified as Ninong or Ninang and stores the full roster name. No match stays `Other guest`; multiple matches are marked `Review` instead of guessing.

The public endpoint includes input validation, a hidden honeypot field, a daily cap of 200 submissions, a one-hour per-address email cooldown, and a daily confirmation cap of 40 emails. Keep the spreadsheet private; it contains guest names, email addresses, and messages. The public website only receives a success/failure message, not the table contents.
