# Ambaji Foods Quality Forms

This folder is ready to publish with GitHub Pages. Submitted form records are stored in a Google Sheet, and a branded PNG copy of every saved record is uploaded to a Google Drive folder.

## 1. Prepare Google storage

The Google Sheet and Google Drive image folder have already been created and configured in `apps-script/Code.gs`.

1. Open [Google Apps Script](https://script.google.com/) and create a new project.
2. Replace the default script with `apps-script/Code.gs`.

## 2. Deploy the Google connection

1. In Apps Script, choose **Deploy → New deployment**.
2. Select **Web app**.
3. Set **Execute as** to **Me**.
4. Select the access level appropriate for your staff.
5. Deploy and authorize access to Sheets and Drive.
6. Copy the web-app URL ending in `/exec`.
7. In `index.html`, replace `PASTE_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE` with that URL.

## 3. Publish on GitHub Pages

1. Create a GitHub repository.
2. Upload `index.html` and this README. The `apps-script` folder may also be kept in the repository for maintenance.
3. Open the repository’s **Settings → Pages**.
4. Choose **Deploy from a branch**, select the `main` branch and root folder, then save.
5. GitHub will provide a public URL such as `https://username.github.io/repository-name/`.

## Storage behavior

- Form data is stored in the `Form Records` tab of the configured Google Sheet.
- A PNG report is uploaded to the configured Drive folder whenever a record is saved or updated.
- CSV export, image download, printing and WhatsApp sharing remain available.
- “Apps By Prateek Agarwal” appears on the selector, forms and generated reports.

## Important

Do not place Google passwords, API keys or OAuth tokens in `index.html` or the GitHub repository. Only the deployed Apps Script web-app URL belongs in the website configuration.
