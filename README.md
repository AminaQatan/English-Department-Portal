# English Department Portal

Staff portal for the English Department, Salalah Vocational College.

English Department © 2026

## Files

| File | What it is |
| --- | --- |
| `index.html` | The whole portal: design, college logos, colours, sounds, birds, reports and calendar. Do not edit. |
| `config.js` | The only file you edit: paste the Google Apps Script web app address here. |
| `apps-script/Code.gs` | Server code. It goes into Google Apps Script, not GitHub Pages. Kept here as a copy. |
| `apps-script/appsscript.json` | Optional Apps Script project settings (Muscat time zone, web app). |

## 1. Put the portal on GitHub Pages

1. On GitHub, create a new **public** repository, for example `english-department-portal`.
2. Press **Add file › Upload files**. Drag in `index.html`, `config.js`, `README.md` and the `apps-script` folder. Press **Commit changes**.
3. Open **Settings › Pages**. Under *Build and deployment*, choose **Deploy from a branch**, then **main** and **/ (root)**, and press **Save**.
4. After a minute or two, the portal is at `https://YOUR-USERNAME.github.io/english-department-portal/`.

At this point it runs as a demo: sample data that stays in each browser.

## 2. Connect it to the Google Sheet (shared data for the whole department)

1. Create a Google Sheet, for example **English Department Portal Data**. Set **File › Settings › Time zone** to **(GMT+04:00) Muscat**.
2. Open **Extensions › Apps Script**, replace everything with `apps-script/Code.gs` and save.
3. Choose the function **setup** and press **Run**. Allow the permissions it asks for.
4. Choose **Deploy › New deployment › Web app**. Set *Execute as* to **Me** and *Who has access* to **Anyone**. Press **Deploy** and copy the web app address (it ends with `/exec`).
   *Who has access* must be **Anyone** for the GitHub copy to reach the Sheet.
5. Back on GitHub, open `config.js`, press the pencil icon, paste the address between the quotes and press **Commit changes**:

   ```js
   window.PORTAL_CONFIG = {
     scriptUrl: "https://script.google.com/macros/s/XXXXXXXX/exec"
   };
   ```

6. After a minute, reload the portal. The demo bar is gone. Sign in as **Head of English Department** with the password **HOD@2026** and change it straight away.
7. In **Admin › Staff accounts**, add every lecturer. In **Admin › Portal text and options**, add the department text, term details and drive links.
8. In the Google Sheet, reload and choose **Portal › Turn on weekly backup**.

## 3. Use the college domain (optional)

To show the portal at an address like `english.svc.edu.om`:

1. In **Settings › Pages › Custom domain**, type the address and press **Save**.
2. Ask the college IT team to add a **CNAME** record for that address pointing to `YOUR-USERNAME.github.io`.
3. When GitHub shows the domain as ready, tick **Enforce HTTPS**.

## Updating

- **New portal page:** upload the new `index.html` over the old one. `config.js` stays as it is, so the connection is kept.
- **New server code:** paste the new `Code.gs` into Apps Script, save, then **Deploy › Manage deployments › Edit › Version: New version › Deploy**. The web app address stays the same.

## Good to know

- The demo's sample entries are not copied to the Google Sheet. The live portal starts with the Head of Department account and the default settings.
- Data and sign-in are in the Google Sheet. Uploads (up to 20 MB each) are saved in the **English Portal Uploads** folder in the Sheet owner's Google Drive.
- GitHub Pages only hosts the page. Without the Google Sheet connection, nothing is shared between staff.
