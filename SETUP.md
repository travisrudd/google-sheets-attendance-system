# Attendance System Setup

The live system is hosted by Google Apps Script. GitHub stores the source code;
your computer does not need to stay on.

## 1. Create the Google Sheet

1. Create a blank Google Sheet.
2. Give it a useful name, such as **Class Attendance**.
3. In the Sheet, open **Extensions → Apps Script**.

## 2. Add the project files

In the Apps Script editor:

1. Replace the contents of `Code.gs` with the supplied `Code.gs`.
2. Add an HTML file named `Index` and paste in `Index.html`.
3. Open **Project Settings**, enable **Show "appsscript.json" manifest file in
   editor**, and replace it with the supplied `appsscript.json`.
4. Save the project.

## 3. Initialize the Sheet

1. In Apps Script, select `initializeAttendanceSystem` and click **Run**.
2. Approve the requested Google permissions.
3. Return to the Sheet and reload it.

The Sheet will now contain `People`, `Attendance`, and `Settings` tabs, plus an
**Attendance System** menu.

In `Settings`, replace **Attendance Check-In** with the organization or class
name you want attendees to see. Keep the timezone as `America/Chicago`.

## 4. Deploy the web app

1. In Apps Script, choose **Deploy → New deployment**.
2. Select **Web app**.
3. Set **Execute as** to **Me**.
4. Set **Who has access** to **Anyone**.
5. Click **Deploy** and approve any prompts.

The `/exec` URL Google provides is the hosted application. It runs in Google's
cloud and does not rely on your local machine.

Copy that working `/exec` URL into the **Web App URL** row in `Settings`. This
ensures personal links always use the intended deployment when the Apps Script
project has more than one active deployment.

> If your Google Workspace administrator does not allow public Apps Script web
> apps, they must enable that option or you must deploy from an account that
> permits it.

## 5. Register people

1. In the `People` sheet, enter each attendee's **Name** and **Phone** on a new
   row. Leave all other columns blank.
2. Use **Attendance System → Set up new people**.
3. If the **Personal Link** column stays blank, use
   **Attendance System → Refresh personal links** after deployment.
4. Send each person only their own link.

Opening a personal link shows the attendee's name, today's Chicago date, and one
**Check In** button. The link works on iPhone, Android, tablets, and computers.
Attendees can add it to their Home Screen.

## 6. Day-to-day use

- Each person can check in once per Chicago calendar date.
- Repeat attempts show **You already checked in today**.
- The raw record is written to `Attendance`.
- A tab for the year is created or refreshed automatically and grouped by date.
- Phone numbers and tokens never appear in annual report tabs.

To deactivate a person, change their `Active` value to `FALSE`. To replace a
compromised link, clear that person's Person ID and Token, then run setup again
on a new row (or have an administrator replace the token in the Sheet and
refresh links).

## Updating the deployed app

After changing code in Apps Script:

1. Choose **Deploy → Manage deployments**.
2. Edit the existing web app deployment.
3. Select **New version** and deploy.

The personal `/exec` links remain the same.

## Optional: sync GitHub with Apps Script using clasp

This is for maintainers who want command-line syncing; it is not required for
the app to run.

1. Install Node.js and then `npm install -g @google/clasp`.
2. Run `clasp login`.
3. In Apps Script **Project Settings**, copy the **Script ID**.
4. Copy `.clasp.json.example` to `.clasp.json` and paste in the Script ID.
5. Use `clasp push` to send repository changes to Apps Script, or `clasp pull`
   to retrieve changes.

Never commit `.clasp.json`; it is excluded by `.gitignore`.
