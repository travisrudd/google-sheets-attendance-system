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
3. Add another HTML file named `Register` and paste in `Register.html`.
4. Open **Project Settings**, enable **Show "appsscript.json" manifest file in
   editor**, and replace it with the supplied `appsscript.json`.
5. Save the project.

## 3. Initialize the Sheet

1. In Apps Script, select `initializeAttendanceSystem` and click **Run**.
2. Approve the requested Google permissions.
3. Return to the Sheet and reload it.

The Sheet will now contain `People`, `Attendance`, and `Settings` tabs, plus an
**Attendance System** menu.

In `Settings`, replace **Attendance Check-In** with the organization or class
name you want attendees to see. Keep the timezone as `America/Chicago`.

The initializer also adds:

- **Self Registration Enabled** — use `TRUE` or `FALSE`.
- **Daily Registration Limit** — defaults to `50` successful registrations per
  Chicago calendar day.

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

The general `/exec` URL is now the self-registration page. A URL containing a
private `?t=...` token remains a person's check-in page.

> If your Google Workspace administrator does not allow public Apps Script web
> apps, they must enable that option or you must deploy from an account that
> permits it.

## 5. Configure delivery

### Email

Email uses the Google account that deployed the web app. The first run may ask
the deployer to approve permission to send email. Google account email quotas
apply.

### SMS with Twilio

Apps Script does not include native SMS delivery. To enable the Text Message
choice, create a Twilio account and obtain an SMS-capable Twilio number. Then in
Apps Script open **Project Settings → Script properties** and add these exact
properties:

| Property | Value |
|---|---|
| `TWILIO_ACCOUNT_SID` | Twilio Account SID |
| `TWILIO_AUTH_TOKEN` | Twilio Auth Token |
| `TWILIO_FROM_NUMBER` | Twilio number in E.164 format, such as `+13125551234` |

Do not put these credentials in the Sheet or source code. Twilio messaging fees,
trial-account restrictions, carrier rules, and consent requirements apply.

If these properties are absent, the registration page tells users that SMS is
not configured and asks them to choose email.

## 6. Self-registration

1. Share the general `/exec` URL, or use **Attendance System → Show
   self-registration link**.
2. The attendee enters their name and chooses email or text delivery.
3. The server validates the information, checks for an existing email/phone,
   creates the Person ID and token, and adds the person to `People`.
4. Their private link and setup instructions are delivered automatically.

The `People` sheet now includes **Email** and **Delivery Method** columns. A
duplicate email address or mobile number is not registered again.

## 7. Register people manually

1. In the `People` sheet, enter each attendee's **Name** and **Phone** on a new
   row. Leave all other columns blank.
2. Use **Attendance System → Set up new people**.
3. If the **Personal Link** column stays blank, use
   **Attendance System → Refresh personal links** after deployment.
4. Send each person only their own link.

Opening a personal link shows the attendee's name, today's Chicago date, and one
**Check In** button. The link works on iPhone, Android, tablets, and computers.
Attendees can add it to their Home Screen.

## 8. Day-to-day use

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

When upgrading from the original version, add `Register.html`, replace
`Code.gs`, run `initializeAttendanceSystem` again, and deploy a **New version**
of the existing web-app deployment. Existing people, tokens, attendance, and
personal links are preserved.

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
