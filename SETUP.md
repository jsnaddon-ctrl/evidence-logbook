# Evidence Logbook – setup guide

Allow about 45 minutes. Do the parts in order. You'll collect three things along the way:

| You get | From | Goes into |
|---|---|---|
| App web address | GitHub (Part 1) | Microsoft and Cloudflare settings |
| Client ID | Microsoft (Part 2) | `config.js` |
| Worker URL | Cloudflare (Part 4) | `config.js` |

In the steps below, replace `YOURNAME` with your GitHub username.

---

## Part 1 – GitHub: put the app online

1. Go to github.com → **New repository**. Name it `evidence-logbook`, set it to **Public**, and click **Create repository**.
   (Free GitHub Pages needs a public repo. Nothing secret is stored in it, because the API key lives on Cloudflare.)
2. Click **uploading an existing file**. Drag in everything from this folder (including the `worker` folder) and click **Commit changes**.
3. Go to **Settings → Pages**. Under **Source**, choose **Deploy from a branch**, select **main** and **/ (root)**, then click **Save**.
4. Wait 1–2 minutes. Your app address is:
   `https://YOURNAME.github.io/evidence-logbook/`
5. **Quick check:** open `https://YOURNAME.github.io/evidence-logbook/?demo=1`. You should see the app in demo mode (nothing is saved).

## Part 2 – Microsoft: allow sign-in and OneDrive

1. Go to **portal.azure.com** and sign in with your personal Microsoft account.
2. Search for **Microsoft Entra ID**, open it, then click **App registrations → New registration**.
   - **If it says you can't register apps or there's no directory:** Microsoft requires personal accounts to have a directory first. Sign up for the free Azure account at azure.microsoft.com/free with the same Microsoft account. It asks for a card to verify your identity, but the free tier isn't charged for this. Then come back to this step.
3. Fill in the form:
   - **Name:** Evidence Logbook
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*
   - **Redirect URI:** choose platform **Single-page application (SPA)** and enter
     `https://YOURNAME.github.io/evidence-logbook/` (include the slash at the end)
4. Click **Register**. Copy the **Application (client) ID** shown on the overview page. This is your **Client ID**.
5. Click **API permissions → Add a permission → Microsoft Graph → Delegated permissions**. Tick **Files.ReadWrite** and click **Add permissions**.
   (User.Read is already there. Apprentices approve these themselves the first time they sign in.)

## Part 3 – Anthropic: the AI key

1. Go to **console.anthropic.com** and sign up.
2. Go to **Billing** and add some credit.
3. **Set a monthly spend limit** under **Limits** so costs can never run away.
4. Go to **API keys → Create key**. Copy it straight away (it's only shown once) and keep it private.

## Part 4 – Cloudflare: the AI helper

1. Go to **dash.cloudflare.com** and sign up (the free plan is enough).
2. Go to **Workers & Pages → Create → Create Worker**. Name it `evidence-ai` and click **Deploy**.
3. Click **Edit code**. Delete what's there, paste in everything from `worker/worker.js`, and click **Deploy**.
4. Go to the Worker's **Settings → Variables and Secrets** and add:

   | Name | Type | Value |
   |---|---|---|
   | `ANTHROPIC_API_KEY` | **Secret** | your key from Part 3 |
   | `ALLOWED_ORIGINS` | Text | `https://YOURNAME.github.io` (no path, no slash at the end) |
   | `ACCESS_CODES` | Text | a code you'll give apprentices, e.g. `symonite-2026` (add more separated by commas) |
   | `MODEL` | Text | optional. Leave it out to use the default, or set a cheaper model |

5. Copy the Worker address shown at the top, e.g. `https://evidence-ai.yourname.workers.dev`. This is your **Worker URL**.

## Part 5 – Connect it all

1. In your GitHub repo, open `config.js` and click the pencil icon to edit.
2. Paste in your **Client ID** and **Worker URL**:
   ```js
   clientId: "1a2b3c4d-....",
   workerUrl: "https://evidence-ai.yourname.workers.dev",
   ```
3. Click **Commit changes** and wait 1–2 minutes.

## Part 6 – Test on a phone

1. On the phone, open `https://YOURNAME.github.io/evidence-logbook/` in **Chrome**.
2. Tap **Sign in with Microsoft** and approve access.
3. Go to **Settings**, enter the **AI access code**, fill in the details, and tap **Save**.
4. Tap Chrome's **⋮ menu → Add to Home screen** so it opens like an app.
5. Add a unit book, take a photo, and check that it appears in OneDrive under **Apprenticeship Evidence**.

---

## Making updates later

- Upload the changed files to the repo, replacing the old ones. Phones get the new version next time the app is opened. If a phone seems stuck on the old version, refresh the page.
- The version number is shown at the bottom of **Settings**.
- Apprentices' data lives in their own OneDrive, so updates never touch it.
- To go back to an earlier version, open the file's **History** in GitHub.

## If something goes wrong

| Problem | Fix |
|---|---|
| Sign-in error mentioning *redirect URI* | The Redirect URI in Part 2 must match the app address exactly, including the slash at the end. |
| Sign-in says personal accounts aren't allowed | In the app registration, open **Authentication** and check the supported account types include personal Microsoft accounts. |
| “The AI helper needs the access code” | Enter a code from `ACCESS_CODES` in the app's Settings. |
| “doesn't recognise this web address” | `ALLOWED_ORIGINS` must be exactly `https://YOURNAME.github.io`. |
| Work or school accounts can't sign in | That organisation's IT may need to approve the app. Personal accounts aren't affected. |
