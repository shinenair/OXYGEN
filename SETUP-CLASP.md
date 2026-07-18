# One-time setup: auto-deploy from GitHub to Apps Script

After this setup, merging any change to `src/` on the `main` branch
automatically updates the Apps Script project AND the live web app —
no more copy-pasting into the editor, no more "New version" clicks.

There are 3 steps only the project owner can do (they involve your
Google account). Total time: about 10 minutes.

---

## Step 1 — Turn on the Apps Script API

1. Open **https://script.google.com/home/usersettings**
2. Switch **Google Apps Script API** to **ON**.

## Step 2 — Get your Script ID (tell it to Claude)

1. Open your OXYGEN project in the Apps Script editor.
2. Click ⚙️ **Project Settings** (left sidebar).
3. Under **IDs**, copy the **Script ID** (a long string of letters/numbers).
4. Give it to Claude, who will put it into `.clasp.json` in this repo.
   (Or edit `.clasp.json` yourself on github.com — replace
   `PASTE_YOUR_SCRIPT_ID_HERE` with it.)

## Step 3 — Create the deployment credentials (clasp login)

This produces a small credentials file that GitHub's robot will use to
act as you. You do it in **Google Cloud Shell** — a free terminal in
your browser, nothing to install:

1. Open **https://shell.cloud.google.com** (sign in with the SAME
   Google account that owns OXYGEN). Wait for the terminal prompt.
2. Paste this and press Enter:

   ```
   npm install -g @google/clasp && clasp login --no-localhost
   ```

3. It prints a long Google URL — click it, choose your account, click
   **Allow**, and copy the code Google shows you.
4. Paste that code back into the terminal and press Enter.
   It should say *"Authorization successful"*.
5. Now print the credentials file, base64-encoded (one copy-safe block —
   raw JSON gets corrupted by terminal line-wrapping when copied):

   ```
   base64 -w 0 ~/.clasprc.json ; echo
   ```

6. Copy the ENTIRE block of letters/digits it prints (it may wrap
   across several lines on screen — that's fine, copy it all).

## Step 4 — Store the credentials as a GitHub secret

1. On github.com, open the **OXYGEN** repo →
   **Settings → Secrets and variables → Actions**.
2. Click **New repository secret**.
3. Name: `CLASPRC_JSON_B64`
4. Secret: paste the base64 block from Step 3.
5. Click **Add secret**.

---

## Done — how it works from now on

- Any merge to `main` that touches `src/` triggers the
  **Deploy to Apps Script** workflow (visible under the **Actions** tab).
- It pushes the code into your Apps Script project, then updates the
  existing web-app deployment — the `/exec` URL stays the same, so the
  GitHub Pages site keeps working untouched.
- You can also trigger it manually: **Actions → Deploy to Apps Script →
  Run workflow**.

## ⚠️ Important: GitHub becomes the source of truth

`clasp push -f` makes the Apps Script project match `src/` in this repo
EXACTLY. Any edit made directly in the Apps Script editor will be
**overwritten** on the next deploy unless it's also in GitHub. From now
on, make changes through GitHub (ask Claude, or edit on github.com) —
not in the Apps Script editor.
