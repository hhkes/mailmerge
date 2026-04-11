# MailMerge Pro — Complete Developer Handover

## What this tool is

A fully self-contained single-page HTML application that lets a user send personalised bulk emails via their own Microsoft Outlook account. No backend, no server, no database. One HTML file hosted as a static site.

Live URL: https://hhkes.github.io/mailmerge
GitHub repo: https://github.com/hhkes/mailmerge

---

## Configuration constants (top of script block)

These are the only values that need to change to adapt the tool for a different user or environment.

```js
const PROOF_EMAIL  = 'hegertonsmith' + '@' + 'cultwines.com';
const MS_CLIENT_ID = 'dd4f1e8b-905c-4d47-8adb-29b071a2aee3';
const MS_TENANT_ID = 'be753f7a-5a0a-4bc2-ac71-5aeba08f237f';
const MS_REDIRECT  = 'https://hhkes.github.io/mailmerge';
const MS_SCOPE     = 'Mail.Send User.Read';
```

**PROOF_EMAIL** — The address that receives a test copy before Send All unlocks. Split across string concatenation to prevent GitHub Pages / Cloudflare from detecting it as an email address and injecting an obfuscation script that corrupts the JavaScript.

**MS_CLIENT_ID** — Azure AD app registration client ID. Find it at portal.azure.com → App registrations → Mass Email Sender.

**MS_TENANT_ID** — Azure AD tenant ID (Cult Wines Ltd). Used in both the /authorize and /token endpoint URLs.

**MS_REDIRECT** — The exact URL where this file is hosted. Must match exactly (no trailing slash) what is registered in Azure AD as a redirect URI. Microsoft does exact string matching.

**MS_SCOPE** — The Graph API permissions requested. `Mail.Send` to send email, `User.Read` to fetch the signed-in user's email address for display.

---

## Azure AD setup (portal.azure.com)

App registration name: **Mass Email Sender**

Required settings:
- Platform: Single-page application
- Redirect URIs registered:
  - `https://hhkes.github.io/mailmerge`
  - `https://regal-sawine-459cfe.netlify.app` (old, can be removed)
  - `https://regal-sawine-459cfe.netlify.app/auth.html` (old, can be removed)
- Authentication > Settings tab:
  - Allow public client flows: **Enabled**
  - Access tokens (implicit flow): **Unchecked**
  - ID tokens (implicit flow): **Unchecked**
- API permissions: `Mail.Send`, `User.Read` (both delegated, Microsoft Graph)

---

## Hosting

Hosted on GitHub Pages from the `main` branch of `https://github.com/hhkes/mailmerge`. The single file `index.html` at the root of the branch is served at `https://hhkes.github.io/mailmerge`.

To deploy an update: upload the new `index.html` to the repository via the GitHub web UI (Add file → Upload files) or via git push.

**Critical**: Do not use Netlify or any CDN that runs HTML post-processing. Both Netlify and GitHub Pages use Cloudflare's email obfuscation system, which detects email address literals in HTML and injects a script tag (`email-decode.min.js`) that splits the app's `<script>` block and breaks JavaScript execution. The PROOF_EMAIL constant is deliberately split as string concatenation to prevent detection.

---

## Authentication — how it works (PKCE OAuth 2.0)

The implicit flow (`response_type=token`) is deprecated by Microsoft for work/org accounts. This tool uses **Authorization Code + PKCE** (RFC 7636).

### Why PKCE without a client secret works

This is a public client SPA. There is no server and no client secret. PKCE replaces the secret: the app generates a random `verifier`, hashes it to produce a `challenge`, sends the challenge to Microsoft in the auth request, then proves it knows the verifier when exchanging the code for a token. Microsoft verifies the hash matches. "Allow public client flows" must be enabled in Azure AD for this to work.

### The popup approach

The tool opens a popup window pointing to Microsoft's `/authorize` endpoint. After the user logs in, Microsoft redirects the popup back to `MS_REDIRECT` (this same page) with `?code=...&state=<verifier>` in the URL.

### Why the verifier is in the `state` parameter

`sessionStorage` is not shared between browser windows. The popup is a separate window — it cannot read the parent's `sessionStorage`. Passing the verifier in the OAuth `state` parameter works because Microsoft echoes `state` back unchanged in the redirect URL, so the popup can read it from its own URL.

### The IIFE redirect handler

At the top of the script block, an immediately-invoked async function `handleOAuthRedirect` runs on every page load. It checks `window.location.search` for `?code=`. If found, it:
1. Shows the auth overlay (dark fullscreen spinner)
2. Hides the normal app UI
3. POSTs to `/token` to exchange the code + verifier for an access token
4. Calls `window.opener.postMessage({ type: 'MS_AUTH_TOKEN', access_token }, MAIN_ORIGIN)`
5. Closes the popup after 200ms

### Why MAIN_ORIGIN must be the bare origin

`window.opener.postMessage(data, targetOrigin)` — the `targetOrigin` must exactly match `window.opener.origin`, which is always scheme + host only, never including a path. For GitHub Pages: `'https://hhkes.github.io'`. Using the full URL `'https://hhkes.github.io/mailmerge'` causes the browser to silently drop the message.

Similarly, the `message` event listener checks `event.origin !== allowedOrigin` where `allowedOrigin = 'https://hhkes.github.io'`. Using `.startsWith()` with a path-including URL would fail because `event.origin` never contains a path.

### Token receipt

The parent page's `window.addEventListener('message', ...)` receives the token. It stores it in the `accessToken` variable and calls `fetchUserEmail()` to hit Graph `/me` and display the user's address in the header.

---

## Data structures

### recipients[]
```js
[{
  idx:        0,          // original row index
  supplier:   'Alias',    // company/supplier name
  name:       'Olivier',  // first name
  surname:    'Cailleau', // last name
  email:      'olivier@...',
  salutation: '',         // custom salutation if present in spreadsheet
  included:   true        // false if unchecked or duplicate
}]
```

### drafts[]
```js
[{
  to:       'olivier@...',
  supplier: 'Alias',
  subject:  'Catch up in Bordeaux - Alias',  // tokens resolved, supplier appended
  body:     'Good afternoon, Olivier,...',    // tokens resolved, plain text
  sigHtml:  '<div>Kind regards...</div>',     // raw HTML of selected signature
  sent:     false,
  failed:   false
}]
```

Note: `body` is stored as plain text with tokens resolved. HTML is built at send time by `buildEmailHtml()` so font/size/colour controls are always read live from the DOM.

### signatures[] (persisted in localStorage as 'mm_signatures')
```js
[{
  id:   '1710000000000',  // String(Date.now()) at creation time
  name: 'Hermione — Cult Wines',
  html: '<div style="font-family:Calibri...">Kind regards,...</div>'
}]
```

### Form state (persisted in localStorage and sessionStorage as 'mm_formstate')
```js
{
  subject: 'Catch up in Bordeaux',
  body:    '{{salutation}} {{name}},\n\nI am thinking about...',
  toggle:  true,     // append supplier name to subject
  cc:      '',       // comma-separated CC addresses
  bcc:     '',       // comma-separated BCC addresses
  font:    "Calibri,'Segoe UI',Arial,sans-serif",
  size:    '11pt',
  colour:  '#000000'
}
```

---

## Key functions reference

### Auth
- `handleOAuthRedirect()` — IIFE. Detects `?code=` in URL. Runs token exchange. Posts token to opener. Closes popup.
- `b64url(u8)` — Base64url-encodes a Uint8Array. Loop-based, no spread operator (avoids stack overflow on large arrays).
- `generatePkce()` — Returns `{ verifier, challenge }`. Verifier = 32 random bytes → base64url. Challenge = SHA-256(verifier) → base64url. All Uint8Array, no string/buffer mixing.
- `handleMsAuth()` — Opens auth popup. Starts watchClose interval to reset button if popup closes without auth.
- `fetchUserEmail()` — Hits Graph `/v1.0/me`. Stores email in `userEmail`. Calls `updateAuthUI(true)`.
- `updateAuthUI(connected)` — Updates header status pill and button text/state.
- `signOut()` — Clears `accessToken` and `userEmail`. Calls `updateAuthUI(false)`.

### File upload
- `handleFileUpload(e)` / `processFile(file)` — FileReader reads as ArrayBuffer. SheetJS parses. Calls `parseRecipients()`.
- `parseRecipients(raw)` — Two-pass fuzzy column matcher. Pass 1: exact lowercase header match. Pass 2: search term as substring of header. Handles NAME, Email Address, First Name, etc. Key: 'name' must be first in name terms list.
- `detectDuplicates()` — Marks second occurrence of duplicate emails as `included: false`.
- `getSalutation(r)` — Returns `r.salutation` if present, else time-based "Good morning"/"Good afternoon".
- `renderTable()` — Renders recipients tbody. Includes greeting preview column.
- `toggleRecipient(i, checked)` — Updates `recipients[i].included`. Calls `updateStats()`.

### Compose
- `insertToken(token)` — Inserts token at textarea cursor position. Preserves cursor.
- `updateSubjectPreview()` — Shows live subject line preview with supplier appended.
- `buildSubject(r)` — Returns subject + " - " + supplier if toggle is on and supplier exists.
- `renderBody(r, rawBody)` — Regex-replaces `{{salutation}}`, `{{name}}`, `{{surname}}`, `{{supplier}}` tokens.
- `buildEmailHtml(bodyText, sigHtml)` — HTML-escapes body, wraps in styled div (font/size/colour from DOM), appends signature with margin but no dividing line.
- `parseEmailList(str)` — Splits comma-separated string into Graph API `[{ emailAddress: { address } }]` array. Filters by `@` presence.

### Signature manager
- `loadSigsFromStorage()` / `saveSigsToStorage()` — localStorage key `mm_signatures`.
- `renderSigDropdown()` — Rebuilds select options. Restores previous selection.
- `onSigSelectChange()` — Updates mini-preview and enables/disables Edit/Delete buttons.
- `getSelectedSigHtml()` — Returns HTML of currently selected signature, or empty string.
- `openSigModal(sigId)` — Opens modal. `sigId = null` = create new. `sigId = string` = edit existing.
- `saveSigFromModal()` — Creates or updates signature. Re-renders dropdown. Selects saved item.
- `sanitiseSigHtml(html)` — Strips `<script>` tags before setting innerHTML.

### Generate & Send
- `generateAndSendProof()` — Validates inputs. Builds `drafts[]`. Calls `renderDraftPreview()`. Sends `[PROOF]` copy of `drafts[0]` to PROOF_EMAIL. Unlocks Send All on success.
- `renderDraftPreview()` — Shows/hides draft preview card. Renders table of all drafts with supplier, email, subject, body snippet.
- `escHtml(s)` — HTML-escapes a string for safe innerHTML insertion.
- `sendAllEmails()` — Iterates drafts, calls `buildEmailHtml()` + `sendRaw()` per draft. Respects `isPaused` (spin-wait) and `abortFlag` (break). Applies send-delay throttle between sends.
- `sendRaw(to, subject, htmlContent)` — Single Graph API POST to `/v1.0/me/sendMail`. Builds attachment array via `fileToBase64()`. Handles 401 by clearing token. Logs result.

### Progress UI
- `showProgress()` — Shows progress section card. Scrolls to it.
- `updateProgress(done, total, current)` — Updates bar, badge, label, count.
- `updateLastSentBanner()` — Shows last sent email and time.
- `logMsg(msg, type)` — Appends span to progress log. Types: ok/err/warn/info.
- `togglePause()` / `abortSend()` — Flip `isPaused` / set `abortFlag`.
- `finaliseSend(total)` — Shows completion banner with sent/failed counts.

### State management
- `saveFormState()` — Saves subject/body/toggle/cc/bcc/font/size/colour to sessionStorage + localStorage.
- `loadFormState()` — Restores from storage on page load.
- `resetAll()` — Clears all state, all DOM, resets step indicators.

### Helpers
- `sleep(ms)` — Returns a Promise that resolves after ms milliseconds.
- `fileToBase64(file)` — FileReader → readAsDataURL → splits off base64 data.
- `getExt(fn)` / `getMime(file)` / `getIcon(fn)` — File type utilities for attachments.

---

## Excel/CSV column mapping

The `parseRecipients` function accepts any column naming. Supported synonyms:

| Field | Accepted headers (case-insensitive, exact or substring) |
|---|---|
| supplier | supplier, company, organisation, organization, account |
| name | name, firstname, first name, forename, given, first |
| surname | surname, lastname, last name, family, last |
| email | email, e-mail, mail |
| salutation | salutation |

Pass 1 tries exact lowercase match. Pass 2 tries substring match (search term appears inside header). 'name' must be first in the name list — if it were after 'firstname', a column called 'firstname' would match 'first' before matching 'firstname'.

---

## Personalisation tokens

Four tokens are replaced per-recipient at draft-build time in `renderBody()`:

| Token | Replaced with |
|---|---|
| `{{salutation}}` | r.salutation, or "Good morning"/"Good afternoon" based on time |
| `{{name}}` | r.name (first name) |
| `{{surname}}` | r.surname |
| `{{supplier}}` | r.supplier |

---

## Cloudflare email obfuscation — the persistent bug

Both Netlify and GitHub Pages run HTML through Cloudflare's email protection system. It detects any `x@y.z` pattern in HTML and injects:
```html
<script data-cfasync="false" src="/cdn-cgi/scripts/.../email-decode.min.js"></script>
```
immediately before the app's `<script>` tag. This closes the script block early and breaks JavaScript.

**The fix** already applied:
1. `PROOF_EMAIL` is written as `'hegertonsmith' + '@' + 'cultwines.com'` — concatenation means no `@` literal in a detectable context.
2. The two visible HTML references to the email are empty `<span>` elements populated at runtime by `window.onload`.

If you add any email address as a string literal anywhere in the HTML or JS, Cloudflare will inject the script again. Always concatenate or construct email strings at runtime.

---

## Send flow diagram

```
User clicks Connect Outlook
  → handleMsAuth() generates PKCE verifier + challenge
  → Opens popup to Microsoft /authorize?...&state=<verifier>
  → User logs in
  → Microsoft redirects popup to https://hhkes.github.io/mailmerge?code=...&state=<verifier>
  → handleOAuthRedirect IIFE fires in popup
  → POSTs to /token with code + verifier
  → Receives access_token
  → window.opener.postMessage({ type:'MS_AUTH_TOKEN', access_token }, 'https://hhkes.github.io')
  → Popup closes
  → Parent message listener receives token
  → accessToken stored, fetchUserEmail() called
  → Header shows connected email address

User uploads Excel
  → processFile() → XLSX.read() → parseRecipients()
  → recipients[] populated
  → Table rendered with greeting previews

User writes subject + body, optionally sets CC/BCC/signature

User clicks Generate & Send Proof
  → drafts[] built with all tokens resolved
  → renderDraftPreview() shows all drafts in table
  → sendRaw(PROOF_EMAIL, '[PROOF] '+subject, html) called
  → On success: proofSent=true, Send All unlocked

User reviews proof in inbox, clicks Send All
  → Confirm dialog
  → Loop over drafts[], sendRaw() per draft
  → Throttle delay between sends
  → Pause/Resume/Abort supported
  → Progress log + bar updated live
  → Completion banner shown
```

---

## Extending the tool

### Adding a new personalisation token
1. Add the token button to `.token-bar` in HTML: `<div class="token" onclick="insertToken('{{newtoken}}')">{{newtoken}}</div>`
2. Add the replacement in `renderBody()`: `.replace(/\{\{newtoken\}\}/g, r.newfield || '')`
3. Add the field to `parseRecipients()`: `newfield: get(row, ['newfield', 'synonyms'])`
4. Add to the recipient object schema

### Adding a new column to the Draft Preview table
Edit `renderDraftPreview()` — add a `<th>` to the header and a `<td>` to the row template.

### Changing the proof email address
Change `PROOF_EMAIL` constant — keep it as string concatenation to avoid Cloudflare detection:
```js
const PROOF_EMAIL = 'newaddress' + '@' + 'domain.com';
```

### Deploying to a different URL
1. Change `MS_REDIRECT` to the new URL
2. Register the new URL in Azure AD as a SPA redirect URI
3. Update `MAIN_ORIGIN` and `allowedOrigin` to the new bare origin (scheme + host only, no path)

### Adding a new Azure AD app registration (different tenant)
1. Go to portal.azure.com → App registrations → New registration
2. Name it, set supported account types to "Single tenant"
3. Add redirect URI as Single-page application: your deployment URL
4. Go to Authentication → Settings → Enable "Allow public client flows"
5. Go to API permissions → Add → Microsoft Graph → Delegated → Mail.Send, User.Read
6. Update MS_CLIENT_ID and MS_TENANT_ID constants

---

## Known limitations

- **Token expiry**: Microsoft access tokens expire after 1 hour. If a send job runs longer, it will fail mid-way with a 401. The tool detects this and shows a reconnect message. Fix: add refresh token support (requires `offline_access` scope, which may conflict with corporate conditional access policies).
- **Attachment size**: 3MB per file, enforced client-side. The Graph API itself accepts up to ~3MB per attachment in a single request before requiring upload sessions.
- **No retry logic**: Failed sends are logged but not retried automatically. The user can see which failed in the log.
- **Single tenant**: The app is registered as single-tenant (Cult Wines only). To allow other organisations, change supported account types in Azure AD to "Multitenant" and update MS_TENANT_ID to `'common'`.
- **Session only**: `accessToken` lives in memory only, not localStorage. Refreshing the page requires reconnecting Outlook.

---

## File structure

There is only one file:

```
index.html   — The entire application (CSS + HTML + JS, ~1700 lines)
```

No build step, no npm, no bundler. Edit the file and upload to GitHub.

---

## Prompt to give a new AI to continue development

```
You are continuing development of MailMerge Pro, a single-file HTML mail merge 
tool that sends personalised bulk emails via Microsoft Outlook using the Graph API.

The tool is live at https://hhkes.github.io/mailmerge and the source is at 
https://github.com/hhkes/mailmerge/blob/main/index.html

KEY TECHNICAL FACTS you must know before making any changes:

1. CLOUDFLARE EMAIL OBFUSCATION: GitHub Pages injects a Cloudflare script 
   before any <script> tag if it detects an email address literal (x@y.z) 
   anywhere in the HTML or JS. This breaks the app. NEVER write a complete 
   email address as a string literal. Always use concatenation:
   'user' + '@' + 'domain.com'

2. POSTMESSAGE ORIGIN: The popup posts the OAuth token to the parent window 
   using postMessage. The targetOrigin MUST be the bare origin only:
   'https://hhkes.github.io'  (NOT 'https://hhkes.github.io/mailmerge')
   event.origin is also always a bare origin. Use exact match (===), not 
   startsWith().

3. PKCE VERIFIER IN STATE PARAM: sessionStorage is not shared between popup 
   and parent windows. The PKCE verifier is passed via the OAuth state 
   parameter and echoed back by Microsoft in the redirect URL.

4. THIS FILE IS ITS OWN REDIRECT TARGET: The same index.html handles both 
   the normal app load and the OAuth redirect. An IIFE at the top of the 
   script checks for ?code= in the URL and runs the token exchange if found.

5. FONT/SIZE/COLOUR are read live from the DOM at send time, not stored in 
   drafts[]. Body text is pre-rendered (tokens resolved) but HTML is built 
   fresh per send.

6. PROOF GATE: Send All is disabled until proofSent=true. proofSent is only 
   set to true after sendRaw() succeeds for the proof email.

Configuration:
- PROOF_EMAIL: hegertonsmith [at] cultwines.com
- MS_CLIENT_ID: dd4f1e8b-905c-4d47-8adb-29b071a2aee3
- MS_TENANT_ID: be753f7a-5a0a-4bc2-ac71-5aeba08f237f
- MS_REDIRECT: https://hhkes.github.io/mailmerge
- MS_SCOPE: Mail.Send User.Read
- Deployment: GitHub Pages, repo hhkes/mailmerge
```
