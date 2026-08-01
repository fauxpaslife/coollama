# Security Policy

## Local-First Boundary

CoOllama is designed as a local-first Outlook add-in. The task pane runs in Outlook, reads the currently selected/open Outlook item through Office.js, and sends prompts to Ollama at `http://localhost:11434`.

The project does not include a hosted backend service. Email content, generated replies, voice samples, and voice profiles are not intentionally sent to any third-party API by this code. The model server is the user's local Ollama instance.

## Data Stored Locally

CoOllama stores these values in the task pane browser `localStorage`:

- selected Ollama model
- light/dark theme preference
- voice profile
- voice sample draft text
- voice sample size
- recent current-item work state, capped to 20 item keys

Use **Clear** in the pane to clear saved work for the current Outlook item. Use **Clear samples** to remove saved voice sample drafts. Browser/Outlook site data controls can also clear this storage.

## Local HTTPS Certificate

New Outlook requires HTTPS resources for add-ins. The included local server creates a self-signed `localhost` certificate in the current user's Windows certificate stores so Outlook can load `https://localhost:8765/pane/taskpane.html`.

The server binds to `127.0.0.1` only. It serves static files from the installed add-in folder and rejects path traversal outside that folder.

## Known Alpha Limits

- This is alpha software. Review generated output before sending email.
- Suspicious Email Check is not a malware scanner, virus scanner, or hidden-header analyzer.
- CoOllama operates on the current Outlook item. Mailbox-wide search/import is deferred to a future Microsoft Graph permission flow.
- Local `localStorage` may contain email-derived notes or generated text until cleared.

## Reporting Issues

For a public repo, open an issue with:

- Outlook version and account type, if relevant
- CoOllama version or commit
- Ollama version and model name
- steps to reproduce

Do not paste private email content, access tokens, certificate private keys, or personal data into public issues.
