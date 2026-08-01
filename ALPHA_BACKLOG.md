# Alpha Backlog

This file tracks the work needed before publishing Ask Ollama / CoOllama as an open GitHub repo for others to fork and improve.

## Current Working Baseline

- New Outlook can load the sideloaded add-in.
- The local HTTPS server works at `https://localhost:8765/pane/taskpane.html`.
- The pane can reach Ollama at `http://localhost:11434`.
- Model selection works.
- Summarize works against selected Inbox email context.
- Draft Reply appears structurally connected.
- Sent Mail also works for basic current-email actions.


## Completed In First Alpha Button-Up Pass

- Rebranded visible app surfaces and manifest text to **CoOllama**.
- Added clean `assets/coollama-logo.png` and regenerated 16/32/80px icons from it.
- Added the CoOllama logo to the task pane header.
- Persisted selected model in `localStorage` and auto-select it on pane load.
- Added Markdown rendering for main model responses.
- Added summary modes: Bulleted Summary, Short Summary, Detailed Review, and Suggested Response.
- Added Draft in Your Voice.
- Added current-email checker.
- Added spell and grammar check.
- Added configurable voice sample size.
- Added a quick Add Current Email button for voice sample testing.

## Completed In Installer Polish Pass

- Reworked `install.bat` around the current New Outlook/manual sideload flow.
- Installer now copies runtime files, README, and backlog to `%LOCALAPPDATA%\OllamaOutlookAddin`.
- Installer creates a Start Menu shortcut: `CoOllama` -> `Start CoOllama Server`.
- Installer checks whether Ollama is on PATH and whether `http://localhost:11434/api/tags` responds.
- Reworked `uninstall.bat` to remove copied files, the Start Menu shortcut, legacy trusted-catalog registry entry, and CoOllama/Ask Ollama localhost certs.
- Server script now uses CoOllama branding and connection timeouts.

## Completed In Startup/Troubleshooting Pass

- Installer now creates a Windows Startup shortcut so the CoOllama local server starts at login.
- Manual Start Menu server shortcut remains available for recovery if the user closes the server.
- Uninstaller removes the Startup shortcut.
- Pane now shows a visible troubleshooting panel when Ollama is unreachable or no models are installed.
- Docs now explain that the Outlook add-in button cannot directly launch the local Windows server process from inside Outlook's sandbox.

## Completed In Suspicious Check Pass

- Added **Suspicious?** current-email action.
- Suspicious check looks for visible phishing, impersonation, fraud, mismatch, urgency, link-text, attachment-pressure, and off-platform conversation cues.
- Prompt explicitly avoids claiming virus scanning or hidden-header/attachment inspection.
- Output asks for Low/Medium/High caution, reasons, red flags, verification steps, and safest next action.

## Completed In Pinning/Compose Pass

- Pinning confirmed after corrected nested v1.0/v1.1 manifest reinstall.

- Added Mail 1.1 manifest `SupportsPinning` for read and compose task pane actions.
- Added `Office.EventType.ItemChanged` handler to refresh context when Outlook keeps a pinned pane open and the selected item changes.
- Added manual **Refresh** button for current email context.
- Added compose-only **Insert** button to place generated output into the draft body at the current selection/cursor.
- Kept this documented as Outlook/client-dependent rather than guaranteed behavior.

## Completed In Icon Pass

- Regenerated command icon assets as a simplified high-contrast CoOllama mark for 16/32/80px Outlook surfaces.
- Added `icon-64.png` as a helper asset for future packaging surfaces.
- Kept `coollama-logo.png` for the larger in-pane logo.
- Still needs visual confirmation in New Outlook because some surfaces choose the 16px command icon by design.

## Completed In Icon URL Fix Pass

- Promoted user-provided 16/32/80 icon assets into the manifest-targeted icon sizes.
- Generated a proper 64x64 app/install icon from the 80px user asset.
- Updated top-level `IconUrl` to `assets/coollama-app-64.png` instead of the 32px command icon.
- Moved command icon URLs to fresh filenames: `coollama-command-16.png`, `coollama-command-32.png`, `coollama-command-80.png`.
- Bumped manifest version to `1.0.2.0` to help with Outlook icon caching.
- Added no-cache headers to the local HTTPS server script for future asset swaps.

## Completed In Theme Toggle Pass

- Added compact Light/Dark toggle in the task pane header.
- Added dark-mode CSS variable palette.
- Persisted theme preference in `localStorage`.

## Completed In Compose Hardening Pass

- Added Read/Compose mode pill to the current email card.
- Hardened compose detection around selection and body insertion APIs.
- Rewrite selection now preserves selection source metadata and refuses stale rewrites after item changes.
- Apply rewrite now reports actionable errors and clears rewrite state after success.
- Insert now uses `item.body.setSelectedDataAsync` so generated drafts go into the compose body instead of accidentally targeting subject selection.
- Generated result insertion is disabled when output belongs to a different Outlook item.
- Insert now converts Markdown preview output to sanitized email HTML, preserving paragraphs/bullets in Outlook, with plain-text fallback. It strips generated `Subject:` lines, horizontal rules, code fences, and placeholder signatures.

## Completed In Sticky State / Model Polish Pass

- Sorted Ollama model list alphabetically.
- Added **Unload** beside the model selector. It calls Ollama `/api/generate` with `keep_alive: 0` for the selected model.
- Added **Clear** for the current item work state.
- Added lightweight recent-item persistence in `localStorage` for reply notes, Ask Anything text, summary mode, and latest result, capped to 20 recent item keys.
- Documented pinning as working but still one-current-Outlook-item-at-a-time for alpha.
## High Priority

1. Done: Persist selected model in `localStorage` and auto-select it on pane load.
2. Done: Add Markdown rendering so responses do not show raw `**bold**` syntax.
3. Done: Fix branding:
   - Decide final public name: **Ask Ollama** or **CoOllama**.
   - Use the CoOllama logo asset in the task pane.
   - Improved manifest icon assets for tiny Outlook surfaces and refreshed the in-pane header logo.
4. Done: Update install/uninstall scripts and README for the current New Outlook flow.
5. Done: Add summarize modes:
   - Bulleted Summary
   - Short Summary
   - Detailed Review
   - Suggested Response / Next Steps

## Medium Priority

1. Done: Add **Draft in your voice**.
2. Done for alpha/manual workflow: Add Sent Mail samples by opening sent messages one at a time with sticky sample drafts. Full automatic Sent-folder import moves to mailbox search/Graph follow-up.
3. Done: Let users choose voice sample size.
4. Done: Add spell and grammar check.
5. Done for current email: Add email checker/reviewer for tone, clarity, missing context, and suggested improvements.
6. Done: Add Suspicious Email Check for visible caution cues. This is not a virus scanner.


## Completed In Voice Sample Persistence Pass

- Renamed **Add current email** to **Add current sample** for the voice profile workflow.
- Voice sample draft text now persists in `localStorage` under `coollamaVoiceSamples`.
- Voice sample size now persists under `coollamaVoiceSampleLimit`.
- Added **Clear samples** and **Copy profile** controls.
- Documented alpha boundary: automatic Sent-folder import requires the later Microsoft Graph/mailbox search path; current alpha supports opening Sent messages one at a time and adding them as samples.

## Completed In Button Contrast Polish Pass

- Tuned primary buttons to match the dark green / light teal status-pill pattern.
- Primary buttons now use a darker green fill, subtle border, and light teal text to avoid the washed-out white-on-aqua look in light mode.

## Completed In Alpha Repo Hygiene / Security Pass

- Added `.gitignore` for local installs, local-only Codex context, screenshots, generated asset variants, logs, and editor/OS clutter.
- Moved nonessential screenshots, old logos, and generated icon variants into ignored `_local-private-assets`.
- Added `SECURITY.md` documenting local-first behavior, localStorage data, localhost certificate behavior, and alpha limits.
- Added `CHANGELOG.md` with `0.1.0-alpha` notes.
- Rewrote `README.md` into a cleaner public alpha handoff and removed the screenshot placeholder reference.
- Removed unnecessary `Access-Control-Allow-Origin: *` from the local static server response headers.
- Re-scanned source for personal paths, names, emails, phone numbers, tokens, and copied real-message content before repo publication.

## Completed In Installer Status Message Pass

- Changed fresh-install Ollama check from a scary warning to clearer status text.
- Installer now distinguishes: Ollama not found, Ollama installed but not responding, Ollama running with no models, and Ollama running with at least one model.
- Replaced the fragile PowerShell JSON probe with a simpler `curl.exe` check against `http://127.0.0.1:11434/api/tags` plus a `findstr` model-name check.
- Smoke-ran `install.bat`; it refreshed the installed copy and correctly detected the running Ollama server.
## Larger Follow-Up

1. Add mailbox/email search.
2. Add Q/A with Inbox or search results instead of only the selected email.
3. Done for tested New Outlook path: Added manifest pinning support plus ItemChanged refresh handling, and pinning is confirmed working after reinstalling the corrected nested manifest.
4. Done for alpha: automatic local server startup through a Windows Startup shortcut. Future polish could replace this with a tray helper or scheduled task.
5. Consider hosting static pane assets on a trusted HTTPS host to avoid local certificate friction.

## Alpha Install Notes

End-user instructions must use generic Windows paths, not developer-specific paths.

Required flow:

1. Install Ollama from `https://ollama.com`.
2. Start Ollama.
3. Pull at least one model, for example:

```bat
ollama pull llama3.2
```

4. Run `install.bat`.
5. Start the local HTTPS server and keep it open:

```bat
%LOCALAPPDATA%\OllamaOutlookAddin\start-server.bat
```

6. Sideload the manifest through:

```text
https://aka.ms/olksideload
```

7. Add from file:

```bat
%LOCALAPPDATA%\OllamaOutlookAddin\manifest.xml
```



















