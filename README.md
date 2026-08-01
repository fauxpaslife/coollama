# CoOllama Outlook Add-in

CoOllama is a local-first Outlook task pane add-in for summarizing email, drafting replies, rewriting selected compose text, and chatting with the current message using models served by Ollama.

> Alpha status: useful for daily testing, but still early. Review generated output before sending email.

## Requirements

- Windows
- New Outlook or Outlook with add-in support
- Ollama installed from `https://ollama.com`
- Ollama running locally at `http://localhost:11434`
- At least one Ollama model installed
- PowerShell, included with Windows

## Install Ollama

1. Visit `https://ollama.com`.
2. Download and install Ollama for Windows.
3. Start Ollama.
4. Pull at least one model, for example:

```bat
ollama pull llama3.2
```

CoOllama lists installed models from `http://localhost:11434/api/tags`.

## Install CoOllama

1. Download or clone this repo.
2. Run `install.bat`.
3. Start the local server from Start Menu -> `CoOllama` -> `Start CoOllama Server`.
4. Keep the server window open while using the add-in.

The server should print:

```text
URL: https://localhost:8765/pane/taskpane.html
```

The installer also creates a Windows Startup shortcut so the local server starts when you log in.

## Add To New Outlook

1. Open `https://aka.ms/olksideload`.
2. Sign in with the same Microsoft account used by New Outlook.
3. Go to `My add-ins` -> `Custom Addins` -> `Add a custom add-in` -> `Add from File`.
4. Select:

```bat
%LOCALAPPDATA%\OllamaOutlookAddin\manifest.xml
```

5. Open New Outlook and open an email or compose window.
6. Look for **CoOllama** under `Apps`, `More apps`, the puzzle-piece icon, or the message/compose `...` menu.
7. Pin the add-in where Outlook shows a pin control.

## Current Features

- Saved model selection, with models sorted alphabetically.
- **Unload** button to request Ollama unload the selected model from memory.
- Light/dark theme toggle with saved preference.
- Multiple summary modes: Bulleted Summary, Short Summary, Detailed Review, and Suggested Response.
- Markdown preview rendering for model responses.
- Draft Reply and Draft in Your Voice.
- Current-email checker for tone, clarity, missing context, risk, and next actions.
- Spell and grammar review.
- Suspicious Email Check for visible phishing, impersonation, fraud, mismatch, and caution cues.
- Compose rewrite/apply support.
- Compose Insert support with sanitized HTML so paragraphs and bullets survive in Outlook.
- Sticky recent-item work state for reply notes, Ask Anything prompt, summary mode, and latest result.
- Voice profile builder with manual Sent sample workflow, sticky sample drafts, Clear Samples, and Copy Profile.

## Known Alpha Limits

- CoOllama works on the current selected/open Outlook item. It is not mailbox-wide memory.
- Pinning works in tested New Outlook, but Outlook still behaves like a one-active-item-at-a-time surface.
- For voice samples, open Sent Mail messages one at a time and click **Add current sample**, or paste samples separated by `---`.
- Automatic Sent-folder import and mailbox-wide Q/A require a future Microsoft Graph permission/auth flow.
- The local HTTPS certificate is self-signed. New Outlook uses Windows trust; Firefox may still warn if you open the pane URL directly.
- Suspicious Email Check is not a virus scanner and cannot inspect hidden headers, attachment contents, or real link destinations beyond visible text.

## Privacy And Storage

CoOllama does not include a hosted backend. Email context is sent by the pane to the local Ollama API at `http://localhost:11434`.

The pane stores these values locally in browser/Outlook WebView storage:

- selected model
- theme preference
- voice profile
- voice sample draft text
- voice sample size
- recent current-item work state, capped to 20 item keys

Use **Clear** to remove saved work for the current item. Use **Clear samples** to remove saved voice sample drafts. See [SECURITY.md](SECURITY.md) for more detail.

## Troubleshooting

If Outlook says installation failed or takes too long, make sure the CoOllama server is running and these URLs open on the same machine:

```text
https://localhost:8765/pane/taskpane.html
https://localhost:8765/assets/coollama-app-64.png
```

During install, CoOllama checks whether Ollama is on PATH and whether `http://127.0.0.1:11434/api/tags` answers. If Ollama is installed but not responding during install, the install can still complete; start Ollama before using the pane.

If the pane loads but actions are disabled, check that Ollama is running and at least one model is installed:

```bat
ollama pull llama3.2
```

The Outlook add-in button cannot silently start a Windows process from inside Outlook's sandbox. CoOllama handles this by creating a Startup shortcut during install and keeping the manual Start Menu path available.

## Compose Testing Checklist

1. Open a compose window and confirm the mode pill says `Compose`.
2. Select draft body text, click **Rewrite selection**, then **Apply rewrite**.
3. Generate a draft reply, place the cursor in the compose body, then click **Insert**.
4. Confirm inserted drafts keep paragraph breaks and bullets, without raw Markdown.
5. Switch pinned-pane items and confirm old generated output is not insertable into the new item.

## Uninstall

Run `uninstall.bat`. It removes copied add-in files, shortcuts, startup entry, and local CoOllama certificates.

If you manually added the add-in through Outlook, also remove it from Outlook's `My add-ins` page:

```text
https://aka.ms/olksideload
```

## License

MIT


