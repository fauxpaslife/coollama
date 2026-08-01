(function () {
  "use strict";

  const OLLAMA_URL = "http://localhost:11434";
  const MODEL_KEY = "coollamaSelectedModel";
  const THEME_KEY = "coollamaTheme";
  const VOICE_KEY = "askOllamaVoiceProfile";
  const VOICE_SAMPLES_KEY = "coollamaVoiceSamples";
  const VOICE_SAMPLE_LIMIT_KEY = "coollamaVoiceSampleLimit";
  const ITEM_STATE_KEY = "coollamaItemStates";
  const MAX_ITEM_STATES = 20;
  const tones = [
    {
      label: "Blunt",
      instruction: "State things directly with no softening language. Cut hedges like 'I think' or 'just wanted to.' Short sentences. It's fine to sound a little terse."
    },
    {
      label: "Neutral",
      instruction: "No tone adjustment - use the base voice as-is."
    },
    {
      label: "Diplomatic",
      instruction: "Soften direct statements slightly. Add brief acknowledgment of the other person's position before making a point. Avoid words that could read as blaming."
    },
    {
      label: "Groveling",
      instruction: "Add extra hedging, apology, and deference. Over-explain reasoning. Use phrases like 'I completely understand if...' and 'no worries at all if...'"
    }
  ];

  const summaryModes = {
    bullets: "Create a bulleted summary. Include key facts, asks, deadlines, decisions, and any needed next action. Keep it compact.",
    short: "Create a short plain-language summary in 2-4 sentences. Focus on what matters most.",
    detailed: "Create a detailed review. Include context, important details, risks, open questions, and recommended next steps when present.",
    suggestions: "Suggest a practical response strategy. Include likely intent, what to answer, what to avoid, and 3-5 concrete reply bullets. Do not draft the full email unless asked."
  };

  const state = {
    item: null,
    isCompose: false,
    email: { subject: "", sender: "", body: "" },
    ollamaReady: false,
    selectedModel: "",
    pendingRewrite: "",
    pendingRewriteItemKey: "",
    rewriteSourceProperty: "",
    lastOutput: "",
    lastOutputItemKey: "",
    itemKey: "",
    officeReady: false
  };

  const el = {};

  document.addEventListener("DOMContentLoaded", () => {
    bindElements();
    bindEvents();
    loadTheme();
    loadVoiceProfile();
    updateTone();
    setActionsEnabled(false);
    setBanner("Checking Ollama...", "muted");
    loadModels();
    initializeOfficeContext();
  });

  function initializeOfficeContext() {
    if (!window.Office || typeof Office.onReady !== "function") {
      useBrowserFallback();
      return;
    }

    const fallbackTimer = window.setTimeout(useBrowserFallback, 2500);
    Office.onReady(() => {
      window.clearTimeout(fallbackTimer);
      if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        useBrowserFallback();
        return;
      }

      state.officeReady = true;
      loadCurrentOfficeItem(false);
      registerItemChangedHandler();
    });
  }


  function registerItemChangedHandler() {
    if (!Office.context.mailbox || typeof Office.context.mailbox.addHandlerAsync !== "function" || !Office.EventType || !Office.EventType.ItemChanged) {
      return;
    }
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      loadCurrentOfficeItem(true);
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("Could not register ItemChanged handler", result.error);
      }
    });
  }

  function loadCurrentOfficeItem(showLoadedMessage) {
    if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
      useBrowserFallback();
      return;
    }

    persistCurrentItemState();
    state.item = Office.context.mailbox.item;
    state.isCompose = isComposeItem(state.item);
    state.itemKey = getCurrentItemKey();
    resetGeneratedState();
    clearRewriteState();
    toggleComposeFeatures();
    updateComposeStatus();
    el.emailSubject.textContent = "Loading...";
    el.emailSender.textContent = "Loading...";
    hydrateEmailContext().then(() => {
      restoreCurrentItemState();
      if (showLoadedMessage) {
        setBanner("Loaded the current Outlook item.", "ok");
      }
      refreshActionState();
    });
  }

  function isComposeItem(item) {
    return Boolean(
      item &&
      typeof item.getSelectedDataAsync === "function" &&
      typeof item.setSelectedDataAsync === "function" &&
      item.body &&
      typeof item.body.setSelectedDataAsync === "function"
    );
  }

  function getCurrentItemKey() {
    if (!state.item) {
      return "";
    }
    return state.item.itemId || state.item.conversationId || `${state.email.subject}|${Date.now()}`;
  }

  function clearRewriteState() {
    state.pendingRewrite = "";
    state.pendingRewriteItemKey = "";
    state.rewriteSourceProperty = "";
    el.rewriteBefore.textContent = "";
    el.rewriteAfter.textContent = "";
    el.applyRewriteBtn.disabled = true;
    el.rewritePanel.classList.add("hidden");
  }

  function resetGeneratedState() {
    state.lastOutput = "";
    state.lastOutputItemKey = "";
    updateInsertResultState();
  }

  function updateComposeStatus() {
    if (!el.composeStatus) {
      return;
    }
    el.composeStatus.textContent = state.isCompose ? "Compose" : "Read";
  }
  function useBrowserFallback() {
    if (state.item) {
      return;
    }
    state.isCompose = false;
    toggleComposeFeatures();
    state.email.subject = "Browser preview";
    state.email.sender = "Outlook context unavailable";
    state.email.body = "Open this pane from Outlook to inject the selected message context.";
    el.emailSubject.textContent = state.email.subject;
    el.emailSender.textContent = state.email.sender;
  }

  function bindElements() {
    [
      "refreshModels", "themeToggle", "statusBanner", "troublePanel", "troubleTitle", "troubleMessage", "troubleSteps", "retryConnectionBtn", "modelSelect", "unloadModelBtn", "summaryMode", "refreshEmailBtn", "emailSubject", "emailSender", "composeStatus",
      "summarizeBtn", "draftBtn", "draftVoiceBtn", "checkBtn", "grammarBtn", "suspiciousBtn", "rewriteBtn", "replyBullets",
      "rewritePanel", "rewriteBefore", "rewriteAfter", "applyRewriteBtn", "resultBox", "insertResultBtn", "clearWorkBtn", "copyResultBtn",
      "chatInput", "chatBtn", "voiceSamples", "voiceSampleLimit", "addCurrentEmailVoiceBtn", "buildVoiceBtn", "clearVoiceSamplesBtn", "copyVoiceProfileBtn",
      "voiceProfile", "toneSlider", "toneLabel"
    ].forEach((id) => {
      el[id] = document.getElementById(id);
    });
  }

  function bindEvents() {
    el.refreshModels.addEventListener("click", loadModels);
    el.unloadModelBtn.addEventListener("click", unloadSelectedModel);
    el.themeToggle.addEventListener("click", toggleTheme);
    el.retryConnectionBtn.addEventListener("click", loadModels);
    el.refreshEmailBtn.addEventListener("click", () => loadCurrentOfficeItem(true));
    el.modelSelect.addEventListener("change", () => {
      state.selectedModel = el.modelSelect.value;
      if (state.selectedModel) {
        localStorage.setItem(MODEL_KEY, state.selectedModel);
      }
      refreshActionState();
    });
    el.summarizeBtn.addEventListener("click", summarizeEmail);
    el.draftBtn.addEventListener("click", () => draftReply(false));
    el.draftVoiceBtn.addEventListener("click", () => draftReply(true));
    el.checkBtn.addEventListener("click", checkEmail);
    el.grammarBtn.addEventListener("click", checkGrammar);
    el.suspiciousBtn.addEventListener("click", checkSuspicious);
    el.rewriteBtn.addEventListener("click", rewriteSelection);
    el.applyRewriteBtn.addEventListener("click", applyRewrite);
    el.insertResultBtn.addEventListener("click", insertResultIntoCompose);
    el.clearWorkBtn.addEventListener("click", clearCurrentItemWork);
    el.copyResultBtn.addEventListener("click", copyResult);
    el.chatBtn.addEventListener("click", sendChat);
    el.addCurrentEmailVoiceBtn.addEventListener("click", addCurrentEmailToVoiceSamples);
    el.buildVoiceBtn.addEventListener("click", buildVoiceProfile);
    el.clearVoiceSamplesBtn.addEventListener("click", clearVoiceSamples);
    el.copyVoiceProfileBtn.addEventListener("click", copyVoiceProfile);
    el.voiceProfile.addEventListener("input", () => {
      localStorage.setItem(VOICE_KEY, el.voiceProfile.value.trim());
    });
    el.voiceSamples.addEventListener("input", saveVoiceSamples);
    el.voiceSampleLimit.addEventListener("input", saveVoiceSampleLimit);
    el.replyBullets.addEventListener("input", persistCurrentItemState);
    el.chatInput.addEventListener("input", persistCurrentItemState);
    el.summaryMode.addEventListener("change", persistCurrentItemState);
    el.toneSlider.addEventListener("input", updateTone);
  }

  async function loadModels() {
    setActionsEnabled(false);
    setBanner("Checking Ollama...", "muted");
    el.modelSelect.innerHTML = '<option value="">Loading models...</option>';

    try {
      const response = await fetch(`${OLLAMA_URL}/api/tags`);
      if (!response.ok) {
        throw new Error(`Ollama returned ${response.status}`);
      }
      const data = await response.json();
      const models = (Array.isArray(data.models) ? data.models : []).slice().sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" }));
      const savedModel = localStorage.getItem(MODEL_KEY) || "";

      el.modelSelect.innerHTML = "";
      if (!models.length) {
        addModelOption("", "No models installed");
        setBanner("Ollama is running, but no models were found.", "error");
        state.ollamaReady = false;
        refreshActionState();
        return;
      }

      models.forEach((model) => addModelOption(model.name, model.name));
      state.selectedModel = models.some((model) => model.name === savedModel) ? savedModel : models[0].name;
      el.modelSelect.value = state.selectedModel;
      localStorage.setItem(MODEL_KEY, state.selectedModel);
      state.ollamaReady = true;
      hideTrouble();
      setBanner("Connected to Ollama.", "ok");
    } catch (error) {
      state.ollamaReady = false;
      state.selectedModel = "";
      el.modelSelect.innerHTML = '<option value="">Is Ollama running?</option>';
      setBanner("Could not reach Ollama at localhost:11434. Is Ollama running?", "error");
      showTrouble("Ollama is not responding", "The CoOllama pane is loaded, but the local Ollama model server is not answering.", ["Start Ollama from the Windows Start Menu or system tray.", "If Ollama is new, install it from ollama.com and pull at least one model.", "Confirm http://localhost:11434/api/tags opens or click Retry here."]);
    }

    refreshActionState();
  }


  function getItemStateMap() {
    try {
      const saved = JSON.parse(localStorage.getItem(ITEM_STATE_KEY) || "{}");
      return saved && typeof saved === "object" ? saved : {};
    } catch (error) {
      return {};
    }
  }

  function setItemStateMap(map) {
    const entries = Object.entries(map)
      .sort((a, b) => (b[1].updatedAt || 0) - (a[1].updatedAt || 0))
      .slice(0, MAX_ITEM_STATES);
    localStorage.setItem(ITEM_STATE_KEY, JSON.stringify(Object.fromEntries(entries)));
  }

  function persistCurrentItemState() {
    if (!state.itemKey || !el.replyBullets || !el.chatInput || !el.resultBox) {
      return;
    }
    const map = getItemStateMap();
    map[state.itemKey] = {
      replyBullets: el.replyBullets.value,
      chatInput: el.chatInput.value,
      summaryMode: el.summaryMode.value,
      lastOutput: state.lastOutputItemKey === state.itemKey ? state.lastOutput : "",
      updatedAt: Date.now()
    };
    setItemStateMap(map);
  }

  function restoreCurrentItemState() {
    const saved = state.itemKey ? getItemStateMap()[state.itemKey] : null;
    el.replyBullets.value = saved && typeof saved.replyBullets === "string" ? saved.replyBullets : "";
    el.chatInput.value = saved && typeof saved.chatInput === "string" ? saved.chatInput : "";
    if (saved && saved.summaryMode) {
      el.summaryMode.value = saved.summaryMode;
    }
    state.lastOutput = saved && typeof saved.lastOutput === "string" ? saved.lastOutput : "";
    state.lastOutputItemKey = state.lastOutput ? state.itemKey : "";
    if (state.lastOutput) {
      renderMarkdown(state.lastOutput, el.resultBox);
    } else {
      setPlainTarget(el.resultBox, "Ready when Outlook and Ollama are.");
    }
    updateInsertResultState();
  }

  function clearCurrentItemWork() {
    if (state.itemKey) {
      const map = getItemStateMap();
      delete map[state.itemKey];
      setItemStateMap(map);
    }
    el.replyBullets.value = "";
    el.chatInput.value = "";
    state.lastOutput = "";
    state.lastOutputItemKey = "";
    clearRewriteState();
    setPlainTarget(el.resultBox, "Cleared saved work for this Outlook item.");
    updateInsertResultState();
  }

  async function unloadSelectedModel() {
    const model = state.selectedModel;
    if (!model) {
      setBanner("Choose a model to unload.", "error");
      return;
    }
    el.unloadModelBtn.disabled = true;
    setBanner(`Unloading ${model} from Ollama memory...`, "muted");
    try {
      const response = await fetch(`${OLLAMA_URL}/api/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model, prompt: "", stream: false, keep_alive: 0 })
      });
      if (!response.ok) {
        throw new Error(`Ollama returned ${response.status}`);
      }
      setBanner(`Requested unload for ${model}. It will load again on the next request.`, "ok");
    } catch (error) {
      setBanner(`Could not unload ${model}: ${error.message || error}`, "error");
    } finally {
      refreshActionState();
    }
  }

  function addModelOption(value, label) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label;
    el.modelSelect.appendChild(option);
  }

  async function hydrateEmailContext() {
    state.email.subject = await getSubject();
    state.email.sender = getSender();
    state.email.body = await getBody();
    if (state.item && !state.item.itemId && !state.item.conversationId) {
      state.itemKey = `${state.email.subject}|${state.email.sender}|${state.isCompose ? "compose" : "read"}`;
      updateInsertResultState();
    }
    el.emailSubject.textContent = state.email.subject || "(No subject)";
    el.emailSender.textContent = state.email.sender || "(Unknown sender)";
  }

  function getSubject() {
    if (!state.item) {
      return Promise.resolve("");
    }
    if (typeof state.item.subject === "string") {
      return Promise.resolve(state.item.subject);
    }
    if (state.item.subject && typeof state.item.subject.getAsync === "function") {
      return promisifyAsync(state.item.subject.getAsync.bind(state.item.subject));
    }
    return Promise.resolve("");
  }

  function getSender() {
    const sender = state.item.from || state.item.sender || {};
    if (typeof sender === "string") {
      return sender;
    }
    return sender.displayName || sender.emailAddress || "";
  }

  function getBody() {
    if (!state.item || !state.item.body || typeof state.item.body.getAsync !== "function") {
      return Promise.resolve("");
    }
    return promisifyAsync((callback) => {
      state.item.body.getAsync(Office.CoercionType.Text, callback);
    });
  }

  function promisifyAsync(fn) {
    return new Promise((resolve) => {
      fn((result) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || "" : "");
      });
    });
  }

  function toggleComposeFeatures() {
    document.querySelectorAll(".compose-only").forEach((node) => {
      node.classList.toggle("hidden", !state.isCompose);
    });
  }

  function refreshActionState() {
    setActionsEnabled(state.ollamaReady && Boolean(state.selectedModel));
  }

  function setActionsEnabled(enabled) {
    [el.summarizeBtn, el.draftBtn, el.draftVoiceBtn, el.checkBtn, el.grammarBtn, el.suspiciousBtn, el.chatBtn, el.buildVoiceBtn, el.addCurrentEmailVoiceBtn].forEach((button) => {
      button.disabled = !enabled;
    });
    el.rewriteBtn.disabled = !enabled || !state.isCompose;
    el.refreshEmailBtn.disabled = !state.officeReady;
    updateInsertResultState();
  }


  function showTrouble(title, message, steps) {
    el.troubleTitle.textContent = title;
    el.troubleMessage.textContent = message;
    el.troubleSteps.innerHTML = "";
    steps.forEach((step) => {
      const li = document.createElement("li");
      li.textContent = step;
      el.troubleSteps.appendChild(li);
    });
    el.troublePanel.classList.remove("hidden");
  }

  function hideTrouble() {
    el.troublePanel.classList.add("hidden");
  }

  function updateInsertResultState() {
    el.insertResultBtn.disabled = !state.isCompose || !state.lastOutput.trim() || state.lastOutputItemKey !== state.itemKey;
  }
  function setBanner(message, type) {
    el.statusBanner.textContent = message;
    el.statusBanner.className = `banner is-${type}`;
  }

  function emailContextBlock() {
    return [
      `Subject: ${state.email.subject || "(No subject)"}`,
      `From: ${state.email.sender || "(Unknown sender)"}`,
      "",
      "Email body:",
      state.email.body || "(No body available)"
    ].join("\n");
  }

  async function summarizeEmail() {
    const mode = el.summaryMode.value || "bullets";
    await runChat([
      { role: "system", content: `${summaryModes[mode]} Do not invent facts. Format the answer in clean Markdown.` },
      { role: "user", content: emailContextBlock() }
    ]);
  }

  async function draftReply(useVoice) {
    const notes = el.replyBullets.value.trim();
    const voice = el.voiceProfile.value.trim();
    const system = useVoice && voice
      ? `${voice}\n\nDraft a concise Outlook email reply in this voice. Use the user's notes as requirements and do not invent facts. Return only the reply body. Do not include a subject line, markdown horizontal rules, code fences, or placeholder signatures like [Your Name]. Use light Markdown only if it improves readability in the preview.`
      : "Draft a concise Outlook email reply. Use the user's notes as requirements and do not invent facts. Return only the reply body. Do not include a subject line, markdown horizontal rules, code fences, or placeholder signatures like [Your Name]. Use light Markdown only if it improves readability in the preview.";

    if (useVoice && !voice) {
      setResult("Build or paste a voice profile first, then try Draft in your voice.");
      return;
    }

    await runChat([
      { role: "system", content: system },
      { role: "user", content: `${emailContextBlock()}\n\nReply notes:\n${notes || "(No extra notes provided.)"}` }
    ]);
  }

  async function checkEmail() {
    await runChat([
      { role: "system", content: "Review this email for clarity, tone, risk, missing context, likely intent, and recommended next actions. Use concise Markdown headings and bullets. Do not invent facts." },
      { role: "user", content: emailContextBlock() }
    ]);
  }

  async function checkGrammar() {
    await runChat([
      { role: "system", content: "Proofread the email text for spelling, grammar, punctuation, and awkward phrasing. Return a corrected version first, then a short bullet list of notable changes. Preserve meaning and facts exactly. Use clean Markdown." },
      { role: "user", content: emailContextBlock() }
    ]);
  }


  async function checkSuspicious() {
    await runChat([
      {
        role: "system",
        content: "Review this email for phishing, impersonation, fraud, and suspicious-request cues. This is not a virus scan. You cannot inspect hidden headers, attachments, or actual destination URLs unless they are visible in the email body. Use only the visible sender, subject, and body content provided. Return clean Markdown with these sections: Caution level: Low, Medium, or High; Reasons; Mismatches or red flags; What to verify before responding; Safest next action. Be specific but do not exaggerate. Do not claim the email is malicious unless the visible evidence strongly supports it."
      },
      { role: "user", content: emailContextBlock() }
    ]);
  }
  async function sendChat() {
    const prompt = el.chatInput.value.trim();
    if (!prompt) {
      setResult("Type a question first.");
      return;
    }
    await runChat([
      { role: "system", content: "Answer using the provided email context. Say when the context does not contain the answer. Format in readable Markdown." },
      { role: "user", content: `${emailContextBlock()}\n\nQuestion:\n${prompt}` }
    ]);
  }

  async function rewriteSelection() {
    if (!state.isCompose) {
      setResult("Open CoOllama from a compose window to rewrite selected text.");
      return;
    }

    const selection = await getSelectedText();
    const selectedText = selection.data || "";
    if (!selectedText.trim()) {
      const where = selection.sourceProperty ? ` Current cursor/selection is in ${selection.sourceProperty}.` : "";
      setResult(`Select text in the compose body or subject first.${where}`);
      return;
    }

    el.rewritePanel.classList.remove("hidden");
    el.rewriteBefore.textContent = selectedText;
    el.rewriteAfter.textContent = "";
    el.applyRewriteBtn.disabled = true;

    const voice = el.voiceProfile.value.trim() || "Write in a style that is clear, professional, concise, and natural.";
    const tone = tones[Number(el.toneSlider.value)];
    await runChat([
      { role: "system", content: `${voice}\n\nTone adjustment: ${tone.instruction}\n\nTone must only affect word choice and softness - never add claims or drop information.` },
      { role: "user", content: `Rewrite the following text to match this voice and tone. Preserve the original meaning and any factual content exactly. Only change tone, phrasing, and structure.\n\n${selectedText}` }
    ], {
      target: el.rewriteAfter,
      renderMarkdown: false,
      onDone: (text) => {
        state.pendingRewrite = text.trim();
        state.pendingRewriteItemKey = state.itemKey;
        state.rewriteSourceProperty = selection.sourceProperty || "selection";
        el.applyRewriteBtn.disabled = !state.pendingRewrite;
      }
    });
  }
  function getSelectedText() {
    if (!state.item || typeof state.item.getSelectedDataAsync !== "function") {
      return Promise.resolve({ data: "", sourceProperty: "" });
    }
    return new Promise((resolve) => {
      state.item.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve({
            data: result.value && typeof result.value.data === "string" ? result.value.data : "",
            sourceProperty: result.value && result.value.sourceProperty ? result.value.sourceProperty : ""
          });
        } else {
          const message = result.error && result.error.message ? result.error.message : "Could not read the selected text.";
          setResult(`Could not read selection: ${message}`);
          resolve({ data: "", sourceProperty: "" });
        }
      });
    });
  }
  function applyRewrite() {
    if (!state.pendingRewrite) {
      setResult("Run Rewrite selection first.");
      return;
    }
    if (state.pendingRewriteItemKey !== state.itemKey) {
      setResult("This rewrite was created for a different Outlook item. Refresh and run Rewrite selection again.");
      clearRewriteState();
      return;
    }
    if (!state.item || typeof state.item.setSelectedDataAsync !== "function") {
      setResult("Open CoOllama from a compose window before applying a rewrite.");
      return;
    }
    state.item.setSelectedDataAsync(state.pendingRewrite, { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setBanner(`Rewrite applied to ${state.rewriteSourceProperty || "selection"}.`, "ok");
        clearRewriteState();
      } else {
        setResult(`Could not apply rewrite: ${result.error && result.error.message ? result.error.message : "Unknown error"}`);
      }
    });
  }
  function addCurrentEmailToVoiceSamples() {
    const body = truncateWords(state.email.body || "", 300);
    if (!body) {
      setResult("No current email body is available to add as a sample.");
      return;
    }
    const existing = el.voiceSamples.value.trim();
    el.voiceSamples.value = existing ? `${existing}\n---\n${body}` : body;
    saveVoiceSamples();
    setResult("Added the current email body to the voice samples. Open Sent Mail items and add several sent messages for best voice results.");
  }

  async function buildVoiceProfile() {
    saveVoiceSamples();
    saveVoiceSampleLimit();
    const samples = prepareSamples(el.voiceSamples.value);
    if (!samples) {
      setResult("Paste sample sent emails first.");
      return;
    }

    await runChat([
      {
        role: "system",
        content: "You are analyzing a collection of emails written by one person to extract their writing voice as a reusable style guide. Do not summarize the content or topics - ignore what they're saying and focus entirely on HOW they say it.\n\nAnalyze the samples for:\n- Sentence length and rhythm (short/punchy vs. long/winding)\n- Formality level (casual contractions vs. formal phrasing)\n- Greeting/sign-off patterns\n- Punctuation habits\n- Vocabulary tendencies (jargon, filler words, recurring phrases)\n- Directness (hedging vs. flat statements)\n- Paragraph structure (blocks vs. short chunks)\n\nOutput ONLY a style guide paragraph (150-250 words) written as instructions for another writer to imitate this voice. Do not include any of the original email content. Do not add commentary. Start directly with \"Write in a style that is...\""
      },
      { role: "user", content: samples }
    ], {
      onDone: (text) => {
        el.voiceProfile.value = text.trim();
        localStorage.setItem(VOICE_KEY, el.voiceProfile.value);
      }
    });
  }

  function prepareSamples(raw) {
    const sampleLimit = clamp(Number(el.voiceSampleLimit.value) || 15, 1, 20);
    const parts = raw
      .split(/\n-{3,}\n/g)
      .map((sample) => truncateWords(sample.trim(), 300))
      .filter(Boolean)
      .slice(0, sampleLimit);
    return parts.length ? parts.join("\n---\n") : "";
  }

  function clamp(value, min, max) {
    return Math.min(max, Math.max(min, value));
  }

  function truncateWords(text, maxWords) {
    const words = text.split(/\s+/).filter(Boolean);
    return words.slice(0, maxWords).join(" ");
  }

  async function runChat(messages, options) {
    const target = options && options.target ? options.target : el.resultBox;
    const onDone = options && options.onDone;
    const shouldRenderMarkdown = !options || options.renderMarkdown !== false;
    let output = "";

    if (!state.selectedModel) {
      setResult("Choose a model first.");
      return "";
    }

    state.lastOutput = "";
    state.lastOutputItemKey = "";
    updateInsertResultState();
    target.textContent = "";
    setActionsEnabled(false);

    try {
      const response = await fetch(`${OLLAMA_URL}/api/chat`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: state.selectedModel, stream: true, messages })
      });

      if (!response.ok || !response.body) {
        throw new Error(`Ollama returned ${response.status}`);
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder();
      let buffer = "";

      while (true) {
        const { value, done } = await reader.read();
        if (done) {
          break;
        }
        buffer += decoder.decode(value, { stream: true });
        const lines = buffer.split("\n");
        buffer = lines.pop() || "";
        lines.forEach((line) => {
          if (!line.trim()) {
            return;
          }
          try {
            const payload = JSON.parse(line);
            const token = payload.message && payload.message.content ? payload.message.content : "";
            output += token;
            target.textContent = output;
          } catch (error) {
            output += "";
          }
        });
      }

      if (buffer.trim()) {
        const payload = JSON.parse(buffer);
        output += payload.message && payload.message.content ? payload.message.content : "";
      }

      if (target === el.resultBox) {
        state.lastOutput = output;
        state.lastOutputItemKey = state.itemKey;
        updateInsertResultState();
      }
      if (shouldRenderMarkdown && target === el.resultBox) {
        renderMarkdown(output, target);
      } else {
        target.textContent = output;
      }
      persistCurrentItemState();
      if (onDone) {
        onDone(output);
      }
    } catch (error) {
      setPlainTarget(target, `Ollama request failed. ${error.message || error}`);
    } finally {
      refreshActionState();
    }

    return output;
  }

  function setResult(message) {
    state.lastOutput = "";
    state.lastOutputItemKey = "";
    setPlainTarget(el.resultBox, message);
    updateInsertResultState();
    persistCurrentItemState();
  }

  function setPlainTarget(target, message) {
    target.textContent = message;
  }


  function insertResultIntoCompose() {
    const text = state.lastOutput.trim();
    if (!text) {
      setResult("There is no result to insert yet.");
      return;
    }
    if (!state.isCompose || !state.item || !state.item.body || typeof state.item.body.setSelectedDataAsync !== "function") {
      setResult("Open CoOllama from a compose window to insert text into a draft.");
      return;
    }
    if (state.lastOutputItemKey !== state.itemKey) {
      setResult("This result was generated for a different Outlook item. Refresh and generate it again before inserting.");
      resetGeneratedState();
      return;
    }
    const insertHtml = markdownToEmailHtml(text);
    const insertText = markdownToPlainEmailText(text);
    state.item.body.setSelectedDataAsync(insertHtml, { coercionType: Office.CoercionType.Html }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setBanner("Inserted result into the draft body.", "ok");
        return;
      }
      state.item.body.setSelectedDataAsync(insertText, { coercionType: Office.CoercionType.Text }, (fallbackResult) => {
        if (fallbackResult.status === Office.AsyncResultStatus.Succeeded) {
          setBanner("Inserted result as plain text.", "ok");
        } else {
          const error = fallbackResult.error || result.error;
          setResult(`Could not insert result: ${error && error.message ? error.message : "Unknown error"}`);
        }
      });
    });
  }
  function cleanMarkdownForEmailInsert(markdown) {
    let text = String(markdown || "").replace(/\r\n/g, "\n").replace(/\u00a0/g, " ");
    text = text.replace(/^\s*\*{0,2}\s*Subject\s*:\s*\*{0,2}.*$/gim, "");
    text = text.replace(/^\s*[-*_]{3,}\s*$/gm, "");
    text = text.replace(/^\s*\[Your Name\]\s*$/gim, "");
    text = text.replace(/```[\s\S]*?```/g, (block) => block.replace(/```[^\n]*\n?/g, "").replace(/```/g, ""));
    text = text.replace(/\n{3,}/g, "\n\n");
    return text.trim();
  }

  function markdownToEmailHtml(markdown) {
    const html = markdownToHtml(cleanMarkdownForEmailInsert(markdown));
    return `<div>${html}</div>`;
  }
  function markdownToPlainEmailText(markdown) {
    let text = cleanMarkdownForEmailInsert(markdown);
    text = text.replace(/^\s*#{1,6}\s+/gm, "");
    text = text.replace(/\*\*([^*]+)\*\*/g, "$1");
    text = text.replace(/__([^_]+)__/g, "$1");
    text = text.replace(/`([^`]+)`/g, "$1");
    text = text.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, "$1 ($2)");
    text = text.replace(/^\s*[-*]\s+/gm, "- ");
    text = text.replace(/\n{3,}/g, "\n\n");
    return text.trim();
  }
  function copyResult() {
    const text = state.lastOutput || el.resultBox.textContent;
    if (navigator.clipboard && text) {
      navigator.clipboard.writeText(text);
    }
  }


  function loadTheme() {
    const savedTheme = localStorage.getItem(THEME_KEY) || "light";
    applyTheme(savedTheme === "dark" ? "dark" : "light");
  }

  function toggleTheme() {
    const nextTheme = document.body.classList.contains("theme-dark") ? "light" : "dark";
    applyTheme(nextTheme);
  }

  function applyTheme(theme) {
    const isDark = theme === "dark";
    document.body.classList.toggle("theme-dark", isDark);
    el.themeToggle.textContent = isDark ? "Light" : "Dark";
    el.themeToggle.title = isDark ? "Switch to light mode" : "Switch to dark mode";
    el.themeToggle.setAttribute("aria-label", el.themeToggle.title);
    localStorage.setItem(THEME_KEY, isDark ? "dark" : "light");
  }
  function loadVoiceProfile() {
    el.voiceProfile.value = localStorage.getItem(VOICE_KEY) || "";
    el.voiceSamples.value = localStorage.getItem(VOICE_SAMPLES_KEY) || "";
    const savedLimit = Number(localStorage.getItem(VOICE_SAMPLE_LIMIT_KEY));
    if (savedLimit) {
      el.voiceSampleLimit.value = String(clamp(savedLimit, 1, 20));
    }
  }

  function saveVoiceSamples() {
    localStorage.setItem(VOICE_SAMPLES_KEY, el.voiceSamples.value.trim());
  }

  function saveVoiceSampleLimit() {
    const sampleLimit = clamp(Number(el.voiceSampleLimit.value) || 15, 1, 20);
    el.voiceSampleLimit.value = String(sampleLimit);
    localStorage.setItem(VOICE_SAMPLE_LIMIT_KEY, String(sampleLimit));
  }

  function clearVoiceSamples() {
    el.voiceSamples.value = "";
    localStorage.removeItem(VOICE_SAMPLES_KEY);
    setResult("Cleared saved voice samples. Your built voice profile is unchanged.");
  }

  function copyVoiceProfile() {
    const text = el.voiceProfile.value.trim();
    if (!text) {
      setResult("No voice profile to copy yet.");
      return;
    }
    if (navigator.clipboard) {
      navigator.clipboard.writeText(text);
      setBanner("Copied voice profile.", "ok");
    }
  }

  function updateTone() {
    el.toneLabel.textContent = tones[Number(el.toneSlider.value)].label;
  }

  function renderMarkdown(markdown, target) {
    target.innerHTML = markdownToHtml(markdown || "");
  }

  function markdownToHtml(markdown) {
    const lines = markdown.replace(/\r\n/g, "\n").split("\n");
    const html = [];
    let listType = "";
    let paragraph = [];
    let inCode = false;
    let codeLines = [];

    const closeParagraph = () => {
      if (paragraph.length) {
        html.push(`<p>${inlineMarkdown(paragraph.join(" "))}</p>`);
        paragraph = [];
      }
    };
    const closeList = () => {
      if (listType) {
        html.push(`</${listType}>`);
        listType = "";
      }
    };

    lines.forEach((rawLine) => {
      const line = rawLine.trimEnd();
      if (line.trim().startsWith("```")) {
        if (inCode) {
          html.push(`<pre><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
          codeLines = [];
          inCode = false;
        } else {
          closeParagraph();
          closeList();
          inCode = true;
        }
        return;
      }
      if (inCode) {
        codeLines.push(rawLine);
        return;
      }
      if (!line.trim()) {
        closeParagraph();
        closeList();
        return;
      }

      const heading = line.match(/^(#{1,3})\s+(.+)$/);
      if (heading) {
        closeParagraph();
        closeList();
        const level = heading[1].length;
        html.push(`<h${level}>${inlineMarkdown(heading[2])}</h${level}>`);
        return;
      }

      const bullet = line.match(/^[-*]\s+(.+)$/);
      if (bullet) {
        closeParagraph();
        if (listType !== "ul") {
          closeList();
          listType = "ul";
          html.push("<ul>");
        }
        html.push(`<li>${inlineMarkdown(bullet[1])}</li>`);
        return;
      }

      const numbered = line.match(/^\d+[.)]\s+(.+)$/);
      if (numbered) {
        closeParagraph();
        if (listType !== "ol") {
          closeList();
          listType = "ol";
          html.push("<ol>");
        }
        html.push(`<li>${inlineMarkdown(numbered[1])}</li>`);
        return;
      }

      closeList();
      paragraph.push(line.trim());
    });

    if (inCode) {
      html.push(`<pre><code>${escapeHtml(codeLines.join("\n"))}</code></pre>`);
    }
    closeParagraph();
    closeList();
    return html.join("") || "<p>No response.</p>";
  }

  function inlineMarkdown(text) {
    let html = escapeHtml(text);
    html = html.replace(/`([^`]+)`/g, "<code>$1</code>");
    html = html.replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>");
    html = html.replace(/__([^_]+)__/g, "<strong>$1</strong>");
    html = html.replace(/\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>');
    return html;
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
})();































