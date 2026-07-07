/* Little Fu — Excel chatbot add-in
   Talks to OpenRouter's chat completions API and only answers Excel questions,
   always as simple step-by-step instructions. */

Office.onReady(() => {
  initApp();
});

const STORAGE_KEY_API = "littlefu_api_key";
const STORAGE_KEY_MODEL = "littlefu_model";
const MAX_HISTORY_MESSAGES = 12; // keep the payload small; older turns are dropped

const SYSTEM_PROMPT = `You are "Little Fu", a friendly fox mascot chatbot that lives inside Microsoft Excel as an add-in.

Rules you must always follow:
1. You ONLY answer questions about Microsoft Excel (formulas, features, formatting, charts, shortcuts, data tools, VBA/macros, Power Query, troubleshooting, etc.).
2. If the user asks about anything unrelated to Excel, politely decline and steer them back, e.g. "I only help with Excel things! Ask me how to do something in a spreadsheet."
3. Always explain in a simple, friendly, non-technical tone, assuming the user may be a beginner.
4. Whenever you explain how to do something, ALWAYS format it as a clear numbered step-by-step list (Step 1, Step 2, ...). Keep each step short and concrete (mention exact menu names, buttons, or formulas).
5. When relevant, give the exact formula or keyboard shortcut in inline code formatting using backticks.
6. Keep responses concise — prefer short steps over long paragraphs. Avoid unnecessary preamble.
7. You may use light, friendly personality (you're a little fox named Fu), but never let personality get in the way of clear instructions.`;

let apiKey = "";
let model = "openai/gpt-4o-mini";
let conversation = []; // {role: 'user'|'assistant', content: string}
let isSending = false;

let els = {};

function initApp() {
  els = {
    settingsBtn: document.getElementById("settings-btn"),
    settingsPanel: document.getElementById("settings-panel"),
    apiKeyInput: document.getElementById("api-key-input"),
    modelSelect: document.getElementById("model-select"),
    saveSettingsBtn: document.getElementById("save-settings-btn"),
    closeSettingsBtn: document.getElementById("close-settings-btn"),
    chatWindow: document.getElementById("chat-window"),
    chatMessages: document.getElementById("chat-messages"),
    chatInput: document.getElementById("chat-input"),
    sendBtn: document.getElementById("send-btn"),
  };

  // Load saved settings
  apiKey = localStorage.getItem(STORAGE_KEY_API) || "";
  model = localStorage.getItem(STORAGE_KEY_MODEL) || model;
  els.apiKeyInput.value = apiKey;
  els.modelSelect.value = model;

  // If no key yet, open settings automatically
  if (!apiKey) {
    els.settingsPanel.classList.remove("hidden");
  }

  els.settingsBtn.addEventListener("click", () => {
    els.settingsPanel.classList.toggle("hidden");
  });

  els.closeSettingsBtn.addEventListener("click", () => {
    els.settingsPanel.classList.add("hidden");
  });

  els.saveSettingsBtn.addEventListener("click", () => {
    apiKey = els.apiKeyInput.value.trim();
    model = els.modelSelect.value;
    localStorage.setItem(STORAGE_KEY_API, apiKey);
    localStorage.setItem(STORAGE_KEY_MODEL, model);
    els.settingsPanel.classList.add("hidden");
    addAssistantNotice("Settings saved. Ask me anything about Excel! 🦊");
  });

  els.sendBtn.addEventListener("click", sendMessage);

  els.chatInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });

  els.chatInput.addEventListener("input", () => {
    els.chatInput.style.height = "auto";
    els.chatInput.style.height = Math.min(els.chatInput.scrollHeight, 90) + "px";
  });
}

function addAssistantNotice(text) {
  const div = document.createElement("div");
  div.className = "message assistant";
  div.innerHTML = `<div class="bubble">${escapeHtml(text)}</div>`;
  els.chatMessages.appendChild(div);
  scrollToBottom();
}

function addMessageBubble(role, htmlContent, isError) {
  const div = document.createElement("div");
  div.className = "message " + role;
  const bubble = document.createElement("div");
  bubble.className = "bubble" + (isError ? " error" : "");
  bubble.innerHTML = htmlContent;
  div.appendChild(bubble);
  els.chatMessages.appendChild(div);
  scrollToBottom();
  return bubble;
}

function scrollToBottom() {
  els.chatWindow.scrollTop = els.chatWindow.scrollHeight;
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/* Very small markdown-ish renderer: bold, inline code, numbered/bulleted lists, line breaks */
function renderMarkdown(text) {
  let escaped = escapeHtml(text);

  // inline code `...`
  escaped = escaped.replace(/`([^`]+)`/g, "<code>$1</code>");
  // bold **...**
  escaped = escaped.replace(/\*\*([^*]+)\*\*/g, "<b>$1</b>");

  const lines = escaped.split("\n");
  let html = "";
  let listType = null; // 'ol' | 'ul' | null

  const closeList = () => {
    if (listType) {
      html += listType === "ol" ? "</ol>" : "</ul>";
      listType = null;
    }
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();
    const numberedMatch = line.match(/^(\d+)[.)]\s+(.*)$/);
    const bulletMatch = line.match(/^[-*•]\s+(.*)$/);

    if (numberedMatch) {
      if (listType !== "ol") {
        closeList();
        html += "<ol>";
        listType = "ol";
      }
      html += `<li>${numberedMatch[2]}</li>`;
    } else if (bulletMatch) {
      if (listType !== "ul") {
        closeList();
        html += "<ul>";
        listType = "ul";
      }
      html += `<li>${bulletMatch[1]}</li>`;
    } else if (line === "") {
      closeList();
      html += "<br/>";
    } else {
      closeList();
      html += line + "<br/>";
    }
  }
  closeList();
  return html;
}

async function sendMessage() {
  if (isSending) return;

  const text = els.chatInput.value.trim();
  if (!text) return;

  if (!apiKey) {
    els.settingsPanel.classList.remove("hidden");
    addAssistantNotice("Please add your OpenRouter API key first (tap ⚙️).");
    return;
  }

  // Show user message
  addMessageBubble("user", escapeHtml(text));
  conversation.push({ role: "user", content: text });
  els.chatInput.value = "";
  els.chatInput.style.height = "auto";

  // Typing indicator
  const typingBubble = addMessageBubble(
    "assistant",
    `<div class="typing-dots"><span></span><span></span><span></span></div>`
  );

  isSending = true;
  els.sendBtn.disabled = true;

  try {
    const reply = await callOpenRouter(conversation);
    typingBubble.innerHTML = renderMarkdown(reply);
    conversation.push({ role: "assistant", content: reply });

    // Trim history so payload stays small
    if (conversation.length > MAX_HISTORY_MESSAGES) {
      conversation = conversation.slice(conversation.length - MAX_HISTORY_MESSAGES);
    }
  } catch (err) {
    typingBubble.parentElement.querySelector(".bubble").classList.add("error");
    typingBubble.innerHTML = escapeHtml(friendlyErrorMessage(err));
  } finally {
    isSending = false;
    els.sendBtn.disabled = false;
    scrollToBottom();
  }
}

function friendlyErrorMessage(err) {
  const msg = (err && err.message) || String(err);
  if (msg.includes("401")) {
    return "Your API key looks invalid or expired. Please check it in Settings (⚙️).";
  }
  if (msg.includes("429")) {
    return "Little Fu is a bit out of breath (rate limit reached). Please wait a moment and try again.";
  }
  if (msg.includes("Failed to fetch") || msg.includes("NetworkError")) {
    return "Couldn't reach OpenRouter. Please check your internet connection and try again.";
  }
  return "Something went wrong: " + msg;
}

async function callOpenRouter(history) {
  const messages = [{ role: "system", content: SYSTEM_PROMPT }, ...history];

  const response = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
      "HTTP-Referer": "https://littlefu.local",
      "X-Title": "Little Fu Excel Add-in",
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.3,
      max_tokens: 700,
    }),
  });

  if (!response.ok) {
    const text = await response.text().catch(() => "");
    throw new Error(`${response.status} ${response.statusText} ${text}`);
  }

  const data = await response.json();
  const content = data?.choices?.[0]?.message?.content;
  if (!content) {
    throw new Error("Empty response from model");
  }
  return content.trim();
}
