import {
  classify,
  countLabel,
  excerpt,
  extensionOf,
  formatBytes,
  groupCounts,
  isInline,
  matchesFilter,
  mimeType,
  safeFileName,
  summarize,
} from "../lib/summary.js";
import { officeText } from "../lib/ziptext.js";

const THUMB_LIMIT = 8 * 1024 * 1024;
const TEXT_CARD_LIMIT = 1024 * 1024;
const OFFICE_LIMIT = 8 * 1024 * 1024;
const PREVIEW_LIMIT = 25 * 1024 * 1024;
const TEXT_READ = 64 * 1024;
const HTML_READ = 256 * 1024;
const FETCH_LIMIT = 3;

const KIND_LABEL = {
  image: "Image",
  pdf: "PDF",
  text: "Text",
  document: "Document",
  audio: "Audio",
  video: "Video",
  archive: "Archive",
  other: "File",
};

const files = new WeakMap();
const liveUrls = new Set();
const queue = [];
let activeFetches = 0;
let renderGeneration = 0;

const state = {
  messages: [],
  truncated: false,
  filter: "all",
  showInline: false,
  preview: null,
  previewUrl: "",
};

const list = document.querySelector("#list");
const status = document.querySelector("#status");
const summaryBox = document.querySelector("#summary");
const filters = document.querySelector("#filters");
const truncated = document.querySelector("#truncated");
const dialog = document.querySelector("#preview");
const previewTitle = document.querySelector("#preview-title");
const previewBody = document.querySelector("#preview-body");
const previewError = document.querySelector("#preview-error");

function setStatus(text) {
  status.hidden = !text;
  status.textContent = text || "";
}

function setPreviewError(text) {
  previewError.hidden = !text;
  previewError.textContent = text || "";
}

function trackUrl(url) {
  liveUrls.add(url);
  return url;
}

function revokeLiveUrls() {
  for (const url of liveUrls) {
    URL.revokeObjectURL(url);
  }
  liveUrls.clear();
}

function clearPreviewUrl() {
  if (state.previewUrl) {
    URL.revokeObjectURL(state.previewUrl);
    state.previewUrl = "";
  }
}

function schedule(task) {
  return new Promise((resolve, reject) => {
    queue.push({ task, resolve, reject });
    pump();
  });
}

function pump() {
  while (activeFetches < FETCH_LIMIT && queue.length) {
    const job = queue.shift();
    activeFetches += 1;
    Promise.resolve()
      .then(job.task)
      .then(job.resolve, job.reject)
      .finally(() => {
        activeFetches -= 1;
        pump();
      });
  }
}

function loadFile(record) {
  if (!record.loading) {
    record.loading = schedule(() => browser.messages.getAttachmentFile(record.messageId, record.partName))
      .then(file => {
        record.file = file;
        return file;
      })
      .catch(error => {
        record.loading = null;
        throw error;
      });
  }
  return record.loading;
}

async function mapPool(items, limit, mapper) {
  const results = new Array(items.length);
  let next = 0;
  async function worker() {
    while (next < items.length) {
      const index = next;
      next += 1;
      results[index] = await mapper(items[index], index);
    }
  }
  const workers = Math.min(limit, items.length);
  await Promise.all(Array.from({ length: workers }, () => worker()));
  return results;
}

function messageIdFromUrl() {
  const raw = new URLSearchParams(location.search).get("messageId");
  if (!raw || !/^\d+$/.test(raw)) {
    return null;
  }
  const id = Number(raw);
  return Number.isSafeInteger(id) ? id : null;
}

function formatWhen(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

function isHtml(record) {
  return mimeType(record.contentType) === "text/html" || ["html", "htm"].includes(extensionOf(record.name));
}

function htmlToText(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ");
}

function looksBinary(text) {
  const sample = String(text || "").slice(0, 2000);
  if (!sample) {
    return false;
  }
  let odd = 0;
  for (let index = 0; index < sample.length; index++) {
    const code = sample.charCodeAt(index);
    if (code === 0) {
      return true;
    }
    if (code < 9 || (code > 13 && code < 32)) {
      odd += 1;
    }
  }
  return odd / sample.length > 0.1;
}

function includedAttachments() {
  const items = [];
  for (const message of state.messages) {
    for (const item of message.attachments) {
      if (!state.showInline && isInline(item)) {
        continue;
      }
      items.push(item);
    }
  }
  return items;
}

function visibleAttachments(message) {
  return message.attachments.filter(item => {
    if (!state.showInline && isInline(item)) {
      return false;
    }
    return matchesFilter(item.kind, state.filter);
  });
}

function needsCardFetch(record) {
  if (record.kind === "image") {
    return record.size === 0 || record.size <= THUMB_LIMIT;
  }
  if (record.kind === "text") {
    return record.size === 0 || record.size <= TEXT_CARD_LIMIT;
  }
  if (record.kind === "document") {
    return record.size === 0 || record.size <= OFFICE_LIMIT;
  }
  return false;
}

function hardenHtml(html) {
  const meta =
    '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; style-src \'unsafe-inline\'; img-src data:; font-src data:;">';
  if (/<head[^>]*>/i.test(html)) {
    return html.replace(/<head[^>]*>/i, match => match + meta);
  }
  return `<!DOCTYPE html><html><head>${meta}<meta charset="utf-8"></head><body>${html}</body></html>`;
}

function renderSummary(items) {
  const summary = summarize(items);
  const groups = groupCounts(summary.byKind);
  document.querySelector("#totals").textContent = `${countLabel(summary.count, "file", "files")} · ${summary.uniqueCount} unique · ${formatBytes(summary.totalBytes)}`;
  const parts = [];
  if (groups.image) {
    parts.push(countLabel(groups.image, "image", "images"));
  }
  if (groups.document) {
    parts.push(countLabel(groups.document, "document", "documents"));
  }
  if (groups.media) {
    parts.push(countLabel(groups.media, "media file", "media files"));
  }
  if (groups.other) {
    parts.push(countLabel(groups.other, "other file", "other files"));
  }
  document.querySelector("#kinds").textContent = parts.join(" · ");
  summaryBox.hidden = summary.count === 0;
  filters.hidden = summary.count === 0;
}

function render() {
  const generation = ++renderGeneration;
  observer.disconnect();
  revokeLiveUrls();
  list.replaceChildren();
  const items = includedAttachments();
  renderSummary(items);
  truncated.hidden = !state.truncated;

  const titled = state.messages.find(message => message.header && message.header.subject);
  const subject = titled ? titled.header.subject : "";
  document.querySelector("#subject").textContent = subject;
  document.title = subject ? `Attachments — ${subject}` : "Attachment gallery";

  let shown = 0;
  for (const message of state.messages) {
    const attachments = visibleAttachments(message);
    if (!attachments.length && !message.errorText) {
      continue;
    }
    shown += attachments.length;
    const section = document.createElement("section");
    section.className = "message";
    const heading = document.createElement("h2");
    heading.textContent = (message.header && message.header.subject) || "(no subject)";
    const byline = document.createElement("p");
    byline.className = "byline";
    byline.textContent = [message.header && message.header.author, message.header && formatWhen(message.header.date)]
      .filter(Boolean)
      .join(" · ");
    section.append(heading, byline);
    if (message.errorText) {
      const error = document.createElement("p");
      error.className = "message-error";
      error.textContent = message.errorText;
      section.append(error);
    }
    if (attachments.length) {
      const grid = document.createElement("div");
      grid.className = "grid";
      for (const record of attachments) {
        grid.append(renderCard(record, generation));
      }
      section.append(grid);
    }
    list.append(section);
  }

  const failed = state.messages.some(message => message.errorText);
  if (!state.messages.length) {
    setStatus("This message could not be read.");
  } else if (!items.length && !failed) {
    setStatus("This conversation has no attachments.");
  } else if (items.length && !shown) {
    setStatus("Nothing in this group.");
  } else {
    setStatus("");
  }
}

function renderCard(record, generation) {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "card";
  const thumb = document.createElement("span");
  thumb.className = "thumb";
  thumb.setAttribute("aria-hidden", "true");
  thumb.textContent = KIND_LABEL[record.kind] || "File";
  const name = document.createElement("span");
  name.className = "name";
  name.textContent = record.name;
  const meta = document.createElement("span");
  meta.className = "meta";
  meta.textContent = `${KIND_LABEL[record.kind] || "File"} · ${formatBytes(record.size)}`;
  const blurb = document.createElement("span");
  blurb.className = "excerpt";
  blurb.hidden = true;
  button.append(thumb, name, meta, blurb);
  files.set(button, record);
  button.addEventListener("click", () => {
    openPreview(record).catch(error => {
      console.error("Attachment gallery could not open a preview.", error);
    });
  });
  if (needsCardFetch(record)) {
    observer.observe(button);
    button.dataset.generation = String(generation);
  }
  return button;
}

const observer = new IntersectionObserver(
  entries => {
    for (const entry of entries) {
      if (!entry.isIntersecting) {
        continue;
      }
      observer.unobserve(entry.target);
      const generation = Number(entry.target.dataset.generation);
      fillCard(entry.target, generation).catch(error => {
        console.error("Attachment gallery could not read a file.", error);
      });
    }
  },
  { rootMargin: "240px" }
);

async function fillCard(button, generation) {
  const record = files.get(button);
  if (!record || button.dataset.filled === "1") {
    return;
  }
  button.dataset.filled = "1";
  const file = await loadFile(record);
  if (generation !== renderGeneration || !button.isConnected) {
    return;
  }
  if (record.kind === "image") {
    const url = trackUrl(URL.createObjectURL(file));
    if (generation !== renderGeneration || !button.isConnected) {
      URL.revokeObjectURL(url);
      liveUrls.delete(url);
      return;
    }
    const image = document.createElement("img");
    image.alt = "";
    image.src = url;
    button.querySelector(".thumb").replaceChildren(image);
    return;
  }

  let text = "";
  if (record.kind === "document") {
    text = await officeText(new Uint8Array(await file.arrayBuffer()), record.name, record.contentType);
  } else {
    const slice = await file.slice(0, TEXT_READ).text();
    text = isHtml(record) ? htmlToText(slice) : slice;
  }
  if (generation !== renderGeneration || !button.isConnected || looksBinary(text)) {
    return;
  }
  const blurb = excerpt(text);
  if (!blurb) {
    return;
  }
  const node = button.querySelector(".excerpt");
  node.hidden = false;
  node.textContent = blurb;
}

function previewNote(text) {
  previewBody.replaceChildren();
  const note = document.createElement("p");
  note.className = "preview-note";
  note.textContent = text;
  previewBody.append(note);
}

async function openPreview(record) {
  state.preview = record;
  clearPreviewUrl();
  setPreviewError("");
  previewTitle.textContent = record.name;
  previewNote("Loading preview…");
  if (!dialog.open) {
    dialog.showModal();
  }
  try {
    await showPreview(record);
  } catch (error) {
    console.error("Attachment gallery could not preview a file.", error);
    previewNote("This file could not be read.");
  }
}

async function showPreview(record) {
  if (state.preview !== record) {
    return;
  }
  if (["archive", "other"].includes(record.kind)) {
    previewNote("No preview for this file type. Save it or open it in another application.");
    return;
  }
  if (record.size > PREVIEW_LIMIT) {
    previewNote("This file is too large to preview here. Save it or open it in another application.");
    return;
  }
  if (record.kind === "document" && record.size > OFFICE_LIMIT) {
    previewNote("This document is too large to preview here. Save it or open it in another application.");
    return;
  }

  const file = await loadFile(record);
  if (state.preview !== record) {
    return;
  }

  if (record.kind === "image" || record.kind === "pdf" || record.kind === "audio" || record.kind === "video") {
    const url = URL.createObjectURL(file);
    state.previewUrl = url;
    previewBody.replaceChildren(mediaElement(record, url));
    return;
  }

  if (record.kind === "document") {
    const text = await officeText(new Uint8Array(await file.arrayBuffer()), record.name, record.contentType);
    if (state.preview !== record) {
      return;
    }
    if (!text) {
      previewNote("No text could be read from this document.");
      return;
    }
    const pre = document.createElement("pre");
    pre.textContent = text;
    previewBody.replaceChildren(pre);
    return;
  }

  const slice = await file.slice(0, isHtml(record) ? HTML_READ : TEXT_READ).text();
  if (state.preview !== record) {
    return;
  }
  if (!isHtml(record) && looksBinary(slice)) {
    previewNote("No preview for this file.");
    return;
  }
  if (isHtml(record)) {
    const frame = document.createElement("iframe");
    frame.sandbox = "";
    frame.referrerPolicy = "no-referrer";
    frame.srcdoc = hardenHtml(slice);
    previewBody.replaceChildren(frame);
    return;
  }
  const pre = document.createElement("pre");
  pre.textContent = slice;
  previewBody.replaceChildren(pre);
}

function mediaElement(record, url) {
  if (record.kind === "image") {
    const image = document.createElement("img");
    image.alt = record.name;
    image.src = url;
    return image;
  }
  if (record.kind === "pdf") {
    const embed = document.createElement("embed");
    embed.type = "application/pdf";
    embed.src = url;
    return embed;
  }
  const media = document.createElement(record.kind === "audio" ? "audio" : "video");
  media.controls = true;
  media.src = url;
  return media;
}

async function saveRecord(record) {
  const file = await loadFile(record);
  const url = URL.createObjectURL(file);
  const link = document.createElement("a");
  link.href = url;
  link.download = safeFileName(file.name || record.name);
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function normalize(messageId, item) {
  if (!item || !item.partName) {
    return null;
  }
  const name = item.name || `part ${item.partName}`;
  const contentType = item.contentType || "application/octet-stream";
  return {
    messageId,
    partName: item.partName,
    name,
    contentType,
    size: Number(item.size) > 0 ? Number(item.size) : 0,
    contentId: item.contentId || "",
    kind: classify(contentType, name),
  };
}

async function threadIds(messageId) {
  try {
    if (!browser.galleryThread) {
      return { ids: [messageId], truncated: false };
    }
    const result = await browser.galleryThread.list(messageId);
    const ids = result && Array.isArray(result.ids) && result.ids.length ? result.ids : [messageId];
    return { ids, truncated: Boolean(result && result.truncated) };
  } catch (error) {
    console.error("Attachment gallery could not list the conversation.", error);
    return { ids: [messageId], truncated: false };
  }
}

async function loadConversation(ids) {
  return mapPool(ids, 4, async id => {
    let header = null;
    try {
      header = await browser.messages.get(id);
    } catch (error) {
      return null;
    }
    let attachments = [];
    let errorText = "";
    try {
      const listed = await browser.messages.listAttachments(id);
      attachments = (listed || []).map(item => normalize(id, item)).filter(Boolean);
    } catch (error) {
      errorText = "Attachments for this message are not available.";
    }
    return { header, attachments, errorText };
  });
}

document.querySelector("#show-inline").addEventListener("change", event => {
  state.showInline = event.target.checked;
  render();
});

for (const button of filters.querySelectorAll("[data-filter]")) {
  button.addEventListener("click", () => {
    state.filter = button.dataset.filter;
    for (const item of filters.querySelectorAll("[data-filter]")) {
      item.setAttribute("aria-pressed", String(item === button));
    }
    render();
  });
}

document.querySelector("#preview-close").addEventListener("click", () => {
  dialog.close();
});

dialog.addEventListener("close", () => {
  state.preview = null;
  clearPreviewUrl();
  previewBody.replaceChildren();
  setPreviewError("");
});

document.querySelector("#preview-save").addEventListener("click", () => {
  if (!state.preview) {
    return;
  }
  setPreviewError("");
  saveRecord(state.preview).catch(error => {
    console.error("Attachment gallery could not save a file.", error);
    setPreviewError("This file could not be saved.");
  });
});

document.querySelector("#preview-open").addEventListener("click", () => {
  if (!state.preview) {
    return;
  }
  setPreviewError("");
  browser.messages.openAttachment(state.preview.messageId, state.preview.partName).catch(error => {
    console.error("Attachment gallery could not open a file externally.", error);
    setPreviewError("Thunderbird could not open this file in another application.");
  });
});

async function start() {
  const messageId = messageIdFromUrl();
  if (messageId == null) {
    setStatus("Open a message, then use Attachment gallery.");
    return;
  }
  const thread = await threadIds(messageId);
  const loaded = await loadConversation(thread.ids);
  state.messages = loaded.filter(Boolean);
  state.truncated = thread.truncated;
  render();
}

start().catch(error => {
  console.error("Attachment gallery could not load.", error);
  setStatus("Attachments could not be loaded.");
});
