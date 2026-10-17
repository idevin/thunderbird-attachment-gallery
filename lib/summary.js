/**
 * Classify attachments and summarize a conversation without reading file bytes.
 */

const OFFICE_EXT = new Set(["doc", "docx", "odt", "xls", "xlsx", "ppt", "pptx", "ods", "odp"]);
const TEXT_EXT = new Set([
  "txt", "csv", "json", "md", "xml", "log", "rtf", "ics", "vcf", "html", "htm", "css", "js", "tsv",
]);
const ARCHIVE_EXT = new Set(["zip", "gz", "tgz", "bz2", "7z", "rar", "tar", "xz"]);
const IMAGE_EXT = new Set(["png", "jpg", "jpeg", "gif", "webp", "bmp", "avif", "svg", "ico"]);
const AUDIO_EXT = new Set(["mp3", "ogg", "wav", "m4a", "flac", "aac", "opus"]);
const VIDEO_EXT = new Set(["mp4", "webm", "ogv", "mov", "mkv"]);

const OFFICE_TYPES = new Set([
  "application/msword",
  "application/vnd.ms-excel",
  "application/vnd.ms-powerpoint",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  "application/vnd.oasis.opendocument.text",
  "application/vnd.oasis.opendocument.spreadsheet",
  "application/vnd.oasis.opendocument.presentation",
]);

export function extensionOf(name) {
  const base = String(name || "").split(/[/\\]/).pop();
  const dot = base.lastIndexOf(".");
  if (dot <= 0) {
    return "";
  }
  return base.slice(dot + 1).toLowerCase();
}

export function mimeType(contentType) {
  return String(contentType || "").split(";")[0].trim().toLowerCase();
}

export function classify(contentType, name) {
  const type = mimeType(contentType);
  const ext = extensionOf(name);
  if (OFFICE_EXT.has(ext) || OFFICE_TYPES.has(type)) {
    return "document";
  }
  if (type === "application/pdf" || ext === "pdf") {
    return "pdf";
  }
  if (type.startsWith("image/") || IMAGE_EXT.has(ext)) {
    return "image";
  }
  if (type.startsWith("audio/") || AUDIO_EXT.has(ext)) {
    return "audio";
  }
  if (type.startsWith("video/") || VIDEO_EXT.has(ext)) {
    return "video";
  }
  if (
    type.startsWith("text/") ||
    type === "application/json" ||
    type === "application/xml" ||
    type.endsWith("+json") ||
    type.endsWith("+xml") ||
    TEXT_EXT.has(ext)
  ) {
    return "text";
  }
  if (
    type === "application/zip" ||
    type === "application/gzip" ||
    type === "application/x-tar" ||
    type === "application/x-7z-compressed" ||
    type === "application/vnd.rar" ||
    ARCHIVE_EXT.has(ext)
  ) {
    return "archive";
  }
  return "other";
}

export function filterGroup(kind) {
  if (kind === "image") {
    return "image";
  }
  if (kind === "pdf" || kind === "text" || kind === "document") {
    return "document";
  }
  if (kind === "audio" || kind === "video") {
    return "media";
  }
  return "other";
}

export function matchesFilter(kind, filter) {
  if (!filter || filter === "all") {
    return true;
  }
  return filterGroup(kind) === filter;
}

export function isInline(attachment) {
  return Boolean(attachment && String(attachment.contentId || "").trim());
}

export function summarize(items) {
  const byKind = {
    image: 0,
    pdf: 0,
    text: 0,
    document: 0,
    audio: 0,
    video: 0,
    archive: 0,
    other: 0,
  };
  let totalBytes = 0;
  const unique = new Set();
  for (const item of items) {
    const kind = item.kind || classify(item.contentType, item.name);
    byKind[kind] = (byKind[kind] || 0) + 1;
    const size = Number(item.size);
    totalBytes += Number.isFinite(size) && size > 0 ? size : 0;
    unique.add(`${item.name}\0${item.size}\0${item.contentType}`);
  }
  return {
    count: items.length,
    uniqueCount: unique.size,
    totalBytes,
    byKind,
  };
}

export function groupCounts(byKind) {
  const kinds = byKind || {};
  return {
    image: kinds.image || 0,
    document: (kinds.pdf || 0) + (kinds.text || 0) + (kinds.document || 0),
    media: (kinds.audio || 0) + (kinds.video || 0),
    other: (kinds.archive || 0) + (kinds.other || 0),
  };
}

export function excerpt(text, limit = 280) {
  const flat = String(text || "").replace(/\s+/g, " ").trim();
  const max = Math.max(1, Number(limit) || 280);
  if (flat.length <= max) {
    return flat;
  }
  return `${flat.slice(0, max).trimEnd()}…`;
}

export function formatBytes(size) {
  let value = Number(size);
  if (!Number.isFinite(value) || value < 0) {
    value = 0;
  }
  const units = ["B", "KB", "MB", "GB"];
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit += 1;
  }
  const digits = unit === 0 || value >= 10 ? 0 : 1;
  return `${value.toFixed(digits)} ${units[unit]}`;
}

export function safeFileName(name) {
  const base = String(name || "attachment").split(/[/\\]/).pop();
  const cleaned = base.replace(/[\u0000-\u001f]/g, "").trim();
  return cleaned || "attachment";
}

export function countLabel(count, singular, plural) {
  return `${count} ${count === 1 ? singular : plural}`;
}
