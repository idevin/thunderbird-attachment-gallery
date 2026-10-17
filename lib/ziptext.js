/**
 * Read a short plain-text preview from Office Open XML and ODT packages.
 * Inflated data is capped so a hostile zip cannot expand without limit.
 */

import { extensionOf, mimeType } from "./summary.js";

const MAX_UNCOMPRESSED = 2 * 1024 * 1024;
const TEXT_CAP = 8000;

function wantedPaths(name, contentType) {
  const ext = extensionOf(name);
  const type = mimeType(contentType);
  if (ext === "docx" || type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
    return ["word/document.xml"];
  }
  if (ext === "odt" || type === "application/vnd.oasis.opendocument.text") {
    return ["content.xml"];
  }
  if (ext === "xlsx" || type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") {
    return ["xl/sharedstrings.xml"];
  }
  if (ext === "pptx" || type === "application/vnd.openxmlformats-officedocument.presentationml.presentation") {
    return Array.from({ length: 8 }, (_, index) => `ppt/slides/slide${index + 1}.xml`);
  }
  return [];
}

function readU16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes, offset) {
  return (
    (bytes[offset] |
      (bytes[offset + 1] << 8) |
      (bytes[offset + 2] << 16) |
      (bytes[offset + 3] << 24)) >>>
    0
  );
}

function findEocd(bytes) {
  const min = Math.max(0, bytes.length - 22 - 65535);
  for (let offset = bytes.length - 22; offset >= min; offset--) {
    if (
      bytes[offset] === 0x50 &&
      bytes[offset + 1] === 0x4b &&
      bytes[offset + 2] === 0x05 &&
      bytes[offset + 3] === 0x06
    ) {
      return offset;
    }
  }
  return -1;
}

function listEntries(bytes) {
  const eocd = findEocd(bytes);
  if (eocd < 0 || eocd + 22 > bytes.length) {
    return [];
  }
  const count = readU16(bytes, eocd + 10);
  let offset = readU32(bytes, eocd + 16);
  const entries = [];
  for (let index = 0; index < count; index++) {
    if (offset + 46 > bytes.length || readU32(bytes, offset) !== 0x02014b50) {
      break;
    }
    const method = readU16(bytes, offset + 10);
    const compSize = readU32(bytes, offset + 20);
    const uncompSize = readU32(bytes, offset + 24);
    const nameLen = readU16(bytes, offset + 28);
    const extraLen = readU16(bytes, offset + 30);
    const commentLen = readU16(bytes, offset + 32);
    const localOffset = readU32(bytes, offset + 42);
    const nameStart = offset + 46;
    const nameEnd = nameStart + nameLen;
    if (nameEnd > bytes.length) {
      break;
    }
    const name = new TextDecoder("utf-8", { fatal: false }).decode(bytes.subarray(nameStart, nameEnd));
    entries.push({ name, method, compSize, uncompSize, localOffset });
    offset = nameEnd + extraLen + commentLen;
  }
  return entries;
}

async function inflate(bytes) {
  if (typeof DecompressionStream !== "function") {
    throw new Error("deflate is not available");
  }
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  const buffer = await new Response(stream).arrayBuffer();
  return new Uint8Array(buffer);
}

async function readEntry(bytes, entry) {
  if (entry.uncompSize > MAX_UNCOMPRESSED || entry.compSize > MAX_UNCOMPRESSED) {
    throw new Error("entry too large");
  }
  if (entry.localOffset + 30 > bytes.length || readU32(bytes, entry.localOffset) !== 0x04034b50) {
    throw new Error("bad local header");
  }
  const nameLen = readU16(bytes, entry.localOffset + 26);
  const extraLen = readU16(bytes, entry.localOffset + 28);
  const dataStart = entry.localOffset + 30 + nameLen + extraLen;
  const dataEnd = dataStart + entry.compSize;
  if (dataEnd > bytes.length) {
    throw new Error("truncated entry");
  }
  const compressed = bytes.subarray(dataStart, dataEnd);
  if (entry.method === 0) {
    return compressed;
  }
  if (entry.method === 8) {
    const inflated = await inflate(compressed);
    if (inflated.length > MAX_UNCOMPRESSED) {
      throw new Error("entry too large");
    }
    return inflated;
  }
  throw new Error("unsupported compression");
}

function xmlToText(xml) {
  return String(xml)
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<w:br\s*\/>/g, "\n")
    .replace(/<\/w:p>/g, "\n")
    .replace(/<\/text:p>/g, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, value) => {
      const code = Number(value);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
    })
    .replace(/&#x([0-9a-f]+);/gi, (_, value) => {
      const code = parseInt(value, 16);
      return code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "";
    })
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();
}

export async function officeText(bytes, name, contentType) {
  const paths = wantedPaths(name, contentType);
  if (!paths.length || !(bytes instanceof Uint8Array)) {
    return "";
  }
  const wanted = new Set(paths);
  const chunks = [];
  let total = 0;
  for (const entry of listEntries(bytes)) {
    const entryName = entry.name.replace(/\\/g, "/").replace(/^\/+/, "").toLowerCase();
    if (!wanted.has(entryName) || entry.uncompSize > MAX_UNCOMPRESSED) {
      continue;
    }
    let raw;
    try {
      raw = await readEntry(bytes, entry);
    } catch (error) {
      continue;
    }
    if (total >= TEXT_CAP) {
      break;
    }
    total += raw.length;
    const text = xmlToText(new TextDecoder("utf-8", { fatal: false }).decode(raw));
    if (text) {
      chunks.push(text);
    }
  }
  return chunks.join("\n\n").slice(0, TEXT_CAP).trim();
}
