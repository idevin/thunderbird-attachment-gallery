import assert from "node:assert/strict";
import { deflateRawSync } from "node:zlib";
import test from "node:test";

import {
  classify,
  countLabel,
  excerpt,
  filterGroup,
  formatBytes,
  groupCounts,
  isInline,
  matchesFilter,
  safeFileName,
  summarize,
} from "../lib/summary.js";
import { officeText } from "../lib/ziptext.js";

function u16(value) {
  return Buffer.from([value & 255, (value >> 8) & 255]);
}

function u32(value) {
  return Buffer.from([value & 255, (value >> 8) & 255, (value >> 16) & 255, (value >> 24) & 255]);
}

function zip(files) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  for (const file of files) {
    const name = Buffer.from(file.name);
    const data = Buffer.from(file.data);
    const method = file.method || 0;
    const uncompressed = file.uncompressed == null ? data.length : file.uncompressed;
    const local = Buffer.concat([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(0),
      u32(data.length),
      u32(uncompressed),
      u16(name.length),
      u16(0),
      name,
      data,
    ]);
    const central = Buffer.concat([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(method),
      u16(0),
      u16(0),
      u32(0),
      u32(data.length),
      u32(uncompressed),
      u16(name.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      name,
    ]);
    locals.push(local);
    centrals.push(central);
    offset += local.length;
  }
  const centralDir = Buffer.concat(centrals);
  const eocd = Buffer.concat([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(files.length),
    u16(files.length),
    u32(centralDir.length),
    u32(offset),
    u16(0),
  ]);
  return new Uint8Array(Buffer.concat([...locals, centralDir, eocd]));
}

test("classifies files by type and extension", () => {
  assert.equal(classify("image/jpeg", "photo.jpg"), "image");
  assert.equal(classify("application/octet-stream", "photo.JPG"), "image");
  assert.equal(classify("application/pdf", "notes.bin"), "pdf");
  assert.equal(classify("text/plain", "notes.txt"), "text");
  assert.equal(classify("application/json", "data.json"), "text");
  assert.equal(classify("application/zip", "report.docx"), "document");
  assert.equal(classify("application/zip", "bundle.zip"), "archive");
  assert.equal(classify("audio/mpeg", "song.mp3"), "audio");
  assert.equal(classify("video/mp4", "clip.mp4"), "video");
  assert.equal(classify("application/octet-stream", "file.bin"), "other");
});

test("groups kinds the way the gallery filters them", () => {
  assert.equal(filterGroup("pdf"), "document");
  assert.equal(filterGroup("text"), "document");
  assert.equal(filterGroup("audio"), "media");
  assert.equal(matchesFilter("image", "all"), true);
  assert.equal(matchesFilter("archive", "other"), true);
  assert.equal(matchesFilter("pdf", "image"), false);
});

test("treats only a real content id as inline", () => {
  assert.equal(isInline({ contentId: "img001" }), true);
  assert.equal(isInline({ contentId: "  " }), false);
  assert.equal(isInline({}), false);
  assert.equal(isInline(null), false);
});

test("summarizes counts, size, and forwarded duplicates", () => {
  const summary = summarize([
    { name: "photo.jpg", size: 1000, contentType: "image/jpeg", kind: "image" },
    { name: "photo.jpg", size: 1000, contentType: "image/jpeg", kind: "image" },
    { name: "notes.pdf", size: 500, contentType: "application/pdf", kind: "pdf" },
  ]);
  assert.equal(summary.count, 3);
  assert.equal(summary.uniqueCount, 2);
  assert.equal(summary.totalBytes, 2500);
  assert.deepEqual(groupCounts(summary.byKind), {
    image: 2,
    document: 1,
    media: 0,
    other: 0,
  });
});

test("formats sizes and file names", () => {
  assert.equal(formatBytes(0), "0 B");
  assert.equal(formatBytes(1024), "1.0 KB");
  assert.equal(formatBytes(1536), "1.5 KB");
  assert.equal(formatBytes(10 * 1024), "10 KB");
  assert.equal(formatBytes(1024 * 1024), "1.0 MB");
  assert.equal(formatBytes(-5), "0 B");
  assert.equal(safeFileName("../../report.pdf"), "report.pdf");
  assert.equal(safeFileName("..\\secret.txt"), "secret.txt");
  assert.equal(safeFileName("   "), "attachment");
  assert.equal(countLabel(1, "file", "files"), "1 file");
  assert.equal(countLabel(2, "file", "files"), "2 files");
});

test("excerpt collapses whitespace and marks a cut", () => {
  assert.equal(excerpt("  hello \n world  "), "hello world");
  assert.equal(excerpt("hello world", 5), "hello…");
  assert.equal(excerpt("", 10), "");
});

test("reads stored office xml as text", async () => {
  const docx = zip([
    {
      name: "word/document.xml",
      data: "<w:document><w:p><w:r><w:t>Hello &amp; goodbye</w:t></w:r></w:p></w:document>",
    },
  ]);
  assert.match(await officeText(docx, "note.docx", ""), /Hello & goodbye/);

  const sheet = zip([
    {
      name: "xl/sharedStrings.xml",
      data: "<sst><si><t>Total</t></si><si><t>Tax</t></si></sst>",
    },
  ]);
  const sheetText = await officeText(sheet, "book.xlsx", "");
  assert.match(sheetText, /Total/);
  assert.match(sheetText, /Tax/);

  const odt = zip([
    {
      name: "content.xml",
      data: "<office:document><office:body><text:p>Plain note</text:p></office:body></office:document>",
    },
  ]);
  assert.match(await officeText(odt, "letter.odt", ""), /Plain note/);
});

test("skips a zip entry that declares a huge uncompressed size", async () => {
  const hostile = zip([
    {
      name: "word/document.xml",
      data: "<w:t>hidden</w:t>",
      uncompressed: 3 * 1024 * 1024,
    },
  ]);
  assert.equal(await officeText(hostile, "bad.docx", ""), "");
  assert.equal(await officeText(new Uint8Array([1, 2, 3, 4]), "note.docx", ""), "");
});

test("reads a deflated docx entry", async () => {
  if (typeof DecompressionStream !== "function") {
    return;
  }
  const xml = "<w:p><w:t>Deflated hello</w:t></w:p>";
  const packed = zip([
    {
      name: "word/document.xml",
      method: 8,
      data: deflateRawSync(Buffer.from(xml)),
      uncompressed: Buffer.byteLength(xml),
    },
  ]);
  assert.equal(await officeText(packed, "note.docx", ""), "Deflated hello");
});
