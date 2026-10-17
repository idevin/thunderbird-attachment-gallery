"use strict";

const { ExtensionCommon } = ChromeUtils.importESModule("resource://gre/modules/ExtensionCommon.sys.mjs");
const { Services } = ChromeUtils.importESModule("resource://gre/modules/Services.sys.mjs");

const MAX_MESSAGES = 200;
const GLODA_TIMEOUT_MS = 4000;
const DELETED = Components.interfaces.nsMsgMessageFlags.IMAPDeleted | Components.interfaces.nsMsgMessageFlags.Expunged;

let closed = false;

function identity(hdr) {
  let folder = "";
  try {
    folder = hdr.folder.URI;
  } catch (error) {
    folder = "";
  }
  return `${folder}|${hdr.messageKey}`;
}

function isAlive(hdr) {
  try {
    return Boolean(hdr && hdr.folder && (hdr.flags & DELETED) === 0);
  } catch (error) {
    return false;
  }
}

function compareDate(left, right) {
  const delta = (Number(left.date) || 0) - (Number(right.date) || 0);
  if (delta) {
    return delta < 0 ? -1 : 1;
  }
  return (Number(left.messageKey) || 0) - (Number(right.messageKey) || 0);
}

function folderThread(hdr) {
  const found = [];
  try {
    const thread = hdr.folder.msgDatabase.getThreadContainingMsgHdr(hdr);
    if (!thread) {
      return found;
    }
    try {
      const root = thread.getRootHdr();
      if (root) {
        found.push(root);
      }
    } catch (error) {
      // The root header can be missing from a damaged index.
    }
    const count = thread.numChildren || 0;
    for (let index = 0; index < count; index++) {
      try {
        const child = thread.getChildHdrAt(index);
        if (child) {
          found.push(child);
        }
      } catch (error) {
        // Skip a child the database cannot materialize.
      }
    }
  } catch (error) {
    // Fall back to the message that was already open.
  }
  return found;
}

function queryGloda(run) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (error, collection) => {
      if (settled) {
        return;
      }
      settled = true;
      clearTimeout(timer);
      if (error) {
        reject(error);
      } else {
        resolve(collection);
      }
    };
    const timer = setTimeout(() => finish(new Error("Gloda query timed out")), GLODA_TIMEOUT_MS);
    try {
      run({
        onItemsAdded() {},
        onItemsModified() {},
        onItemsRemoved() {},
        onQueryCompleted(collection) {
          finish(null, collection);
        },
      });
    } catch (error) {
      finish(error);
    }
  });
}

async function glodaThread(hdr) {
  if (closed) {
    return [];
  }
  try {
    if (!Services.prefs.getBoolPref("mailnews.database.global.indexer.enabled", false)) {
      return [];
    }
  } catch (error) {
    return [];
  }

  let Gloda;
  try {
    ({ Gloda } = ChromeUtils.importESModule("resource:///modules/gloda/Gloda.sys.mjs"));
  } catch (error) {
    return [];
  }

  let indexed;
  try {
    indexed = await queryGloda(listener => Gloda.getMessageCollectionForHeader(hdr, listener));
  } catch (error) {
    return [];
  }
  if (closed || !indexed || !indexed.items || !indexed.items.length) {
    return [];
  }

  const glodaMessage = indexed.items[0];
  if (!glodaMessage.conversation || typeof glodaMessage.conversation.getMessagesCollection !== "function") {
    return [];
  }

  let conversation;
  try {
    conversation = await queryGloda(listener => glodaMessage.conversation.getMessagesCollection(listener));
  } catch (error) {
    return [];
  }
  if (closed || !conversation || !conversation.items) {
    return [];
  }

  const headers = [];
  for (const item of conversation.items) {
    try {
      if (item && item.folderMessage) {
        headers.push(item.folderMessage);
      }
    } catch (error) {
      // The indexed row no longer points at a message.
    }
  }
  return headers;
}

function limitHeaders(headers, anchor) {
  const sorted = headers.slice().sort(compareDate);
  if (sorted.length <= MAX_MESSAGES) {
    return { headers: sorted, truncated: false };
  }
  const newest = sorted.slice(sorted.length - MAX_MESSAGES);
  const anchorId = identity(anchor);
  if (!newest.some(item => identity(item) === anchorId)) {
    newest[0] = anchor;
  }
  newest.sort(compareDate);
  return { headers: newest, truncated: true };
}

async function listThread(context, messageId) {
  const anchor = context.extension.messageManager.get(messageId);
  const byId = new Map();
  const add = hdr => {
    if (!isAlive(hdr)) {
      return;
    }
    const key = identity(hdr);
    if (!byId.has(key)) {
      byId.set(key, hdr);
    }
  };

  add(anchor);
  for (const hdr of folderThread(anchor)) {
    add(hdr);
  }
  for (const hdr of await glodaThread(anchor)) {
    if (closed) {
      break;
    }
    add(hdr);
  }

  const limited = limitHeaders([...byId.values()], anchor);
  const ids = [];
  const seen = new Set();
  for (const hdr of limited.headers) {
    if (closed) {
      break;
    }
    try {
      const converted = context.extension.messageManager.convert(hdr);
      if (converted && !seen.has(converted.id)) {
        seen.add(converted.id);
        ids.push(converted.id);
      }
    } catch (error) {
      // Skip a header the WebExtension message manager cannot expose.
    }
  }
  if (!ids.length) {
    ids.push(messageId);
  }
  return { ids, truncated: limited.truncated };
}

this.galleryThread = class extends ExtensionCommon.ExtensionAPI {
  onShutdown() {
    closed = true;
  }

  getAPI(context) {
    return {
      galleryThread: {
        async list(messageId) {
          if (closed) {
            return { ids: [messageId], truncated: false };
          }
          try {
            return await listThread(context, messageId);
          } catch (error) {
            throw new ExtensionCommon.ExtensionError("Could not list this conversation.");
          }
        },
      },
    };
  }
};
