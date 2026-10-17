/**
 * Open the attachment gallery for the message the user is reading.
 * messages.get registers that message so the experiment can resolve its id.
 */

async function selectedMessage(tab) {
  const tabs = [];
  if (tab && tab.id != null) {
    tabs.push(tab);
  }
  try {
    const mailTabs = await browser.tabs.query({ mailTab: true });
    for (const mailTab of mailTabs) {
      if (!tabs.some(item => item.id === mailTab.id)) {
        tabs.push(mailTab);
      }
    }
  } catch (error) {
    console.error("Attachment gallery could not list mail tabs.", error);
  }

  for (const candidate of tabs) {
    try {
      const displayed = await browser.messageDisplay.getDisplayedMessage(candidate.id);
      if (displayed) {
        return displayed;
      }
    } catch (error) {
      // This tab is not displaying a message.
    }
    try {
      const listed = await browser.mailTabs.getSelectedMessages(candidate.id);
      if (listed && listed.messages && listed.messages.length) {
        return listed.messages[0];
      }
    } catch (error) {
      // This tab is not a mail tab.
    }
  }
  return null;
}

async function openGallery(tab) {
  const message = await selectedMessage(tab);
  const page = new URL(browser.runtime.getURL("gallery/gallery.html"));
  if (message) {
    try {
      await browser.messages.get(message.id);
    } catch (error) {
      console.error("Attachment gallery could not read the selected message.", error);
    }
    page.searchParams.set("messageId", String(message.id));
  }
  await browser.tabs.create({ url: page.href });
}

browser.action.onClicked.addListener(tab => {
  openGallery(tab).catch(error => {
    console.error("Attachment gallery could not open.", error);
  });
});

browser.messageDisplayAction.onClicked.addListener(tab => {
  openGallery(tab).catch(error => {
    console.error("Attachment gallery could not open.", error);
  });
});
