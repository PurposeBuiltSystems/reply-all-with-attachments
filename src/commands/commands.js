/*
 * Reply All with Attachments — ribbon button command.
 *
 * Outlook intentionally drops the original attachments when you Reply or
 * Reply All (unlike Forward). This command restores that behavior: it reads
 * the attachments on the currently open message, fetches each one's bytes,
 * and opens a Reply All compose form with those attachments re-attached.
 *
 * Flow:
 *   1. Read item.attachments (skip inline images — Outlook keeps those in the
 *      quoted body automatically).
 *   2. For each, getAttachmentContentAsync -> bytes in one of several formats.
 *   3. Convert to ReplyFormAttachment objects.
 *   4. displayReplyAllFormAsync({ attachments }) to open the populated reply.
 *
 * Requirement set: Mailbox 1.9 (1.8 for getAttachmentContentAsync,
 * 1.9 for the async displayReplyAllFormAsync).
 */

/* global Office */

Office.onReady(function () {
  // Register the function so the manifest's <FunctionName> can call it.
  // associate() is the modern, cross-platform registration mechanism.
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("replyAllWithAttachments", replyAllWithAttachments);
  }
});

/**
 * Entry point invoked by the ribbon button.
 * @param {Office.AddinCommands.Event} event
 */
async function replyAllWithAttachments(event) {
  try {
    const item = Office.context.mailbox.item;
    const attachments = (item.attachments || []).filter(function (a) {
      // Keep real file/cloud/item attachments; drop inline images, which the
      // quoted reply body preserves on its own.
      return !a.isInline;
    });

    if (attachments.length === 0) {
      // Nothing to carry over — just open a normal Reply All.
      item.displayReplyAllFormAsync("", function () {
        finish(event);
      });
      return;
    }

    // Fetch every attachment's content in parallel.
    const results = await Promise.all(
      attachments.map(function (a) {
        return getAttachmentContent(item, a.id).then(function (content) {
          return toReplyFormAttachment(a, content);
        });
      })
    );

    const replyAttachments = results.filter(Boolean);
    const skipped = attachments.length - replyAttachments.length;

    const formData = {
      htmlBody: "",
      attachments: replyAttachments,
    };

    item.displayReplyAllFormAsync(formData, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        notify("error", "Couldn't open the reply: " + asyncResult.error.message);
      } else if (skipped > 0) {
        notify(
          "info",
          skipped + " attachment(s) couldn't be carried over and were skipped."
        );
      }
      finish(event);
    });
  } catch (err) {
    notify("error", "Reply All with Attachments failed: " + (err && err.message ? err.message : err));
    finish(event);
  }
}

/**
 * Promisified getAttachmentContentAsync.
 * @returns {Promise<Office.AttachmentContent>}
 */
function getAttachmentContent(item, attachmentId) {
  return new Promise(function (resolve, reject) {
    item.getAttachmentContentAsync(attachmentId, function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

/**
 * Convert an Office attachment + its fetched content into a ReplyFormAttachment.
 * Returns null for content we can't safely re-attach (it gets skipped).
 * @param {Office.AttachmentDetails} attachment
 * @param {Office.AttachmentContent} content
 */
function toReplyFormAttachment(attachment, content) {
  const Formats = Office.MailboxEnums.AttachmentContentFormat;
  const Types = Office.MailboxEnums.AttachmentType;

  switch (content.format) {
    case Formats.Base64:
      // Regular file attachments (PDF, images, docs, …).
      return {
        type: Types.Base64,
        name: attachment.name,
        base64file: content.content,
        inLine: false,
      };

    case Formats.Url:
      // Cloud attachment — re-attach by its URL.
      return {
        type: Types.File,
        name: attachment.name,
        url: content.content,
        inLine: false,
      };

    case Formats.Eml:
      // An attached email message. Wrap its MIME as a base64 .eml file.
      return {
        type: Types.Base64,
        name: ensureExt(attachment.name, ".eml"),
        base64file: utf8ToBase64(content.content),
        inLine: false,
      };

    case Formats.ICalendar:
      // An attached calendar item. Wrap its text as a base64 .ics file.
      return {
        type: Types.Base64,
        name: ensureExt(attachment.name, ".ics"),
        base64file: utf8ToBase64(content.content),
        inLine: false,
      };

    default:
      return null;
  }
}

/** Append an extension if the name doesn't already end with it. */
function ensureExt(name, ext) {
  if (!name) return "attachment" + ext;
  return name.toLowerCase().endsWith(ext) ? name : name + ext;
}

/** UTF-8 safe string -> base64. */
function utf8ToBase64(str) {
  // encodeURIComponent handles multibyte chars; unescape rebuilds the byte
  // string that btoa expects.
  return btoa(unescape(encodeURIComponent(str)));
}

/** Show a transient notification on the item (best effort). */
function notify(kind, message) {
  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.notificationMessages) return;
    item.notificationMessages.replaceAsync("raAttach", {
      type:
        kind === "error"
          ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
          : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message.substring(0, 150),
      icon: "Icon.16",
      persistent: false,
    });
  } catch (e) {
    // Notifications are unavailable in some contexts; ignore.
  }
}

/** Signal Outlook the command is done so the ribbon button is re-enabled. */
function finish(event) {
  if (event && typeof event.completed === "function") {
    event.completed();
  }
}
