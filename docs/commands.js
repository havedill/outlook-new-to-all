/* global Office */

(function () {
  "use strict";

  Office.onReady(function () {
    Office.actions.associate("newToAll", newToAll);
  });

  function normalizeEmail(email) {
    return (email || "").trim().toLowerCase();
  }

  function addUniqueRecipient(list, seen, email, selfLower) {
    const trimmed = (email || "").trim();
    if (!trimmed) {
      return;
    }
    const lower = normalizeEmail(trimmed);
    if (!lower || lower === selfLower || seen.has(lower)) {
      return;
    }
    seen.add(lower);
    list.push(trimmed);
  }

  function newToAll(event) {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;
    const selfLower = normalizeEmail(mailbox.userProfile.emailAddress);

    if (!item || !item.from || !item.from.emailAddress) {
      event.completed();
      return;
    }

    item.to.getAsync(function (toResult) {
      if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed();
        return;
      }

      item.cc.getAsync(function (ccResult) {
        if (ccResult.status !== Office.AsyncResultStatus.Succeeded) {
          event.completed();
          return;
        }

        const seen = new Set();
        const toRecipients = [];

        addUniqueRecipient(toRecipients, seen, item.from.emailAddress, selfLower);

        (toResult.value || []).forEach(function (recipient) {
          addUniqueRecipient(toRecipients, seen, recipient.emailAddress, selfLower);
        });

        const ccRecipients = [];
        (ccResult.value || []).forEach(function (recipient) {
          addUniqueRecipient(ccRecipients, seen, recipient.emailAddress, selfLower);
        });

        mailbox.displayNewMessageForm({
          toRecipients: toRecipients,
          ccRecipients: ccRecipients,
        });

        event.completed();
      });
    });
  }
})();
