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

  function completeEvent(event) {
    try {
      event.completed();
    } catch (e) {
      // Office may throw if completed was already called.
    }
  }

  /**
   * MessageRead: item.to / item.cc are EmailAddressDetails[] (see MessageRead docs).
   * MessageCompose: they are Recipients objects with getAsync.
   */
  function getToAndCcLists(item, callback) {
    if (Array.isArray(item.to)) {
      callback(item.to, Array.isArray(item.cc) ? item.cc : []);
      return;
    }

    if (item.to && typeof item.to.getAsync === "function") {
      item.to.getAsync(function (toResult) {
        if (toResult.status !== Office.AsyncResultStatus.Succeeded) {
          callback([], []);
          return;
        }
        const toList = toResult.value || [];
        if (item.cc && typeof item.cc.getAsync === "function") {
          item.cc.getAsync(function (ccResult) {
            const ccList =
              ccResult.status === Office.AsyncResultStatus.Succeeded
                ? ccResult.value || []
                : [];
            callback(toList, ccList);
          });
        } else {
          callback(toList, Array.isArray(item.cc) ? item.cc : []);
        }
      });
      return;
    }

    callback([], Array.isArray(item.cc) ? item.cc : []);
  }

  function openNewToAll(mailbox, item, toList, ccList) {
    const selfLower = normalizeEmail(mailbox.userProfile.emailAddress);

    if (!item.from || !item.from.emailAddress) {
      return;
    }

    const seen = new Set();
    const toRecipients = [];

    addUniqueRecipient(toRecipients, seen, item.from.emailAddress, selfLower);

    (toList || []).forEach(function (recipient) {
      addUniqueRecipient(toRecipients, seen, recipient.emailAddress, selfLower);
    });

    const ccRecipients = [];
    (ccList || []).forEach(function (recipient) {
      addUniqueRecipient(ccRecipients, seen, recipient.emailAddress, selfLower);
    });

    mailbox.displayNewMessageForm({
      toRecipients: toRecipients,
      ccRecipients: ccRecipients,
    });
  }

  function newToAll(event) {
    try {
      const mailbox = Office.context.mailbox;
      const item = mailbox.item;

      if (!item) {
        completeEvent(event);
        return;
      }

      getToAndCcLists(item, function (toList, ccList) {
        try {
          openNewToAll(mailbox, item, toList, ccList);
        } catch (err) {
          // Swallow errors so the ribbon command still completes.
        } finally {
          completeEvent(event);
        }
      });
    } catch (err) {
      completeEvent(event);
    }
  }
})();
