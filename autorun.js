Office.initialize = function (reason) { };

/**
 * Handles the OnNewMessageCompose event.
 */
function onNewMessageComposeHandler() {
    fetch("https://codetwodev.github.io/sample-signatures-addin/signatures/signature.html")
    .then(response => {
      return response.text();
    })
    .then(signature => {
      Office.context.mailbox.item.sessionData.setAsync("signature", signature);
    });

    fetch("https://codetwodev.github.io/sample-signatures-addin/signatures/disclaimer.html")
    .then(response => {
      return response.text();
    })
    .then(disclaimer => {
      Office.context.mailbox.item.sessionData.setAsync("disclaimer", disclaimer);
    });
}

/**
 * Handles the OnMessageRecipientsChanged event.
 */
function onMessageRecipientsChangedHandler() {

  Office.context.mailbox.item.to.getAsync(
    function (asyncResult) {
      var recipients = asyncResult.value;

      var hasExternal = recipients.some(r => r.recipientType !== Office.MailboxEnums.RecipientType.User);
      
      insertSignature(hasExternal);
    }
  );
}

/** Inserts the signature stored in the session data */
function insertSignature(withDisclaimer) {
  Office.context.mailbox.item.sessionData.getAllAsync(
    function (asyncResult) {
      var sessionData = asyncResult.value;
      var signature = sessionData.signature;

      if (withDisclaimer) {
        signature += sessionData.disclaimer;
      }

      Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: Office.CoercionType.Html });
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onMessageRecipientsChangedHandler", onMessageRecipientsChangedHandler);