Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}
function onNewMessageComposeHandler(event) {
    setSubject(event);
}
function onNewAppointmentComposeHandler(event) {
    setSubject(event);
}
function setSubject(event) {
    Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
            "asyncContext": event
        },
        function (asyncResult) {
            // Handle success or error.
            if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
                console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
            }

            // Call event.completed() to signal to the Outlook client that the add-in has completed processing the event.
            asyncResult.asyncContext.completed();
        });
}

// IMPORTANT: To ensure your add-in is supported in Outlook, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);