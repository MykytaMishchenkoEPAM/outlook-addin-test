'use strict';

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            const item = Office.context.mailbox.item;
            loadItemProps(item, item.attachments.length > 0);

            for (let i = 0; i < item.attachments.length; i++) {
                if (item.attachments[i].contentType == "application/pdf") {
                    const options = { asyncContext: { name: item.attachments[i].name } };
                    item.getAttachmentContentAsync(item.attachments[i].id, options, handleAttachmentsCallback);
                }
            }

            item.displayReplyForm({});
        });
    });

    function loadItemProps(item, isAttachmentsPresent) {
        // Write message property values to the task pane
        $('#item-id').text(item.itemId);
        $('#item-subject').text(item.subject);
        $('#item-internetMessageId').text(item.internetMessageId);
        $('#item-from').html(item.from.displayName + " &lt;" + item.from.emailAddress + "&gt;");
        $('#item-attachmentsIds').text(isAttachmentsPresent ? item.attachments.map(attachment => attachment.id).join() : "No attachments");
    }

    function handleAttachmentsCallback(result) {
        switch (result.value.format) {
            case Office.MailboxEnums.AttachmentContentFormat.Base64:
                console.log("Base64");
                sendRequestWithFile(result.asyncContext.name, result.value.content)
                break;
            case Office.MailboxEnums.AttachmentContentFormat.Eml:
                console.log("Eml");
                break;
            case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
                console.log("ICalendar");
                break;
            case Office.MailboxEnums.AttachmentContentFormat.Url:
                console.log("Url");
                break;
            default:
                console.error("Attachment isn't supported");
        }
    }

    function sendRequestWithFile(name, content) {
        fetch("https://localhost:7180/file/upload", {
            method: "POST",
            body: JSON.stringify({
                    name: name,
                    base64: content
                }),
            headers: {
                "Content-type": "application/json; charset=UTF-8"
            }
        })
        .then(response => console.log(response));
    }
})();