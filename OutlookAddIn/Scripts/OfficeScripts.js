Office.initialize = function (reason) {
    Office.context.mailbox.item.loadCustomPropertiesAsync(customLoadCallback);
};

function customLoadCallback(asyncResult) {
    var customProps = asyncResult.value;
    var emailData = customProps.get("emailData");
    if (emailData == null) {
        emailData = [];
    }
}

Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewEmail);

function loadNewEmail(eventArgs) {
    var item = Office.context.mailbox.item;
    var emailSchema = {
        from: item.from,
        to: item.to,
        date: item.dateTimeCreated,
        subject: item.subject,
        body: item.body
    };
    Office.context.ui.displayDialogAsync('HTML/Popup.html', { height: 30, width: 20, displayInIframe: true }, function (asyncResult) {
        var dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
            if (arg.message === 'saveData') {
                emailData.push(emailSchema);
                customProps.set("emailData", emailData);
                customProps.saveAsync(function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        console.log("Failed to save data. Error: " + asyncResult.error.message);
                    }
                });
            }
            dialog.close();
        });
    });
}

document.getElementById('ribbonButton').onclick = function () {
    loadNewEmail();
};