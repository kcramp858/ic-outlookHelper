```javascript
Office.initialize = function (reason) {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemChanged, loadNewEmail);
};

function loadNewEmail(eventArgs) {
    var emailMetadata = document.getElementById('emailMetadata');
    var item = Office.context.mailbox.item;
    emailMetadata.innerHTML = `From: ${item.from.displayName} <br> To: ${item.to.displayName} <br> Date: ${item.dateTimeCreated} <br> Subject: ${item.subject} <br> Body: ${item.body.getAsync('text')}`;
}

document.getElementById('ribbonButton').addEventListener('click', function () {
    Office.context.ui.displayDialogAsync('https://localhost:3000/Popup.html', { height: 30, width: 20, displayInIframe: true }, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error(asyncResult.error.message);
        }
    });
});

document.getElementById('cancelButton').addEventListener('click', cancelPopup);
document.getElementById('submitButton').addEventListener('click', submitPopup);

function cancelPopup() {
    Office.context.ui.messageParent('cancel');
}

function submitPopup() {
    var explanation = document.getElementById('explanationBox').value;
    var tasks = document.getElementById('tasksBox').value;
    var emailData = {
        metadata: document.getElementById('emailMetadata').innerHTML,
        explanation: explanation,
        tasks: tasks
    };
    Office.context.ui.messageParent(JSON.stringify(emailData));
}
```