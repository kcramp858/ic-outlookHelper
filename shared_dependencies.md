1. Exported Variables:
   - `emailData`: This variable will hold the email metadata and user inputs from the popup. It will be used across the JavaScript files and saved in the `emailData.json` file.

2. Data Schemas:
   - `EmailSchema`: This schema will define the structure of the email metadata (from, to, date, subject, body) and will be used in `OfficeScripts.js` and `emailData.json`.
   - `UserInputSchema`: This schema will define the structure of the user inputs (explanation, tasks) and will be used in `PopupScripts.js` and `emailData.json`.

3. ID Names of DOM Elements:
   - `ribbonButton`: The ID of the button in the ribbon that triggers the popup. Used in `RibbonButton.html` and `OfficeScripts.js`.
   - `emailMetadata`: The ID of the left column in the popup that displays the email metadata. Used in `Popup.html` and `PopupScripts.js`.
   - `explanationBox`: The ID of the top text box in the popup. Used in `Popup.html` and `PopupScripts.js`.
   - `tasksBox`: The ID of the bottom text box in the popup. Used in `Popup.html` and `PopupScripts.js`.
   - `submitButton`: The ID of the submit button in the popup. Used in `Popup.html` and `PopupScripts.js`.
   - `cancelButton`: The ID of the cancel button in the popup. Used in `Popup.html` and `PopupScripts.js`.

4. Message Names:
   - `showPopup`: This message will be sent from `OfficeScripts.js` to `PopupScripts.js` to display the popup.
   - `saveData`: This message will be sent from `PopupScripts.js` to `OfficeScripts.js` to save the user inputs and email metadata in `emailData.json`.

5. Function Names:
   - `showPopup()`: This function in `OfficeScripts.js` will display the popup when the ribbon button is clicked.
   - `saveData()`: This function in `OfficeScripts.js` will save the user inputs and email metadata in `emailData.json`.
   - `cancelPopup()`: This function in `PopupScripts.js` will close the popup without saving when the cancel button is clicked.
   - `submitPopup()`: This function in `PopupScripts.js` will send the `saveData` message when the submit button is clicked.