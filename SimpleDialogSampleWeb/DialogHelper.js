/*Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
4  See LICENSE in the project root for license information */

var dialog;

function dialogCallback(asyncResult) {
    if (asyncResult.status == "failed") {

        // In addition to general system errors, there are 3 specific errors for 
        // displayDialogAsync that you can handle individually.
        switch (asyncResult.error.code) {
            case 12004:
                showNotification("Domain is not trusted");
                break;
            case 12005:
                showNotification("HTTPS is required");
                break;
            case 12007:
                showNotification("A dialog is already opened.");
                break;
            default:
                showNotification(asyncResult.error.message);
                break;
        }
    }
    else {
        dialog = asyncResult.value;
        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
    }
}


function messageHandler(arg) {
    dialog.close();
//        showNotification(arg.message);

    //    var billValues = arg.message.split('|');
    // insertLISRef(billValues[0], billValues[1], billValues[2]);
    var message = JSON.parse(arg.message);

    switch (message.methodName) {
        case "LISRef":
            showNotification("Inserting LIS Reference");
            insertLISRef(message.LIS.billType, message.LIS.billNum, message.LIS.congressYear);
            break;
        case "CRSRef":
            showNotification("Inserting CRS Prod Reference");
            insertCRSRef(message.CrsProduct, message.displayText);
            break;
        case "InsertAuthors":
            showNotification("Inserting Authors");
            addAuthors();
//            insertCRSRef(message.CrsProduct, message.displayText);
            break;
    }
 
}


function eventHandler(arg) {

    // In addition to general system errors, there are 2 specific errors 
    // and one event that you can handle individually.
    switch (arg.error) {
        case 12002:
            showNotification("Cannot load URL, no such page or bad URL syntax.");
            break;
        case 12003:
            showNotification("HTTPS is required.");
            break;
        case 12006:
            // The dialog was closed, typically because the user the pressed X button.
            showNotification("Dialog closed by user");
            break;
        default:
            showNotification("Undefined error in dialog window");
            break;
    }
}

function openLISRefDialog() {
    Office.context.ui.displayDialogAsync("https://khalidk7.github.io/SimpleDialogSampleWeb/DialogLIS.html",
        { height: 30, width: 30 }, dialogCallback);
}

function openCRSRefDialog() {
    Office.context.ui.displayDialogAsync("https://khalidk7.github.io/SimpleDialogSampleWeb/DialogCRS.html",
        { height: 35, width: 33 }, dialogCallback);
}

function openInsertAuthorsDialog() {
    Office.context.ui.displayDialogAsync("https://khalidk7.github.io/SimpleDialogSampleWeb/DialogAuthors.html",
        { height: 40, width: 33 }, dialogCallback);
}

function openDialogAsIframe() {
    //IMPORTANT: IFrame mode only works in Online (Web) clients. Desktop clients (Windows, IOS, Mac) always display as a pop-up inside of Office apps. 
    Office.context.ui.displayDialogAsync("https://khalidk7.github.io/SimpleDialogSampleWeb/DialogLIS.html",
        { height: 50, width: 50, displayInIframe: true }, dialogCallback);
}


