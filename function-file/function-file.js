'use strict';
var dialog;
var sendEvent;

Office.initialize = function (reason) {
  // init here
};

function onSendEvent(event) {
  sendEvent = event;
  // dispaly a dialog in a frame
  Office.context.ui.displayDialogAsync('https://localhost:3000/function-file/dialog.html',
      { height: 20, width: 30, displayInIframe: true },
      function (asyncResult) {
          dialog = asyncResult.value;
          // callbacks from the parent
          dialog.addEventHandler(Office.EventType.DialogEventReceived, processMessage);
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      });
}

function processMessage(arg) {
    // close the dialog
    dialog.close();
    // procress the result
    if(arg.error == 12006) {  
      // user clicked the (X) on the dialog 
      sendEvent.completed({ allowEvent: false }); 
    } else {
      if(arg.message=="Yes") {
        // user clicked yes
        sendEvent.completed({ allowEvent: true });
      } else {
        // user clicked no
        sendEvent.completed({ allowEvent: false });
      }
    }
} 

