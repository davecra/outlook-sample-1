/// <reference path="../node_modules/officejs.dialogs/dialogs.js" />
/// <reference path="../node_modules/easyews/easyews.js" />
'use strict';

/** @type {Event} */
var sendEvent;
/** @type {string[]} */
var groups = [];
/** @type {Array} */
var externals = [];
/** @type {completedCallbackDelegate} */ 
var completedCallback;
/** @type {string} */
var addinName = "";
/** @type {Office.Message} */
var item;
/** @type {string} */
var thisUser= "";
/** @type {string} */
var domain = "";
/** @type {Office.MessageCompose} */
var composeItem;

/**
 * Office Initializes here
 * @param {Office.InitializationReason} reason 
 */
Office.initialize = function (reason) {
  // init here
  item = Office.context.mailbox.item;
  composeItem = Office.cast.item.toItemCompose(item);
  thisUser = Office.context.mailbox.userProfile.emailAddress;
  domain = getDomain(thisUser);
  addinName = "Outlook Blocking Dialog with ExpandDL Sample";
};

/**
 * The Manifest points to this event when it detects 
 * the user pressing Send on an email message. See:
 * https://docs.microsoft.com/en-us/outlook/add-ins/outlook-on-send-addins?product=outlook
 * https://theofficecontext.com/2017/08/10/deploying-onsend-outlook-add-ins/
 * @param {*} event 
 */
function onSendEvent(event) {
  sendEvent = event; // grab this so it does not get cleaned up
  // show progress notification
  showProgress("The Outlook Demo add-in is processing this message. Please wait..."); 
  // check the To/CC/BCC lines for users not in this domain
  // split all groups and then evaluate, identify all external
  // users that are listed in all the groups...
  getExternalRecipientsAsync(function(){
    // were there any outside users?
    if(externals.length > 0) {
      // we have outside users found, so we need to
      // ask the user for Yes or No
      getResponseFromUser();
    } else {
      // there are no external users
      removeProgress();
      // Send - all internal
      sendEvent.completed({ allowEvent: true });
    } // end-if
  }, function(error) {
      removeProgress();
      showError(error, function() {
        // BLOCK THE SEND
        sendEvent.completed({ allowEvent: false }); 
      });
  }); 
}

/**
 * Displays an error to the user
 * @param {string} error 
 */
function showError(error, callback) {
  // uses the OfficeJS.dialogs Alert. See:
  // https://github.com/davecra/OfficeJS.dialogs
  // an error occurred trying to get all the emails on To/CC/BCC
  Alert.Show("Unable to process TO/CC/BCC: " + error, function() { 
      // Notification Message (error)
      Office.context.mailbox.item.notificationMessages.addAsync("error", {
      type: "errorMessage",
      message : "The Outlook Demo add-in failed to process this message."
    });
  }, callback); // Alert.Show
}

/**
 * Display a progress message to the user
 * @param {string} msg The message to display
 */
function showProgress(msg) {
  // Notify the user the message is being processed just in case
  // there are a LOT of groups and alot of user accounts
  Office.context.mailbox.item.notificationMessages.addAsync("progress", {
    type: "progressIndicator",
    message : msg
  });
}

/** 
 * Removes the progress message from the notifications area
*/
function removeProgress() {
  item.notificationMessages.removeAsync("progress");
}


function showInformation(msg) {
  item.notificationMessages.addAsync("information", {
    type: "informationalMessage",
    message : msg,
    icon : "icon16",
    persistent: false});
}

/**
 * Notified the user that there are external users and then
 * gets their response - yes (ok to send) or no (stop).
 * @param {sendCallbackDelegate} result 
 */
function getResponseFromUser() {
  /** @type {string} */
  var message = "There are users that are outside your organization on the To/CC/BCC.\n" +
                "Are you sure you want to send this message?"
  /** @type {string} */
  var title = "Blocking Send";
  // uses the OfficeJS.dialogs MessageBox. See:
  // https://github.com/davecra/OfficeJS.dialogs
  MessageBox.Show(message, title, MessageBoxButtons.YesNoCancel, 
                  MessageBoxIcons.Question, false, false, 
                  // callback when the user presses a button on the dialog
                  function(button) {
                    // did the user click Yes
                    if(button == "Yes") {
                      removeProgress();
                      // SEND
                      sendEvent.completed({ allowEvent: true });
                    } else {
                      removeProgress();
                      showInformation(addinName + " has stopped the message from being sent.");
                      // STOP THE SEND
                      sendEvent.completed({ allowEvent: false });
                    }
                  });
}

/**
 * Gets the domain portion of an email address. For example:
 *  - user@exchange.contoso.com = contoso.com
 *  - user@constoso.com = contoso.com
 * @param {string} user The email address of the user
 * @returns {string} domain name returned
 */
function getDomain(user) {
  /** @type {string} */
  var fullDomain = user.split("@")[1];
  /** @type {string[]} */
  var parts = fullDomain.split(".");
  /** @type {string} */
  var domain = parts[0] + "." + parts[1];
  if(parts.length > 2) {
    domain = parts[parts.length-2] + "." + parts[parts.length-1];
  }
  return domain;
}

/**
 * Gets all the recipients from the To/CC/BCC lines
 * @param {completedCallbackDelegate} successCallback
 * @param {errorCallbackDelegate} errorCallback
 */
function getExternalRecipientsAsync(successCallback, errorCallback) {
  // use for later
  completedCallback = successCallback;

  // get the TO line
  composeItem.to.getAsync(function(toAsyncResult) {
    if(toAsyncResult.error) {
      errorCallback(error);
    } else {
      /** @type {Office.Recipients} */
      var recipients = toAsyncResult.value;
      // if there are results, add them to the return array
      if(recipients.length > 0) { 
        recipients.forEach(function(recip, index) {
          if(recip.recipientType == Office.MailboxEnums.RecipientType.ExternalUser) {
            externals.push(recip.emailAddress);
          } else if(recip.recipientType == Office.MailboxEnums.RecipientType.DistributionList) {
            groups.push(recip.emailAddress);
          }
        });
      }
      // get the CC line
      composeItem.cc.getAsync(function(ccAsyncResult) {
        if(ccAsyncResult.error) {
          errorCallback(error);
        } else {
          /** @type {Office.Recipients} */
          var recipients = ccAsyncResult.value;
          // if we have results
          if(recipients.length > 0) {
            recipients.forEach(function(recip, index) {
              // only add unique/new items
              if(recip.recipientType == Office.MailboxEnums.RecipientType.ExternalUser) {
                externals.push(recip.emailAddress);
              } else if(recip.recipientType == Office.MailboxEnums.RecipientType.DistributionList) {
                groups.push(recip.emailAddress);
              }
            }); // forEach ccAsyncResult
          } // end-if ccAsyncResult.value.length

          // get the BCC line
          composeItem.bcc.getAsync(function(bccAsyncResult) {
            if(bccAsyncResult.error) {
              errorCallback(error);
            } else {
              /** @type {Office.Recipients} */
              var recipients = bccAsyncResult.value;
              if(recipients.length > 0) {
                recipients.forEach(function(recip, index) {
                  if(recip.recipientType == Office.MailboxEnums.RecipientType.ExternalUser) {
                    externals.push(recip.emailAddress);
                  } else if(recip.recipientType == Office.MailboxEnums.RecipientType.DistributionList) {
                    groups.push(recip.emailAddress);
                  } // end-if
                }); // forEach
                // call this function async, when it finished recursively calling
                // itself and splitting all groups it find, it will issue a callack
                // to this function callback defined now globally as completedCallback
                splitGroupsAndFindExternalsRecursivelyAsync();
              } else {
                splitGroupsAndFindExternalsRecursivelyAsync();
              } // end-if(bccAsyncResult.value.length > 0)
            } // end-if(bccAsyncResult.error)
          }); //composeItem.bcc.getAsync
        } // end-if(ccAsyncResult.error)
      }); // composeItem.cc.getAsync
    } // end-if(toAsyncResult.error)
  }); // to.getAsync
}

/**
 * Splits a group and calls the completed function
 */
function splitGroupsAndFindExternalsRecursivelyAsync() {
  if(groups.length == 0) { 
    // if no groups stop
    completedCallback(); 
  } else {
    /** @type {string} */
    var group = groups.pop();
    // call expandGroup to get users
    easyEws.expandGroup(group, function(groupUsers) {
      groupUsers.forEach(function(groupUser, index){
        if(groupUser.MailboxType() == "PublicDL") {
          groups.push(groupUser);
        } else {
          /** @type {string} */
          var emailDomain = getDomain(groupUser.Address());
          if(emailDomain != domain) {
            externals.push(groupUser.Address());
          }
        }
      }); // groupUsers.forEach
      splitGroupsAndFindExternalsRecursivelyAsync(); // recursive
    }, function(error) {
      console.log(error);
      // just fail
      completedCallback();
    }); // easyEws.expandGroup
  } // end-if
}

/**
 * This is the error callback
 * @callback errorCallbackDelegate
 * @param {string} error
 * @returns {void}
 */
var errorCallbackDelegate = function(error) { };
/**
 * This is the completed callback
 * @callback completedCallbackDelegate
 * @returns {void}
 */
var completedCallbackDelegate = function() { };