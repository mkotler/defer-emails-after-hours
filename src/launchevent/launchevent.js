/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

/* global Office */

// Import shared utilities
const { 
  delaySendEnabled: sharedDelaySendEnabled,
  businessStartHour: sharedBusinessStartHour,
  businessEndHour: sharedBusinessEndHour,
  loadSettings: sharedLoadSettings,
  isAfterBusinessHours,
  calculateNextBusinessDayStart,
  removeDelay: sharedRemoveDelay,
  toggleDelaySetting
} = require('../shared/utils.js');

// Global variables
let delaySendEnabled = sharedDelaySendEnabled;
let businessStartHour = sharedBusinessStartHour;
let businessEndHour = sharedBusinessEndHour;

// Wait for Office to initialize before using any Office APIs
Office.onReady(() => {
  console.log("Office is ready");
  // Register our event handlers after Office initialization
  Office.actions.associate("checkAfterHours", checkAfterHours);
  Office.actions.associate("toggleDelaySend", toggleDelaySend);
  Office.actions.associate("removeDelay", removeDelay);

  // Initialize from saved settings
  loadSettings();
});

/**
 * Loads user settings from roaming settings
 */
function loadSettings() {
  // Use the shared loadSettings function
  const settings = sharedLoadSettings();
  
  // Update local variables
  delaySendEnabled = settings.delaySendEnabled;
  businessStartHour = settings.businessStartHour;
  businessEndHour = settings.businessEndHour;
}

/**
 * Handler for new message compose event
 * This runs automatically when a new email is being composed
 */
function checkAfterHours(event) {
  console.log("checkAfterHours called");
  
  // Check if delay send is enabled and it's after hours
  const afterHours = isAfterBusinessHours();
  console.log("Is after business hours:", afterHours, "Business hours:", businessStartHour, "to", businessEndHour);
  
  if (delaySendEnabled && afterHours) {
    // Automatically set the delay delivery time and show notification
    setDelayedDelivery(event);
  } else {
    // Complete the event handling
    event.completed();
  }
}

/**
 * Sets delay delivery time to the start of the next business day
 * This happens automatically when a new message is composed after hours
 */
function setDelayedDelivery(event) {
  const delayTime = calculateNextBusinessDayStart();
  const options = { 
    weekday: 'long', 
    month: 'long', 
    day: 'numeric', 
    hour: 'numeric', 
    minute: 'numeric' 
  };
  const formattedDate = delayTime.toLocaleString('en-US', options);
  
  // Set the delayed delivery time
  Office.context.mailbox.item.delayDeliveryTime.setAsync(delayTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      event.completed(); // Complete the event if setting delay failed
      return;
    }
    
    // Show notification about the scheduled delay - no action buttons
    const notification = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Email scheduled to send at ${formattedDate} (next business day).`,
      icon: "DelaySend.16x16",
      persistent: true 
    };

    // Add the notification
    Office.context.mailbox.item.notificationMessages.addAsync("afterHoursNotification", notification, function(result) {
      // Complete the event handling
      event.completed();
    });
  });
}

/**
 * Toggles the delay send feature on/off
 */
function toggleDelaySend(event) {
  // Use the shared toggleDelaySetting function with a callback
  delaySendEnabled = toggleDelaySetting(undefined, function(asyncResult) {
    // Show a notification to the user
    const message = delaySendEnabled 
      ? "Delay Send feature is now enabled" 
      : "Delay Send feature is now disabled";

    // Remove any existing notification
    Office.context.mailbox.item.notificationMessages.removeAsync("toggleNotification", function() {
      // Add new notification after removing the old one
      Office.context.mailbox.item.notificationMessages.addAsync("toggleNotification", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "DelaySend.16x16",
        persistent: false
      });
      
      // Complete the command
      if (event) event.completed();
    });
  });
}

/**
 * Removes the delay from the current message
 */
function removeDelay(event) {
  // Use the shared removeDelay function with a callback
  sharedRemoveDelay(function() {
    // Complete the command
    if (event) event.completed();
  });
}

