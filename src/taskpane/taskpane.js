/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Import shared utilities
const { 
  delaySendEnabled: sharedDelaySendEnabled,
  businessStartHour: sharedBusinessStartHour,
  businessEndHour: sharedBusinessEndHour,
  loadSettings: sharedLoadSettings,
  saveSettings,
  toggleDelaySetting,
  formatHourForDisplay
} = require('../shared/utils.js');

// Global variables to track settings (initialized from shared values)
let delaySendEnabled = sharedDelaySendEnabled;
let businessStartHour = sharedBusinessStartHour;
let businessEndHour = sharedBusinessEndHour;

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Initialize UI based on current settings
    loadSettings();    // Add event listeners
    document.getElementById("enable-delay").addEventListener("change", toggleDelaySend);
    document.getElementById("disable-delay").addEventListener("change", toggleDelaySend);
    document.getElementById("save-hours").addEventListener("click", saveBusinessHours);
    
    // Add event listener for when the taskpane gets focus
    window.addEventListener("focus", refreshSettings);
       
    // Also refresh settings on click events
    document.addEventListener("click", function() {
      setTimeout(refreshSettings, 100); // Refresh shortly after any click in the taskpane
    });
    
    // Force a refresh whenever the window visibility changes
    document.addEventListener("visibilitychange", function() {
      if (!document.hidden) {
        refreshSettings();
      }
    });
  }
});

/**
 * Loads user settings from roaming settings and updates the UI
 */
function loadSettings() {
  // Use the shared loadSettings function
  const settings = sharedLoadSettings();
  
  // Update local variables
  delaySendEnabled = settings.delaySendEnabled;
  businessStartHour = settings.businessStartHour;
  businessEndHour = settings.businessEndHour;
  
  // Update UI to reflect current settings
  if (delaySendEnabled) {
    document.getElementById("enable-delay").checked = true;
  } else {
    document.getElementById("disable-delay").checked = true;
  }
  
  document.getElementById("start-hour").value = businessStartHour;
  document.getElementById("end-hour").value = businessEndHour;
  
  // Update the current hours display
  updateCurrentHoursDisplay();
}

/**
 * Toggles the delay send feature on/off
 */
function toggleDelaySend() {
  console.log("toggleDelaySend called in taskpane.js");
  
  // Get the new state from the radio button
  const newState = document.getElementById("enable-delay").checked;
  
  // Use the shared toggle function with a callback
  delaySendEnabled = toggleDelaySetting(newState, function(asyncResult) {
    // Show notification
    const message = delaySendEnabled 
      ? "Delay Send feature is now enabled" 
      : "Delay Send feature is now disabled";
      
    // Remove any existing notification about a delay
    Office.context.mailbox.item.notificationMessages.removeAsync("toggleNotification", function() {
      // Show the notification
      Office.context.mailbox.item.notificationMessages.addAsync("toggleNotification", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "DelaySend.16x16",
        persistent: false
      });
    });
  });
}

/**
 * Refreshes the settings from roaming settings
 * This is called periodically and when the taskpane gets focus
 * to ensure the UI stays in sync with the actual settings
 */
function refreshSettings() {
  console.log("Refreshing settings in taskpane");
  
  // Only refresh if Office context is available
  if (Office && Office.context && Office.context.roamingSettings) {
    try {
      // Check for changes in all settings
      
      // Check delay enabled setting
      const savedDelaySetting = Office.context.roamingSettings.get("delaySendEnabled");
      if (savedDelaySetting !== undefined && savedDelaySetting !== delaySendEnabled) {
        console.log("Updating delaySendEnabled setting:", savedDelaySetting);
        delaySendEnabled = savedDelaySetting;
        
        // Update UI to reflect current setting
        if (delaySendEnabled) {
          document.getElementById("enable-delay").checked = true;
        } else {
          document.getElementById("disable-delay").checked = true;
        }
      }
      
      // Check business start hour setting
      const savedStartHour = Office.context.roamingSettings.get("businessStartHour");
      if (savedStartHour !== undefined && savedStartHour !== businessStartHour) {
        console.log("Updating businessStartHour setting:", savedStartHour);
        businessStartHour = savedStartHour;
        document.getElementById("start-hour").value = businessStartHour;
      }
      
      // Check business end hour setting
      const savedEndHour = Office.context.roamingSettings.get("businessEndHour");
      if (savedEndHour !== undefined && savedEndHour !== businessEndHour) {
        console.log("Updating businessEndHour setting:", savedEndHour);
        businessEndHour = savedEndHour;
        document.getElementById("end-hour").value = businessEndHour;
      }
    } catch (error) {
      console.error("Error refreshing settings:", error);
    }
  }
}

/**
 * Saves business hours settings
 */
function saveBusinessHours() {
  // Get values from dropdowns
  const newStartHour = parseInt(document.getElementById("start-hour").value);
  const newEndHour = parseInt(document.getElementById("end-hour").value);
  
  // Validate the values
  if (newStartHour >= newEndHour) {
    // Show error notification
    Office.context.mailbox.item.notificationMessages.addAsync("hoursError", {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Start time must be earlier than end time.",
      icon: "DelaySend.16x16",
      persistent: false
    });
    return;
  }
  
  // Update the local variables
  businessStartHour = newStartHour;
  businessEndHour = newEndHour;
  
  // Use the shared saveSettings function with a callback
  saveSettings(undefined, businessStartHour, businessEndHour, function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to save business hours:", asyncResult.error.message);
      
      // Show error notification
      Office.context.mailbox.item.notificationMessages.addAsync("saveError", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Failed to save business hours. Please try again.",
        icon: "DelaySend.16x16",
        persistent: false
      });
    } else {
      console.log("Successfully saved business hours:", businessStartHour, "to", businessEndHour);
      
      // Update the current hours display
      updateCurrentHoursDisplay();
      
      // Format times for display for the notification
      const startAmPm = businessStartHour < 12 ? "AM" : "PM";
      const startHour = businessStartHour > 12 ? businessStartHour - 12 : businessStartHour;
      
      const endAmPm = businessEndHour < 12 ? "AM" : "PM";
      const endHour = businessEndHour > 12 ? businessEndHour - 12 : businessEndHour;
      
      // Show success notification
      Office.context.mailbox.item.notificationMessages.addAsync("hoursUpdated", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `Business hours updated to ${startHour} ${startAmPm} - ${endHour} ${endAmPm}.`,
        icon: "DelaySend.16x16",
        persistent: false
      });
    }
  });
}

/**
 * Updates the display of current business hours in a human-readable format
 */
function updateCurrentHoursDisplay() {
  // Use the shared formatHourForDisplay function
  const startFormatted = formatHourForDisplay(businessStartHour);
  const endFormatted = formatHourForDisplay(businessEndHour);
  
  // Update the display if the element exists
  const currentHoursElement = document.getElementById("current-hours");
  if (currentHoursElement) {
    currentHoursElement.textContent = `${startFormatted} - ${endFormatted}, Monday-Friday`;
  }
}
