/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

// Log that the shared utils module is loaded
console.log("Shared utils module loaded");

// Global variables to track settings
let delaySendEnabled = true; // Default to enabled
let businessStartHour = 7; // Default to 7 AM
let businessEndHour = 18;  // Default to 6 PM (18:00)

/**
 * Loads user settings from roaming settings
 * @returns {Object} Current settings object
 */
function loadSettings() {
  console.log("loadSettings called");
  
  // Load delay send enabled setting
  if (Office.context.roamingSettings.get("delaySendEnabled") !== undefined) 
    delaySendEnabled = Office.context.roamingSettings.get("delaySendEnabled");
  else
    delaySendEnabled = true; // Default to enabled if not set 
    
  // Load business start hour (default: 7 AM)
  if (Office.context.roamingSettings.get("businessStartHour") !== undefined)
    businessStartHour = Office.context.roamingSettings.get("businessStartHour");
  else
    businessStartHour = 7; // Default to 7 AM if not set
    
  // Load business end hour (default: 6 PM / 18:00)
  if (Office.context.roamingSettings.get("businessEndHour") !== undefined)
    businessEndHour = Office.context.roamingSettings.get("businessEndHour");
  else
    businessEndHour = 18; // Default to 6 PM if not set
    
  console.log("Settings loaded - Delay enabled:", delaySendEnabled, 
              "Business hours:", businessStartHour, "to", businessEndHour);
              
  return { delaySendEnabled, businessStartHour, businessEndHour };
}

/**
 * Saves user settings to roaming settings
 * @param {boolean} delaySendEnabled - Whether delay send is enabled
 * @param {number} businessStartHour - Business start hour (0-23)
 * @param {number} businessEndHour - Business end hour (0-23)
 * @param {Function} callback - Optional callback function after save completes
 */
function saveSettings(newDelaySendEnabled, newBusinessStartHour, newBusinessEndHour, callback) {
  console.log("saveSettings called"); 
  
  // Update the global variables
  if (newDelaySendEnabled !== undefined) delaySendEnabled = newDelaySendEnabled;
  if (newBusinessStartHour !== undefined) businessStartHour = newBusinessStartHour;
  if (newBusinessEndHour !== undefined) businessEndHour = newBusinessEndHour;
  
  // Save all settings
  Office.context.roamingSettings.set("delaySendEnabled", delaySendEnabled);
  Office.context.roamingSettings.set("businessStartHour", businessStartHour);
  Office.context.roamingSettings.set("businessEndHour", businessEndHour);
  
  // Save asynchronously
  Office.context.roamingSettings.saveAsync(function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to save settings:", asyncResult.error.message);
    } else {
      console.log("Successfully saved all settings to roaming settings");
    }
    
    // Call the callback if provided
    if (callback && typeof callback === "function") {
      callback(asyncResult);
    }
  });
}

/**
 * Toggles the delay send feature on/off
 * @param {boolean} newState - The new state for delay send (if not provided, it will toggle)
 * @param {Function} callback - Optional callback function after save completes
 * @returns {boolean} New state of delaySendEnabled
 */
function toggleDelaySetting(newState, callback) {
  // Toggle or set the setting
  if (newState !== undefined) {
    delaySendEnabled = newState;
  } else {
    delaySendEnabled = !delaySendEnabled;
  }
  
  console.log("Delay send toggled:", delaySendEnabled);
  
  // Save the setting with callback
  saveSettings(delaySendEnabled, undefined, undefined, callback);
  
  return delaySendEnabled;
}

/**
 * Checks if the current time is outside of normal business hours
 * Business hours defined by businessStartHour and businessEndHour variables
 * @returns {boolean} true if current time is after business hours
 */
function isAfterBusinessHours() {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
  const hours = now.getHours();
  
  // Weekend check
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return true;
  }
  
  // Weekday time check (before businessStartHour AM or after businessEndHour PM)
  if (hours < businessStartHour || hours >= businessEndHour) {
    return true;
  }
  
  return false;
}

/**
 * Checks if a date is a US holiday
 * Includes: New Year's Day, MLK Day, Presidents Day, Memorial Day, 
 * Independence Day, Labor Day, Thanksgiving & day after, Christmas Eve & Christmas
 * @param {Date} date - The date to check
 * @returns {boolean} True if it's a holiday, false otherwise
 */
function isUSHoliday(date) {
  const year = date.getFullYear();
  const month = date.getMonth(); // 0-11
  const day = date.getDate();
  const dayOfWeek = date.getDay(); // 0-6, 0 is Sunday
  
  // New Year's Day (January 1)
  if (month === 0 && day === 1) {
    return true;
  }
  
  // Martin Luther King Jr. Day (3rd Monday in January)
  if (month === 0 && dayOfWeek === 1 && day >= 15 && day <= 21) {
    return true;
  }
  
  // Presidents Day (3rd Monday in February)
  if (month === 1 && dayOfWeek === 1 && day >= 15 && day <= 21) {
    return true;
  }
  
  // Memorial Day (Last Monday in May)
  if (month === 4 && dayOfWeek === 1) {
    // Check if it's the last Monday by seeing if the following Monday is in June
    const nextWeek = new Date(date);
    nextWeek.setDate(day + 7);
    if (nextWeek.getMonth() !== month) {
      return true;
    }
  }
  
  // Independence Day (July 4)
  if (month === 6 && day === 4) {
    return true;
  }
  
  // Labor Day (1st Monday in September)
  if (month === 8 && dayOfWeek === 1 && day <= 7) {
    return true;
  }
  
  // Thanksgiving Day (4th Thursday in November)
  if (month === 10 && dayOfWeek === 4 && day >= 22 && day <= 28) {
    return true;
  }
  
  // Day after Thanksgiving (4th Friday in November)
  if (month === 10 && dayOfWeek === 5 && day >= 23 && day <= 29) {
    return true;
  }
  
  // Christmas Eve (December 24)
  if (month === 11 && day === 24) {
    return true;
  }
  
  // Christmas Day (December 25)
  if (month === 11 && day === 25) {
    return true;
  }
  
  return false;
}

/**
 * Calculates the next business day at the configured start hour
 * Business days are Monday-Friday, excluding US holidays
 * @returns {Date} Date object representing the business start time on the next business day
 */
function calculateNextBusinessDayStart() {
  const now = new Date();
  const nextBusinessDay = new Date(now);
  
  // Set time to the configured business start hour
  nextBusinessDay.setHours(businessStartHour, 0, 0, 0);
  
  // If it's already past the business start hour, move to the next day
  if (now.getHours() >= businessStartHour && now.getMinutes() >= 0) {
    nextBusinessDay.setDate(nextBusinessDay.getDate() + 1);
  }
  
  // Loop until we find a valid business day (not a weekend or holiday)
  let isBusinessDay = false;
  while (!isBusinessDay) {
    const dayOfWeek = nextBusinessDay.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday
    
    // Check if it's a weekend
    if (dayOfWeek === 0) { // Sunday
      nextBusinessDay.setDate(nextBusinessDay.getDate() + 1); // Move to Monday
      continue;
    } else if (dayOfWeek === 6) { // Saturday
      nextBusinessDay.setDate(nextBusinessDay.getDate() + 2); // Move to Monday
      continue;
    }
    
    // Check if it's a holiday
    if (isUSHoliday(nextBusinessDay)) {
      nextBusinessDay.setDate(nextBusinessDay.getDate() + 1);
      continue;
    }
    
    // If we reach here, we found a business day
    isBusinessDay = true;
  }
  
  return nextBusinessDay;
}

/**
 * Removes the delay from the current message
 * @param {Function} callback - Optional callback function after removing delay
 */
function removeDelay(callback) {
  // Remove any scheduled delay
  const noDateTime = new Date(0); // Set to 0 to remove delay
  Office.context.mailbox.item.delayDeliveryTime.setAsync(noDateTime, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      if (callback && typeof callback === "function") {
        callback(asyncResult);
      }
      return;
    }
    
    // Show confirmation
    Office.context.mailbox.item.notificationMessages.addAsync("delayRemoved", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Delay has been removed. Email will send immediately when you click Send.",
      icon: "DelaySend.16x16",
      persistent: false
    });
    
    // Remove any existing notification about a delay
    Office.context.mailbox.item.notificationMessages.removeAsync("afterHoursNotification");
    
    // Call the callback if provided
    if (callback && typeof callback === "function") {
      callback(asyncResult);
    }
  });
}

/**
 * Format hours in AM/PM format for display
 * @param {number} hour - Hour in 24-hour format (0-23)
 * @returns {string} Formatted hour string (e.g., "7 AM")
 */
function formatHourForDisplay(hour) {
  const amPm = hour < 12 ? "AM" : "PM";
  const hourDisplay = hour > 12 ? hour - 12 : (hour === 0 ? 12 : hour);
  return `${hourDisplay} ${amPm}`;
}

// Export the functions and variables so they can be used in other files
module.exports = {
  delaySendEnabled,
  businessStartHour,
  businessEndHour,
  loadSettings,
  saveSettings,
  toggleDelaySetting,
  isAfterBusinessHours,
  isUSHoliday,
  calculateNextBusinessDayStart,
  removeDelay,
  formatHourForDisplay
};
