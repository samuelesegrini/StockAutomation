/**
 * =====================================================
 * STOCK MARKET DATA UPDATE TRIGGERS
 * =====================================================
 * 
 * Copyright (c) 2024 Samuele Segrini
 * Licensed under MIT License
 * 
 * This script manages automated triggers for stock market data updates,
 * coordinating with different market timings across European and US exchanges.
 * 
 * @author Samuele Segrini <samuele.segrini@gmail.com>
 * @version 1.0.0
 * @license MIT
 * @lastModified 2024-10-22
 */

/**
 * TRIGGER MANAGEMENT SYSTEM
 * ========================
 * This system creates and manages time-based triggers for updating stock data
 * according to different market schedules.
 * 
 * Market Update Schedule (All times in CET - Central European Time):
 * 
 * European Markets:
 * - 09:00 - Market Open Update
 * - 12:00 - Mid-Day Update
 * - 17:00 - Market Close Update
 * 
 * US Markets:
 * - 15:30 - US Market Open Update
 * - 19:00 - US Mid-Day Update
 * - 22:00 - US Market Close Update
 * 
 * Time and markets be modified with whathever you like, remeber to
 * 1. Modify the triggers file
 * 2. Modify the CONFIG section
 * 
 * please refer to Google Documentation:
 * https://developers.google.com/apps-script/reference/script/clock-trigger-builder
 */

/**
 * Creates all required time-based triggers for stock market updates.
 * This is the main function that should be run manually to set up the automation.
 * 
 * WARNING: Running this function will delete all existing triggers and create new ones.
 * 
 * @throws {Error} If trigger creation fails
 */
function createStockUpdateTriggers() {
  try {
    Logger.log('Starting trigger setup process...');
    
    // Remove any existing triggers to prevent duplicates
    deleteExistingTriggers();
    
    // European market triggers (CET times)
    createDailyTrigger('updateGlobalStocks', 9, 0);   // Market open
    createDailyTrigger('updateGlobalStocks', 12, 0);  // Mid-day
    createDailyTrigger('updateGlobalStocks', 17, 0);  // Market close

    // US market triggers (CET times)
    createDailyTrigger('updateGlobalStocks', 15, 30); // US market open
    createDailyTrigger('updateGlobalStocks', 19, 0);  // US mid-day
    createDailyTrigger('updateGlobalStocks', 22, 0);  // US market close
    
    Logger.log('All triggers have been created successfully.');
  } catch (error) {
    Logger.log('Failed to create triggers: ' + error.toString());
    throw error;
  }
}

/**
 * Creates a single daily trigger for a specific time
 * 
 * @param {string} functionName - Name of the function to trigger
 * @param {number} hour - Hour in 24-hour format (0-23)
 * @param {number} minute - Minute (0-59)
 * @throws {Error} If trigger creation fails
 */
function createDailyTrigger(functionName, hour, minute) {
  try {
    ScriptApp.newTrigger(functionName)
      .timeBased()
      .atHour(hour)
      .nearMinute(minute)
      .everyDays(1)
      .create();
    
    Logger.log(`Created trigger for ${functionName} at ${hour}:${minute}`);
  } catch (error) {
    Logger.log(`Failed to create trigger for ${hour}:${minute}: ${error}`);
    throw error;
  }
}

/**
 * Removes all existing triggers from the project
 * 
 * @throws {Error} If trigger deletion fails
 */
function deleteExistingTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
    Logger.log(`Deleted ${triggers.length} existing triggers`);
  } catch (error) {
    Logger.log('Failed to delete existing triggers: ' + error.toString());
    throw error;
  }
}
