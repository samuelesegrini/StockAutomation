/**
 * =====================================================
 * AUTOMATIC STOCK MARKET DATA UPDATER
 * =====================================================
 * 
 * Copyright (c) 2024 Samuele Segrini
 * Licensed under MIT License
 * 
 * @author Samuele Segrini <samuele.segrini@gmail.com>
 * @version 1.0.0
 * @license MIT
 * @lastModified 2024-10-22
 * 
 * IMPORTANT: Before using this script, you must set up triggers!
 * Run the createStockUpdateTriggers() function in the triggers.gs file
 * to set up automatic updates. See README for detailed instructions.
 */

/**
 * CONFIGURATION SECTION
 * ====================
 * 
 * This section contains all the settings you can customize.
 * Each setting is explained in detail below.
 */
const CONFIG = {
  // SHEET NAMES AND LAYOUT
  // ---------------------
  sheets: {
    // The sheet containing your list of stocks
    sourceSheetName: 'Recup',
    
    // The sheet where prices will be saved
    targetSheetName: 'Dati',
    
    // Which columns contain what information in your stock list sheet
    sourceColumns: {
      ticker: 3,    // Column C: Stock symbols (e.g., AAPL)
      exchange: 9   // Column I: Exchange names (e.g., NASDAQ)
    },
    
    // Where different information should be stored in your price history sheet
    targetColumns: {
      price: 1,        // Column A: Stock price formula
      exchange: 3,     // Column C: Exchange name
      ticker: 4,       // Column D: Stock symbol
      datetime: 5,     // Column E: Date and time
      priceResult: 6   // Column F: Actual price value
    }
  },
  
  // SUPPORTED STOCK EXCHANGES
  // ------------------------
  exchanges: {
    // European stock exchanges
    european: [
      'MIL',    // Milan
      'LSE',    // London
      'XETRA',  // Frankfurt Electronic
      'ETR',    // Frankfurt
      'BIT'     // Borsa Italiana
    ],
    
    // American stock exchanges
    american: [
      'NASDAQ', // NASDAQ
      'NYSE'    // New York Stock Exchange
    ]
  },
  
  // ERROR TRACKING AND DEBUGGING
  // ---------------------------
  debug: {
    // Set to false to disable detailed logging
    enabled: true,
    
    // Should errors be saved to a separate sheet?
    logToSheet: true,
    
    // Name of the sheet where errors will be logged
    logSheetName: 'Logs',
    
    // Should you receive emails when errors occur?
    emailOnError: false,
    
    // Email address for error notifications
    emailAddress: ''  // Add your email here if you want notifications
  }
};

/**
 * ERROR LOGGING SYSTEM
 * ===================
 * 
 * This system keeps track of what the script is doing and any problems it encounters.
 * It's like having a detailed diary of the script's activities.
 */
class Logger {
  /**
   * Creates a new logging system
   * @param {Object} debugConfig - Settings for how logging should work
   */
  constructor(debugConfig) {
    this.config = debugConfig;
    this.logs = [];
    // Store 50 logs before writing them to the sheet
    this.batchSize = 50;
  }

  /**
   * Records a normal message about what the script is doing
   * @param {string} message - The information to record
   */
  log(message) {
    if (this.config.enabled) {
      // Show the message in the script console
      console.log(message);
      
      // Save the message with timestamp
      this.logs.push({
        timestamp: new Date(),
        type: 'INFO',
        message: message
      });
      
      // Write to log sheet if we've collected enough messages
      if (this.logs.length >= this.batchSize && this.config.logToSheet) {
        this.writeToLogSheet();
      }
    }
  }

  /**
   * Records an error message when something goes wrong
   * @param {string} message - Description of what went wrong
   * @param {Error} error - The actual error that occurred
   */
  error(message, error) {
    // Always show errors in the console
    console.error(message, error);
    
    // Save the error with timestamp
    this.logs.push({
      timestamp: new Date(),
      type: 'ERROR',
      message: `${message} ${error.toString()}`
    });
    
    // Write to log sheet immediately if enabled
    if (this.config.logToSheet) {
      this.writeToLogSheet();
    }
  }

  /**
   * Saves all logged messages to the log sheet
   * Creates the log sheet if it doesn't exist
   */
  writeToLogSheet() {
    if (!this.logs.length) return;

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = ss.getSheetByName(this.config.logSheetName);
      
      // Create log sheet if it doesn't exist
      if (!logSheet) {
        logSheet = ss.insertSheet(this.config.logSheetName);
        logSheet.appendRow(['Time', 'Type', 'Message']);
      }

      // Prepare all logs to be written at once
      const logData = this.logs.map(log => [
        log.timestamp,
        log.type,
        log.message
      ]);

      // Write all logs in one operation
      logSheet.getRange(
        logSheet.getLastRow() + 1,
        1,
        logData.length,
        3
      ).setValues(logData);

      // Clear the log buffer
      this.logs = [];
    } catch (error) {
      console.error('Could not write to log sheet:', error);
    }
  }
}

/**
 * MAIN FUNCTION: UPDATES STOCK PRICES
 * ==================================
 * 
 * This is the main function you'll run to update your stock prices.
 * It coordinates the entire process from reading stocks to saving prices.
 * 
 * @param {Object} customConfig - Optional: Your custom settings to override defaults
 */
function updateGlobalStocks(customConfig = {}) {
  // Combine your custom settings with the defaults
  const config = { ...CONFIG, ...customConfig };
  const logger = new Logger(config.debug);
  
  // Don't run on weekends when markets are closed
  if (!isWeekday()) {
    logger.log('Today is a weekend - No updates needed');
    return;
  }

  try {
    logger.log('Starting stock price update process...');
    
    // Get references to your spreadsheets
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = {
      source: ss.getSheetByName(config.sheets.sourceSheetName),
      target: ss.getSheetByName(config.sheets.targetSheetName)
    };

    // Make sure both sheets exist
    if (!sheets.source || !sheets.target) {
      throw new Error(
        'Could not find required sheets. ' +
        'Please check that both your stock list and price history sheets exist ' +
        'and are named correctly in the CONFIG section.'
      );
    }

    // Process all stocks and track results
    const stats = processStocksWithStats(sheets, config, logger);
    
    // Log a summary of what was done
    logger.log(`
      Update Complete:
      - Total stocks checked: ${stats.total}
      - Prices updated: ${stats.updated}
      - Duplicates skipped: ${stats.duplicates}
      - Errors encountered: ${stats.errors}
    `);
  } catch (error) {
    logger.error('A serious error occurred:', error);
    
    // Send email notification if enabled
    if (config.debug.emailOnError && config.debug.emailAddress) {
      sendErrorEmail(error, config.debug.emailAddress);
    }
  }
}

/**
 * UPDATES ALL STOCKS AND TRACKS STATISTICS
 * ======================================
 * 
 * This function handles the actual work of updating stock prices
 * and keeps track of how many were updated, skipped, etc.
 */
function processStocksWithStats(sheets, config, logger) {
  // Initialize counters
  const stats = {
    total: 0,      // Total stocks processed
    updated: 0,    // Successfully updated
    duplicates: 0, // Skipped as duplicate
    errors: 0      // Errors encountered
  };
  
  try {
    // Get list of stocks to update
    const stockData = getStockData(sheets.source, config, logger);
    
    // Get existing entries to avoid duplicates
    const existingEntries = getExistingEntries(sheets.target, config, logger);
    const existingSet = new Set(
      existingEntries.map(entry => `${entry[0]}:${entry[1]}:${entry[2]}`)
    );
    
    stats.total = stockData.length;
    if (!stats.total) return stats;

    // Prepare updates in batches for better performance
    const batchUpdates = [];
    const currentDateTime = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm"
    );

    // Process each stock
    stockData.forEach(stock => {
      try {
        // Skip unsupported exchanges
        if (!isStockSupported(stock.exchange, config)) {
          logger.log(`Skipping unsupported exchange: ${stock.exchange}`);
          return;
        }

        // Check for duplicates
        const entryKey = `${stock.exchange}:${stock.ticker}:${currentDateTime}`;
        if (existingSet.has(entryKey)) {
          stats.duplicates++;
          return;
        }

        // Add to batch update list
        batchUpdates.push({
          symbol: `${stock.exchange}:${stock.ticker}`,
          exchange: stock.exchange,
          ticker: stock.ticker,
          dateTime: currentDateTime
        });
        
      } catch (error) {
        stats.errors++;
        logger.error(`Problem with stock ${stock.ticker}:`, error);
      }
    });

    // Update all stocks in one batch
    if (batchUpdates.length) {
      updateStockPrices(batchUpdates, sheets.target, config);
      stats.updated = batchUpdates.length;
    }

  } catch (error) {
    logger.error('Error during batch processing:', error);
    stats.errors++;
  }

  return stats;
}

/**
 * UPDATES PRICES FOR A BATCH OF STOCKS
 * ==================================
 * 
 * This function writes the new price data to your spreadsheet
 * in an efficient way by updating multiple stocks at once.
 */
function updateStockPrices(entries, targetSheet, config) {
  const lastRow = targetSheet.getLastRow() + 1;
  const cols = config.sheets.targetColumns;
  
  // Prepare Google Finance formulas
  const formulas = entries.map(entry => 
    [`=GOOGLEFINANCE("${entry.symbol}")`]
  );
  
  // Prepare other data
  const data = entries.map(entry => [
    entry.exchange,
    entry.ticker,
    entry.dateTime
  ]);

  // Write everything to the sheet
  targetSheet.getRange(lastRow, cols.price, entries.length, 1)
    .setFormulas(formulas);
  
  targetSheet.getRange(lastRow, cols.exchange, entries.length, 3)
    .setValues(data);
}

/**
 * READS STOCK LIST FROM SOURCE SHEET
 * ================================
 * 
 * Gets the list of stocks to update from your stock list sheet.
 */
function getStockData(sourceSheet, config, logger) {
  try {
    const lastRow = sourceSheet.getLastRow();
    
    // Read all stock data at once
    const data = sourceSheet.getRange(
      2,  // Start from row 2 (skip header)
      Math.min(config.sheets.sourceColumns.exchange, config.sheets.sourceColumns.ticker),
      lastRow - 1,
      Math.abs(config.sheets.sourceColumns.exchange - config.sheets.sourceColumns.ticker) + 1
    ).getValues();

    // Convert to easy-to-use format
    return data.map(row => ({
      exchange: row[config.sheets.sourceColumns.exchange - 
        Math.min(config.sheets.sourceColumns.exchange, config.sheets.sourceColumns.ticker)],
      ticker: row[config.sheets.sourceColumns.ticker - 
        Math.min(config.sheets.sourceColumns.exchange, config.sheets.sourceColumns.ticker)]
    }));
  } catch (error) {
    logger.error('Could not read stock list:', error);
    return [];
  }
}

/**
 * GETS EXISTING PRICE ENTRIES
 * =========================
 * 
 * Reads existing entries to avoid duplicates
 */
function getExistingEntries(targetSheet, config, logger) {
  try {
    const lastRow = targetSheet.getLastRow();
    return lastRow <= 1 ? [] : 
      targetSheet.getRange(2, 3, lastRow - 1, 4).getValues();
  } catch (error) {
    logger.error('Could not read existing entries:', error);
    return [];
  }
}

/**
 * CHECKS IF STOCK EXCHANGE IS SUPPORTED
 * ==================================
 * 
 * Verifies if we can get prices from this exchange
 */
function isStockSupported(exchange, config) {
  const allExchanges = [
    ...config.exchanges.european,
    ...config.exchanges.american
  ];
  return allExchanges.includes(exchange);
}

/**
 * CHECKS IF TODAY IS A WEEKDAY
 * ==========================
 * 
 * Returns true for Monday through Friday
 */
const isWeekday = () => {
  const day = new Date().getDay();
  return day > 0 && day < 6;
};

/**
 * SENDS ERROR NOTIFICATION EMAIL
 * ============================
 * 
 * Sends an email when something goes wrong (if enabled in settings)
 * 
 * @param {Error} error - The error that occurred
 * @param {string} emailAddress - Where to send the notification
 */
function sendErrorEmail(error, emailAddress) {
  // Don't send if no email address is configured
  if (!emailAddress) return;
  
  MailApp.sendEmail(
    emailAddress,
    'Stock Price Update Error Alert',
    `
    Dear User,

    The stock price update script encountered a problem at ${new Date().toISOString()}.

    Error Details:
    ${error.toString()}

    Please check your spreadsheet's log sheet for more information.

    This is an automated message.
    `
  );
}
