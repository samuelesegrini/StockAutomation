# Automatic Stock Market Data Updater
## Version 1.0.0

Copyright (c) 2024 Samuele Segrini  
Licensed under MIT License

## Overview
This Google Apps Script project automatically updates stock prices from multiple exchanges (European and US markets) and stores them in a Google Spreadsheet. It's designed to run automatically at specific times during market hours.

## Features
- Automatic price updates for European and US markets (Add more in the CONFIG section)
- Support for multiple stock exchanges
- Duplicate entry prevention
- Error logging and notification system
- Configurable update schedule
- Weekend detection (no updates on non-trading days)

## Supported Stock Exchanges (Add more in the CONFIG section)
### European Markets
- MIL (Milan Stock Exchange)
- LSE (London Stock Exchange)
- XETRA (Frankfurt Electronic Trading)
- ETR (Frankfurt Stock Exchange)
- BIT (Borsa Italiana)

### American Markets
- NASDAQ
- NYSE

## Prerequisites
1. Google Account with access to Google Sheets and Google Apps Script
2. Basic understanding of Google Sheets
3. Stocks list in a properly formatted spreadsheet

## Installation
1. Create a new Google Spreadsheet
2. Create two sheets:
   - "Recup" (for stock list)
   - "Dati" (for price history)
3. Open Script Editor: Tools > Script Editor
4. Create two new script files:
   - `stockUpdater.gs`: Copy the main updater code
   - `triggers.gs`: Copy the triggers code
5. Save both files

## Sheet Setup
### Stock List Sheet ("Recup")
- Column C: Stock symbols (e.g., AAPL, GOOGL)
- Column I: Exchange names (e.g., NASDAQ, NYSE)

### Price History Sheet ("Dati")
Will be populated automatically with:
- Column A: Stock price formula
- Column C: Exchange name
- Column D: Stock symbol
- Column E: Date and time
- Column F: Actual price value

IMPORTANT: If you want to change the colums or sheets please modify the `CONFIG` section

## Initial Setup
1. Open Script Editor
2. Run the `createStockUpdateTriggers` function once
3. Authorize the script when prompted
4. Check the execution log to verify trigger creation

## Update Schedule (CET Time)
### European Markets
- 09:00 - Market Open
- 12:00 - Mid-Day
- 17:00 - Market Close

### US Markets
- 15:30 - Market Open
- 19:00 - Mid-Day
- 22:00 - Market Close

IMPORTANT: If you want to edit times (you can setup whenever) please look at Google Documentation and edit the `triggers.gs` file

## Troubleshooting
### Common Issues
1. "Sheet not found" error:
   - Verify sheet names match CONFIG settings
   - Check case sensitivity
2. No updates occurring:
   - Confirm triggers are set up
   - Check execution logs
   - Verify it's a weekday
3. Invalid stock symbols:
   - Check symbol format
   - Verify exchange names
   - Remove extra spaces

## Contributing
1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a Pull Request

## License
This project is licensed under the MIT License - see the LICENSE section in the code files for details.

## Author
Samuele Segrini <samuele.segrini@gmail.com>

## Version History
- 1.0.0 (2024-10-22): Initial release
