# ad-hoc-gas
__Google Apps Scripts for use with CARL Ad Hoc Reporting__

These sample Google Apps Scripts (GAS), created by staff at the Somerset County Library System of New Jersey, are intended for use with TLC's CARL integrated library system product and ad hoc reporting. Feel free to copy and modify; just make sure to use your own report server address and credentials.

__Circ Transaction Trends__
 * Unbound - script can be saved as a standalone file
 * Add a time-based trigger to schedule the script to run automatically (ours runs daily)
 * Moderate query
 * Looks for a file matching a particular year + name combination
   * If it finds the file, it appends the data to what's already there
   * If if doesn't find the file, it creates it and then adds the data

__Check ISBNs__
 * Unbound - script can be saved as a standalone file
 * Runs as an add on (see https://developers.google.com/apps-script/add-ons/ for more info)
 * Simple query
 * A loop runs the query a number of times based on the ISBNs that are already in the spreadsheet
 * Adds data to whatever is in the open spreadsheet

__Long In Transit__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Add a time-based trigger to schedule the script to run automatically (our runs weekly)
 * Moderate query
 * Includes lots of script logic to slice and dice data by branch
 * Replaces all existing data in every tab with new data

__High Holds__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Runs on demand via menu option
 * VERY complicated query - you'll want to do extensive testing before you embed in the script
   * See https://github.com/sclsnj/ad-hoc-gas/blob/master/High%20Holds%20annotated.sql for more in-depth info
 * Script includes some added logic to fine-tune hold ratios
 * Inserts new data into a new tab, instead of overwriting existing data

__Vendor Spending__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Includes a trigger to runs whenever the file is opened
 * Complicated query - you'll want to do some adaptation and testing before you embed in the script
 * Compares existing spreadsheet data with new data and overwrites when necessary

