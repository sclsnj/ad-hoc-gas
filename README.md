# ad-hoc-gas
__Google Apps Scripts for use with CARL Ad Hoc Reporting__

These sample Google Apps Scripts (GAS) are intended for use with TLC's CARL integrated library system product and ad hoc reporting. Feel free to copy and modify, and make sure to use your own report server address and credentials.

__High holds__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Runs on demand via menu option
 * VERY complicated query - you'll want to do extensive testing before you embed in the script
 * Script includes logic around hold ratios
 * Inserts new data into a new tab, instead of overwriting existing data

__Check ISBNs__
 * Unbound - script can be saved as a standalone file
 * Runs as an add on (see https://developers.google.com/apps-script/add-ons/ for more info)
 * Simple query
 * Loop runs the query a number of times based on the ISBNs that are already in the spreadsheet
 * Adds data to whatever is in the open spreadsheet

__Vendor spending__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Includes a trigger to runs whenever the file is opened
 * Complicated query - you'll want to do some adaptation and testing before you embed in the script
 * Compares existing spreadsheet data with new data and overwrites when necessary

__Long in transit__
 * Bound - script needs to be added to the Script Editor for a specific Google Sheet in order to run
 * Runs on a schedule (weekly)
 * Moderate query
 * Includes lots of script logic to slice and dice data by branch
 * Replaces all existing data in every tab with new data
