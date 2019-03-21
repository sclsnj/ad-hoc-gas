# ad-hoc-gas/Collection Tools
__Google Apps Scripts for SCLSNJ's Collection Tools add on__

This set of Google Apps Scripts (GAS) and HTML files constitute the Collection Tools add on created by staff at the Somerset County Library System of New Jersey for use with TLC's CARL integrated library system product and ad hoc reporting. Feel free to copy and modify; just make sure to use your own report server address and credentials.

NOTE: This add on is recycled from an older version that got its data totally differently, which is why the primary code.gs file is separate from the CARLcode.gs file containing the ad hoc logic.

__code.gs__
 * Primary add on script
 * Loads add on, creates menu option, launches sidebar (CARLsidebar.html), formats data
 * Runs as an add on (see https://developers.google.com/apps-script/add-ons/ for more info)

__CARLcode.gs__
 * Called by the add on menu
 * Uses variables from the side bar to construct and run a query
 * Puts raw data into the active Google Sheet file

__CARLcollchecksidebar.html__
 * Client-side user interaction
 * Includes logic that allows the user to set default values for custom flags
