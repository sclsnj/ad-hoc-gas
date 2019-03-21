# ad-hoc-gas/Collection Tools
__Google Apps Scripts for SCLSNJ's Collection Tools add on__

This set of Google Apps Scripts (GAS) and HTML files constitute the Collection Tools add on created by staff at the Somerset County Library System of New Jersey for use with TLC's CARL integrated library system product and ad hoc reporting. Feel free to copy and modify; just make sure to use your own report server address and credentials.

NOTE: This add on is recycled from an older version that got its data totally differently, which is why the primary code.gs file which creates the menu item and formats the report is separate from the CARLcode.gs file containing the ad hoc logic.

__code.gs__
 * Primary add on script
 * Loads add on, creates menu option, launches sidebar (CARLsidebar.html), formats data
 * Runs as an add on (see https://developers.google.com/apps-script/add-ons/ for more info)

__CARLcode.gs__
 * Called by the add on sidebar
 * Uses variables from the side bar to construct and run a query
 * Puts raw data into the active Google Sheet file

__CARLcollchecksidebar.html__
 * Client-side user interaction
 * Includes logic that allows the user to set default values for custom flags

-----

The basic path this add on takes is:
 * code.gs: onInstall() -- Only invoked when the add on is first installed.
 * code.gs: onOpen()
 * code.gs: showCARLSidebar() -- Triggered by selecting CARL Collection Check from the add on menu.
 * CARLcollchecksidebar.html -- Presents form.
 * CARLcollchecksidebar.html: function() -- Runs when the form loads to get any preferences from the server side.
 * code.gs: getPreferences() -- Grabs any existing user preferences and sends them back into the sidebar.
 * CARLcollchecksidebar.html: loadPreferences() -- If the user had preferences, pre-populate form fields accordingly.
 * CARLcollchecksidebar.html: formatCollCheck() -- Triggered on form submit; returns form data to the server side.
 * CARLcode.gs: formatCARLCollCheck() -- The meat of the add on, constructs and runs the query.
 * CARLcode.gs: formatDateForSQL() -- Invoked in the main function to do some date normalizing.
 * code.gs: formatReport() -- Called at the end of the main function to handle formatting.
