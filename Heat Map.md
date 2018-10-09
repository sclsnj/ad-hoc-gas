## Getting a Google Sheet to be a dynamically-updatable heat map

This is a two-step process
1. First get the data that corresponds to what you're heat mapping and/or matches user-defined criteria -- DSUM
2. Format the heat map based on the relationships among the values -- Conditional Formatting

### Using the DSUM function
`=dsum(Data!$A:$E, "Transactions", ({{"Time";"=9"},{"Day";"=2"},{"Date";">="&$F$1},{"Date";"<="&$F$2},{"Branch";if(NOT(ISBLANK($B$1)),"="&$B$1,"")}}))`

This treats the source sheet as a database and allows you to sum across data that meets particular criteria. 
It takes the form **DSUM(database, field, criteria)** where:
- *database* is the array or range containing the data to consider, structured in such a way that the first row contains the labels for each column's values
- *field* indicates which column in database contains the values to be extracted and operated on
- *criteria* represents an array or range containing zero or more criteria to filter the database values by before operating

In this case, the database is the columns containing data from the Data sheet.
The field is what's in the Transactions column.
The criteria is actually a whole array of things:
```
{
  {"Time";"=9"},
  {"Day";"=2"},
  {"Date";">="&$F$1},
  {"Date";"<="&$F$2},
  {"Branch";if(NOT(ISBLANK($B$1)),"="&$B$1,"")}
}
```
  - Time is the hour of the day recorded for the transaction.
  - Day is the day of the week recorded for the transaction.
  - There are two date criteria to set up the date range, each one with an absolute reference to the cell that the user can update.
  - Branch looks to see if there's anything in the branch cell that the use can update. If there is, it uses that. If there isn't, it leaves that value empty so that the function works across all of the data.

Each cell in the heat map has a unique formula in it, corresponding to the appropriate hour and day of the week.

### Conditional Formatting

Google Sheets has built-in conditional formatting that colors each cell in a range to highlight the relationship among the values in the range.

1. Select the range you want to format.
2. Open the **Conditional formatting** dialog (either through the Format menu or by right-clicking).
3. At the top of the dialog, choose the "Color scale" option.
4. You can use a default color scale, or you can choose your own min and max colors.
    * The default is to spread the colors evenly across the scale, but by specifying a midpoint percentile, you can weight the colors low or high, depending on what's more useful for you.
5. When you're ready, save the formatting rule.

As the values in the cells update, the color scale will shift accordingly.
