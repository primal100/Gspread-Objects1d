# Gspread-Objects1d
Work with 1-dimensional rows, columns and non-adjacent cells in Google Sheets

To use, go to Resources > Library in the Google Sheets script editor and enter the following project key:
MS8m0Cjb7P7PEESdMy2qrTYoHEfXZl5VD

```javascript
function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = Objects1d.getRow(sheet, 1, 1);     //Get row 1, header size is 1
  var values = row.getValues();                //Returns 1d array of values, excluding the header
  row.setValue(1, 'x');                        //Value not set in spreadsheet yet (uncomitted)
  row.setValue(2, 'y');                        //Value not set in spreadsheet yet (uncomitted)
  row.commit();                                //Values from previous two lines set in single request to gspread API (more efficent).
  values[4] = 'no';
  values[6] = 'yes';
  row.setValues(values);                       //Set values using a 1d array (committed)
  row.getCells();                              //Returns 1d array of cell objects
  row.count();                                 //List of values not including headers
  row.addValue('value');                       //Appends the value to the next empty column in the row (uncomitted)
  row.addValueIfNotExists('value')             //Adds a value to a row if it doesn't already exist (uncomitted)
  row.addValueIfNotExists(['value1', 'value2'])//Runs add valueIfNotExists for each value in the array (uncomitted)
  row.getCell(1);                              //Returns a cell object by its index in the row
  row.getHeader();                             //Returns a cell object of the row header
  var column = Objects1d.getColumn(sheet, 1, 'Name', 1, 1); //Returns a column by the header name 'Name', with headers in Row 1. Header Size: 1 and Column 1 contains headers to enable ORM
  var name = column.date //Returns cell in column with row header 'date'
}
```

In addition to rows and columns, objects1d also allows interations with Cell Collections (non-adjacent cells).

Lots more features to be documented soon, in the meantime check out the code.
