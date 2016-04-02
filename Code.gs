function errorIfNegative(number, message, includeZero){
  includeZero = includeZero || false;
  if ((number < 0) || (includeZero && number == 0)){  
    throw new Error('argument to' + message + ' must not be a negative number'); 
  }
}
  
/* Filters a 1d array (row, column or cell collection) by value
*
* @param {Array} valueList Array of values
* @param {Value} value Value to filter by
* @return {Array} indexes Array of integers representing indexes in the valueList who's value matches 
*/
function filterByValue(valueList, value){
  var indexes = [], i = -1;
  while ((i = valueList.indexOf(value, i+1)) != -1){
        indexes.push(i);
    }
  return indexes;   
}

function splitRangeIntoCellIndexes(range){
	var indexes = [];
	var firstCell = range.getCell(1,1);
	var firstRow = firstCell.getRow();
	var firstCol = firstCell.getColumn();
    var height = range.getHeight();
    var width = range.getWidth();
	for ( var row = 0; row < height; row++ ){
		for ( var col = 0; col < width; col++ ){
			indexes.push([row + firstRow, col + firstCol]);
		}
	}
    return indexes;	
}

function columnLetterToNumber(sheet, columnLetter){
 return sheet.getRange(columnLetter + "1").getColumn(); 
}

/* Object for storing a 2d range with extra functionality
*
* @param {Sheet} sheet The Sheet object the range belongs to
* @param {Range} range The Range object
* @param {Integer} rowHeaderSize The number of rows containing headers
* @param {Integer} columnHeaderSize The number of columns containing headers
*/
function Range2d(sheet, range, rowHeaderSize, columnHeaderSize){
  this.sheet = sheet;
  this.range = range;
  this.startRow = range.getRow();
  this.startColumn = range.getColumn;
  this.rowHeaderSize = rowHeaderSize || 0;
  this.columnHeaderSize = columnHeaderSize || 0;
  this.getValue = function(index, row, column){ 
   this.values = this.values || this.getValues();
   return this.values[row][column];
  }
  this.getValues = function(){
    var values = this.getRange().getValues();
    this.values = values;
    return values;
  }
  this.setValue = function(row, column, value){
    this.values = this.values || this.getValues();
    this.values[row][column] = value; 
    return this
  }
  this.setValues = function(values){
    var numRows = values.length;
    var numColumns = values[0].length;
    this.sheet.getRange(this.startRow, this.startColumn, numRows, numColumns).setValues(values);
    this.values = values;
    return this;
  }  
  this.commit = function(){
    this.setValues(this.values);
    return this;
  }
  this.getRange = function(){
    return this.range;
  }
}

/* Prototype for row, column and non-adjacent cell collection
*/
function Range1d(){
  this.sheet = false;
  this.headerSize = 0;
  this.getValue = function(index){ 
   this.values = this.values || this.getValues();
   return this.values[index];
  }
  this.setValue = function(index, value){
    if (typeof index === "string"){
      var index = getIndexByHeader(index);
    }
    this.values = this.values || this.getValues()
    this.values[index] = value 
    return this;
  }
  this.commit = function(){
    this.setValues(this.values);
    return this;
  }
  this.getList = function(func){
   var items = [];
    for ( i=0; i<= this.count(); i++ ){
      items[i] = func(i);
    }
    return items
  }
  this.getCells = function(){
    return this.getList(this.getCell);
  }
  this.fillArray = function(value){
    arr = [];
    for ( i=0; i < this.count(); i++ ){
      array[i] = value;
    }
    return array;
  }
  this.getBackgrounds = function(){
    return this.to1d(this.getRange().getBackgrounds());
  }
  this.getDataValidations = function(){
    return this.to1d(this.getRange().getDataValidations());
  }
  this.getFontColors = function(){
    return this.to1d(this.getRange().getFontColors()); 
  }
  this.getFontFamilies = function(){
    return this.to1d(this.getRange().getFontFamilies());
  }
  this.getFontSizes = function(){
    return this.to1d(this.getRange().getFontSizes());   
  }
  this.getFontStyles = function(){
    return this.to1d(this.getRange().getFontStyles());  
  }
  this.getFontWeights = function(){
    return this.to1d(this.getRange().getFontWeights());
  }
  this.getFormulas = function(){
    return this.to1d(this.getRange().getFormulas());
  }
  this.getFormulasR1C1 = function(){
    return this.to1d(this.getRange().getFormulasR1C1()); 
  }
  this.getHorizontelAlignments = function(){
    return this.to1d(this.getRange().getHorizontelAlignment());  
  }  
  this.getNotes = function(){
    return this.to1d(this.getRange().getNotes()); 
  }  
  this.getNumberFormats = function(){
    return this.to1d(this.getRange().getNumberFormats());
  } 
  this.getValues = function(){
    var values = this.to1d(this.getRange().getValues());
    this.values = values;
    return values;
  }
  this.getVerticalAlignments = function(){
    return this.to1d(this.getRange().getVerticalAlignments());   
  }  
  this.getWraps = function(){
    return this.to1d(this.getRange().getWraps());
  } 
  this.setBackgrounds = function(values1d){
    return this.getRange().setBackgrounds(this.to2d(values1d, this.get2dParamater('Backgrounds')));
  }
  this.setDataValidations = function(values1d){
    return this.getRange().setDataValidations(this.to2d(values1d, this.get2dParamater('DataValidations')));
  }
  this.setFontColors = function(values1d){
    return this.getRange().setFontColors(this.to2d(values1d, this.get2dParamater('FontColors')));
  }
  this.setFontFamilies = function(values1d){
    return this.getRange().setFontFamilies(this.to2d(values1d, this.get2dParamater('FontFamilies')));
  }  
  this.setFontLines = function(values1d){
    return this.getRange().setFontLines(this.to2d(values1d, this.get2dParamater('FontLines')));
  }
  this.setFontSizes = function(values1d){
    return this.getRange().setFontSizes(this.to2d(values1d, this.get2dParamater('FontSizes')));
  }   
  this.setFontStyles = function(values1d){
    return this.getRange().setFontStyles(this.to2d(values1d, this.get2dParamater('FontStyles')));
  }
  this.setFontWeights = function(values1d){
    return this.getRange().setFontWeights(this.to2d(values1d, this.get2dParamater('FontWeights')));
  }   
  this.setFormulas = function(values1d){
    return this.getRange().Formulas(this.to2d(values1d, this.get2dParamater('Formula')));
  }
  this.setFormulasR1C1 = function(values1d){
    return this.getRange().setFormulasR1C1(this.to2d(values1d, this.get2dParamater('FormulasR1C1')));
  }  
  this.setHorizontalAlignments = function(values1d){
    return this.getRange().setHorizontalAlignments(this.to2d(values1d, this.get2dParamater('HorizontalAlignments')));
  }
  this.setNotes = function(values1d){
    return this.getRange().setNotes(this.to2d(values1d, this.get2dParamater('Notes')));
  }
  this.setNumberFormats = function(values1d){
    return this.getRange().setNumberFormats(this.to2d(values1d, this.get2dParamater('NumberFormats')));
  }
  this.setValues = function(values1d){
    var newValues = this.to2d(values1d, this.get2dParamater('Values'));
    var range = this.getRange().setValues(newValues);
    this.values = newValues;
    return range;
  }  
  this.setVerticalAlignments = function(values1d){
    return this.getRange().setVerticalAlignments(this.to2d(values1d, this.get2dParamater('VerticalAlignments')));
  }
  this.setWraps = function(values1d){
    return this.getRange().setWraps(this.to2d(values1d, this.get2dParamater('Wraps')));
  } 
  this.indexes2Cells = function(indexes){
     var cells = [];
     for ( var i in indexes ){
        cellIndexes[i] = this.getCellIndex(indexes[i]);
     }
     var cellCollection = getCellCollection(this.sheet, indexes);       
     return cellCollection;   
  }
  this.filter = function(value, useCachedValues, filterType){
    var useCachedValues = useCachedValues || false;
    var filterType = filterType || 0;

    if (useCachedValues){
      var values = this.values || this.getValues();
    }
    else{
      var values = this.getValues();
    }
    if (filterType == 0){
      var indexes = filterByValue(values, value);
    }
    return indexes;
  }
  this.first = function(value, useCahcedValues, filterType){
    var useCachedValues = useCachedValues || false;
    var filterType = filterType || 0;
    var indexes = this.filter(value);
    if ( indexes.length > 0){
      var index = indexes[0];
      return index;
    }
    else{
      throw new Error(value + ' not found in this range'); 
    }  
}
}

/* Object for storing a non-adjacent collection of cells. Inherits from Range1d
*
* @param {Sheet} sheet The Sheet object the cells belong to
* @param {Array} cellIndexes A 2d array containing row and column indexes
*/
function CellCollection(sheet, cellIndexes){
  this.sheet = sheet;
  this.indexes = cellIndexes;
  this.addCell = function(row, col){
    this.indexes.push(row, col); 
    this.range = this.indexes2range();
    if ( this.values !== undefined ){
      var length = this.values.length;
      var value = this.getCell(length - 1).getValue();
      this.values.push(value);
    }
    return this
  }
  this.indexes2Range = function(){
    var rows = [];
    var columns = [];
    for ( var i in this.indexes ){
      rows.push(this.indexes[i][0]);
      columns.push(this.indexes[i][1]);
    }
    rows.sort();
    columns.sort();
    var range = this.sheet.getRange(rows[0], columns[0], rows.slice(-1)[0], columns.slice(-1)[0]);
    return range; 
  }
  this.range = this.indexes2Range();
  this.getRange = function(){
   return this.range; 
  }
  this.getIndex = function(row, col){
    var index = filterByValue(this.indexes, [row, col]);
    return index;
  }
  this.getCell = function(index){
    var cellIndex = this.indexes[index];
    var cell = this.range.getCell(cellIndex[0], cellIndex[1]);
    return cell;
  }
  this.getCellIndex = function(index){
    return this.indexes[index];
  }
  this.count = function(){
    return this.indexes.length;
  }
  this.to1d = function(array2d){
    var array1d = [];
    for ( var i in this.indexes ){
      var row = this.indexes[i][0] -1;
      var column = this.indexes[i][1] - 1;
      array1d[i] = array2d[row][column];
    }
    return array1d;
  }
  this.to2d = function(array1d, rangeValues){
    for ( var i in this.indexes ){
      var row = this.indexes[i][0] -1;
      var column = this.indexes[i][1] - 1;
      rangeValues[row][column] = array1d[i];
    }
    return rangeValues;
  }  
  this.get2dParamater = function(dataType){
    return eval('this.getRange().get' + dataType + '()');
  }
  this.getA1Notation = function(){
    var a1Array = [];
    for ( i=0; i<= this.count(); i++ ){
      cells[i] = this.getCell(i).getA1Notation();
    }
    return cells;
  }
  this.isBlank = function(){
    var values = this.getValues().join("");
    if ( values == "" ){
      return true;
    }
    else{
      return false;  
    }
  }
}

CellCollection.prototype = new Range1d();

/* Prototype for columns and rows. Inherits from Range1d
*/
function ColumnOrRow(){
  this.startPoint = function(){
    return this.headerSize + 1;
  }
  this.count = function(){
    var count = this.last() - this.startPoint() + 1;
    return count
  }
  this.last = function(){
    if ( this.values === undefined ){
      var last = this.lastOnSheet() - this.headerSize;
    }
    else{
      var last = this.values.length;
    }
    var last = last || 1;
    return last;
  }
  this.getActual = function(index){
    var actual = index + this.headerSize + 1;
    return actual
  }
  this.addValue = function(value){
    this.values == this.values || this.getValues();
    var index = this.values.length;  
    this.values[index] = value;
    return index;
  }
  this.addValueIfNotExists = function(value, filterType){
    var filterType = filterType || 0;
    var valueExists = this.filter(value, true, filterType);
    if (valueExists.length > 0){
      return valueExists[0];
    }
    else{
      var row = this.addValue(value);
      return row;
    }
  }
  this.addValuesIfNotExists = function(values, filterType){
    var filterType = filterType || 0;
    var indexes = []
    for ( var i  in values ){
      indexes.push(addValueIfNotExists(values[i], filterType));
      last ++;
    }
    return indexes;
  }
  this.to1d = function(values2d){
    var values1d = [];
    for(var i = 0; i < values2d.length; i++){
      values1d = values1d.concat(values2d[i]);
    }
    return values1d;
  }
  this.to2d = function(values1d, spliceby){
    var values1dcopy = values1d.slice();
    var values2d = [];
    while(values1dcopy.length) values2d.push(values1dcopy.splice(0, spliceby));
    return values2d;
  }
  this.getCell = function(index){
    if (typeof index === "string"){
      var index = this.getIndexByHeader(index);
    }
    var cellIndexes = this.getRangeOrder(index + 1, 1);
    var cell = this.getRange().getCell(cellIndexes[0], cellIndexes[1]);
    return cell;
  }
  this.getValue = function(index){
    if (typeof index === "string"){
      var index = this.getIndexByHeader(index);
    }
    var value = this.getCell(index).getValue();
    this.values[index] = value;
    return value;
  }
  this.getHeader = function(headerIndex){
    var headerIndex = headerIndex || 1;
    var cellIndexes = this.getRangeOrder(headerIndex);
    var header = this.sheet.getRange.getCell(cellIndexes[0], cellIndexes[1]);
    return header;
  }
  this.getIndexByHeader = function(headerName){
    var headerRange = this.headerRange || this.getHeaderRange();
    var index = this.headerRange.filter(headerName)[0];
    return index;
  }
}

ColumnOrRow.prototype = new Range1d();

/* Object for managing a column as a 1d array. Inherits from ColumnOrRow
*
* @param {Sheet} sheet The Sheet object the column belongs to
* @param {Integer} headerSize The number of rows to ignore when getting and setting values
* @param {Integer} columnNumber The column number from which to create the column object
* @param {Integer} headerColumn The column containing the header names to enable ORM-like functionality. Optional. If not set the first column in the sheet is used.
*/
function Column(sheet, headerSize, columnNumber, headerColumn){
  this.sheet = sheet;
  this.column = columnNumber;
  this.header = headerColumn || 1;
  this.headerSize = headerSize;
  this.id = this.column;
  this.lastOnSheet = function(){
    return this.sheet.getLastRow();
  }
  this.getRangeOrder = function(row, col){
    var col = col || this.id;
    return [row, col];
  }
  this.getRange = function(){
    var numRows = this.last();
    var range = this.sheet.getRange(this.startPoint(), this.column, numRows);
    return range;
  }
  this.get2dParamater = function(getFunc){
    return 1;
  }
  this.getHeaderRange = function(){
    this.headerRange = this.getColumn(this.sheet, this.headerSize, this.header);
    return this.headerRange;
  }
  this.setValues = function(values1d){
    var values2d = this.to2d(values1d, 1);
    var range = this.sheet.getRange(this.startPoint(), this.column, values1d.length).setValues(values2d);
    this.values = values2d;
    return range;
  }
  this.sortSheet = function(descending){
    var descending = descending || false;
    var ascending = ! descending;
    var range = this.sheet.getRange(headerSize + 1, 1, this.sheet.getLastRow(), this.sheet.getLastColumn()) 
    range.sort({column: this.column, ascending: ascending});
  }
}

Column.prototype = new ColumnOrRow();

/* Object for managing a row as a 1d array. Inherits from ColumnOrRow
*
* @param {Sheet} sheet The Sheet object the row belongs to
* @param {Integer} headerSize The number of columns to ignore when getting and setting values
* @param {Integer} rowNumber The row number from which to create th row object
* @param {Integer} headerRow The row containing the header names to enable ORM-like functionality. Optional. If not set the first row in the sheet is used.
*/
function Row(sheet, headerSize, rowNumber, headerRow){
  this.sheet = sheet;
  this.row = rowNumber;
  this.header = headerRow || 1;
  this.headerSize = headerSize;
  this.id = this.row;
  this.getCellRange = function(column){
    return [this.row, column];
  }
  this.lastOnSheet = function(){
    return this.sheet.getLastColumn();
  }
  this.getRangeOrder = function(col, row){
    var row = row || this.id;
    return [row, col];
  }
  this.getRange = function(){
    var numColumns = this.last();
    var range = this.sheet.getRange(this.row, this.startPoint(), this.row, numColumns);
    return range;
  }
  this.getHeaderRange = function(){
    this.headerRange = this.getRow(this.sheet, this.headerSize, this.header);
    return this.headerRange;
  }
  this.get2dParamater = function(getFunc){
    return this.count();
  }
  this.setValues = function(values1d){
    var numColumns = values1d.length;
    var values2d = this.to2d(values1d, numColumns);
    var range = this.sheet.getRange(this.row, this.startPoint(), this.row, numColumns).setValues(values2d);
    this.values = values2d;
    return range;
  }
}

Row.prototype = new ColumnOrRow();
                                                   
/**
* Converts a 2d range to a 1d cells collection object
*
* @param {Sheet} sheet The sheet the cells belong to
* @param {Array} indexes An array of cells defined by rows and columns 
* @return {CellCollection} the collection of cells
*/
function getCellCollection(sheet, indexes){
  	if (typeof indexes[0] == "string"){
		var cellIndexes = [];
		for ( var i in indexes ){
			var range = sheet.getRange(indexes[i]);
			var individualIndexes = splitRangeIntoCellIndexes(range);
            for ( var i in individualIndexes ){
				cellIndexes.push(individualIndexes[i]);
			}
		}
	}
	else{
		var cellIndexes = cells;	
	}                                              
  var cells = new CellCollection(sheet, cellIndexes);
  return cells;
}

/* Returns a Column object
*
* @param {Sheet} sheet the sheet object containing the column
* @param {Integer} headerSize the number of rows containing headers. Defaults to 0.
* @param {ColumnIdentifier} columnIdentifier the column id by number or string (which matches the header of the column). This paramater is optional, if not given the next available empty column will be used.
* @param {Integer} headerRowIdentifier If columnIdenitifer is a string, this is the row which will be searched to find a column whos' header matches the string. Defaults to 1.
* @param {Integer} headerColumn The column containing the header names to enable ORM-like functionality. Optional. If not set the first column in the sheet is used.
* @return {Column} the column object
*/
function getColumn(sheet, headerSize, columnIdentifier, headerRowIdentifier, headerColumn){
  var headerSize = headerSize || 0;
  var headerColumn = headerColumn || 1;
  var columnIdentifier = columnIdentifier || sheet.getLastColumn() + 1;
  if (typeof columnIdentifier === "string"){
    var headerRowIdentifier = headerRowIdentifier || 1;
    var headerRow = getRow(sheet, 0, headerRowIdentifier);
    var columnIdentifier = headerRow.first(columnIdentifier);
  }
  else{
    errorIfNegative(headerSize, 'getColumn(): headerSize');
    }
  return new Column(sheet, headerSize, columnIdentifier, headerColumn);
}

/* Returns a row object
*
* @param {Sheet} sheet the sheet object containing the row
* @param {Integer} headerSize the number of columns containing headers. Defaults to 0.
* @param {RowIdentifier} rowIdentifier the row id by number or string (which matches the header of the row). This paramater is optional, if not given the next available row will be used.
* @param {Integer} headerColumnIdentifier If rowIdenitifer is a string, this is the column which will be searched to find a row whos' header matches the string. Defaults to 1.
* @param {Integer} headerRow The row containing the header names to enable ORM-like functionality. Optional. If not set the first row in the sheet is used.
* @return {Column} the row object
*/
function getRow(sheet, headerSize, rowIdentifier, headerColumnIdentifier, headerRow){
  var headerRow = headerRow || 1;
  var headerSize = headerSize || 0;
  var rowIdentifier = rowIdentifier || sheet.getLastRow() + 1;
  if (typeof headerColumnIdentifier === "string"){
    var headerColIdentifier = headerRowIdentifier || 1;
    var headerCol = getColumn(sheet, 0, headerRowIdentifier);
    var rowIdentifier = headerCol.first(columnIdentifier);
  }
  else{
  errorIfNegative(headerSize, 'getColumn(): headerSize');
  }

  return new Row(sheet, headerSize, rowIdentifier, headerRow);
}

/* Returns a cell by matching row and column headers
*
* @param {Sheet} sheet the sheet object containing the cell
* @param {String} rowHeader the value of the cell's row header
* @param {String} columnHeader the value of the cell's column header
* @param {Integer} rowHeader the row containing the headers. Defaults to 1.
* @param {Integer} columnNumber the column containing the headers. Defaults to 1.
*/
function getCellByHeaders(sheet, rowHeader, columnHeader, rowNumber, columnNumber){
  var rowNumber = rowNumber || 1;
  var columnNumber = columnNumber || 1;
  var cell = getRow(sheet, 0, rowHeader, rowNumber).getByHeader(columnHeader, columnNumber); 
  return cell;
}
