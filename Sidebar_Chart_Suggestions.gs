function onOpen() {
  SpreadsheetApp.getUi() 
      .createMenu('Custom Menu')
      .addItem('Create', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('Present dataset')
      .setWidth(300);
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}

function newChart(range_ip, chartNumber) {
  var sheet = SpreadsheetApp.getActiveSheet(); 
  var range = sheet.getRange(range_ip); 
  var chartBuilder = sheet.newChart();
  chartBuilder.addRange(range)
      .setPosition(5, 5, 0, 0)
     
  if(chartNumber==0){
    chartBuilder.setChartType(Charts.ChartType.LINE).setOption('title', 'Line Chart');
  }
  else if(chartNumber==1){
    chartBuilder.setChartType(Charts.ChartType.BAR).setOption('title', 'Bar Chart');
  }
  else if (chartNumber==2){
    chartBuilder.setChartType(Charts.ChartType.COLUMN).setOption('title', 'Column Chart');
  }
  else if (chartNumber==3){
    chartBuilder.setChartType(Charts.ChartType.PIE).setOption('title', 'Pie Chart');
  }
  sheet.insertChart(chartBuilder.build());
}

//merge isNumber and isString functions into isPrimitive() 
function isNumberRow(data, row_idx){ 
  Logger.log('number row'+row_idx); 
  var count = 0;
  var row_length = data[row_idx].length;
  for(var j=0; j<row_length; j++){ 
      if(typeof(data[row_idx][j])==='undefined'){
         return false;
      }
      Logger.log(typeof(data[row_idx][j]));
      if(typeof(data[row_idx][j])==='number'){
          count++;   
      }
      Logger.log(count);
      if(count>=10 || count==row_length){
          return true;
      }
  }  
  return false;
}

function isNumberCol(data, col_idx){
  Logger.log('no col '+col_idx);
  var count = 0;
  var col_length = data.length;
  for (var i = 0; i <col_length; i++){
    if(typeof(data[i])==='undefined' || typeof(data[i][col_idx])==='undefined'){
         return false;
    }
    if(typeof(data[i][col_idx])==='number'){
          count++;   
     }
    if(count>=10 || count==col_length){
      return true;
    }
  }
  return false;
}

function isStringRow(data, row_idx){
  Logger.log('string row '+row_idx);  
  var count = 0;
  var row_length = data[row_idx].length;
  for(var j=0; j<row_length; j++){ 
      if(typeof(data[row_idx][j])==='undefined'){
         return false;
      }
      if(typeof(data[row_idx][j])==='string'){
          count++;   
      }
      Logger.log(count);
      if(count>=10 || count==row_length){ 
          return true;
      }
  }  
  return false;
}

function isStringCol(data, col_idx){
  Logger.log('string col '+ col_idx);
  var count = 0;
  var col_length = data.length;
  for (var i = 0; i <col_length; i++){
    if(typeof(data[i])==='undefined' || typeof(data[i][col_idx])==='undefined'){
         return false;
    }
    if(typeof(data[i][col_idx])==='string'){
          count++;   
     }
    if(count>=10 || count==col_length){
      return true;
    }
  }
  return false;
}

function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date);
}

function isTimeDateRow(data){
   var count = 0;
   var row_length = data[0].length;
   for(var j=0; j<row_length; j++){ 
//      if(typeof(data[0][j])==='undefined'){
//         return false;
//      }
      if(isValidDate(data[0][j])){
          count++;   
      }
      if(count>=10 || count==row_length){
          return true;
      }
   }  
   return false;
}

function isTimeDateCol(data){
  var count = 0;
  var col_length = data.length;
  for (var i = 0; i <col_length; i++){
//    if(typeof(data[i])==='undefined'){
//         return false;
//    }
    if(isValidDate(data[i][0])){
          count++;   
    }
    if(count>=10 || count==col_length){
          return true;
    }
  }
  return false;
}

function isPieCol(data){
  var sum=0;
  var lastCidx = data.length-1; //last elt index in a column
  var no_of_cols = data[0].length;
  var no_of_rows = data.length;
  if( no_of_cols == 2 && (isStringCol(data,0) && isNumberCol(data,1) || isStringCol(data,0) && isNumberCol(data,1))){
    return 1;
  }
  for (var i = 0; i <data.length-1; i++){
    var ratio = (Number(data[i][0])/Number(data[lastCidx][0]));
    if(ratio>=0 && ratio<=1){
      sum += Number(data[i][0]);
    }
    else{
      return 0;
    }
  }
  if(sum==data[lastCidx][0]){
    return 2;
  }
  return 0;
}

function isPieRow(data){
  Logger.log('pie2');
  var sum=0;
  var lastRidx = data[0].length-1; //last elt index in a row
  var no_of_cols = data[0].length;
  var no_of_rows = data.length;
  if( no_of_rows == 2 && isStringRow(data,0) && isNumberRow(data,1)){
    return true;
  }
  for (var i = 0; i <lastRidx; i++){
    var ratio = (Number(data[0][i])/Number(data[0][lastRidx]));
    if(ratio>=0 && ratio<=1){
      sum += Number(data[0][i]);
    }
    else{
      return false;
    }
  }
  if(sum==data[0][lastRidx]){
    return true;
  }
  return false;
}

function GiveSuggestion(range){
  Logger.log('suggestion fn');
  Logger.log(range);
   var sheet = SpreadsheetApp.getActiveSheet();
  var searchRange = sheet.getRange(range); 
  var data = searchRange.getValues();
  var suggest = [];
  
  //traverse first row of range
  var sr = isStringRow(data,0);
  if(sr){
    suggest.push("Add first row as domain/on major axis\n");
  }
  //traverse first col of range 
  var sc = isStringCol(data,0);
  Logger.log(sc);
  if(sc){
    suggest.push("Add first column as domain/on major axis\n");
  }
  
  // if Column/Row has Date/Time suggest Timeline Chart
  var tr = isTimeDateRow(data)
  Logger.log(tr);
  if(tr){
    Logger.log('Try Timeline Graph with first row as major axis');
    suggest.push("Try Timeline Graph with first row as major axis");
  }
  var tc = isTimeDateCol(data)
  Logger.log(tc);
  if(tc){
    Logger.log('Try Timeline Graph with first column as major axis');
    suggest.push("Try Timeline Graph with first column as major axis");
  }
  
  var pieC = isPieCol(data);
  Logger.log(pieC);
  //if data[0].length==2, if for all i, typeof(data[i][0])= string and typeof(data[i][1])== number or reverse values then return true
  if(pieC==1){
    Logger.log('Try PieChart');
    suggest.push("Try PieChart");
  }
  // if Sum(firstCol, lastColumn-2)==value at lastColumn-1 then return true
  else if(pieC==2){
    Logger.log('Try PieChart with first column as labels');
    suggest.push("Try PieChart with first column as labels");
  }
  // if Sum(firstRow, lastRow-2)==value at lastRow-1 then return true
  var pieR = isPieRow(data);
  if(pieR){
    Logger.log('Try PieChart with first row as labels');
    suggest.push("Try PieChart with first row as labels");
  }
  return suggest;
}
