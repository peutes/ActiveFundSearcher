function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:G').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('G1'));
  spreadsheet.getActiveSheet().sort(7, false);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getActiveSheet().sort(2, true);
  spreadsheet.getRange('D2:D76').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('D2:D76'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('applyAggregateData', 0)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', '「分類」')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(10, 10, 302, 12)
  .build();
  sheet.insertChart(chart);
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N:N').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('N1'));
  spreadsheet.getActiveSheet().sort(14, false);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('B1'));
  spreadsheet.getActiveSheet().sort(2, true);
  spreadsheet.getRange('D2:D76').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asPieChart()
  .addRange(spreadsheet.getRange('D2:D76'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('applyAggregateData', 0)
  .setOption('bubble.stroke', '#000000')
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', '「分類」')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(10, 10, 302, 12)
  .build();
  sheet.insertChart(chart);
};
