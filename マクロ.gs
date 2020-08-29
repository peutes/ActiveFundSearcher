function myFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('G:G').activate();
  spreadsheet.getActiveSheet().sort(7, false);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.getActiveSheet().sort(2, true);
};

function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('D2:D76').activate();
  var sheet = spreadsheet.getActiveSheet();
  var chart = sheet.newChart()
  .asLineChart()
  .addRange(spreadsheet.getRange('D2:D76'))
  .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
  .setTransposeRowsAndColumns(false)
  .setNumHeaders(0)
  .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH)
  .setOption('applyAggregateData', 0)
  .setOption('useFirstColumnAsDomain', true)
  .setOption('isStacked', 'false')
  .setOption('title', '「カテゴリ」のカウント数')
  .setXAxisTitle('「カテゴリ」のカウント数')
  .setOption('series.0.aggregateFunction', 'count')
  .setPosition(17, 5, 196, 12)
  .build();
  sheet.insertChart(chart);
  var charts = sheet.getCharts();
  chart = charts[charts.length - 1];
  sheet.removeChart(chart);
  chart = sheet.newChart()
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
  .setOption('title', '「カテゴリ」のカウント数')
  .setOption('annotations.domain.textStyle.color', '#808080')
  .setOption('textStyle.color', '#000000')
  .setOption('legend.textStyle.color', '#191919')
  .setOption('titleTextStyle.color', '#757575')
  .setOption('annotations.total.textStyle.color', '#808080')
  .setPosition(17, 5, 196, 12)
  .build();
  sheet.insertChart(chart);
};

function myFunction2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('N:N').activate();
  spreadsheet.getActiveSheet().sort(14, false);
  spreadsheet.getRange('B:B').activate();
  spreadsheet.getActiveSheet().sort(2, true);
};