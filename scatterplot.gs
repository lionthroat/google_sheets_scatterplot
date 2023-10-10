function createScatterplotChart() {
  // Get the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // The name of the sheet where you want to create the new chart
  var sheetName = 'Practice Test Scores'; // Use the correct sheet name

  // Get the sheet with the specified name
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }

  // Get the named ranges
  var testDatesRange = spreadsheet.getRangeByName('TestDates');
  var testScoresRange = spreadsheet.getRangeByName('TestScores');

  if (!testDatesRange || !testScoresRange) {
    Logger.log('Named ranges not found.');
    return;
  }

  // Get the values from the named ranges
  var testDates = testDatesRange.getValues();
  var testScores = testScoresRange.getValues();

  // Filter out rows with empty values in both date and score columns
  var filteredData = testDates.reduce(function(result, date, index) {
    if (date[0] !== "" && testScores[index][0] !== "") {
      result.dates.push(date[0]);
      result.scores.push(testScores[index][0]);
    }
    return result;
  }, { dates: [], scores: [] });

  // Extract filtered dates and scores
  var filteredDates = filteredData.dates;
  var filteredScores = filteredData.scores;

  // Create scatterplot chart
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var chart = sheet.newChart()
    .setOption('chartArea.backgroundColor', {
      fill: 'white',   // does work as intended
      strokeWidth: -1, // doesn't work
      stroke: 'white'  // doesn't work
    })
    .asScatterChart()
    .addRange(sheet.getRange(7, 2, filteredDates.length, 2)) // Assumes date in column B, score in column C
    .setOption('title', 'Ham Radio Technician License - Practice Tests, Fall 2023')
    .setOption('titleTextStyle', {
      bold: true,
      color: '#a2e0a5',
      fontSize: 14,
      fontName: 'Open Sans',
      alignment: 'center'
    })
    .setOption('width', 700)
    .setOption('height', 400)
    .setPosition(1,6,0,0)
    .setTransposeRowsAndColumns(false)
    .setOption('legend', 'none')
    .setOption('trendlines', [{ 
      type: 'polynomial',
      degree: 3,
      color: '#e7f780' // Color for the trend line
    }])
    // X-axis settings (Dates)
    .setOption('hAxis.ticks', filteredDates)
    .setOption('hAxis.format', 'MMM d')
    .setOption('hAxis.slantedTextAngle', 45)
    .setOption('hAxis.gridlines', { color: '#eee', count: 10 }) // Exactly 10 days' worth of scores
    .setOption('hAxis.textStyle', {
      fontSize: 10,
      fontName: 'Open Sans',
      color: '#363636' // Dark-ish grey
    })
    // Y-axis settings (Scores)
    .setOption('vAxis.scaleType', 'log')
    .setOption('vAxis.viewWindow.min', 0.35)
    .setOption('vAxis.viewWindow.max', 1.0)
    .setOption('vAxis.gridlines', { color: '#eee', count: 14 }) // Put ticks every 5% on y-axis, from 35% to 100%
    .setOption('vAxis.textStyle', {
      fontSize: 11,
      fontName: 'Open Sans',
      color: '#929292' // Medium-light grey
    })
    .build();
  // Define separate arrays to store colors
  var colors = ['#b4aa5e','#dde0a2', '#a2e0a5', '#76c34a'];

  // Loop through data points and assign colors based on index
  for (var i = 0; i < filteredScores.length; i++) {
    var score = filteredScores[i];
    if (score < .50) {
      chart = chart.modify().setOption('series.0.items.' + i + '.color', colors[0]).build();
    } else if (score > .50 && score < .74) {
      chart = chart.modify().setOption('series.0.items.' + i + '.color', colors[1]).build();
    } else if (score < .91 && score >= .74) {
      chart = chart.modify().setOption('series.0.items.' + i + '.color', colors[2]).build();
    } else {
      chart = chart.modify().setOption('series.0.items.' + i + '.color', colors[3]).build();
    }
  }
  sheet.insertChart(chart);

  Logger.log('Scatterplot chart created.');
}