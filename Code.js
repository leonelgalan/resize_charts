/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var SIDEBAR_TITLE = 'Resize Charts';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  console.info('onOpen', 'e.authMode', e && e.authMode)
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Show Sidebar', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  console.info('onInstall', 'e.authMode', e && e.authMode)
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Resizes all charts in current sheet
 *
 * @param {Number} width
 * @param {Number} height
 */
function resizeAllCharts(width, height) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var charts = sheet.getCharts();
  for (var i in charts) {
    var chart = charts[i];
    var newChart = chart.modify()
      .setOption('width', Number(width))
      .setOption('height', Number(height))
      .build();
    sheet.updateChart(newChart);
  }
}
