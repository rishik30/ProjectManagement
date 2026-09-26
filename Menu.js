/**
 * ===========================================================
 * Menu.gs
 * ===========================================================
 */

function onOpen() {
	const ui = SpreadsheetApp.getUi();

	const menu = ui
		.createMenu('Production System')
		.addItem('Daily Production', 'openProductionSidebar')
		.addItem('Machine Setup', 'openMachineSetupSidebar')
		.addItem('Sales', 'showSalesSidebar')
		.addItem('Order Completion Calculator', 'showOrderCompletionSidebar')
		.addSeparator()
		.addItem('Customer Master', 'showCustomerSidebar')
		.addItem('Stock Adjustments', 'showStockAdjustmentsSidebar')
		.addSeparator()
		.addItem('Dashboard', 'openDashboard')
		.addItem('Sales Dashboard', 'openSalesDashboard');

	// This menu item is deliberately omitted for unauthorised users.
	if (isAmountAnalysisAuthorized()) {
		menu.addItem('Amount Analysis', 'openAmountAnalysisSidebar');
	}

	menu
		.addSubMenu(
			ui
				.createMenu('Reports')
				.addItem('Product Report', 'generateProductReport')
				.addItem('Machine Report', 'generateMachineReport')
				.addItem('Daily Report', 'generateDailyReport')
				.addSubMenu(
					ui
						.createMenu('Sales Reports')
						.addItem('Sales Summary', 'generateSalesSummaryReport')
						.addItem('Sales Order Details', 'generateSalesOrderDetailReport'),
				),
		)
		.addSeparator()
		.addItem('Export Current Report (PDF)', 'exportCurrentReportPdf')
		.addItem('Validate System', 'validateSystem')
		.addSeparator()
		.addItem('About', 'aboutSystem')
		.addToUi();

	// Initialize Inventory Engine
	initializeInventory();
}

function openProductionSidebar() {
	openSidebar('production');
}

function openMachineSetupSidebar() {
	openSidebar('machine');
}

function openSidebar(activeTab) {
	const template = HtmlService.createTemplateFromFile('Sidebar');

	template.activeTab = activeTab;

	const html = template.evaluate().setTitle(SIDEBAR.TITLE);

	SpreadsheetApp.getUi().showSidebar(html);
}

function generateReport() {
	SpreadsheetApp.getUi().alert('Module 5');
}

function aboutSystem() {
	SpreadsheetApp.getUi().alert(APP_NAME + '\n\nVersion : 1.0');
}

/**
 * Opens the restricted Amount Analysis sidebar.
 */
function openAmountAnalysisSidebar() {
	assertAmountAnalysisAuthorized();

	const html = HtmlService.createTemplateFromFile('AmountAnalysisSidebar')
		.evaluate()
		.setTitle('Amount Analysis');

	SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens Customer Master Sidebar
 */
function showCustomerSidebar() {
	const html = HtmlService.createTemplateFromFile('CustomerSidebar')
		.evaluate()
		.setTitle('Customer Master');

	SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Opens Sales Sidebar
 */
function showSalesSidebar() {
	const html = HtmlService.createTemplateFromFile('SalesSidebar')
		.evaluate()
		.setTitle('Sales');

	SpreadsheetApp.getUi().showSidebar(html);
}

function showStockAdjustmentsSidebar() {
	const html = HtmlService.createTemplateFromFile('StockAdjustmentSidebar')
		.evaluate()
		.setTitle('Stock Adjustments');

	SpreadsheetApp.getUi().showSidebar(html);
}

function showOrderCompletionSidebar() {
	const html = HtmlService.createTemplateFromFile('OrderCompletionSidebar')
		.evaluate()
		.setTitle('Order Completion');

	SpreadsheetApp.getUi().showSidebar(html);
}
