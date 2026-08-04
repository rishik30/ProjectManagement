/**
 * ===========================================================
 * Menu.gs
 * ===========================================================
 */

function onOpen() {
	const ui = SpreadsheetApp.getUi();

	ui.createMenu('Production System')

		.addItem('Daily Production', 'openProductionSidebar')

		.addItem('Machine Setup', 'openMachineSetupSidebar')

		.addItem('Sales', 'showSalesSidebar')
		.addItem('Painting & Packing', 'showInventoryProcessingSidebar')

		.addSeparator()

		.addItem('Customer Master', 'showCustomerSidebar')

		.addItem('Stock Adjustments', 'showStockAdjustmentsSidebar')

		.addSeparator()

		.addItem('Dashboard', 'openDashboard')

		.addSubMenu(
			ui
				.createMenu('Reports')
				.addItem('Product Report', 'generateProductReport')
				.addItem('Machine Report', 'generateMachineReport')
				.addItem('Daily Report', 'generateDailyReport'),
		)

		.addSeparator()

		.addItem('Export Current Report (PDF)', 'exportCurrentReportPdf')

		.addItem('Validate System', 'validateSystem')

		.addSeparator()

		.addItem('About', 'aboutSystem')

		.addToUi();

	// Initialize Inventory Engine
	initializeInventory();
	initializeStagedInventory();
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

function showInventoryProcessingSidebar() {
	const html = HtmlService.createTemplateFromFile('InventoryProcessingSidebar')
		.evaluate()
		.setTitle('Painting & Packing');

	SpreadsheetApp.getUi().showSidebar(html);
}
