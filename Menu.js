/**
 * ===========================================================
 * Menu.gs
 * ===========================================================
 */

function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu('Production System')

		.addItem('Daily Production', 'openProductionSidebar')

		.addItem('Machine Setup', 'openMachineSetupSidebar')

		.addSeparator()

		.addItem('Product Report', 'generateProductReport')

		.addItem('Machine Report', 'generateMachineReport')

		.addItem('Daily Report', 'generateDailyReport')

		.addItem('Dashboard', 'openDashboard')

		.addSeparator()

		.addItem('Export Current Report (PDF)', 'exportCurrentReportPdf')

		.addItem('Validate System', 'validateSystem')

		.addSeparator()

		.addItem('About', 'aboutSystem')

		.addToUi();
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
