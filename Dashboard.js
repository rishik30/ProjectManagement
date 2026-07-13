/**
 * ===========================================================
 * Dashboard.gs
 * Module 4A - Part 1
 * ===========================================================
 */

function openDashboard() {
	refreshDashboard();
}

/**
 * ===========================================================
 * Dashboard Refresh Engine
 * ===========================================================
 */

function refreshDashboard() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	let sheet = ss.getSheetByName(SHEETS.DASHBOARD);

	if (!sheet) {
		sheet = ss.insertSheet(SHEETS.DASHBOARD);
	}

	const rebuilt = ensureDashboardInfrastructure(sheet);

	applyDashboardValidation(sheet);

	if (!rebuilt) {
		const uiFilters = getDashboardFilterValues(sheet);

		const filters = normalizeDashboardFilters(uiFilters);

		refreshDashboardData(sheet, filters);
	}

	ss.setActiveSheet(sheet);
}

function buildDashboardLayout(sheet) {
	sheet.clear();

	sheet.setHiddenGridlines(true);

	buildDashboardTitle(sheet);

	buildFilterSection(sheet);

	applyDashboardValidation(sheet);

	buildKpiSection(sheet);
}

function buildDashboardTitle(sheet) {
	sheet
		.getRange('A1:H1')
		.merge()
		.setValue('FACTORY PRODUCTION DASHBOARD')
		.setFontSize(18)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');

	sheet.setRowHeight(1, 40);

	sheet
		.getRange('A2:F2')
		.merge()
		.setValue('Ready')
		.setFontStyle('italic')
		.setHorizontalAlignment('left');
}

function buildFilterSection(sheet) {
	const filters = getDashboardFilters();

	/*
	 * Labels
	 */

	const labels = [
		['Date Range'],
		['From Date'],
		['To Date'],
		['Machine'],
		['Operator'],
		['Product'],
	];

	sheet
		.getRange(3, 1, labels.length, 1)
		.setValues(labels)
		.setFontWeight('bold');

	/*
	 * Default Values
	 */

	const today = new Date();

	const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);

	sheet.getRange('B3').setValue('Current Month');

	sheet.getRange('B4').setValue(firstDay);

	sheet.getRange('B5').setValue(today);

	sheet.getRange('B4:B5').setNumberFormat('dd/MM/yyyy');

	sheet.getRange('B6').setValue('All');

	sheet.getRange('B7').setValue('All');

	sheet.getRange('B8').setValue('All');

	/*
	 * Hidden Lists
	 */

	const startColumn = 20; // T

	sheet.getRange(1, startColumn).setValue('DateRanges');

	sheet
		.getRange(2, startColumn, 7, 1)
		.setValues([
			['Today'],
			['This Week'],
			['This Month'],
			['Last Month'],
			['Last 3 Months'],
			['Financial Year'],
			['Custom'],
		]);

	sheet.getRange(1, startColumn + 1).setValue('Machines');

	const machines = [['All']].concat(filters.machines.map((m) => [m.name]));

	sheet.getRange(2, startColumn + 1, machines.length, 1).setValues(machines);

	sheet.getRange(1, startColumn + 2).setValue('Operators');

	const operators = [['All']].concat(filters.operators.map((o) => [o.name]));

	sheet.getRange(2, startColumn + 2, operators.length, 1).setValues(operators);

	sheet.getRange(1, startColumn + 3).setValue('Products');

	const products = [['All']].concat(filters.products.map((p) => [p.name]));

	sheet.getRange(2, startColumn + 3, products.length, 1).setValues(products);
}

function buildDashboardPlaceholder(sheet) {
	sheet
		.getRange('A11:H11')
		.merge()
		.setValue('Dashboard will be generated here')
		.setHorizontalAlignment('center')
		.setFontWeight('bold');
}

/**
 * ===========================================================
 * Dashboard Filter Data
 * ===========================================================
 */

function getDashboardFilters() {
	return {
		machines: getMachines(),
		operators: getOperators(),
		products: getAllProducts(),
	};
}

function applyDashboardValidation(sheet) {
	const machineSheet = getSheet(SHEETS.MACHINE_MASTER);
	const userSheet = getSheet(SHEETS.USER_MASTER);
	const productSheet = getSheet(SHEETS.PRODUCT_MASTER);

	/*
	 * Date Range
	 */

	const dateRangeRule = SpreadsheetApp.newDataValidation()
		.requireValueInList(DASHBOARD.DATE_RANGES, true)
		.build();

	/*
	 * Machine
	 */

	const machineRule = SpreadsheetApp.newDataValidation()
		.requireValueInRange(
			machineSheet.getRange(
				2,
				MACHINE_MASTER_COLUMNS.NAME,
				Math.max(machineSheet.getLastRow() - 1, 1),
				1,
			),
			true,
		)
		.build();

	/*
	 * Operator
	 */

	const operatorRule = SpreadsheetApp.newDataValidation()
		.requireValueInRange(
			userSheet.getRange(
				2,
				USER_MASTER_COLUMNS.NAME,
				Math.max(userSheet.getLastRow() - 1, 1),
				1,
			),
			true,
		)
		.build();

	/*
	 * Product
	 */

	const productRule = SpreadsheetApp.newDataValidation()
		.requireValueInRange(
			productSheet.getRange(
				2,
				PRODUCT_MASTER_COLUMNS.NAME,
				Math.max(productSheet.getLastRow() - 1, 1),
				1,
			),
			true,
		)
		.build();

	/*
	 * Apply validations
	 */

	sheet.getRange('B3').setDataValidation(dateRangeRule);

	sheet.getRange('B6').setDataValidation(machineRule);

	sheet.getRange('B7').setDataValidation(operatorRule);

	sheet.getRange('B8').setDataValidation(productRule);
}

/**
 * ===========================================================
 * KPI Layout
 * ===========================================================
 */

function buildKpiSection(sheet) {
	sheet
		.getRange('A11:F11')
		.merge()
		.setValue('Factory KPIs')
		.setFontWeight('bold')
		.setFontSize(14);

	DASHBOARD.KPI_CARDS.forEach((card) => {
		buildKpiCard(sheet, card.title, card.cell);
	});
}

function buildKpiCard(sheet, title, startCell) {
	const range = sheet.getRange(startCell);

	const row = range.getRow();

	const col = range.getColumn();

	sheet
		.getRange(row, col, 1, 2)
		.merge()
		.setValue(title)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBackground('#E8F0FE');

	sheet
		.getRange(row + 1, col, 2, 2)
		.merge()
		.setValue('-')
		.setFontSize(18)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');

	formatKpiCard(sheet, row, col);
}

function formatKpiCard(sheet, row, col) {
	sheet.getRange(row, col, 3, 2).setBorder(true, true, true, true, true, true);

	sheet.setRowHeight(row, 24);

	sheet.setRowHeight(row + 1, 28);

	sheet.setRowHeight(row + 2, 28);
}

function updateKpiSection(sheet, kpi) {
	DASHBOARD.KPI_CARDS.forEach((card) => {
		let value = kpi[card.key];

		if (value === undefined) {
			value = '-';
		}

		if (typeof value === 'number') {
			value = value.toFixed(card.decimals);
		}

		const range = sheet.getRange(card.cell);

		sheet
			.getRange(range.getRow() + 1, range.getColumn(), 2, 2)
			.merge()
			.setValue(value);
	});
}

function testDashboardLayout() {
	const sheet = getSheet(SHEETS.DASHBOARD);
	ensureDashboardInfrastructure(sheet);
}

function testDashboardData() {
	const sheet = getSheet(SHEETS.DASHBOARD);
	refreshDashboardData(sheet);
}

function ensureDashboardLayout(sheet) {
	const title = sheet.getRange('A1').getValue();

	const hasFilters = sheet.getRange('A3').getValue() === 'Date Range';

	if (title === 'FACTORY PRODUCTION DASHBOARD' && hasFilters) {
		return false;
	}

	buildDashboardLayout(sheet);

	return true;
}

function getDashboardFilterValues(sheet) {
	return {
		dateRange: sheet.getRange('B3').getValue(),
		fromDate: sheet.getRange('B4').getValue(),
		toDate: sheet.getRange('B5').getValue(),
		machineId: sheet.getRange('B6').getValue(),
		operatorId: sheet.getRange('B7').getValue(),
		productId: sheet.getRange('B8').getValue(),
	};
}

function handleDashboardEdit(e) {
	const sheet = e.range.getSheet();

	if (sheet.getName() !== SHEETS.DASHBOARD) {
		return;
	}

	const a1 = e.range.getA1Notation();

	const dashboardFilterCells = ['B3', 'B4', 'B5', 'B6', 'B7', 'B8'];

	if (!dashboardFilterCells.includes(a1)) {
		return;
	}

	/*
	 * Date Range Preset Changed
	 */

	if (a1 === 'B3') {
		applyDateRangeSelection(sheet);
	}

	/*
	 * Read final values after any updates
	 */

	const uiFilters = getDashboardFilterValues(sheet);

	const filters = normalizeDashboardFilters(uiFilters);

	refreshDashboardData(sheet, filters);
}

function ensureDashboardInfrastructure(sheet) {
	return ensureDashboardLayout(sheet);
}

function initializeDashboardFilters(sheet) {
	applyDateRangeSelection(sheet);
	const uiFilters = getDashboardFilterValues(sheet);
	const filters = normalizeDashboardFilters(uiFilters);
	refreshDashboardData(sheet, filters);
}

/**
 * ===========================================================
 * Dashboard Data Refresh
 * ===========================================================
 */

function refreshDashboardData(sheet, filters) {
	try {
		setDashboardLoading(sheet);
		if (!filters) {
			const uiFilters = getDashboardFilterValues(sheet);
			filters = normalizeDashboardFilters(uiFilters);
		}

		const dataset = getProductionDataset(filters);
		const kpi = calculateKPIs(dataset);
		updateKpiSection(sheet, kpi);
		SpreadsheetApp.flush();
		setDashboardReady(sheet);
	} catch (error) {
		setDashboardError(sheet, error);
		throw error;
	}
}

function applyDateRangeSelection(sheet) {
	const option = sheet.getRange('B3').getValue();

	if (option === 'Custom') {
		return;
	}

	const range = resolveDateRange(option);

	sheet.getRange('B4').setValue(range.fromDate);
	sheet.getRange('B5').setValue(range.toDate);
}

function resolveDateRange(option) {
	const today = new Date();

	today.setHours(0, 0, 0, 0);

	let fromDate;
	let toDate;

	switch (option) {
		case 'Today':
			fromDate = new Date(today);
			toDate = new Date(today);
			break;

		case 'This Week':
			fromDate = getStartOfWeek(today);
			toDate = new Date(today);
			break;

		case 'This Month':
			fromDate = new Date(today.getFullYear(), today.getMonth(), 1);

			toDate = new Date(today);
			break;

		case 'Last Month':
			fromDate = new Date(today.getFullYear(), today.getMonth() - 1, 1);

			toDate = new Date(today.getFullYear(), today.getMonth(), 0);

			break;

		case 'Last 3 Months':
			fromDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);

			toDate = new Date(today);
			break;

		case 'Financial Year':
			return getFinancialYear(today);

		default:
			fromDate = new Date(today.getFullYear(), today.getMonth(), 1);

			toDate = new Date(today);
	}

	fromDate.setHours(0, 0, 0, 0);
	toDate.setHours(23, 59, 59, 999);

	return {
		fromDate,
		toDate,
	};
}

function setDashboardStatus(sheet, message) {
	sheet.getRange('A2:F2').merge().setValue(message);
}

function setDashboardReady(sheet) {
	const now = Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		'dd/MM/yyyy hh:mm:ss a',
	);

	setDashboardStatus(sheet, '🟢 Ready | Last Updated: ' + now);
}

function setDashboardLoading(sheet) {
	setDashboardStatus(sheet, '🟡 Refreshing Dashboard...');
}

function setDashboardError(sheet, error) {
	setDashboardStatus(sheet, '🔴 ' + error.message);
}
