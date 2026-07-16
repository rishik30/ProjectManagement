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
	const machineNames = getMachines().map((m) => m.name);

	const operatorNames = getOperators().map((o) => o.name);

	const productNames = getAllProducts().map((p) => p.name);

	const dateRangeRule = SpreadsheetApp.newDataValidation()
		.requireValueInList(DASHBOARD.DATE_RANGES, true)
		.build();

	const machineRule = SpreadsheetApp.newDataValidation()
		.requireValueInList(['All', ...machineNames], true)
		.build();

	const operatorRule = SpreadsheetApp.newDataValidation()
		.requireValueInList(['All', ...operatorNames], true)
		.build();

	const productRule = SpreadsheetApp.newDataValidation()
		.requireValueInList(['All', ...productNames], true)
		.build();

	const validations = [
		['B3', dateRangeRule],
		['B6', machineRule],
		['B7', operatorRule],
		['B8', productRule],
	];

	validations.forEach(([cell, rule]) => {
		const range = sheet.getRange(cell);

		// Remove old validation first
		range.clearDataValidations();

		// Reapply
		range.setDataValidation(rule);

		// Preserve current value if still valid
		const value = range.getDisplayValue();

		if (!value) {
			range.setValue('All');
		}
	});
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
		const range = sheet.getRange(card.cell);

		const valueRange = sheet.getRange(
			range.getRow() + 1,
			range.getColumn(),
			2,
			2,
		);

		valueRange.breakApart();

		let value = kpi[card.key];

		if (value === undefined || value === null || value === '') {
			value = '-';
		}

		/*
		 * Highest & Lowest Production Cards
		 */
		if (
			card.key === 'highestProductionPieces' ||
			card.key === 'lowestProductionPieces'
		) {
			const dateKey =
				card.key === 'highestProductionPieces'
					? 'highestProductionDate'
					: 'lowestProductionDate';

			let displayText = '-';

			if (value !== '-') {
				if (typeof value === 'number') {
					value = value.toFixed(card.decimals);
				}

				let dateText = '-';

				if (kpi[dateKey]) {
					dateText = Utilities.formatDate(
						new Date(kpi[dateKey]),
						Session.getScriptTimeZone(),
						'dd/MM/yyyy',
					);
				}

				displayText = value + '\n' + dateText;
			}

			valueRange
				.merge()
				.setValue(displayText)
				.setWrap(true)
				.setHorizontalAlignment('center')
				.setVerticalAlignment('middle');

			return;
		}

		/*
		 * All Other KPI Cards
		 */
		if (typeof value === 'number') {
			value = value.toFixed(card.decimals);
		}

		valueRange
			.merge()
			.setValue(value)
			.setHorizontalAlignment('center')
			.setVerticalAlignment('middle');
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
		updateProductionHighlights(sheet, kpi);

		updateSummaryTable(
			sheet,
			summarizeByProduct(dataset),
			DASHBOARD.SUMMARY_TABLES.PRODUCT,
		);

		updateSummaryTable(
			sheet,
			summarizeByMachine(dataset),
			DASHBOARD.SUMMARY_TABLES.MACHINE,
		);

		updateSummaryTable(
			sheet,
			summarizeByOperator(dataset),
			DASHBOARD.SUMMARY_TABLES.OPERATOR,
		);

		updateSummaryTable(
			sheet,
			summarizeByDate(dataset),
			DASHBOARD.SUMMARY_TABLES.DAILY,
		);

		SpreadsheetApp.flush();
		setDashboardReady(sheet);
	} catch (error) {
		logError('refreshDashboardData', error, filters);

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
			return getFinancialYearObject(today);

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

/**
 * ===========================================================
 * Product Summary Table
 * ===========================================================
 */

function updateSummaryTable(sheet, summary, config) {
	const startRow = config.startRow;
	const startCol = config.startColumn;

	sheet.getRange(startRow, startCol, 1000, 2).clearContent();

	sheet
		.getRange(startRow, startCol, 1, 2)
		.merge()
		.setValue(config.title)
		.setFontWeight('bold')
		.setFontSize(14);

	sheet.getRange(startRow + 1, startCol).setValue(config.nameHeader);

	sheet.getRange(startRow + 1, startCol + 1).setValue(config.valueHeader);

	sheet.getRange(startRow + 1, startCol, 1, 2).setFontWeight('bold');

	if (summary.length === 0) {
		sheet.getRange(startRow + 2, startCol).setValue('No production found.');

		return;
	}

	if (config === DASHBOARD.SUMMARY_TABLES.DAILY) {
		summary.sort((a, b) => a.productionDateOnly - b.productionDateOnly);
	} else {
		summary.sort((a, b) => b[config.valueField] - a[config.valueField]);
	}

	const values = summary.map((item) => {
		let name = item[config.nameField];

		if (name instanceof Date) {
			name = Utilities.formatDate(
				name,
				Session.getScriptTimeZone(),
				'dd/MM/yyyy',
			);
		}

		return [name, item[config.valueField]];
	});

	sheet.getRange(startRow + 2, startCol, values.length, 2).setValues(values);

	sheet
		.getRange(startRow + 2, startCol + 1, values.length, 1)
		.setNumberFormat('#,##0');
}

function updateProductionHighlights(sheet, kpi) {
	const layout = DASHBOARD_LAYOUT.HIGHLIGHTS;

	clearProductionHighlights(sheet);

	sheet
		.getRange(layout.TITLE_ROW, 1, 1, 10)
		.merge()
		.setValue('Production Highlights')
		.setFontWeight('bold')
		.setFontSize(14)
		.setHorizontalAlignment('center');

	drawProductionHighlight(
		sheet,
		layout.LEFT,
		'Highest Production',
		kpi.highestProductionPieces,
		kpi.highestProductionDate,
	);

	drawProductionHighlight(
		sheet,
		layout.RIGHT,
		'Lowest Production',
		kpi.lowestProductionPieces,
		kpi.lowestProductionDate,
	);
}

function clearProductionHighlights(sheet) {
	const layout = DASHBOARD_LAYOUT.HIGHLIGHTS;
	sheet.getRange(layout.TITLE_ROW, 1, 5, 10).clearContent();
}

function drawProductionHighlight(sheet, layout, title, pieces, date) {
	sheet
		.getRange(layout.ROW, layout.COLUMN, 1, layout.WIDTH)
		.merge()
		.setBorder(true, true, true, true, true, true)
		.setValue(title)
		.setBackground('#E8F0FE')
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');

	const body = sheet.getRange(
		layout.ROW + 1,
		layout.COLUMN,
		layout.HEIGHT - 1,
		layout.WIDTH,
	);

	body
		.merge()
		.setBorder(true, true, true, true, true, true)
		.setFontWeight('bold')
		.setFontSize(14)
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle')
		.setWrap(true);

	if (!pieces || !date) {
		body.setValue('-');
		return;
	}

	body.setValue(
		Utilities.formatString(
			'%s Pieces\n\n%s',
			Number(pieces).toLocaleString(),
			Utilities.formatDate(
				new Date(date),
				Session.getScriptTimeZone(),
				'dd/MM/yyyy',
			),
		),
	);
}
