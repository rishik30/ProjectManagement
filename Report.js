/**
 * ===========================================================
 * Report Engine
 * Phase 2
 * Sprint 2.1
 * ===========================================================
 */

/**
 * Returns the report sheet.
 */
function getReportSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	let sheet = ss.getSheetByName(SHEETS.REPORT);

	if (!sheet) {
		sheet = ss.insertSheet(SHEETS.REPORT);
	}

	return sheet;
}

/**
 * Reads current dashboard filters.
 */
function getCurrentDashboardFilters() {
	const dashboard = getSheet(SHEETS.DASHBOARD);

	if (!dashboard) {
		throw new Error('Dashboard sheet not found.');
	}

	const uiFilters = getDashboardFilterValues(dashboard);

	return normalizeDashboardFilters(uiFilters);
}

/**
 * Clears report sheet.
 */
function clearReportSheet(sheet) {
	sheet.clearContents();
	sheet.clearFormats();
	sheet.clearConditionalFormatRules();
	sheet.setFrozenRows(0);
	sheet.setFrozenColumns(0);
}

/**
 * Writes report title.
 */
function writeReportTitle(sheet, title) {
	sheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(16);

	sheet
		.getRange(2, 1)
		.setValue(
			'Generated : ' +
				Utilities.formatDate(
					new Date(),
					Session.getScriptTimeZone(),
					'dd/MM/yyyy hh:mm:ss a',
				),
		);

	sheet.setFrozenRows(4);

	sheet.getRange('A1').setVerticalAlignment('middle');

	sheet.getRange('A2').setFontStyle('italic');
}

/**
 * Writes report table.
 */
function writeReportTable(sheet, headers, values) {
	sheet
		.getRange(4, 1, 1, headers.length)
		.setValues([headers])
		.setFontWeight('bold');

	if (values.length > 0) {
		sheet.getRange(5, 1, values.length, headers.length).setValues(values);
	}

	sheet.autoResizeColumns(1, headers.length);

	return headers.length;
}

/**
 * Opens report sheet.
 */
function openReportSheet(sheet) {
	SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}

/**
 * ===========================================================
 * Generic Report Generator
 * ===========================================================
 */

function generateSummaryReport(config) {
	const filters = getCurrentDashboardFilters();

	const dataset = getProductionDataset(filters);

	const summary = config.summaryFunction(dataset);

	if (config.sortFunction) {
		summary.sort(config.sortFunction);
	} else {
		summary.sort((a, b) => b[config.valueField] - a[config.valueField]);
	}

	const values = summary.map((item) =>
		config.columns.map((column) => {
			let value = item[column.field];

			if (config.valueTransformers && config.valueTransformers[column.field]) {
				value = config.valueTransformers[column.field](value, item);
			}

			return value;
		}),
	);

	const sheet = getReportSheet();

	clearReportSheet(sheet);

	writeReportTitle(sheet, config.title);

	writeReportFilters(sheet, filters);

	const totalColumns = writeReportTable(
		sheet,
		config.columns.map((column) => column.header),
		values,
	);

	formatSummaryReport(
		sheet,
		values.length,
		totalColumns,
		config.numberColumns || [],
	);

	makeReportPrintable(sheet);
	openReportSheet(sheet);
}

/**
 * ===========================================================
 * Generic Report Formatting
 * ===========================================================
 */

function formatSummaryReport(sheet, rowCount, totalColumns, numberColumns) {
	if (rowCount === 0) {
		return;
	}

	numberColumns.forEach((column) => {
		sheet.getRange(5, column, rowCount, 1).setNumberFormat('#,##0');
	});

	sheet
		.getRange(4, 1, rowCount + 1, totalColumns)
		.setBorder(true, true, true, true, true, true);

	sheet
		.getRange(4, 1, 1, totalColumns)
		.setBackground('#E8F0FE')
		.setFontWeight('bold')
		.setHorizontalAlignment('center');

	sheet
		.getRange(5, 1, rowCount, totalColumns)
		.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

/**
 * ===========================================================
 * Product Report
 * Phase 2
 * Sprint 2.2
 * ===========================================================
 */

function generateProductReport() {
	generateSummaryReport({
		title: 'PRODUCT PRODUCTION REPORT',
		summaryFunction: summarizeByProduct,
		valueField: 'pieces',
		columns: [
			{
				header: 'Product ID',
				field: 'productId',
			},
			{
				header: 'Product Name',
				field: 'productName',
			},
			{
				header: 'Pieces',
				field: 'pieces',
			},
		],

		numberColumns: [3],
	});
}

/**
 * ===========================================================
 * Report Helpers
 * ===========================================================
 */

function writeReportFilters(sheet, filters) {
	const formatDate = (date) =>
		Utilities.formatDate(
			new Date(date),
			Session.getScriptTimeZone(),
			'dd/MM/yyyy',
		);

	const rows = [
		['From Date', formatDate(filters.fromDate)],
		['To Date', formatDate(filters.toDate)],
		['Machine', filters.machineId],
		['Operator', filters.operatorId],
		['Product', filters.productId],
	];

	sheet.getRange(1, 10, rows.length, 2).setValues(rows);

	sheet.getRange(1, 10, rows.length, 1).setFontWeight('bold');

	sheet.autoResizeColumns(10, 2);
}

/**
 * ===========================================================
 * Machine Report
 * Phase 2
 * Sprint 2.3
 * ===========================================================
 */

function generateMachineReport() {
	generateSummaryReport({
		title: 'MACHINE PRODUCTION REPORT',
		summaryFunction: summarizeByMachine,
		valueField: 'pieces',
		columns: [
			{
				header: 'Machine ID',
				field: 'machineId',
			},
			{
				header: 'Machine Name',
				field: 'machineName',
			},
			{
				header: 'Pieces',
				field: 'pieces',
			},
		],
		numberColumns: [3],
	});
}

function generateDailyReport() {
	generateSummaryReport({
		title: 'DAILY PRODUCTION REPORT',
		summaryFunction: summarizeByDate,
		valueField: 'pieces',
		sortFunction: (a, b) => a.productionDateOnly - b.productionDateOnly,
		columns: [
			{
				header: 'Date',
				field: 'productionDateOnly',
			},
			{
				header: 'Pieces',
				field: 'pieces',
			},
		],
		numberColumns: [2],
	});
}

/**
 * ===========================================================
 * Printable Report Formatting
 * Phase 2
 * Sprint 2.6
 * ===========================================================
 */

function makeReportPrintable(sheet) {
	const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

	spreadsheet.setActiveSheet(sheet);

	const lastRow = sheet.getLastRow();
	const lastColumn = sheet.getLastColumn();

	// Remove any existing filter
	if (sheet.getFilter()) {
		sheet.getFilter().remove();
	}

	// Freeze report header
	sheet.setFrozenRows(4);

	// Ensure all rows/columns are visible
	sheet.showRows(1, lastRow);
	sheet.showColumns(1, lastColumn);

	// Apply basic print settings
	const spreadsheetId = spreadsheet.getId();
	const sheetId = sheet.getSheetId();

	const printUrl =
		'https://docs.google.com/spreadsheets/d/' +
		spreadsheetId +
		'/export?' +
		[
			'format=pdf',
			'gid=' + sheetId,
			'portrait=false',
			'fitw=true',
			'sheetnames=false',
			'printtitle=false',
			'pagenumbers=true',
			'gridlines=false',
			'fzr=true',
			'top_margin=0.50',
			'bottom_margin=0.50',
			'left_margin=0.50',
			'right_margin=0.50',
		].join('&');

	// Store for future PDF export
	PropertiesService.getDocumentProperties().setProperty(
		'LAST_REPORT_EXPORT_URL',
		printUrl,
	);
}

/**
 * ===========================================================
 * PDF Export
 * Phase 2
 * Sprint 2.7
 * ===========================================================
 */

const REPORT_FOLDER_NAME = 'Reports';

function exportCurrentReportPdf() {
	const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = getReportSheet();

	const spreadsheetId = spreadsheet.getId();
	const sheetId = sheet.getSheetId();

	const title = sheet.getRange('A1').getDisplayValue().trim();

	const timestamp = Utilities.formatDate(
		new Date(),
		Session.getScriptTimeZone(),
		'yyyyMMdd_HHmmss',
	);

	const fileName = `${title}_${timestamp}.pdf`;

	const url =
		'https://docs.google.com/spreadsheets/d/' +
		spreadsheetId +
		'/export?' +
		[
			'format=pdf',
			'gid=' + sheetId,
			'portrait=false',
			'fitw=true',
			'sheetnames=false',
			'printtitle=false',
			'pagenumbers=true',
			'gridlines=false',
			'fzr=true',
			'top_margin=0.50',
			'bottom_margin=0.50',
			'left_margin=0.50',
			'right_margin=0.50',
		].join('&');

	const token = ScriptApp.getOAuthToken();

	const response = UrlFetchApp.fetch(url, {
		headers: {
			Authorization: 'Bearer ' + token,
		},
	});

	const folder = getOrCreateReportFolder();

	const pdf = response.getBlob().setName(fileName);

	const file = folder.createFile(pdf);

	SpreadsheetApp.getUi().alert(
		'Report exported successfully.\n\n' +
			file.getName() +
			'\n\n' +
			file.getUrl(),
	);
}
/**
 * Returns the Reports folder.
 */
function getOrCreateReportFolder() {
	const folders = DriveApp.getFoldersByName(REPORT_FOLDER_NAME);

	if (folders.hasNext()) {
		return folders.next();
	}

	return DriveApp.createFolder(REPORT_FOLDER_NAME);
}
