/**
 * ===========================================================
 * Sales Reports
 * ===========================================================
 * Sales entries currently store quantity and pieces, but no selling price.
 * Reports therefore focus on volume, order activity, customers, and products.
 */

function getSalesReportSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(SHEETS.SALES_REPORT);
	if (!sheet) sheet = ss.insertSheet(SHEETS.SALES_REPORT);
	return sheet;
}

function getSalesReportFilters() {
	const dashboard = getSheet(SHEETS.SALES_DASHBOARD);
	if (dashboard) {
		return normalizeSalesDashboardFilters(
			getSalesDashboardFilterValues(dashboard),
		);
	}
	const range = resolveDateRange('This Month');
	return {
		fromDate: normalizeDate(range.fromDate),
		toDate: normalizeDate(range.toDate),
		customerName: 'All',
		productName: 'All',
	};
}

function clearSalesReportSheet(sheet) {
	sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
	sheet.clear();
	sheet.clearConditionalFormatRules();
	sheet.setHiddenGridlines(true);
	sheet.setFrozenRows(0);
	sheet.setFrozenColumns(0);
}

function generateSalesSummaryReport() {
	const filters = getSalesReportFilters();
	const dataset = getSalesDashboardDataset(filters);
	const sheet = getSalesReportSheet();

	clearSalesReportSheet(sheet);
	writeSalesReportHeader(sheet, 'SALES SUMMARY REPORT', filters);
	writeSalesReportKpis(sheet, calculateSalesDashboardKpis(dataset));

	const productSummary = summarizeSales(dataset, 'productName');
	const customerSummary = summarizeSales(dataset, 'customerName');
	const dailySummary = summarizeSales(dataset, 'date', true);
	let row = 15;
	row = writeSalesReportTable(
		sheet,
		row,
		'Product Performance',
		['Product', 'Pieces Sold'],
		productSummary.map((item) => [item.label, item.pieces]),
	);
	row += 2;
	row = writeSalesReportTable(
		sheet,
		row,
		'Customer Performance',
		['Customer', 'Pieces Sold'],
		customerSummary.map((item) => [item.label, item.pieces]),
	);
	row += 2;
	writeSalesReportTable(
		sheet,
		row,
		'Daily Sales',
		['Date', 'Pieces Sold'],
		dailySummary.map((item) => [item.label, item.pieces]),
		true,
	);

	finalizeSalesReport(sheet);
}

function generateSalesOrderDetailReport() {
	const filters = getSalesReportFilters();
	const dataset = getSalesDashboardDataset(filters);
	const sheet = getSalesReportSheet();

	clearSalesReportSheet(sheet);
	writeSalesReportHeader(sheet, 'SALES ORDER DETAIL REPORT', filters);
	writeSalesReportKpis(sheet, calculateSalesDashboardKpis(dataset));

	const rows = dataset
		.slice()
		.sort(
			(a, b) =>
				b.date - a.date || String(a.salesId).localeCompare(String(b.salesId)),
		)
		.map((record) => [
			record.salesId,
			record.date,
			record.customerName,
			record.productName,
			record.unit,
			record.quantity,
			record.pieces,
		]);
	writeSalesOrderDetailTable(sheet, 15, rows);
	finalizeSalesReport(sheet);
}

function writeSalesReportHeader(sheet, title, filters) {
	sheet
		.getRange('A1:G1')
		.merge()
		.setValue(title)
		.setFontSize(18)
		.setFontWeight('bold')
		.setFontColor('#FFFFFF')
		.setBackground('#1F4E78')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');
	sheet.setRowHeight(1, 36);
	sheet
		.getRange('A2:G2')
		.merge()
		.setValue(
			'Generated: ' +
				Utilities.formatDate(
					new Date(),
					Session.getScriptTimeZone(),
					'dd/MM/yyyy hh:mm:ss a',
				),
		)
		.setFontStyle('italic')
		.setFontColor('#666666');

	const date = (value) =>
		Utilities.formatDate(
			new Date(value),
			Session.getScriptTimeZone(),
			'dd/MM/yyyy',
		);
	const filtersTable = [
		['From Date', date(filters.fromDate), 'To Date', date(filters.toDate)],
		['Customer', filters.customerName, 'Product', filters.productName],
	];
	sheet
		.getRange('A4:D5')
		.setValues(filtersTable)
		.setBorder(true, true, true, true, true, true);
	sheet.getRange('A4:A5').setFontWeight('bold').setBackground('#E8F0FE');
	sheet.getRange('C4:C5').setFontWeight('bold').setBackground('#E8F0FE');
	[180, 120, 210, 120, 120, 100, 110].forEach((width, index) =>
		sheet.setColumnWidth(index + 1, width),
	);
}

function writeSalesReportKpis(sheet, kpis) {
	const cards = [
		['A7', 'Sales Orders', formatSalesNumber(kpis.salesOrders)],
		['C7', 'Pieces Sold', formatSalesNumber(kpis.totalPieces)],
		['E7', 'Customers', formatSalesNumber(kpis.customers)],
		['G7', 'Products', formatSalesNumber(kpis.products)],
	];
	cards.forEach(([cell, label, value]) => {
		const anchor = sheet.getRange(cell);
		const row = anchor.getRow();
		const column = anchor.getColumn();
		sheet
			.getRange(row, column, 1, 1)
			.setValue(label)
			.setFontWeight('bold')
			.setBackground('#E8F0FE')
			.setHorizontalAlignment('center');
		sheet
			.getRange(row + 1, column, 2, 1)
			.merge()
			.setValue(value)
			.setFontSize(15)
			.setFontWeight('bold')
			.setHorizontalAlignment('center')
			.setVerticalAlignment('middle');
		sheet
			.getRange(row, column, 3, 1)
			.setBorder(true, true, true, true, true, true);
	});
	sheet.setRowHeight(7, 24);
	sheet.setRowHeights(8, 2, 27);
}

/** Writes a compact two-column report section and returns its final row. */
function writeSalesReportTable(sheet, row, title, headers, rows, hasDates) {
	sheet
		.getRange(row, 1, 1, 2)
		.merge()
		.setValue(title)
		.setFontWeight('bold')
		.setBackground('#E8F0FE');
	sheet
		.getRange(row + 1, 1, 1, 2)
		.setValues([headers])
		.setFontWeight('bold')
		.setBackground('#1F4E78')
		.setFontColor('#FFFFFF');
	sheet.getRange(row + 1, 1).setHorizontalAlignment('left');
	sheet.getRange(row + 1, 2).setHorizontalAlignment('right');
	if (!rows.length) {
		sheet
			.getRange(row + 2, 1, 1, 2)
			.merge()
			.setValue('No sales found for the selected filters.');
		return row + 2;
	}
	sheet.getRange(row + 2, 1, rows.length, 2).setValues(rows);
	if (hasDates)
		sheet.getRange(row + 2, 1, rows.length, 1).setNumberFormat('dd/MM/yyyy');
	sheet.getRange(row + 2, 1, rows.length, 1).setHorizontalAlignment('left');
	sheet
		.getRange(row + 2, 2, rows.length, 1)
		.setHorizontalAlignment('right')
		.setNumberFormat('#,##0');
	const totalRow = row + 2 + rows.length;
	sheet
		.getRange(totalRow, 1, 1, 2)
		.setValues([
			['Total', rows.reduce((sum, item) => sum + Number(item[1] || 0), 0)],
		])
		.setFontWeight('bold')
		.setBackground('#E8F0FE');
	sheet.getRange(totalRow, 1).setHorizontalAlignment('left');
	sheet
		.getRange(totalRow, 2)
		.setHorizontalAlignment('right')
		.setNumberFormat('#,##0');
	sheet
		.getRange(row, 1, totalRow - row + 1, 2)
		.setBorder(true, true, true, true, true, true);
	return totalRow;
}

function writeSalesOrderDetailTable(sheet, row, rows) {
	const headers = [
		'Sales ID',
		'Date',
		'Customer',
		'Product',
		'Unit',
		'Qty',
		'Pieces',
	];
	sheet
		.getRange(row, 1, 1, headers.length)
		.merge()
		.setValue('Sales Order Lines')
		.setFontWeight('bold')
		.setBackground('#E8F0FE');
	sheet
		.getRange(row + 1, 1, 1, headers.length)
		.setValues([headers])
		.setFontWeight('bold')
		.setBackground('#1F4E78')
		.setFontColor('#FFFFFF');
	sheet.getRange(row + 1, 1, 1, 5).setHorizontalAlignment('left');
	sheet.getRange(row + 1, 6, 1, 2).setHorizontalAlignment('right');
	if (!rows.length) {
		sheet
			.getRange(row + 2, 1, 1, headers.length)
			.merge()
			.setValue('No sales found for the selected filters.');
		return;
	}
	sheet.getRange(row + 2, 1, rows.length, headers.length).setValues(rows);
	sheet.getRange(row + 2, 2, rows.length, 1).setNumberFormat('dd/MM/yyyy');
	sheet.getRange(row + 2, 6, rows.length, 2).setNumberFormat('#,##0');
	sheet.getRange(row + 2, 1, rows.length, 5).setHorizontalAlignment('left');
	sheet.getRange(row + 2, 6, rows.length, 2).setHorizontalAlignment('right');
	const totalRow = row + 2 + rows.length;
	const totalPieces = rows.reduce((sum, item) => sum + Number(item[6] || 0), 0);
	sheet
		.getRange(totalRow, 1, 1, headers.length)
		.setValues([['Total', '', '', '', '', '', totalPieces]])
		.setFontWeight('bold')
		.setBackground('#E8F0FE');
	sheet
		.getRange(totalRow, 7)
		.setNumberFormat('#,##0')
		.setHorizontalAlignment('right');
	sheet
		.getRange(row, 1, totalRow - row + 1, headers.length)
		.setBorder(true, true, true, true, true, true);
}

function finalizeSalesReport(sheet) {
	makeReportPrintable(sheet);
	// Keep filters and KPI strip visible while reviewing the sales report.
	sheet.setFrozenRows(6);
	SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
}
