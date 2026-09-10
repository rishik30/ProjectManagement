/**
 * ===========================================================
 * Sales Dashboard
 * ===========================================================
 * Uses Sales_Header and Sales_Details only. The system does not store prices,
 * so every metric is based on sale orders and product pieces—not revenue.
 */

const SALES_DASHBOARD_FILTER_CELLS = ['B3', 'B4', 'B5', 'B6', 'B7'];
const SALES_DASHBOARD_DATE_RANGES = [
	'Today',
	'This Week',
	'This Month',
	'Last Month',
	'Last 3 Months',
	'Financial Year',
	'Custom',
];

function openSalesDashboard() {
	refreshSalesDashboard();
}

function refreshSalesDashboard() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(SHEETS.SALES_DASHBOARD);
	if (!sheet) sheet = ss.insertSheet(SHEETS.SALES_DASHBOARD);

	ensureSalesDashboardLayout(sheet);
	applySalesDashboardValidation(sheet);
	refreshSalesDashboardData(sheet);
	ss.setActiveSheet(sheet);
}

function ensureSalesDashboardLayout(sheet) {
	if (
		sheet.getRange('A1').getValue() === 'SALES DASHBOARD' &&
		sheet.getRange('A3').getValue() === 'Date Range'
	) {
		return false;
	}

	sheet.clear();
	sheet.clearConditionalFormatRules();
	sheet.setHiddenGridlines(true);
	sheet.setFrozenRows(2);
	buildSalesDashboardLayout(sheet);
	return true;
}

function buildSalesDashboardLayout(sheet) {
	const today = new Date();
	const firstDay = new Date(today.getFullYear(), today.getMonth(), 1);

	sheet
		.getRange('A1:J1')
		.merge()
		.setValue('SALES DASHBOARD')
		.setFontSize(18)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle')
		.setBackground('#1F4E78')
		.setFontColor('#FFFFFF');
	sheet.setRowHeight(1, 38);
	sheet.getRange('A2:J2').merge().setValue('Ready').setFontStyle('italic');

	sheet
		.getRange('A3:A7')
		.setValues([
			['Date Range'],
			['From Date'],
			['To Date'],
			['Customer'],
			['Product'],
		])
		.setFontWeight('bold');
	sheet.getRange('B3').setValue('This Month');
	sheet.getRange('B4').setValue(firstDay).setNumberFormat('dd/MM/yyyy');
	sheet.getRange('B5').setValue(today).setNumberFormat('dd/MM/yyyy');
	sheet.getRange('B6:B7').setValue('All');

	buildSalesKpiCard(sheet, 'A10', 'Sales Orders');
	buildSalesKpiCard(sheet, 'C10', 'Pieces Sold');
	buildSalesKpiCard(sheet, 'E10', 'Customers');
	buildSalesKpiCard(sheet, 'G10', 'Products Sold');
	buildSalesKpiCard(sheet, 'I10', 'Avg Pieces / Order');
	buildSalesKpiCard(sheet, 'A14', 'Avg Pieces / Day');
	buildSalesKpiCard(sheet, 'C14', 'Top Customer');
	buildSalesKpiCard(sheet, 'E14', 'Top Product');
	buildSalesKpiCard(sheet, 'G14', 'Latest Sale');

	setSalesKpiRowHeights(sheet);
	setSalesDashboardColumnWidths(sheet);
}

function setSalesKpiRowHeights(sheet) {
	// Each value spans two rows. These heights fit two-line labels such as
	// "Top Customer" without text overflowing its bordered KPI card.
	[10, 14].forEach((titleRow) => {
		sheet.setRowHeight(titleRow, 26);
		sheet.setRowHeights(titleRow + 1, 2, 34);
	});
}

function setSalesDashboardColumnWidths(sheet) {
	[165, 110, 165, 110, 165, 110, 165, 110, 165, 110].forEach((width, index) =>
		sheet.setColumnWidth(index + 1, width),
	);
}

function buildSalesKpiCard(sheet, cell, title) {
	const anchor = sheet.getRange(cell);
	const row = anchor.getRow();
	const column = anchor.getColumn();
	sheet
		.getRange(row, column, 1, 2)
		.merge()
		.setValue(title)
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBackground('#E8F0FE');
	sheet
		.getRange(row + 1, column, 2, 2)
		.merge()
		.setValue('-')
		.setFontSize(16)
		.setFontWeight('bold')
		.setWrap(true)
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');
	sheet
		.getRange(row, column, 3, 2)
		.setBorder(true, true, true, true, true, true);
}

function getSalesDashboardFilterValues(sheet) {
	return {
		dateRange: sheet.getRange('B3').getValue(),
		fromDate: sheet.getRange('B4').getValue(),
		toDate: sheet.getRange('B5').getValue(),
		customerName: sheet.getRange('B6').getValue(),
		productName: sheet.getRange('B7').getValue(),
	};
}

function normalizeSalesDashboardFilters(filters) {
	const range =
		filters.dateRange === 'Custom'
			? {
					fromDate: normalizeDate(filters.fromDate),
					toDate: normalizeDate(filters.toDate),
				}
			: resolveDateRange(filters.dateRange || 'This Month');
	if (!range.fromDate || !range.toDate || range.fromDate > range.toDate) {
		throw new Error('From Date cannot be after To Date.');
	}
	return {
		fromDate: normalizeDate(range.fromDate),
		toDate: normalizeDate(range.toDate),
		customerName: filters.customerName || 'All',
		productName: filters.productName || 'All',
	};
}

function applySalesDashboardValidation(sheet) {
	const masters = getSalesDashboardMasters();
	const validation = (values) =>
		SpreadsheetApp.newDataValidation()
			.requireValueInList(values, true)
			.setAllowInvalid(false)
			.build();
	sheet
		.getRange('B3')
		.setDataValidation(validation(SALES_DASHBOARD_DATE_RANGES));
	sheet
		.getRange('B6')
		.setDataValidation(validation(['All'].concat(masters.customers)));
	sheet
		.getRange('B7')
		.setDataValidation(validation(['All'].concat(masters.products)));
}

function getSalesDashboardMasters() {
	const customers = new Set();
	const products = new Set();
	const headerSheet = getSheet(SHEETS.SALES_HEADER);
	const detailSheet = getSheet(SHEETS.SALES_DETAILS);
	if (headerSheet && headerSheet.getLastRow() > 1) {
		headerSheet
			.getRange(2, 4, headerSheet.getLastRow() - 1, 1)
			.getValues()
			.forEach((row) => {
				if (row[0]) customers.add(String(row[0]));
			});
	}
	if (detailSheet && detailSheet.getLastRow() > 1) {
		detailSheet
			.getRange(2, 3, detailSheet.getLastRow() - 1, 1)
			.getValues()
			.forEach((row) => {
				if (row[0]) products.add(String(row[0]));
			});
	}
	return {
		customers: Array.from(customers).sort(),
		products: Array.from(products).sort(),
	};
}

function getSalesDashboardDataset(filters) {
	const headerSheet = getSheet(SHEETS.SALES_HEADER);
	const detailSheet = getSheet(SHEETS.SALES_DETAILS);
	if (!headerSheet || !detailSheet) return [];
	const headers = headerSheet.getDataRange().getValues();
	const details = detailSheet.getDataRange().getValues();
	if (headers.length <= 1 || details.length <= 1) return [];
	const saleMap = {};
	headers.slice(1).forEach((row) => {
		saleMap[row[0]] = {
			salesId: row[0],
			date: normalizeDate(row[1]),
			customerId: row[2],
			customerName: row[3],
		};
	});
	return details
		.slice(1)
		.map((row) => {
			const sale = saleMap[row[0]];
			if (!sale) return null;
			return {
				salesId: sale.salesId,
				date: sale.date,
				customerId: sale.customerId,
				customerName: sale.customerName,
				productId: row[1],
				productName: row[2],
				unit: row[3],
				quantity: Number(row[4] || 0),
				pieces: Number(row[5] || 0),
			};
		})
		.filter(Boolean)
		.filter(
			(record) =>
				record.date >= filters.fromDate &&
				record.date <= filters.toDate &&
				(filters.customerName === 'All' ||
					record.customerName === filters.customerName) &&
				(filters.productName === 'All' ||
					record.productName === filters.productName),
		);
}

function calculateSalesDashboardKpis(dataset) {
	const sales = new Set();
	const customers = new Set();
	const products = new Set();
	const days = new Set();
	const customerTotals = {};
	const productTotals = {};
	let totalPieces = 0;
	let latestSale = null;

	dataset.forEach((record) => {
		sales.add(record.salesId);
		customers.add(record.customerId);
		products.add(record.productId);
		days.add(record.date.getTime());
		totalPieces += record.pieces;
		customerTotals[record.customerName] =
			(customerTotals[record.customerName] || 0) + record.pieces;
		productTotals[record.productName] =
			(productTotals[record.productName] || 0) + record.pieces;
		if (!latestSale || record.date > latestSale) latestSale = record.date;
	});
	const top = (totals) => Object.entries(totals).sort((a, b) => b[1] - a[1])[0];
	return {
		salesOrders: sales.size,
		totalPieces,
		customers: customers.size,
		products: products.size,
		averagePiecesPerOrder: sales.size ? totalPieces / sales.size : 0,
		averagePiecesPerDay: days.size ? totalPieces / days.size : 0,
		topCustomer: top(customerTotals),
		topProduct: top(productTotals),
		latestSale,
	};
}

function refreshSalesDashboardData(sheet) {
	try {
		setSalesDashboardStatus(sheet, '🟡 Refreshing Sales Dashboard...');
		const filters = normalizeSalesDashboardFilters(
			getSalesDashboardFilterValues(sheet),
		);
		const dataset = getSalesDashboardDataset(filters);
		const kpis = calculateSalesDashboardKpis(dataset);
		updateSalesDashboardKpis(sheet, kpis);
		writeSalesDashboardTables(sheet, dataset);
		setSalesDashboardStatus(
			sheet,
			'🟢 Ready | Last Updated: ' +
				Utilities.formatDate(
					new Date(),
					Session.getScriptTimeZone(),
					'dd/MM/yyyy hh:mm:ss a',
				),
		);
	} catch (error) {
		setSalesDashboardStatus(sheet, '🔴 ' + error.message);
		logError('refreshSalesDashboardData', error);
		throw error;
	}
}

function setSalesDashboardStatus(sheet, message) {
	sheet.getRange('A2:J2').breakApart().merge().setValue(message);
}

function updateSalesDashboardKpis(sheet, kpis) {
	setSalesKpiRowHeights(sheet);
	const values = {
		A10: kpis.salesOrders,
		C10: kpis.totalPieces,
		E10: kpis.customers,
		G10: kpis.products,
		I10: kpis.averagePiecesPerOrder,
		A14: kpis.averagePiecesPerDay,
		C14: kpis.topCustomer
			? kpis.topCustomer[0] +
				'\n' +
				formatSalesNumber(kpis.topCustomer[1]) +
				' pcs'
			: '-',
		E14: kpis.topProduct
			? kpis.topProduct[0] +
				'\n' +
				formatSalesNumber(kpis.topProduct[1]) +
				' pcs'
			: '-',
		G14: kpis.latestSale
			? Utilities.formatDate(
					kpis.latestSale,
					Session.getScriptTimeZone(),
					'dd/MM/yyyy',
				)
			: '-',
	};
	Object.entries(values).forEach(([cell, value]) => {
		const anchor = sheet.getRange(cell);
		const range = sheet.getRange(anchor.getRow() + 1, anchor.getColumn(), 2, 2);
		range
			.breakApart()
			.merge()
			.setValue(typeof value === 'number' ? formatSalesNumber(value) : value)
			.setHorizontalAlignment('center')
			.setVerticalAlignment('middle')
			.setWrap(true);
	});
}

function writeSalesDashboardTables(sheet, dataset) {
	clearSalesDashboardTables(sheet);
	const productSummary = summarizeSales(dataset, 'productName');
	const customerSummary = summarizeSales(dataset, 'customerName');
	const dailySummary = summarizeSales(dataset, 'date', true);
	writeSalesDashboardSummary(
		sheet,
		19,
		1,
		'Product Performance',
		['Product', 'Pieces'],
		productSummary,
	);
	writeSalesDashboardSummary(
		sheet,
		19,
		4,
		'Customer Performance',
		['Customer', 'Pieces'],
		customerSummary,
	);
	writeSalesDashboardSummary(
		sheet,
		19,
		7,
		'Daily Sales',
		['Date', 'Pieces'],
		dailySummary,
	);
	const recentSalesRow = Math.max(
		42,
		23 +
			Math.max(
				productSummary.length,
				customerSummary.length,
				dailySummary.length,
			),
	);
	writeRecentSalesTable(sheet, dataset, recentSalesRow);
}

function clearSalesDashboardTables(sheet) {
	// Remove prior title/empty-state merges before rewriting the tables.
	sheet.getRange('A19:J1000').breakApart().clearContent().clearFormat();
}

function summarizeSales(dataset, field, dateField) {
	const totals = {};
	dataset.forEach((record) => {
		const key = dateField ? record.date.getTime() : record[field];
		if (!totals[key])
			totals[key] = {
				label: dateField ? record.date : record[field],
				pieces: 0,
			};
		totals[key].pieces += record.pieces;
	});
	return Object.values(totals).sort((a, b) =>
		dateField ? a.label - b.label : b.pieces - a.pieces,
	);
}

function writeSalesDashboardSummary(sheet, row, column, title, headers, items) {
	const titleRange = sheet.getRange(row, column, 1, 2);
	const headerRange = sheet.getRange(row + 1, column, 1, 2);

	titleRange
		.merge()
		.setValue(title)
		.setFontWeight('bold')
		.setBackground('#E8F0FE')
		.setHorizontalAlignment('left')
		.setVerticalAlignment('middle')
		.setBorder(true, true, false, true, false, false);
	headerRange
		.setValues([headers])
		.setFontWeight('bold')
		.setBackground('#1F4E78')
		.setFontColor('#FFFFFF')
		.setVerticalAlignment('middle')
		.setBorder(false, true, true, true, false, true);
	headerRange.getCell(1, 1).setHorizontalAlignment('left');
	headerRange.getCell(1, 2).setHorizontalAlignment('right');
	if (!items.length) {
		sheet
			.getRange(row + 2, column, 1, 2)
			.merge()
			.setValue('No sales found.')
			.setBorder(false, true, true, true, false, false);
		return;
	}
	const values = items.map((item) => [
		item.label instanceof Date
			? Utilities.formatDate(
					item.label,
					Session.getScriptTimeZone(),
					'dd/MM/yyyy',
				)
			: item.label,
		item.pieces,
	]);
	const dataRange = sheet.getRange(row + 2, column, values.length, 2);
	dataRange
		.setValues(values)
		.setVerticalAlignment('middle')
		.setBorder(false, true, false, true, false, true);
	sheet
		.getRange(row + 2, column, values.length, 1)
		.setHorizontalAlignment('left');
	sheet
		.getRange(row + 2, column + 1, values.length, 1)
		.setHorizontalAlignment('right')
		.setNumberFormat('#,##0');
	sheet
		.getRange(row + 2 + values.length, column, 1, 2)
		.setValues([['Total', items.reduce((sum, item) => sum + item.pieces, 0)]])
		.setFontWeight('bold')
		.setBackground('#E8F0FE')
		.setBorder(false, true, true, true, false, true);
	sheet
		.getRange(row + 2 + values.length, column, 1, 1)
		.setHorizontalAlignment('left');
	sheet
		.getRange(row + 2 + values.length, column + 1, 1, 1)
		.setHorizontalAlignment('right')
		.setNumberFormat('#,##0');
}

function writeRecentSalesTable(sheet, dataset, row) {
	const records = Object.values(
		dataset.reduce((sales, record) => {
			if (!sales[record.salesId])
				sales[record.salesId] = {
					salesId: record.salesId,
					date: record.date,
					customerName: record.customerName,
					pieces: 0,
				};
			sales[record.salesId].pieces += record.pieces;
			return sales;
		}, {}),
	)
		.sort((a, b) => b.date - a.date)
		.slice(0, 10);
	const titleRange = sheet.getRange(row, 1, 1, 4);
	const headerRange = sheet.getRange(row + 1, 1, 1, 4);
	titleRange
		.merge()
		.setValue('Recent Sales Orders')
		.setFontWeight('bold')
		.setBackground('#E8F0FE')
		.setBorder(true, true, false, true, false, false);
	headerRange
		.setValues([['Sales ID', 'Date', 'Customer', 'Pieces']])
		.setFontWeight('bold')
		.setBackground('#1F4E78')
		.setFontColor('#FFFFFF')
		.setBorder(false, true, true, true, false, true);
	sheet.getRange(row + 1, 1, 1, 3).setHorizontalAlignment('left');
	headerRange.getCell(1, 4).setHorizontalAlignment('right');
	if (!records.length) {
		sheet
			.getRange(row + 2, 1, 1, 4)
			.merge()
			.setValue('No sales found.')
			.setBorder(false, true, true, true, false, false);
		return;
	}
	const dataRange = sheet
		.getRange(row + 2, 1, records.length, 4)
		.setValues(
			records.map((record) => [
				record.salesId,
				record.date,
				record.customerName,
				record.pieces,
			]),
		)
		.setBorder(false, true, true, true, false, true)
		.setVerticalAlignment('middle');
	sheet.getRange(row + 2, 1, records.length, 3).setHorizontalAlignment('left');
	sheet.getRange(row + 2, 2, records.length, 1).setNumberFormat('dd/MM/yyyy');
	sheet
		.getRange(row + 2, 4, records.length, 1)
		.setHorizontalAlignment('right')
		.setNumberFormat('#,##0');
}

function formatSalesNumber(value) {
	return Number(value || 0).toLocaleString(undefined, {
		maximumFractionDigits: 2,
	});
}

function handleSalesDashboardEdit(e) {
	const sheet = e.range.getSheet();
	if (
		sheet.getName() !== SHEETS.SALES_DASHBOARD ||
		!SALES_DASHBOARD_FILTER_CELLS.includes(e.range.getA1Notation())
	)
		return;
	if (
		e.range.getA1Notation() === 'B3' &&
		sheet.getRange('B3').getValue() !== 'Custom'
	) {
		const range = resolveDateRange(sheet.getRange('B3').getValue());
		sheet.getRange('B4').setValue(range.fromDate);
		sheet.getRange('B5').setValue(range.toDate);
	}
	refreshSalesDashboardData(sheet);
}
