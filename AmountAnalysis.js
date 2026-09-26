/**
 * ===========================================================
 * Amount Analysis
 * -----------------------------------------------------------
 * Production Amount = Production Pieces × Product Price
 * Sales Amount      = Sales Pieces × Product Price
 * ===========================================================
 */

const AMOUNT_ANALYSIS_ALLOWED_EMAILS = [
	'rishabhk3003@gmail.com',
	'devanshu189@gmail.com',
];

function isAmountAnalysisAuthorized() {
	const email = String(Session.getActiveUser().getEmail() || '')
		.trim()
		.toLowerCase();

	return (
		email !== '' &&
		AMOUNT_ANALYSIS_ALLOWED_EMAILS.map((item) =>
			String(item).trim().toLowerCase(),
		).includes(email)
	);
}

function assertAmountAnalysisAuthorized() {
	if (!isAmountAnalysisAuthorized()) {
		throw new Error('You are not authorized to access Amount Analysis.');
	}
}

function getAmountAnalysisFormData() {
	assertAmountAnalysisAuthorized();

	syncProductPricingSheet_();

	return {
		dateRanges: DASHBOARD.DATE_RANGES,
		products: getAllProducts(),
		machines: getMachines(),
	};
}

function getAmountAnalysisDateRange(option) {
	assertAmountAnalysisAuthorized();

	const range = resolveDateRange(option || 'This Month');

	return {
		fromDate: Utilities.formatDate(
			range.fromDate,
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		),
		toDate: Utilities.formatDate(
			range.toDate,
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		),
	};
}

function calculateAmountAnalysis(filters) {
	assertAmountAnalysisAuthorized();

	const normalized = normalizeAmountAnalysisFilters_(filters);
	const prices = getProductPriceMap_();

	const dataset =
		normalized.type === 'production'
			? getProductionAmountDataset_(normalized)
			: getSalesAmountDataset_(normalized);

	const rows = [];
	const missingPrices = new Set();

	const totals = {
		pieces: 0,
		wastePieces: 0,
		netPieces: 0,
		amount: 0,
		wasteAmount: 0,
		adjustedAmount: 0,
	};

	dataset.forEach((record) => {
		const price = prices[record.productId];

		if (price === null || price === undefined || price === '') {
			missingPrices.add(record.productId + '|' + record.productName);
			return;
		}

		const numericPrice = Number(price);

		if (!isFinite(numericPrice) || numericPrice < 0) {
			missingPrices.add(record.productId + '|' + record.productName);
			return;
		}

		const pieces = Number(record.pieces || 0);

		let wastePieces = 0;
		let netPieces = pieces;
		let amount = 0;
		let wasteAmount = 0;
		let adjustedAmount = amount;

		if (normalized.type === 'production') {
			wastePieces = pieces * 0.01;
			netPieces = pieces - wastePieces;

			wasteAmount = wastePieces * numericPrice;
			amount = netPieces * numericPrice;
			adjustedAmount = amount - wasteAmount;
		} else {
			netPieces = pieces;
			amount = pieces * numericPrice;
		}

		totals.pieces += pieces;
		totals.wastePieces += wastePieces;
		totals.netPieces += netPieces;
		totals.amount += amount;
		totals.wasteAmount += wasteAmount;
		totals.adjustedAmount += adjustedAmount;

		const existing = rows.find((item) => item.productId === record.productId);

		if (existing) {
			existing.pieces += pieces;
			existing.wastePieces += wastePieces;
			existing.netPieces += netPieces;
			existing.amount += amount;
			existing.wasteAmount += wasteAmount;
			existing.adjustedAmount += adjustedAmount;
		} else {
			rows.push({
				productId: record.productId,
				productName: record.productName,
				pieces,
				wastePieces,
				netPieces,
				price: numericPrice,
				amount,
				wasteAmount,
				adjustedAmount,
			});
		}
	});

	if (missingPrices.size > 0) {
		const missing = Array.from(missingPrices).map((item) => {
			const parts = item.split('|');
			return {
				productId: parts[0],
				productName: parts.slice(1).join('|'),
			};
		});

		return {
			success: false,
			errorType: 'MISSING_PRICES',
			missingPrices: missing,
			totals,
			rows,
		};
	}

	rows.sort((a, b) => b.amount - a.amount);

	return {
		success: true,
		type: normalized.type,
		filters: {
			type: normalized.type,
			dateRange: normalized.dateRange,
			fromDate: Utilities.formatDate(
				normalized.fromDate,
				Session.getScriptTimeZone(),
				'yyyy-MM-dd',
			),
			toDate: Utilities.formatDate(
				normalized.toDate,
				Session.getScriptTimeZone(),
				'yyyy-MM-dd',
			),
			productId: normalized.productId,
			machineId: normalized.machineId,
		},
		totals,
		rows,
	};
}

function normalizeAmountAnalysisFilters_(filters) {
	filters = filters || {};

	const type = filters.type === 'sales' ? 'sales' : 'production';
	const dateRange = String(filters.dateRange || 'This Month');

	let fromDate;
	let toDate;

	if (dateRange === 'Custom') {
		if (!filters.fromDate || !filters.toDate) {
			throw new Error('Please select both From Date and To Date.');
		}

		fromDate = normalizeDate(filters.fromDate);
		toDate = normalizeDate(filters.toDate);
		toDate.setHours(23, 59, 59, 999);
	} else {
		const range = resolveDateRange(dateRange);
		fromDate = range.fromDate;
		toDate = range.toDate;
	}

	if (fromDate > toDate) {
		throw new Error('From Date cannot be after To Date.');
	}

	return {
		type,
		dateRange,
		fromDate,
		toDate,
		productId: filters.productId || 'All',
		machineId: type === 'production' ? filters.machineId || 'All' : 'All',
	};
}

function getProductionAmountDataset_(filters) {
	return getProductionDataset({
		machineId: filters.machineId,
		operatorId: 'All',
		productId: filters.productId,
		fromDate: filters.fromDate,
		toDate: filters.toDate,
	});
}

function getSalesAmountDataset_(filters) {
	const headerSheet = getSheet(SHEETS.SALES_HEADER);
	const detailSheet = getSheet(SHEETS.SALES_DETAILS);

	if (!headerSheet || !detailSheet) {
		return [];
	}

	const headers = headerSheet.getDataRange().getValues();
	const details = detailSheet.getDataRange().getValues();

	if (headers.length <= 1 || details.length <= 1) {
		return [];
	}

	const salesMap = {};

	headers.slice(1).forEach((row) => {
		salesMap[row[0]] = {
			salesId: row[0],
			date: getSalesDateOnly(row[1]),
		};
	});

	return details
		.slice(1)
		.map((row) => {
			const sale = salesMap[row[0]];

			if (!sale) {
				return null;
			}

			return {
				salesId: sale.salesId,
				date: sale.date,
				productId: row[1],
				productName: row[2],
				pieces: Number(row[5] || 0),
			};
		})
		.filter(Boolean)
		.filter((record) => {
			if (
				filters.productId !== 'All' &&
				record.productId !== filters.productId
			) {
				return false;
			}

			return record.date >= filters.fromDate && record.date <= filters.toDate;
		});
}

function getProductPriceMap_() {
	const sheet = getSheet(SHEETS.PRODUCT_PRICING);

	if (!sheet || sheet.getLastRow() <= 1) {
		return {};
	}

	const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
	const map = {};

	values.forEach((row) => {
		const productId = String(row[0] || '').trim();

		if (productId) {
			map[productId] = row[1];
		}
	});

	return map;
}

function syncProductPricingSheet_() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(SHEETS.PRODUCT_PRICING);

	if (!sheet) {
		sheet = ss.insertSheet(SHEETS.PRODUCT_PRICING);
		sheet.getRange(1, 1, 1, 2).setValues([['Product ID', 'Price']]);
		sheet
			.getRange('A1:B1')
			.setFontWeight('bold')
			.setBackground('#1F4E78')
			.setFontColor('#FFFFFF');
		sheet.setFrozenRows(1);
		sheet.setColumnWidth(1, 160);
		sheet.setColumnWidth(2, 120);
		sheet.getRange('B:B').setNumberFormat('₹#,##0.00');
	}

	const existingIds = new Set();

	if (sheet.getLastRow() > 1) {
		sheet
			.getRange(2, 1, sheet.getLastRow() - 1, 1)
			.getValues()
			.forEach((row) => {
				const id = String(row[0] || '').trim();
				if (id) existingIds.add(id);
			});
	}

	const products = getAllProducts();
	const missingRows = products
		.filter((product) => !existingIds.has(String(product.id)))
		.map((product) => [product.id, '']);

	if (missingRows.length > 0) {
		sheet
			.getRange(sheet.getLastRow() + 1, 1, missingRows.length, 2)
			.setValues(missingRows);
	}

	sheet.getRange('B:B').setNumberFormat('₹#,##0.00');

	return sheet;
}

/**
 * Run once manually after deploying this feature.
 * It creates Product_Pricing and populates missing Product IDs.
 */
function initializeProductPricing() {
	syncProductPricingSheet_();
	SpreadsheetApp.getUi().alert(
		'Product_Pricing sheet is ready. Enter the price for each Product ID in column B.',
	);
}
