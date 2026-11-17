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

const PRODUCT_PRICING_JSON = [
	{ productId: 'P001', productName: 'UMBRELLA DUCK', ratePerPc: 13.9 },
	{ productId: 'P002', productName: 'NAUGHTY MONKEY', ratePerPc: 11.6 },
	{ productId: 'P003', productName: 'DANA DUCK', ratePerPc: 13.9 },
	{ productId: 'P004', productName: 'PET PUPPY', ratePerPc: 12.8 },
	{ productId: 'P005', productName: 'NUMERICAL BALL', ratePerPc: 16.08 },
	{ productId: 'P006', productName: '4 YAAR', ratePerPc: 11.2 },
	{ productId: 'P007', productName: 'KING ELEPHANT', ratePerPc: 14.97 },
	{ productId: 'P008', productName: '12 PC CHU CHU', ratePerPc: 7.6 },
	{ productId: 'P009', productName: 'FAIRY SET', ratePerPc: 13.9 },
	{ productId: 'P010', productName: 'LOVELY CAT', ratePerPc: 16.05 },
	{ productId: 'P011', productName: 'BALLOON ANIMAL', ratePerPc: 14.97 },
	{ productId: 'P012', productName: 'CALLY SHEEP', ratePerPc: 14.2 },
	{ productId: 'P013', productName: 'VEGETABLES BIG', ratePerPc: 12.6 },
	{ productId: 'P014', productName: 'BOTTLE', ratePerPc: 25.6 },
	{ productId: 'P015', productName: 'BABY BIG', ratePerPc: 25.6 },
	{ productId: 'P016', productName: 'LOVE COUPLE', ratePerPc: 14.97 },
	{ productId: 'P017', productName: 'DASHING TEDDY', ratePerPc: 19.25 },
	{ productId: 'P018', productName: 'JOLLY RABBIT', ratePerPc: 14.97 },
	{ productId: 'P019', productName: 'PARI DUCK', ratePerPc: 12.6 },
	{ productId: 'P020', productName: 'DUCK BABY', ratePerPc: 14.97 },
	{ productId: 'P021', productName: 'PARTY TIME', ratePerPc: 14.97 },
	{ productId: 'P022', productName: 'FRENCH FRIES', ratePerPc: 14.97 },
	{ productId: 'P023', productName: 'HAPPY TEDDY', ratePerPc: 11.6 },
	{ productId: 'P024', productName: 'LOVE N LILY', ratePerPc: 19.25 },
	{ productId: 'P025', productName: 'CUTE BEAR$DOG', ratePerPc: 13.9 },
	{ productId: 'P026', productName: 'DUCK FRIEND', ratePerPc: 14.97 },
	{ productId: 'P027', productName: 'VALLEY FRIENDS', ratePerPc: 16.05 },
	{ productId: 'P028', productName: 'FARM FRIENDS', ratePerPc: 13.9 },
	{ productId: 'P029', productName: 'RAINBOW FISH', ratePerPc: 11.6 },
	{ productId: 'P030', productName: 'FUNNY CLOWN', ratePerPc: 19.25 },
	{ productId: 'P031', productName: 'FLOWER BABY', ratePerPc: 22.2 },
	{ productId: 'P032', productName: 'CRAZY BIRD', ratePerPc: 14.97 },
	{ productId: 'P033', productName: 'COOKIE DUCK', ratePerPc: 13.55 },
	{ productId: 'P034', productName: 'WILD ANIMAL', ratePerPc: 11.6 },
	{ productId: 'P035', productName: 'YELLOW DUCK', ratePerPc: 10.5 },
	{ productId: 'P036', productName: 'MEDIUM FRUIT', ratePerPc: 10.5 },
	{ productId: 'P037', productName: 'TRANSPORT 1109', ratePerPc: 14.5 },
	{ productId: 'P038', productName: 'M B SWIMIMG DUCK', ratePerPc: 11.5 },
	{ productId: 'P039', productName: 'SMALL FRUIT', ratePerPc: 10.5 },
	{ productId: 'P040', productName: 'BIG FRUIT', ratePerPc: 12.0 },
	{ productId: 'P041', productName: 'DALMATION DOG', ratePerPc: 14.97 },
	{ productId: 'P042', productName: 'VOLLEY FRIENDS', ratePerPc: 16.05 },
	{ productId: 'P043', productName: 'PUP BUDDIES', ratePerPc: 16.05 },
	{ productId: 'P044', productName: 'MINI ZOO', ratePerPc: 12.83 },
	{ productId: 'P045', productName: 'HENNY EGG', ratePerPc: 13.55 },
	{ productId: 'P046', productName: 'BABY BASKET', ratePerPc: 18.83 },
	{ productId: 'P047', productName: 'FUNNY ANIMAL', ratePerPc: 13.9 },
	{ productId: 'P048', productName: 'PARTY TIME BOY', ratePerPc: 14.97 },
	{ productId: 'P049', productName: 'ANIMAL TOKREY', ratePerPc: 16.05 },
	{ productId: 'P050', productName: 'DOG SET', ratePerPc: 13.9 },
	{ productId: 'P051', productName: 'POCKET PETS', ratePerPc: 12.0 },
	{ productId: 'P052', productName: 'PENGUIN 6 PC', ratePerPc: 11.23 },
	{ productId: 'P053', productName: 'M BANSAL KHILONA', ratePerPc: 14.0 },
	{ productId: 'P054', productName: 'LEOPLAST 1105', ratePerPc: 14.5 },
	{ productId: 'P055', productName: 'LEOPLAST 1104', ratePerPc: 14.5 },
	{ productId: 'P056', productName: 'LEOPLAST 1106', ratePerPc: 14.5 },
	{ productId: 'P057', productName: 'LEOPLAST 1107', ratePerPc: 0.0 },
	{ productId: 'P058', productName: 'LEOPLAST 1108', ratePerPc: 11.5 },
	{ productId: 'P059', productName: 'LEOPLAST 1109', ratePerPc: 14.5 },
	{ productId: 'P060', productName: 'LEOPLAST 1111', ratePerPc: 12.16 },
	{ productId: 'P061', productName: 'LEOPLAST 1112', ratePerPc: 14.5 },
	{ productId: 'P062', productName: 'LEOPLAST 1113', ratePerPc: 15.5 },
	{ productId: 'P063', productName: 'LEOPLAST 1117', ratePerPc: 15.0 },
	{ productId: 'P064', productName: 'LEOPLAST 1118', ratePerPc: 15.0 },
	{ productId: 'P065', productName: 'LEOPLAST 1119', ratePerPc: 15.0 },
	{ productId: 'P066', productName: 'LEOPLAST 1120', ratePerPc: 15.0 },
	{ productId: 'P067', productName: 'LEOPLAST 1121', ratePerPc: 0.0 },
	{ productId: 'P068', productName: 'LEOPLAST 1122', ratePerPc: 13.5 },
	{ productId: 'P069', productName: 'LEOPLAST 1123', ratePerPc: 0.0 },
	{ productId: 'P070', productName: 'LEOPLAST 1124', ratePerPc: 15.5 },
	{ productId: 'P071', productName: 'LEOPLAST 1125', ratePerPc: 15.0 },
	{ productId: 'P072', productName: 'LEOPLAST 1126', ratePerPc: 15.33 },
	{ productId: 'P073', productName: 'LEOPLAST 1127', ratePerPc: 15.33 },
	{ productId: 'P074', productName: 'LEOPLAST 1128', ratePerPc: 15.33 },
	{ productId: 'P075', productName: 'LEOPLAST 1129', ratePerPc: 15.33 },
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
	const map = {};

	PRODUCT_PRICING_JSON.forEach((row) => {
		const productId = row.productId ? String(row.productId).trim() : null;

		if (productId) {
			map[productId] =
				row.ratePerPc !== undefined && row.ratePerPc !== null
					? Number(row.ratePerPc)
					: null;
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
