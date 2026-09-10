/**
 * Sales.js
 * -----------------------------------------------------------
 * File        : Sales.js
 * Description : Sales related functions
 * ===========================================================
 * {
	date: '',
	customerId: '',
	customerName: '',
	remarks: '',
	products: [
		{
			productId: '',
			productName: '',
			unit: 'Box',
			quantity: 2,
			pieces: 120
		}
	]
}
 */

function initializeSales() {
	createSheetIfMissing(SHEETS.SALES_HEADER, [
		'Sales ID',
		'Date',
		'Customer ID',
		'Customer Name',
		'Remarks',
	]);

	createSheetIfMissing(SHEETS.SALES_DETAILS, [
		'Sales ID',
		'Product ID',
		'Product Name',
		'Unit',
		'Quantity',
		'Pieces',
	]);
}

function generateSalesId() {
	const sheet = getSheet(SHEETS.SALES_HEADER);

	const lastRow = sheet.getLastRow();

	if (lastRow <= 1) {
		return 'SAL0001';
	}

	const lastId = sheet.getRange(lastRow, 1).getValue();

	const number = Number(lastId.replace('SAL', ''));

	return 'SAL' + String(number + 1).padStart(4, '0');
}

/**
 * Validates a sales object.
 *
 * @param {Object} sale
 */
function validateSale(sale) {
	try {
		if (!sale) {
			throw new Error('Sales data is required.');
		}

		if (!sale.date) {
			throw new Error('Date is required.');
		}

		if (!sale.customerId) {
			throw new Error('Customer is required.');
		}

		if (!sale.products || sale.products.length === 0) {
			throw new Error('At least one product is required.');
		}

		const productIds = new Set();

		sale.products.forEach((item) => {
			if (!item.productId) {
				throw new Error('Product is required.');
			}

			if (productIds.has(item.productId)) {
				throw new Error('Duplicate products are not allowed.');
			}

			productIds.add(item.productId);

			if (isNaN(item.quantity) || Number(item.quantity) <= 0) {
				throw new Error('Invalid quantity for ' + item.productName);
			}

			if (isNaN(item.pieces) || Number(item.pieces) <= 0) {
				throw new Error('Invalid piece conversion for ' + item.productName);
			}
		});

		return true;
	} catch (error) {
		logError('validateSale', error, sale);
		throw error;
	}
}

/**
 * Saves a sales entry.
 *
 * @param {Object} sale
 */
function addSale(sale) {
	try {
		validateSale(sale);

		validateSaleStock(sale);

		const salesId = generateSalesId();

		const headerSheet = getSheet(SHEETS.SALES_HEADER);

		headerSheet.appendRow([
			salesId,
			sale.date,
			sale.customerId,
			sale.customerName,
			sale.remarks || '',
		]);

		const detailSheet = getSheet(SHEETS.SALES_DETAILS);

		const rows = sale.products.map((item) => [
			salesId,
			item.productId,
			item.productName,
			item.unit,
			Number(item.quantity),
			Number(item.pieces),
		]);

		detailSheet
			.getRange(detailSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
			.setValues(rows);

		try {
			const inventoryData = buildSalesInventoryData(salesId, sale);

			appendLedgerRows(inventoryData.ledgerRows);

			updateCurrentStockBatch(inventoryData.stockChanges);
		} catch (error) {
			rollbackSale(salesId);

			throw error;
		}

		return {
			success: true,
			salesId,
		};
	} catch (error) {
		logError('addSale', error, sale);
		throw error;
	}
}

/**
 * Returns a sales entry by Sales ID.
 *
 * @param {string} salesId
 * @returns {Object|null}
 */
function getSale(salesId) {
	try {
		const headerSheet = getSheet(SHEETS.SALES_HEADER);
		const detailSheet = getSheet(SHEETS.SALES_DETAILS);

		const headers = headerSheet.getDataRange().getValues();
		const details = detailSheet.getDataRange().getValues();

		const header = headers.find(
			(row, index) => index > 0 && row[0] === salesId,
		);

		if (!header) {
			return null;
		}

		const products = details
			.filter((row, index) => index > 0 && row[0] === salesId)
			.map((row) => ({
				productId: row[1],
				productName: row[2],
				unit: row[3],
				quantity: Number(row[4]),
				pieces: Number(row[5]),
			}));

		return JSON.parse(
			JSON.stringify({
				id: header[0],
				date: formatSalesDateForInput(header[1]),
				customerId: header[2],
				customerName: header[3],
				remarks: header[4],
				products: products,
			}),
		);
	} catch (error) {
		logError('getSale', error, salesId);
		throw error;
	}
}

/**
 * Searches sales entries.
 *
 * @param {string} searchText
 * @returns {Array}
 */
function searchSales(criteria = {}) {
	try {
		const sheet = getSheet(SHEETS.SALES_HEADER);

		const data = sheet.getDataRange().getValues();

		if (typeof criteria === 'string') {
			criteria = { salesId: criteria };
		}

		const salesId = String(criteria.salesId || '').toLowerCase().trim();
		const customerId = String(criteria.customerId || '').trim();
		const fromDate = criteria.fromDate
			? getSalesDateOnly(criteria.fromDate)
			: null;
		const toDate = criteria.toDate ? getSalesDateOnly(criteria.toDate) : null;

		if (fromDate && toDate && fromDate > toDate) {
			throw new Error('Date From cannot be after Date To.');
		}

		const results = data
			.slice(1)
			.filter((row) => {
				const date = getSalesDateOnly(row[1]);
				return (
					(!salesId || String(row[0]).toLowerCase().includes(salesId)) &&
					(!customerId || String(row[2]) === customerId) &&
					(!fromDate || date >= fromDate) &&
					(!toDate || date <= toDate)
				);
			})
			.map((row) => ({
				id: row[0],
				date: formatSalesDateForInput(row[1]),
				customerId: row[2],
				customerName: row[3],
				remarks: row[4],
			}));

		return JSON.parse(JSON.stringify(results));
	} catch (error) {
		logError('searchSales', error, criteria);
		throw error;
	}
}

function getSalesDateOnly(value) {
	const date = new Date(value);
	date.setHours(0, 0, 0, 0);
	return date;
}

function formatSalesDateForInput(value) {
	return Utilities.formatDate(
		new Date(value),
		Session.getScriptTimeZone(),
		'yyyy-MM-dd',
	);
}

/**
 * Updates a sales entry.
 *
 * @param {String} salesId
 * @param {Object} sale
 */
function updateSale(salesId, sale) {
	try {
		const oldSale = getSale(salesId);

		if (!oldSale) {
			throw new Error('Sales entry not found.');
		}

		validateSale(sale);

		validateSaleStock(sale, salesId);

		reverseTransaction(salesId);

		try {
			updateSaleWithoutInventory(salesId, sale);

			const inventoryData = buildSalesInventoryData(salesId, sale);

			appendLedgerRows(inventoryData.ledgerRows);

			updateCurrentStockBatch(inventoryData.stockChanges);

			return true;
		} catch (error) {
			restoreSale(salesId, oldSale);

			throw error;
		}
	} catch (error) {
		logError('updateSale', error, sale);

		throw error;
	}
}

/**
 * Deletes a sales entry.
 *
 * @param {String} salesId
 */
function deleteSale(salesId) {
	try {
		const oldSale = getSale(salesId);

		if (!oldSale) {
			throw new Error('Sales entry not found.');
		}

		reverseTransaction(salesId);

		try {
			rollbackSale(salesId);

			return true;
		} catch (error) {
			restoreSale(salesId, oldSale);

			throw error;
		}
	} catch (error) {
		logError('deleteSale', error, salesId);

		throw error;
	}
}

/**
 * Validates stock availability.
 *
 * @param {Object} sale
 * @param {String} [salesId]
 */
function validateSaleStock(sale, salesId = '') {
	try {
		const stock = getCurrentStockMap();

		// If editing an existing sale, temporarily restore its stock
		if (salesId) {
			const existingSale = getSale(salesId);

			if (!existingSale) {
				throw new Error('Existing sale not found.');
			}

			existingSale.products.forEach((item) => {
				if (!stock[item.productId]) {
					stock[item.productId] = 0;
				}

				stock[item.productId] += Number(item.pieces);
			});
		}

		sale.products.forEach((item) => {
			const available = Number(stock[item.productId] || 0);

			const required = Number(item.pieces);

			if (available < required) {
				throw new Error(
					`Insufficient stock for ${item.productName}. Available: ${available}, Required: ${required}`,
				);
			}
		});

		return true;
	} catch (error) {
		logError('validateSaleStock', error, {
			salesId,
			sale,
		});

		throw error;
	}
}

/**
 * Builds inventory data for a sales entry.
 *
 * @param {String} salesId
 * @param {Object} sale
 */
function buildSalesInventoryData(salesId, sale) {
	const ledgerRows = [];

	const stockChanges = {};

	sale.products.forEach((item) => {
		ledgerRows.push({
			date: sale.date,
			transactionType: 'Sales',
			referenceId: salesId,
			productId: item.productId,
			productName: item.productName,
			qtyIn: 0,
			qtyOut: Number(item.pieces),
			status: 'Active',
			remarks: sale.remarks || '',
		});

		stockChanges[item.productId] = {
			productName: item.productName,
			quantity: -Number(item.pieces),
		};
	});

	return {
		ledgerRows,
		stockChanges,
	};
}

/**
 * Removes a newly created sale.
 * Used when addSale() fails after writing data.
 */
function rollbackSale(salesId) {
	try {
		const headerSheet = getSheet(SHEETS.SALES_HEADER);
		const detailSheet = getSheet(SHEETS.SALES_DETAILS);

		// Delete Details
		const detailData = detailSheet.getDataRange().getValues();

		for (let i = detailData.length - 1; i >= 1; i--) {
			if (detailData[i][0] === salesId) {
				detailSheet.deleteRow(i + 1);
			}
		}

		// Delete Header
		const headerData = headerSheet.getDataRange().getValues();

		for (let i = headerData.length - 1; i >= 1; i--) {
			if (headerData[i][0] === salesId) {
				headerSheet.deleteRow(i + 1);
				break;
			}
		}
	} catch (error) {
		logError('rollbackSale', error, salesId);
	}
}

/**
 * Restores an existing sale.
 *
 * Used if update fails after reversing inventory.
 */
function restoreSale(salesId, oldSale) {
	try {
		updateSaleWithoutInventory(salesId, oldSale);

		const inventoryData = buildSalesInventoryData(salesId, oldSale);

		appendLedgerRows(inventoryData.ledgerRows);

		updateCurrentStockBatch(inventoryData.stockChanges);
	} catch (error) {
		logError('restoreSale', error, salesId);

		throw error;
	}
}

/**
 * Updates Sales Header and Sales Details only.
 * No inventory updates are performed.
 *
 * @param {string} salesId
 * @param {Object} sale
 */
function updateSaleWithoutInventory(salesId, sale) {
	try {
		const headerSheet = getSheet(SHEETS.SALES_HEADER);
		const detailSheet = getSheet(SHEETS.SALES_DETAILS);

		const headerData = headerSheet.getDataRange().getValues();

		let headerRow = -1;

		for (let i = 1; i < headerData.length; i++) {
			if (headerData[i][0] === salesId) {
				headerRow = i + 1;
				break;
			}
		}

		if (headerRow === -1) {
			throw new Error('Sales entry not found.');
		}

		headerSheet
			.getRange(headerRow, 2, 1, 4)
			.setValues([
				[sale.date, sale.customerId, sale.customerName, sale.remarks || ''],
			]);

		const detailData = detailSheet.getDataRange().getValues();

		for (let i = detailData.length - 1; i >= 1; i--) {
			if (detailData[i][0] === salesId) {
				detailSheet.deleteRow(i + 1);
			}
		}

		const rows = sale.products.map((item) => [
			salesId,
			item.productId,
			item.productName,
			item.unit,
			Number(item.quantity),
			Number(item.pieces),
		]);

		if (rows.length) {
			detailSheet
				.getRange(detailSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
				.setValues(rows);
		}

		return true;
	} catch (error) {
		logError('updateSaleWithoutInventory', error, sale);

		throw error;
	}
}

function testAddSale() {
	Logger.log(
		addSale({
			date: '2026-07-25',

			customerId: 'CUST0001',

			customerName: 'ABC Traders',

			remarks: 'Test Sale',

			products: [
				{
					productId: 'PRD0001',

					productName: 'Product A',

					unit: 'Box',

					quantity: 2,

					pieces: 120,
				},
				{
					productId: 'PRD0002',

					productName: 'Product B',

					unit: 'Box',

					quantity: 1,

					pieces: 60,
				},
			],
		}),
	);
}

function testGetSale() {
	const salesId = 'SAL0001';

	Logger.log(JSON.stringify(getSale(salesId), null, 2));
}

function testSearchSalesAll() {
	Logger.log(JSON.stringify(searchSales(), null, 2));
}

function testSearchSalesByCustomer() {
	Logger.log(JSON.stringify(searchSales('ABC'), null, 2));
}

function testUpdateSale() {
	Logger.log(
		updateSale('SAL0001', {
			date: '2026-07-25',

			customerId: 'CUST0001',

			customerName: 'ABC Traders',

			remarks: 'Updated Sale',

			products: [
				{
					productId: 'PRD0001',

					productName: 'Product A',

					unit: 'Box',

					quantity: 3,

					pieces: 180,
				},

				{
					productId: 'PRD0002',

					productName: 'Product B',

					unit: 'Box',

					quantity: 2,

					pieces: 120,
				},
			],
		}),
	);
}

function testDeleteSale() {
	Logger.log(deleteSale('SAL0001'));
}

function testCurrentStockMap() {
	Logger.log(JSON.stringify(getCurrentStockMap(), null, 2));
}

function testValidateSaleStock() {
	const sale = getSale('SAL0001');
	Logger.log(validateSaleStock(sale));
}

function testBuildSalesInventoryData() {
	const sale = getSale('SAL0001');
	Logger.log(JSON.stringify(buildSalesInventoryData('SAL0001', sale), null, 2));
}
