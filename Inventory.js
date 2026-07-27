/**
 * =====================================================
 * INVENTORY ENGINE
 * =====================================================
 */

function initializeInventory() {
	createSheetIfMissing(SHEETS.STOCK_LEDGER, [
		'Ledger ID',
		'Date',
		'Transaction Type',
		'Reference ID',
		'Product ID',
		'Product Name',
		'Qty In',
		'Qty Out',
		'Remarks',
	]);

	createSheetIfMissing(SHEETS.CURRENT_STOCK, [
		'Product ID',
		'Product Name',
		'Available Qty',
	]);
}

/**
 * Builds inventory data for a Production Entry.
 *
 * @param {string} entryId
 * @param {Object} production
 * @returns {Object}
 */
function buildProductionInventoryData(entryId, production) {
	if (!entryId) {
		throw new Error('Entry ID is required.');
	}

	if (!production || !production.products) {
		throw new Error('Invalid production data.');
	}

	const ledgerRows = [];
	const stockChanges = {};

	production.products.forEach((item) => {
		const quantity = Number(
			item.pieces ?? Number(item.mould) * Number(production.rounds),
		);

		ledgerRows.push({
			date: production.date,
			transactionType: 'Production',
			referenceId: entryId,
			productId: item.productId,
			productName: item.productName,
			qtyIn: quantity,
			qtyOut: 0,
			remarks: '',
		});

		stockChanges[item.productId] = {
			productName: item.productName,
			quantity: (stockChanges[item.productId]?.quantity || 0) + quantity,
		};
	});

	return {
		ledgerRows,
		stockChanges,
	};
}

function postSalesStock(sale) {
	// TODO
}

function postStockAdjustment(adjustment) {
	// TODO
}

/**
 * Reverses all inventory transactions for a Reference ID.
 *
 * @param {string} referenceId
 * @param {string} remarks
 */
function reverseTransaction(referenceId, remarks = 'Transaction Reversed') {
	if (!referenceId) {
		throw new Error('Reference ID is required.');
	}

	try {
		const sheet = getSheet(SHEETS.STOCK_LEDGER);
		const values = sheet.getDataRange().getValues();

		if (values.length <= 1) return;

		const reversalRows = [];
		const stockChanges = [];
		const rowsToUpdate = [];

		const today = new Date();

		for (let i = 1; i < values.length; i++) {
			const row = values[i];

			if (row[3] !== referenceId) continue;
			if (row[8] !== 'Active') continue;

			const transactionType = row[2];

			if (transactionType.endsWith('Reversal')) {
				continue;
			}

			const productId = row[4];
			const productName = row[5];
			const qtyIn = Number(row[6] || 0);
			const qtyOut = Number(row[7] || 0);

			reversalRows.push({
				date: today,
				transactionType: `${transactionType} Reversal`,
				referenceId,
				productId,
				productName,
				qtyIn: qtyOut,
				qtyOut: qtyIn,
				status: 'Active',
				remarks,
			});

			stockChanges.push({
				productId,
				productName,
				quantity: qtyOut - qtyIn,
			});

			rowsToUpdate.push(i + 1);
		}

		if (reversalRows.length === 0) {
			throw new Error('No active inventory transaction found.');
		}

		// Build stock batch

		const batch = {};

		stockChanges.forEach((item) => {
			if (!batch[item.productId]) {
				batch[item.productId] = {
					productName: item.productName,
					quantity: 0,
				};
			}

			batch[item.productId].quantity += item.quantity;
		});

		// Transactional posting

		try {
			appendLedgerRows(reversalRows);

			updateCurrentStockBatch(batch);

			rowsToUpdate.forEach((rowNumber) => {
				sheet.getRange(rowNumber, 9).setValue('Cancelled');
			});

			return true;
		} catch (error) {
			// Undo stock update if it already happened
			rollbackStockChanges(batch);

			// Remove reversal ledger rows
			rollbackReverseTransaction(referenceId);

			throw error;
		}
	} catch (error) {
		logError('reverseTransaction', error, {
			referenceId,
			remarks,
		});

		throw error;
	}
}

function getCurrentStock(productName) {
	// TODO
}

/**
 * Rebuilds Current Stock from the Stock Ledger.
 */
function rebuildCurrentStock() {
	const ledgerSheet = getSheet(SHEETS.STOCK_LEDGER);
	const stockSheet = getSheet(SHEETS.CURRENT_STOCK);

	const ledger = ledgerSheet.getDataRange().getValues();

	// Clear Current Stock (keep header)
	if (stockSheet.getLastRow() > 1) {
		stockSheet.getRange(2, 1, stockSheet.getLastRow() - 1, 3).clearContent();
	}

	if (ledger.length <= 1) {
		return;
	}

	const totals = {};

	for (let i = 1; i < ledger.length; i++) {
		const row = ledger[i];

		const productId = row[4];
		const productName = row[5];
		const status = row[8];

		if (status !== 'Active') {
			continue;
		}

		const qtyIn = Number(row[6] || 0);
		const qtyOut = Number(row[7] || 0);

		if (!totals[productId]) {
			totals[productId] = {
				productName: productName,
				quantity: 0,
			};
		}

		totals[productId].quantity += qtyIn - qtyOut;
	}

	updateCurrentStockBatch(totals);
}

/* ===========================
   Private Helpers
=========================== */

/**
 * Appends one or more inventory transactions to the Stock Ledger.
 *
 * @param {Object[]} rows
 */
function appendLedgerRows(rows) {
	if (!rows || rows.length === 0) return;

	const sheet = getSheet(SHEETS.STOCK_LEDGER);
	let sequence = getNextLedgerSequence(rows[0].date);

	const values = rows.map((row) => {
		const ledgerId = generateLedgerId(row.date, sequence++);

		return [
			ledgerId,
			row.date,
			row.transactionType,
			row.referenceId,
			row.productId,
			row.productName,
			Number(row.qtyIn || 0),
			Number(row.qtyOut || 0),
			row.status || 'Active',
			row.remarks || '',
		];
	});

	sheet
		.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length)
		.setValues(values);
}

/**
 * Updates the available stock for a product.
 *
 * @param {string} productId
 * @param {string} productName
 * @param {number} quantityChange
 */
function updateCurrentStock(productId, productName, quantityChange) {
	const sheet = getSheet(SHEETS.CURRENT_STOCK);

	const lastRow = sheet.getLastRow();

	// Sheet contains only headers
	if (lastRow < 2) {
		sheet.appendRow([productId, productName, Number(quantityChange)]);

		return;
	}

	const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

	for (let i = 0; i < data.length; i++) {
		if (data[i][0] === productId) {
			const row = i + 2;

			const newQty = Number(data[i][2]) + Number(quantityChange);

			sheet.getRange(row, 3).setValue(newQty);

			return;
		}
	}

	// Product not found
	sheet.appendRow([productId, productName, Number(quantityChange)]);
}

/**
 * Batch updates CurrentStock for multiple products.
 *
 * @param {Object} stockChanges
 *
 * Example:
 * {
 *   "PRD001": {
 *      productName: "Duck Tile",
 *      quantity: 540
 *   },
 *   "PRD002": {
 *      productName: "Fish Tile",
 *      quantity: 360
 *   }
 * }
 */
function updateCurrentStockBatch(stockChanges) {
	const sheet = getSheet(SHEETS.CURRENT_STOCK);

	const lastRow = sheet.getLastRow();

	// Empty sheet
	if (lastRow < 2) {
		const values = Object.entries(stockChanges).map(([productId, stock]) => [
			productId,

			stock.productName,

			Number(stock.quantity),
		]);

		if (values.length) {
			sheet.getRange(2, 1, values.length, 3).setValues(values);
		}

		return;
	}

	const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

	const productIndex = {};

	data.forEach((row, index) => {
		productIndex[row[0]] = index;
	});

	const newRows = [];

	Object.entries(stockChanges).forEach(([productId, stock]) => {
		if (productIndex.hasOwnProperty(productId)) {
			const rowIndex = productIndex[productId];

			data[rowIndex][2] = Number(data[rowIndex][2]) + Number(stock.quantity);
		} else {
			newRows.push([productId, stock.productName, Number(stock.quantity)]);
		}
	});

	// Update existing rows
	sheet.getRange(2, 1, data.length, 3).setValues(data);

	// Append new products
	if (newRows.length) {
		sheet
			.getRange(sheet.getLastRow() + 1, 1, newRows.length, 3)
			.setValues(newRows);
	}
}

/**
 * Returns the next sequence number for a given date.
 */
function getNextLedgerSequence(transactionDate) {
	const sheet = getSheet(SHEETS.STOCK_LEDGER);
	const values = sheet.getDataRange().getValues();

	const targetDate = Utilities.formatDate(
		new Date(transactionDate),
		Session.getScriptTimeZone(),
		'yyyyMMdd',
	);

	let maxSequence = 0;

	for (let i = 1; i < values.length; i++) {
		const ledgerId = values[i][0];

		if (!ledgerId || !ledgerId.startsWith(`LED-${targetDate}-`)) continue;

		const sequence = parseInt(ledgerId.split('-')[2], 10);

		if (sequence > maxSequence) {
			maxSequence = sequence;
		}
	}

	return maxSequence + 1;
}

/**
 * Generates Ledger ID
 *
 * Format:
 * LED-YYYYMMDD-0001
 */
function generateLedgerId(transactionDate, sequence) {
	const datePart = Utilities.formatDate(
		new Date(transactionDate),
		Session.getScriptTimeZone(),
		'yyyyMMdd',
	);

	return `LED-${datePart}-${String(sequence).padStart(4, '0')}`;
}

/**
 * Returns current stock indexed by Product ID.
 */
function getCurrentStockMap() {
	try {
		const sheet = getSheet(SHEETS.CURRENT_STOCK);

		const data = sheet.getDataRange().getValues();

		const stock = {};

		for (let i = 1; i < data.length; i++) {
			stock[data[i][0]] = Number(data[i][2] || 0);
		}

		return stock;
	} catch (error) {
		logError('getCurrentStockMap', error);

		throw error;
	}
}

/**
 * Removes reversal ledger rows for a transaction.
 *
 * @param {String} referenceId
 */
function rollbackReverseTransaction(referenceId) {
	try {
		const sheet = getSheet(SHEETS.STOCK_LEDGER);

		const data = sheet.getDataRange().getValues();

		for (let i = data.length - 1; i >= 1; i--) {
			const transactionType = String(data[i][2]);

			if (data[i][3] === referenceId && transactionType.endsWith('Reversal')) {
				sheet.deleteRow(i + 1);
			}
		}
	} catch (error) {
		logError('rollbackReverseTransaction', error, referenceId);
	}
}

/**
 * Reverts stock changes.
 *
 * @param {Object} stockChanges
 */
function rollbackStockChanges(stockChanges) {
	try {
		const reverseChanges = {};

		Object.keys(stockChanges).forEach((productId) => {
			reverseChanges[productId] = {
				productName: stockChanges[productId].productName,

				quantity: -stockChanges[productId].quantity,
			};
		});

		updateCurrentStockBatch(reverseChanges);
	} catch (error) {
		logError('rollbackStockChanges', error, stockChanges);
	}
}

/**
 * Testing
 */
function testAppendLedger() {
	appendLedgerRows([
		{
			date: new Date(),
			transactionType: 'Production',
			referenceId: 'PRD-TEST-001',
			product: 'Test Product A',
			qtyIn: 100,
			qtyOut: 0,
			remarks: 'Testing',
		},
		{
			date: new Date(),
			transactionType: 'Production',
			referenceId: 'PRD-TEST-001',
			product: 'Test Product B',
			qtyIn: 50,
			qtyOut: 0,
			remarks: 'Testing',
		},
	]);
}

function testCurrentStock() {
	updateCurrentStock('Test Product A', 100);
	updateCurrentStock('Test Product B', 50);

	updateCurrentStock('Test Product A', 25);

	updateCurrentStock('Test Product B', -10);
}

function testPostProductionStock() {
	postProductionStock({
		header: {
			entryId: 'PRD-TEST-001',

			date: new Date(),
		},

		details: [
			{
				product: 'Duck',

				totalPieces: 500,
			},

			{
				product: 'Fish',

				totalPieces: 300,
			},
		],
	});
}

/**
 * One-time migration of historical Production data into Inventory.
 */
function migrateProductionToInventory() {
	const ledgerSheet = getSheet(SHEETS.STOCK_LEDGER);
	const stockSheet = getSheet(SHEETS.CURRENT_STOCK);

	if (ledgerSheet.getLastRow() > 1 || stockSheet.getLastRow() > 1) {
		throw new Error(
			'Inventory already contains data. Clear StockLedger and CurrentStock before running migration.',
		);
	}

	const productionHeaders = getSheet(SHEETS.DAILY_HEADER)
		.getDataRange()
		.getValues();

	if (productionHeaders.length <= 1) {
		return;
	}

	for (let i = 1; i < productionHeaders.length; i++) {
		const entryId = productionHeaders[i][0];

		const production = getProductionEntry(entryId);
		const productionData = {
			date: production.header.date,
			machineId: production.header.machineId,
			machineName: production.header.machineName,
			operatorId: production.header.operatorId,
			operatorName: production.header.operatorName,
			bags: production.header.bags,
			rounds: production.header.rounds,
			products: production.details.map((item) => ({
				productId: item.productId,
				productName: item.productName,
				mould: item.mould,
				pieces: Number(item.pieces),
			})),
		};

		const inventory = buildProductionInventoryData(entryId, productionData);

		appendLedgerRows(inventory.ledgerRows);

		updateCurrentStockBatch(inventory.stockChanges);
	}
}
