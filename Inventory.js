/**
 * =====================================================
 * INVENTORY ENGINE
 * =====================================================
 */

const INVENTORY_TRANSACTION_TYPES = {
	PRODUCTION: 'Production',
	SALE: 'Sales',
	OPENING_STOCK: 'Opening Stock',
	ADJUSTMENT: 'Adjustment',
	ALLOCATION: 'Stage Allocation',
	PAINTING: 'Painting',
	PACKING: 'Packing',
};

const INVENTORY_STAGES = {
	LOOSE: 'Loose',
	PAINTED: 'Painted',
	PACKED: 'Packed',
	UNALLOCATED: 'Unallocated',
};

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
		'Status',
		'Remarks',
	]);

	createSheetIfMissing(SHEETS.CURRENT_STOCK, [
		'Product ID',
		'Product Name',
		'Available Qty',
	]);

	ensureCurrentStockStageColumns();
}

/** Initializes stage columns in CurrentStock and the bundle definition sheet. */
function initializeStagedInventory() {
	ensureLedgerStageColumns();
	ensureCurrentStockStageColumns();
	createSheetIfMissing(SHEETS.BUNDLE_BOM, [
		'Bundle ID',
		'Bundle Name',
		'Component ID',
		'Component Name',
		'Qty Per Bundle',
	]);
	ensureProductMasterColumns();
}

function ensureCurrentStockStageColumns() {
	const sheet = getSheet(SHEETS.CURRENT_STOCK);
	if (!sheet) return;
	let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	['Loose Qty', 'Painted Qty', 'Packed Qty'].forEach((header) => {
		if (headers.indexOf(header) !== -1) return;
		sheet.getRange(1, headers.length + 1).setValue(header);
		headers.push(header);
	});
	if (headers.indexOf('Unallocated Qty') === -1) {
		const column = headers.length + 1;
		sheet.getRange(1, column).setValue('Unallocated Qty');
		// Existing Available Qty is deliberately left unclassified until the user
		// allocates it to Loose, Painted, or Packed through Inventory Movements.
		if (sheet.getLastRow() > 1) {
			const existingTotals = sheet
				.getRange(2, 3, sheet.getLastRow() - 1, 1)
				.getValues();
			sheet
				.getRange(2, column, existingTotals.length, 1)
				.setValues(existingTotals);
		}
	}
}

function ensureLedgerStageColumns() {
	const sheet = getSheet(SHEETS.STOCK_LEDGER);
	if (!sheet) return;
	const headers = sheet
		.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
		.getValues()[0];
	if (headers.indexOf('From Stage') === -1)
		sheet.getRange(1, headers.length + 1).setValue('From Stage');
	const refreshedHeaders = sheet
		.getRange(1, 1, 1, sheet.getLastColumn())
		.getValues()[0];
	if (refreshedHeaders.indexOf('To Stage') === -1)
		sheet.getRange(1, refreshedHeaders.length + 1).setValue('To Stage');
}

function ensureProductMasterColumns() {
	const sheet = getSheet(SHEETS.PRODUCT_MASTER);
	if (!sheet || sheet.getLastRow() < 1) return;
	let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	if (headers.indexOf('Requires Painting') === -1) {
		sheet.getRange(1, headers.length + 1).setValue('Requires Painting');
		headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
		if (sheet.getLastRow() > 1)
			sheet
				.getRange(2, headers.length, sheet.getLastRow() - 1, 1)
				.setValue('Yes');
	}
	if (headers.indexOf('Product Type') === -1) {
		sheet.getRange(1, headers.length + 1).setValue('Product Type');
		if (sheet.getLastRow() > 1)
			sheet
				.getRange(2, headers.length + 1, sheet.getLastRow() - 1, 1)
				.setValue('Standard');
	}
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
			toStage: INVENTORY_STAGES.LOOSE,
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
	validateStockAdjustment(adjustment);

	const quantity = Number(adjustment.quantity);

	const qtyIn = quantity > 0 ? quantity : 0;
	const qtyOut = quantity < 0 ? Math.abs(quantity) : 0;

	const ledgerRows = [
		{
			date: adjustment.date,
			transactionType: INVENTORY_TRANSACTION_TYPES.ADJUSTMENT,
			referenceId: null,
			productId: adjustment.productId,
			productName: adjustment.productName,
			qtyIn,
			qtyOut,
			toStage: adjustment.stage || INVENTORY_STAGES.PACKED,
			remarks: adjustment.remarks || '',
		},
	];

	const stockChanges = {
		[adjustment.productId]: {
			productName: adjustment.productName,
			quantity,
		},
	};

	try {
		appendLedgerRows(ledgerRows);

		updateCurrentStockBatch(stockChanges);
		updateStageStockBatch([
			{
				productId: adjustment.productId,
				productName: adjustment.productName,
				stage: adjustment.stage || INVENTORY_STAGES.PACKED,
				quantity,
			},
		]);

		return JSON.parse(
			JSON.stringify({
				success: true,
			}),
		);
	} catch (error) {
		rollbackStockChanges(stockChanges);

		throw error;
	}
}

function validateStockAdjustment(adjustment) {
	if (!adjustment) {
		throw new Error('Adjustment data is required.');
	}

	if (!adjustment.productId) {
		throw new Error('Please select a product.');
	}

	const qty = Number(adjustment.quantity);

	if (!qty) {
		throw new Error('Quantity cannot be zero.');
	}

	if (qty < 0) {
		const stock = getStageStockMap(adjustment.stage || INVENTORY_STAGES.PACKED);

		const available = Number(stock[adjustment.productId] || 0);

		if (available < Math.abs(qty)) {
			throw new Error(`Insufficient stock. Available : ${available}`);
		}
	}
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
		const stageChanges = [];
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
			const fromStage = row[10] || '';
			const toStage = row[11] || '';

			reversalRows.push({
				date: today,
				transactionType: `${transactionType} Reversal`,
				referenceId,
				productId,
				productName,
				qtyIn: qtyOut,
				qtyOut: qtyIn,
				fromStage: toStage,
				toStage: fromStage,
				status: 'Active',
				remarks,
			});

			stockChanges.push({
				productId,
				productName,
				quantity: qtyOut - qtyIn,
			});

			rowsToUpdate.push(i + 1);

			if (fromStage && qtyOut)
				stageChanges.push({
					productId,
					productName,
					stage: fromStage,
					quantity: qtyOut,
				});
			if (toStage && qtyIn)
				stageChanges.push({
					productId,
					productName,
					stage: toStage,
					quantity: -qtyIn,
				});
		}

		if (reversalRows.length === 0) {
			throw new Error('No active inventory transaction found.');
		}

		// Do not reverse a production or packing record once its output has
		// already moved to a later stage or been sold.
		validateStageAvailability(stageChanges);

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

			updateStageStockBatch(stageChanges);

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
			row.fromStage || '',
			row.toStage || '',
		];
	});

	sheet
		.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length)
		.setValues(values);
}

/** Returns product-master column positions while remaining compatible with old sheets. */
function getProductColumnMap(headers) {
	const column = (name, fallback) => {
		const index = headers.indexOf(name);
		return index === -1 ? fallback : index;
	};
	return {
		id: column('Product ID', 0),
		name: column('Product Name', 1),
		defaultMould: column('Default Mould', 2),
		piecesPerPacket: column('Pieces Per Packet', 3),
		packetsPerBox: column('Packets Per Box', 4),
		category: column('Category', 5),
		active: column('Active', 6),
		requiresPainting: column('Requires Painting', -1),
		productType: column('Product Type', -1),
	};
}

function isPaintingRequired(value) {
	if (value === undefined || value === null || value === '') return true;
	return !['no', 'false', '0'].includes(String(value).trim().toLowerCase());
}

function getProductCatalog() {
	const sheet = getSheet(SHEETS.PRODUCT_MASTER);
	if (!sheet || sheet.getLastRow() <= 1) return [];
	const values = sheet.getDataRange().getValues();
	const columns = getProductColumnMap(values.shift());
	return values
		.filter(
			(row) =>
				row[columns.active] === true ||
				String(row[columns.active]).toUpperCase() === 'TRUE',
		)
		.map((row) => ({
			id: row[columns.id],
			name: row[columns.name],
			requiresPainting: isPaintingRequired(row[columns.requiresPainting]),
			productType: String(row[columns.productType] || 'Standard'),
			packetsPerBox: Number(row[columns.packetsPerBox] || 1),
			piecesPerBox: Math.max(
				1,
				Number(row[columns.piecesPerPacket] || 1) *
					Number(row[columns.packetsPerBox] || 1),
			),
		}));
}

function getProductInfo(productId) {
	const product = getProductCatalog().find((item) => item.id === productId);
	if (!product) throw new Error('Product not found or inactive: ' + productId);
	return product;
}

function getBundleComponents(bundleId) {
	const sheet = getSheet(SHEETS.BUNDLE_BOM);
	if (!sheet || sheet.getLastRow() <= 1) return [];
	return sheet
		.getRange(2, 1, sheet.getLastRow() - 1, 5)
		.getValues()
		.filter((row) => row[0] === bundleId)
		.map((row) => ({
			productId: row[2],
			productName: row[3],
			quantity: Number(row[4]),
		}));
}

function getBundleBomFormData() {
	const products = getProductCatalog();
	return {
		bundles: products.filter(
			(product) => product.productType.toLowerCase() === 'bundle',
		),
		components: products.filter(
			(product) => product.productType.toLowerCase() !== 'bundle',
		),
		bomByBundle: getBundleComponentsByBundle(),
	};
}

/** Replaces one bundle's recipe after validating all component products. */
function saveBundleBom(data) {
	if (!data || !data.bundleId) throw new Error('Select a bundle SKU.');
	if (!Array.isArray(data.components) || data.components.length === 0) {
		throw new Error('Add at least one component product.');
	}

	return withInventoryLock(() => {
		const products = getProductCatalog();
		const byId = {};
		products.forEach((product) => {
			byId[product.id] = product;
		});
		const bundle = byId[data.bundleId];
		if (!bundle || bundle.productType.toLowerCase() !== 'bundle') {
			throw new Error('The selected product is not an active bundle SKU.');
		}

		const componentIds = new Set();
		const components = data.components.map((component) => {
			const product = byId[component.productId];
			if (!product || product.productType.toLowerCase() === 'bundle') {
				throw new Error('Every component must be an active standard product.');
			}
			if (componentIds.has(product.id)) {
				throw new Error('A component can be included only once in a bundle.');
			}
			if (Number(component.quantity) <= 0) {
				throw new Error('Component quantity must be greater than zero.');
			}
			componentIds.add(product.id);
			return [
				bundle.id,
				bundle.name,
				product.id,
				product.name,
				Number(component.quantity),
			];
		});

		const sheet = getSheet(SHEETS.BUNDLE_BOM);
		const existing =
			sheet.getLastRow() > 1
				? sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
				: [];
		const retained = existing.filter((row) => row[0] !== bundle.id);
		const rows = retained.concat(components);

		if (sheet.getLastRow() > 1) {
			sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
		}
		if (rows.length) sheet.getRange(2, 1, rows.length, 5).setValues(rows);

		return { success: true, bundleId: bundle.id };
	});
}

function getCurrentStockColumnMap(headers) {
	const column = (name, fallback) => {
		const index = headers.indexOf(name);
		return index === -1 ? fallback : index;
	};
	return {
		productId: column('Product ID', 0),
		productName: column('Product Name', 1),
		available: column('Available Qty', 2),
		loose: column('Loose Qty', -1),
		painted: column('Painted Qty', -1),
		packed: column('Packed Qty', -1),
		unallocated: column('Unallocated Qty', -1),
	};
}

function getStageStockColumn(stage, columns) {
	return {
		[INVENTORY_STAGES.LOOSE]: columns.loose,
		[INVENTORY_STAGES.PAINTED]: columns.painted,
		[INVENTORY_STAGES.PACKED]: columns.packed,
		[INVENTORY_STAGES.UNALLOCATED]: columns.unallocated,
	}[stage];
}

function getStageStockMap(stage) {
	const sheet = getSheet(SHEETS.CURRENT_STOCK);
	const result = {};
	if (!sheet || sheet.getLastRow() <= 1) return result;
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const columns = getCurrentStockColumnMap(headers);
	const stageColumn = getStageStockColumn(stage, columns);
	if (stageColumn === -1) return result;
	sheet
		.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
		.getValues()
		.forEach((row) => {
			if (row[columns.productId])
				result[row[columns.productId]] = Number(row[stageColumn] || 0);
		});
	return result;
}

/** Reads CurrentStock stage columns once, indexed as { productId: { Loose, Painted, Packed } }. */
function getStageStockIndex() {
	const sheet = getSheet(SHEETS.CURRENT_STOCK);
	const index = {};
	if (!sheet || sheet.getLastRow() <= 1) return index;
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const columns = getCurrentStockColumnMap(headers);
	sheet
		.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
		.getValues()
		.forEach((row) => {
			if (!row[columns.productId]) return;
			index[row[columns.productId]] = {
				[INVENTORY_STAGES.LOOSE]: Number(row[columns.loose] || 0),
				[INVENTORY_STAGES.PAINTED]: Number(row[columns.painted] || 0),
				[INVENTORY_STAGES.PACKED]: Number(row[columns.packed] || 0),
				[INVENTORY_STAGES.UNALLOCATED]: Number(row[columns.unallocated] || 0),
			};
		});
	return index;
}

function getBundleComponentsByBundle() {
	const sheet = getSheet(SHEETS.BUNDLE_BOM);
	const bundles = {};
	if (!sheet || sheet.getLastRow() <= 1) return bundles;
	sheet
		.getRange(2, 1, sheet.getLastRow() - 1, 5)
		.getValues()
		.forEach((row) => {
			if (!row[0] || !row[2]) return;
			if (!bundles[row[0]]) bundles[row[0]] = [];
			bundles[row[0]].push({
				productId: row[2],
				productName: row[3],
				quantity: Number(row[4]),
			});
		});
	return bundles;
}

function getAvailableStageStock(productId, stage) {
	return Number(getStageStockMap(stage)[productId] || 0);
}

/** Applies an array of changes: {productId, productName, stage, quantity}. */
function updateStageStockBatch(changes) {
	if (!changes || !changes.length) return;
	const sheet = getSheet(SHEETS.CURRENT_STOCK);
	const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
	const columns = getCurrentStockColumnMap(headers);
	const lastRow = sheet.getLastRow();
	const existing =
		lastRow > 1
			? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
			: [];
	const index = {};
	existing.forEach((row, rowIndex) => {
		index[row[columns.productId]] = rowIndex;
	});
	const newRows = [];
	changes.forEach((change) => {
		const stageColumn = getStageStockColumn(change.stage, columns);
		if (stageColumn === -1)
			throw new Error('CurrentStock stage columns are missing.');
		if (Object.prototype.hasOwnProperty.call(index, change.productId)) {
			const row = existing[index[change.productId]];
			row[stageColumn] =
				Number(row[stageColumn] || 0) + Number(change.quantity);
		} else {
			const row = Array(headers.length).fill('');
			row[columns.productId] = change.productId;
			row[columns.productName] = change.productName;
			row[columns.available] = 0;
			row[stageColumn] = Number(change.quantity);
			index[change.productId] = existing.length + newRows.length;
			newRows.push(row);
		}
	});
	if (existing.length)
		sheet
			.getRange(2, 1, existing.length, sheet.getLastColumn())
			.setValues(existing);
	if (newRows.length)
		sheet
			.getRange(
				sheet.getLastRow() + 1,
				1,
				newRows.length,
				sheet.getLastColumn(),
			)
			.setValues(newRows);
}

function validateStageAvailability(changes) {
	const staged = {};
	changes.forEach((change) => {
		const key = change.productId + '|' + change.stage;
		const current = Object.prototype.hasOwnProperty.call(staged, key)
			? staged[key]
			: getAvailableStageStock(change.productId, change.stage);
		staged[key] = current + Number(change.quantity);
		if (staged[key] < 0)
			throw new Error(
				`Insufficient ${change.stage} stock for ${change.productName}. Available: ${getAvailableStageStock(change.productId, change.stage)}`,
			);
	});
}

/** Moves a known historical balance from Unallocated into a real stock stage. */
function postStageAllocation(data) {
	if (!data || !data.productId || Number(data.quantity) <= 0) {
		throw new Error('Product and positive quantity are required.');
	}
	if (
		![
			INVENTORY_STAGES.LOOSE,
			INVENTORY_STAGES.PAINTED,
			INVENTORY_STAGES.PACKED,
		].includes(data.stage)
	) {
		throw new Error('Select Loose, Painted, or Packed as the target stage.');
	}
	return withInventoryLock(() => {
		const product = getProductInfo(data.productId);
		const quantity = Number(data.quantity);
		const changes = [
			{
				productId: product.id,
				productName: product.name,
				stage: INVENTORY_STAGES.UNALLOCATED,
				quantity: -quantity,
			},
			{
				productId: product.id,
				productName: product.name,
				stage: data.stage,
				quantity,
			},
		];
		validateStageAvailability(changes);
		const referenceId = 'ALG-' + Utilities.getUuid().slice(0, 8).toUpperCase();
		appendLedgerRows([
			{
				date: data.date || new Date(),
				transactionType: INVENTORY_TRANSACTION_TYPES.ALLOCATION,
				referenceId,
				productId: product.id,
				productName: product.name,
				qtyIn: quantity,
				qtyOut: quantity,
				fromStage: INVENTORY_STAGES.UNALLOCATED,
				toStage: data.stage,
				remarks: data.remarks || 'Historical stock allocation',
			},
		]);
		updateStageStockBatch(changes);
		return { success: true, referenceId };
	});
}

function postPainting(data) {
	return withInventoryLock(() => postPaintingUnsafe(data));
}

function postPaintingUnsafe(data) {
	if (!data || !data.productId || Number(data.quantity) <= 0)
		throw new Error('Product and positive quantity are required.');
	const product = getProductInfo(data.productId);
	if (product.productType.toLowerCase() === 'bundle')
		throw new Error('Bundles cannot be painted.');
	if (!product.requiresPainting)
		throw new Error(product.name + ' is configured for direct packing.');
	const changes = [
		{
			productId: product.id,
			productName: product.name,
			stage: INVENTORY_STAGES.LOOSE,
			quantity: -Number(data.quantity),
		},
		{
			productId: product.id,
			productName: product.name,
			stage: INVENTORY_STAGES.PAINTED,
			quantity: Number(data.quantity),
		},
	];
	validateStageAvailability(changes);
	const referenceId = 'PNT-' + Utilities.getUuid().slice(0, 8).toUpperCase();
	appendLedgerRows([
		{
			date: data.date || new Date(),
			transactionType: INVENTORY_TRANSACTION_TYPES.PAINTING,
			referenceId,
			productId: product.id,
			productName: product.name,
			qtyIn: Number(data.quantity),
			qtyOut: Number(data.quantity),
			fromStage: INVENTORY_STAGES.LOOSE,
			toStage: INVENTORY_STAGES.PAINTED,
			remarks: data.remarks || '',
		},
	]);
	updateStageStockBatch(changes);
	return { success: true, referenceId };
}

function postPacking(data) {
	return withInventoryLock(() => postPackingUnsafe(data));
}

function postPackingUnsafe(data) {
	if (!data || !data.productId || Number(data.boxes) <= 0)
		throw new Error('Product and positive box quantity are required.');
	const product = getProductInfo(data.productId);
	const boxes = Number(data.boxes);
	const referenceId = 'PKG-' + Utilities.getUuid().slice(0, 8).toUpperCase();
	const components =
		product.productType.toLowerCase() === 'bundle'
			? getBundleComponents(product.id)
			: [
					{
						productId: product.id,
						productName: product.name,
						quantity: product.piecesPerBox,
					},
				];
	if (!components.length)
		throw new Error('No bundle BOM is configured for ' + product.name + '.');
	const changes = [];
	const ledgerRows = [];
	components.forEach((component) => {
		if (!component.productId || Number(component.quantity) <= 0)
			throw new Error(
				'Every bundle BOM row must have a product and positive quantity.',
			);
		const componentProduct = getProductInfo(component.productId);
		const sourceStage = componentProduct.requiresPainting
			? INVENTORY_STAGES.PAINTED
			: INVENTORY_STAGES.LOOSE;
		const bundlesPerBox = product.packetsPerBox;
		const required = Number(component.quantity) * bundlesPerBox * boxes;
		changes.push({
			productId: component.productId,
			productName: componentProduct.name,
			stage: sourceStage,
			quantity: -required,
		});
		ledgerRows.push({
			date: data.date || new Date(),
			transactionType: INVENTORY_TRANSACTION_TYPES.PACKING,
			referenceId,
			productId: component.productId,
			productName: componentProduct.name,
			qtyIn: 0,
			qtyOut: required,
			fromStage: sourceStage,
			remarks:
				'Packed into ' +
				product.name +
				(data.remarks ? ': ' + data.remarks : ''),
		});
	});
	// Packed stock is recorded in selling units. Standard products use pieces; bundle products use boxes.
	const packedQuantity =
		product.productType.toLowerCase() === 'bundle'
			? boxes
			: boxes * product.piecesPerBox;
	changes.push({
		productId: product.id,
		productName: product.name,
		stage: INVENTORY_STAGES.PACKED,
		quantity: packedQuantity,
	});
	ledgerRows.push({
		date: data.date || new Date(),
		transactionType: INVENTORY_TRANSACTION_TYPES.PACKING,
		referenceId,
		productId: product.id,
		productName: product.name,
		qtyIn: packedQuantity,
		qtyOut: 0,
		toStage: INVENTORY_STAGES.PACKED,
		remarks: data.remarks || '',
	});
	validateStageAvailability(changes);
	appendLedgerRows(ledgerRows);
	const aggregateChanges = {};
	ledgerRows.forEach((row) => {
		const quantity = Number(row.qtyIn || 0) - Number(row.qtyOut || 0);
		if (!aggregateChanges[row.productId])
			aggregateChanges[row.productId] = {
				productName: row.productName,
				quantity: 0,
			};
		aggregateChanges[row.productId].quantity += quantity;
	});
	updateCurrentStockBatch(aggregateChanges);
	updateStageStockBatch(changes);
	return { success: true, referenceId, packedQuantity };
}

function getPackingFormData() {
	const products = getProductCatalog();
	return {
		products,
		availability: getInventoryAvailability(
			products,
			getStageStockIndex(),
			getBundleComponentsByBundle(),
		),
	};
}

function withInventoryLock(callback) {
	const lock = LockService.getDocumentLock();
	lock.waitLock(30000);
	try {
		return callback();
	} finally {
		lock.releaseLock();
	}
}

function getInventoryAvailability(products, stageStock, bundleComponents) {
	const stageIndex = stageStock || getStageStockIndex();
	const bundleIndex = bundleComponents || getBundleComponentsByBundle();
	const productIndex = {};
	products.forEach((product) => {
		productIndex[product.id] = product;
	});
	const quantityAt = (productId, stage) =>
		Number((stageIndex[productId] || {})[stage] || 0);

	return products.map((product) => {
		const packedQuantity = quantityAt(product.id, INVENTORY_STAGES.PACKED);
		const packed =
			product.productType.toLowerCase() === 'bundle'
				? packedQuantity
				: Math.floor(packedQuantity / product.piecesPerBox);
		let packable = 0;
		if (product.productType.toLowerCase() === 'bundle') {
			const components = bundleIndex[product.id] || [];
			packable = components.length
				? Math.min(
						...components.map((component) => {
							const componentProduct = productIndex[component.productId];
							if (!componentProduct || Number(component.quantity) <= 0)
								return 0;
							const stage = componentProduct.requiresPainting
								? INVENTORY_STAGES.PAINTED
								: INVENTORY_STAGES.LOOSE;
							return Math.floor(
								quantityAt(component.productId, stage) / component.quantity,
							);
						}),
					)
				: 0;
		} else {
			const stage = product.requiresPainting
				? INVENTORY_STAGES.PAINTED
				: INVENTORY_STAGES.LOOSE;
			packable = Math.floor(
				quantityAt(product.id, stage) / product.piecesPerBox,
			);
		}
		return {
			productId: product.id,
			productName: product.name,
			productType: product.productType,
			packed,
			packedQuantity,
			canPackNow: Math.max(0, packable),
			canFulfilAfterPacking: packed + Math.max(0, packable),
		};
	});
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
