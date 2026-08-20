/**
 * ===========================================================
 * Database.gs
 * ===========================================================
 */

function getSheet(sheetName) {
	return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Returns JSON object required by sidebar
 */
function getProductionFormData() {
	return {
		machines: getMachines(),

		operators: getOperators(),
	};
}

/**
 * Returns all active products for Sales UI.
 */
function getActiveProducts() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
		SHEETS.PRODUCT_MASTER,
	);

	const values = sheet.getDataRange().getValues();

	const headers = values.shift();
	const columns = getProductColumnMap(headers);

	return values
		.filter(
			(row) => row[columns.active] === true || row[columns.active] === 'TRUE',
		)
		.map((row) => {
			const productType = row[columns.productType] || 'Standard';
			const isBundle = String(productType).toLowerCase() === 'bundle';
			return {
				id: row[columns.id],
				name: row[columns.name],
				defaultMould: row[columns.defaultMould],
				// A bundle's selling unit is always one packed box.
				piecesPerPacket: isBundle
					? 1
					: Number(row[columns.piecesPerPacket]) || 1,
				packetsPerBox: isBundle ? 1 : Number(row[columns.packetsPerBox]) || 1,
				category: row[columns.category],
				requiresPainting: isPaintingRequired(row[columns.requiresPainting]),
				productType,
			};
		});
}

function getSalesFormData() {
	return {
		customers: getActiveCustomers(),

		products: getActiveProducts(),
	};
}

/** Product-only payload for inventory forms; avoids loading unrelated customers. */
function getInventoryProductFormData() {
	return {
		products: getActiveProducts(),
	};
}
