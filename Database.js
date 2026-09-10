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

	values.shift(); // Remove header

	return values
		.filter((row) => row[6] === true || row[6] === 'TRUE')
		.map((row) => ({
			id: row[0],
			name: row[1],
			defaultMould: row[2],
			piecesPerPacket: Number(row[3]) || 0,
			packetsPerBox: Number(row[4]) || 0,
			category: row[5],
		}));
}

function getSalesFormData() {
	return {
		customers: getActiveCustomers(),

		products: getActiveProducts(),

		stock: getCurrentStockMap(),
	};
}
