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
