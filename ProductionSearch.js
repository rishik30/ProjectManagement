/**
 * ===========================================================
 * ProductionSearch.gs
 * Phase 3 - Sprint 3.1B
 * ===========================================================
 */

/**
 * Searches production entries.
 *
 * Supported filters:
 * - entryId
 * - date
 * - machineId
 * - operatorId
 *
 * Returns matching header records only.
 */
function searchProductionEntries(criteria) {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	const values = sheet.getDataRange().getValues();

	const headers = values.shift();

	const columnMap = buildHeaderColumnMap(headers);

	const results = values.filter((row) =>
		matchesSearchCriteria(row, columnMap, criteria),
	);

	const response = results.map((row) => ({
		entryId: row[columnMap.entryId],
		timestamp: row[columnMap.timestamp],
		date: row[columnMap.productionDate],
		machineId: row[columnMap.machineId],
		machineName: row[columnMap.machineName],
		operatorId: row[columnMap.operatorId],
		operatorName: row[columnMap.operatorName],
		bags: row[columnMap.bags],
		rounds: row[columnMap.rounds],
	}));

	return JSON.parse(JSON.stringify(response));
}

/**
 * Builds header column index map.
 */
function buildHeaderColumnMap(headers) {
	return {
		entryId: headers.indexOf('Entry ID'),
		timestamp: headers.indexOf('Timestamp'),
		productionDate: headers.indexOf('Production Date'),
		machineId: headers.indexOf('Machine ID'),
		machineName: headers.indexOf('Machine Name'),
		operatorId: headers.indexOf('Operator ID'),
		operatorName: headers.indexOf('Operator Name'),
		bags: headers.indexOf('Bags'),
		rounds: headers.indexOf('Rounds'),
	};
}

/**
 * Returns true if a row matches the supplied search criteria.
 */
function matchesSearchCriteria(row, columnMap, criteria) {
	if (criteria.entryId) {
		const sheetEntryId = String(row[columnMap.entryId]).trim().toUpperCase();
		const searchEntryId = String(criteria.entryId).trim().toUpperCase();

		if (sheetEntryId !== searchEntryId) {
			return false;
		}
	}

	if (criteria.date) {
		const rowDate = Utilities.formatDate(
			new Date(row[columnMap.productionDate]),
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		const searchDate = Utilities.formatDate(
			new Date(criteria.date),
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		if (rowDate !== searchDate) {
			return false;
		}
	}

	if (criteria.machineId) {
		if (String(row[columnMap.machineId]) !== String(criteria.machineId)) {
			return false;
		}
	}

	if (criteria.operatorId) {
		if (String(row[columnMap.operatorId]) !== String(criteria.operatorId)) {
			return false;
		}
	}

	return true;
}

/**
 * ===========================================================
 * Returns complete production transaction
 * ===========================================================
 */
function getProductionEntry(entryId) {
	const header = getProductionHeader(entryId);

	if (!header) {
		return null;
	}

	return JSON.parse(
		JSON.stringify({
			header: header,
			details: getProductionDetails(entryId),
		}),
	);
}

/**
 * Returns production header.
 */
function getProductionHeader(entryId) {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) {
		return null;
	}

	const headers = values.shift();

	const map = buildHeaderColumnMap(headers);

	for (const row of values) {
		if (row[map.entryId] === entryId) {
			return {
				entryId: row[map.entryId],
				timestamp: row[map.timestamp],
				date: row[map.productionDate],
				machineId: row[map.machineId],
				machineName: row[map.machineName],
				operatorId: row[map.operatorId],
				operatorName: row[map.operatorName],
				bags: row[map.bags],
				rounds: row[map.rounds],
			};
		}
	}

	return null;
}

/**
 * Returns production detail rows.
 */
function getProductionDetails(entryId) {
	const sheet = getSheet(SHEETS.DAILY_DETAIL);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) {
		return [];
	}

	const headers = values.shift();

	const entryIndex = headers.indexOf('Entry ID');
	const productIdIndex = headers.indexOf('Product ID');
	const productNameIndex = headers.indexOf('Product Name');
	const mouldIndex = headers.indexOf('Mould');
	const piecesIndex = headers.indexOf('Pieces');

	return values
		.filter((row) => row[entryIndex] === entryId)
		.map((row) => ({
			productId: row[productIdIndex],
			productName: row[productNameIndex],
			mould: row[mouldIndex],
			pieces: row[piecesIndex],
		}));
}

/**
 * Test helper.
 */
function testSearchProductionEntries() {
	Logger.log(
		JSON.stringify(
			searchProductionEntries({
				entryId: 'TXN-20260701-0001',
			}),
			null,
			2,
		),
	);
}

function testGetProductionEntry() {
	Logger.log(JSON.stringify(getProductionEntry('TXN-20260715-0001'), null, 2));
}

function debugDailyHeaderColumns() {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	Logger.log(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]);
}

function debugFirstHeaderRow() {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	const values = sheet.getDataRange().getValues();

	Logger.log(JSON.stringify(values[1], null, 2));
}

function debugEntryIds() {
	const sheet = getSheet(SHEETS.DAILY_HEADER);
	const values = sheet.getDataRange().getValues();

	values.slice(1).forEach((row, index) => {
		Logger.log(index + ' -> [' + row[0] + ']');
	});
}

function debugSingleComparison() {
	const sheet = getSheet(SHEETS.DAILY_HEADER);
	const values = sheet.getDataRange().getValues();

	const row = values[1];

	Logger.log('Sheet  : [' + row[0] + ']');
	Logger.log('Search : [TXN-20260701-0001]');
	Logger.log(String(row[0]).trim().toUpperCase() === 'TXN-20260701-0001');
}
