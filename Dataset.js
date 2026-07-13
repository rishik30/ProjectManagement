/**
 * ===========================================================
 * Dataset Engine
 * Module 4A - Part 3
 * ===========================================================
 */

function getProductionDataset(filters) {
	const headerSheet = getSheet(SHEETS.DAILY_HEADER);
	const detailSheet = getSheet(SHEETS.DAILY_DETAIL);

	const headerValues = headerSheet.getDataRange().getValues();
	const detailValues = detailSheet.getDataRange().getValues();

	if (headerValues.length <= 1 || detailValues.length <= 1) {
		return [];
	}

	headerValues.shift();
	detailValues.shift();

	const headerMap = buildHeaderMap(headerValues);

	let dataset = buildDataset(detailValues, headerMap);

	dataset = applyFilters(dataset, filters);

	return dataset;
}

function buildHeaderMap(headers) {
	const map = {};

	headers.forEach((row) => {
		map[row[0]] = {
			entryId: row[0],
			timestamp: row[1],
			date: row[2],
			machineId: row[3],
			machineName: row[4],
			operatorId: row[5],
			operatorName: row[6],
			bags: Number(row[7]),
			rounds: Number(row[8]),
			enteredBy: row[9],
			financialYear: row[10],
			month: row[11],
			week: row[12],
		};
	});

	return map;
}

function buildDataset(details, headerMap) {
	const dataset = [];

	details.forEach((row) => {
		const header = headerMap[row[0]];

		if (!header) {
			return;
		}

		dataset.push({
			entryId: header.entryId,
			timestamp: header.timestamp,
			date: header.date,
			machineId: header.machineId,
			machineName: header.machineName,
			operatorId: header.operatorId,
			operatorName: header.operatorName,
			bags: header.bags,
			rounds: header.rounds,
			financialYear: header.financialYear,
			month: header.month,
			week: header.week,
			productId: row[1],
			productName: row[2],
			mould: Number(row[3]),
			pieces: Number(row[4]),
		});
	});

	return dataset;
}

function applyFilters(dataset, filters) {
	if (!filters) {
		return dataset;
	}

	return dataset.filter((record) => {
		if (
			filters.machineId &&
			filters.machineId !== 'All' &&
			record.machineId !== filters.machineId
		) {
			return false;
		}

		if (
			filters.operatorId &&
			filters.operatorId !== 'All' &&
			record.operatorId !== filters.operatorId
		) {
			return false;
		}

		if (
			filters.productId &&
			filters.productId !== 'All' &&
			record.productId !== filters.productId
		) {
			return false;
		}

		if (filters.fromDate) {
			if (record.date < filters.fromDate) {
				return false;
			}
		}

		if (filters.toDate) {
			if (record.date > filters.toDate) {
				return false;
			}
		}

		return true;
	});
}

function testDataset() {
	const filters = {
		machineId: 'All',

		operatorId: 'All',

		productId: 'All',

		fromDate: null,

		toDate: null,
	};

	const dataset = getProductionDataset(filters);

	Logger.log(JSON.stringify(dataset, null, 2));
}

/**
 * ===========================================================
 * Transaction Summary Engine
 * Module 4B Part 1
 * ===========================================================
 */

function getTransactionSummary(dataset) {
	const transactionMap = {};

	dataset.forEach((record) => {
		if (!transactionMap[record.entryId]) {
			transactionMap[record.entryId] = {
				entryId: record.entryId,
				date: record.date,
				machineId: record.machineId,
				machineName: record.machineName,
				operatorId: record.operatorId,
				operatorName: record.operatorName,
				bags: record.bags,
				rounds: record.rounds,
				pieces: 0,
			};
		}

		transactionMap[record.entryId].pieces += record.pieces;
	});

	return Object.values(transactionMap);
}

function calculateKPIs(dataset) {
	const transactions = getTransactionSummary(dataset);

	const kpi = {
		totalPieces: 0,
		totalBags: 0,
		totalRounds: 0,
		activeMachines: 0,
		avgPiecesPerBag: 0,
		avgPiecesPerRound: 0,
	};

	const machineSet = new Set();

	transactions.forEach((txn) => {
		kpi.totalPieces += txn.pieces;
		kpi.totalBags += txn.bags;
		kpi.totalRounds += txn.rounds;
		machineSet.add(txn.machineId);
	});

	kpi.activeMachines = machineSet.size;

	if (kpi.totalBags > 0) {
		kpi.avgPiecesPerBag = kpi.totalPieces / kpi.totalBags;
	}

	if (kpi.totalRounds > 0) {
		kpi.avgPiecesPerRound = kpi.totalPieces / kpi.totalRounds;
	}

	return kpi;
}

function testKPIs() {
	const dataset = getProductionDataset({
		machineId: 'All',

		operatorId: 'All',

		productId: 'All',

		fromDate: null,

		toDate: null,
	});

	Logger.log(
		JSON.stringify(
			calculateKPIs(dataset),

			null,

			2,
		),
	);
}
