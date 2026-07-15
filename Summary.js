/**
 * ===========================================================
 * Summary Engine
 * Module 4C Part 1
 * ===========================================================
 */

function summarize(dataset, keyFields, numericFields) {
	const summaryMap = {};

	dataset.forEach((record) => {
		const key = keyFields.map((field) => record[field]).join('|');

		if (!summaryMap[key]) {
			summaryMap[key] = {};

			keyFields.forEach((field) => {
				summaryMap[key][field] = record[field];
			});

			numericFields.forEach((field) => {
				summaryMap[key][field] = 0;
			});
		}

		numericFields.forEach((field) => {
			summaryMap[key][field] += Number(record[field]) || 0;
		});
	});

	return Object.values(summaryMap);
}

function summarizeByProduct(dataset) {
	return summarize(dataset, ['productId', 'productName'], ['pieces']);
}

function summarizeByMachine(dataset) {
	return summarize(dataset, ['machineId', 'machineName'], ['pieces']);
}

function summarizeByOperator(dataset) {
	return summarize(dataset, ['operatorId', 'operatorName'], ['pieces']);
}

function summarizeByDate(dataset) {
	return summarize(dataset, ['productionDateOnly'], ['pieces']);
}

function testProductSummary() {
	const dataset = getProductionDataset({
		machineId: 'All',

		operatorId: 'All',

		productId: 'All',

		fromDate: null,

		toDate: null,
	});

	Logger.log(
		JSON.stringify(
			summarizeByProduct(dataset),

			null,

			2,
		),
	);
}
