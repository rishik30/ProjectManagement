const MONTHLY_CHART_PERIODS = ['3 Months', 'Quarter', 'Half Year', 'Full Year'];

function getMonthlyChartRange_(period) {
	const today = new Date();
	today.setHours(23, 59, 59, 999);

	const year = today.getFullYear();
	const month = today.getMonth();

	let startDate;

	switch (period) {
		case 'Quarter': {
			const financialMonth = (month - 3 + 12) % 12;
			const quarterIndex = Math.floor(financialMonth / 3);

			const financialYearStart = month >= 3 ? year : year - 1;

			startDate = new Date(financialYearStart, 3 + quarterIndex * 3, 1);
			break;
		}

		case 'Half Year':
			// Rolling six months, including the current month.
			startDate = new Date(year, month - 5, 1);
			break;

		case 'Full Year':
			startDate = new Date(month >= 3 ? year : year - 1, 3, 1);
			break;

		case '3 Months':
		default:
			startDate = new Date(year, month - 2, 1);
			break;
	}

	startDate.setHours(0, 0, 0, 0);

	return {
		fromDate: startDate,
		toDate: today,
	};
}

function getMonthlyChartLabels_(range) {
	const labels = [];
	const current = new Date(
		range.fromDate.getFullYear(),
		range.fromDate.getMonth(),
		1,
	);

	const last = new Date(range.toDate.getFullYear(), range.toDate.getMonth(), 1);

	while (current <= last) {
		labels.push({
			key: current.getFullYear() + '-' + current.getMonth(),
			label: Utilities.formatDate(
				current,
				Session.getScriptTimeZone(),
				'MMM yyyy',
			),
		});

		current.setMonth(current.getMonth() + 1);
	}

	return labels;
}

function aggregateMonthlyProductionPieces_(range) {
	const dataset = getProductionDataset({
		machineId: 'All',
		operatorId: 'All',
		productId: 'All',
		fromDate: range.fromDate,
		toDate: range.toDate,
	});

	const totals = {};

	dataset.forEach((record) => {
		const date = record.productionDateOnly;
		const key = date.getFullYear() + '-' + date.getMonth();

		totals[key] = (totals[key] || 0) + Number(record.pieces || 0);
	});

	return totals;
}

function aggregateMonthlySalesPieces_(range) {
	const dataset = getSalesDashboardDataset({
		fromDate: range.fromDate,
		toDate: range.toDate,
		customerName: 'All',
		productName: 'All',
	});

	const totals = {};

	dataset.forEach((record) => {
		const date = record.date;
		const key = date.getFullYear() + '-' + date.getMonth();

		totals[key] = (totals[key] || 0) + Number(record.pieces || 0);
	});

	return totals;
}

function writeMonthlyColumnChart_(
	sheet,
	title,
	period,
	monthlyValues,
	chartRow,
	valueTitle,
) {
	const labels = getMonthlyChartLabels_(getMonthlyChartRange_(period));

	const values = [['Month', valueTitle]];

	labels.forEach((item) => {
		values.push([item.label, Number(monthlyValues[item.key] || 0)]);
	});

	const dataColumn = 26; // Z
	const dataRange = sheet.getRange(
		1,
		dataColumn,
		values.length,
		values[0].length,
	);

	dataRange.clearContent().setValues(values);

	sheet.hideColumns(dataColumn, 2);

	sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

	const chart = sheet
		.newChart()
		.setChartType(Charts.ChartType.COLUMN)
		.addRange(dataRange)
		.setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.SHOW_BOTH)
		.setNumHeaders(1)
		.setPosition(chartRow, 1, 0, 0)
		.setOption('title', title)
		.setOption('legend', { position: 'none' })
		.setOption('hAxis', { title: 'Month' })
		.setOption('vAxis', { title: valueTitle, minValue: 0 })
		.setOption('width', 900)
		.setOption('height', 360)
		.build();

	sheet.insertChart(chart);
}

function updateMonthlyChartPeriodValidation_(sheet, cell) {
	const rule = SpreadsheetApp.newDataValidation()
		.requireValueInList(MONTHLY_CHART_PERIODS, true)
		.setAllowInvalid(false)
		.build();

	sheet.getRange(cell).setDataValidation(rule);
}
