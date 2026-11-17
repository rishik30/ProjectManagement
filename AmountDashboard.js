const AMOUNT_DASHBOARD_SHEET = 'Amount Dashboard';

function openAmountDashboard() {
	assertAmountAnalysisAuthorized();
	refreshAmountDashboard();
}

function refreshAmountDashboard() {
	assertAmountAnalysisAuthorized();

	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(AMOUNT_DASHBOARD_SHEET);

	if (!sheet) {
		sheet = ss.insertSheet(AMOUNT_DASHBOARD_SHEET);
	}

	buildAmountDashboardLayout_(sheet);
	updateAmountDashboardCharts_(sheet);

	ss.setActiveSheet(sheet);
}

function buildAmountDashboardLayout_(sheet) {
	sheet.clear();
	sheet.getCharts().forEach((chart) => sheet.removeChart(chart));
	sheet.setHiddenGridlines(true);

	sheet
		.getRange('A1:H1')
		.merge()
		.setValue('MONTHLY AMOUNT DASHBOARD')
		.setFontSize(18)
		.setFontWeight('bold')
		.setHorizontalAlignment('center');

	sheet.getRange('A3').setValue('Chart Period').setFontWeight('bold');
	sheet.getRange('B3').setValue('3 Months');

	updateMonthlyChartPeriodValidation_(sheet, 'B3');

	sheet
		.getRange('A5:H5')
		.merge()
		.setValue('Production and Sales Amount')
		.setFontWeight('bold')
		.setFontSize(14);
}

function updateAmountDashboardCharts_(sheet) {
	assertAmountAnalysisAuthorized();

	const period = sheet.getRange('B3').getValue() || '3 Months';
	const range = getMonthlyChartRange_(period);
	const labels = getMonthlyChartLabels_(range);

	const rows = [['Month', 'Production Amount', 'Sales Amount']];

	labels.forEach((item, index) => {
		const monthStart = new Date(
			range.fromDate.getFullYear(),
			range.fromDate.getMonth() + index,
			1,
		);

		const monthEnd = new Date(
			monthStart.getFullYear(),
			monthStart.getMonth() + 1,
			0,
		);

		monthEnd.setHours(23, 59, 59, 999);

		const fromDate = Utilities.formatDate(
			monthStart,
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		const toDate = Utilities.formatDate(
			monthEnd > range.toDate ? range.toDate : monthEnd,
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		const production = calculateAmountAnalysis({
			type: 'production',
			dateRange: 'Custom',
			fromDate: fromDate,
			toDate: toDate,
			productId: 'All',
			machineId: 'All',
		});

		const sales = calculateAmountAnalysis({
			type: 'sales',
			dateRange: 'Custom',
			fromDate: fromDate,
			toDate: toDate,
			productId: 'All',
			machineId: 'All',
		});

		if (!production.success || !sales.success) {
			console.log('Error calculating amounts:', production, sales);
			throw new Error(
				'Amount Dashboard could not calculate monthly amounts. Check Product_Pricing for missing prices.',
			);
		}

		rows.push([
			item.label,
			Number(production.totals.amount || 0),
			Number(sales.totals.amount || 0),
		]);
	});

	const dataRange = sheet.getRange(1, 26, rows.length, 3);
	dataRange.clearContent().setValues(rows);
	sheet.hideColumns(26, 3);

	sheet.getCharts().forEach((chart) => sheet.removeChart(chart));

	const chart = sheet
		.newChart()
		.setChartType(Charts.ChartType.COLUMN)
		.addRange(dataRange)
		.setNumHeaders(1)
		.setPosition(7, 1, 0, 0)
		.setOption('title', 'Monthly Production Amount vs Sales Amount')
		.setOption('legend', { position: 'bottom' })
		.setOption('hAxis', { title: 'Month' })
		.setOption('vAxis', { title: 'Amount (₹)', minValue: 0 })
		.setOption('width', 1000)
		.setOption('height', 450)
		.build();

	sheet.insertChart(chart);
}

function handleAmountDashboardEdit(e) {
	const sheet = e.range.getSheet();

	if (
		sheet.getName() !== AMOUNT_DASHBOARD_SHEET ||
		e.range.getA1Notation() !== 'B3'
	) {
		return;
	}

	assertAmountAnalysisAuthorized();
	updateAmountDashboardCharts_(sheet);
}
