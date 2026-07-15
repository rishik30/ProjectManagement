/**
 * ===========================================================
 * Dashboard Configuration
 * ===========================================================
 */

const DASHBOARD = {
	DATE_RANGES: [
		'Today',

		'This Week',

		'This Month',

		'Last Month',

		'Last 3 Months',

		'Financial Year',

		'Custom',
	],
	KPI_CARDS: [
		{
			key: 'totalPieces',
			title: 'Total Pieces',
			cell: 'A13',
			decimals: 0,
		},

		{
			key: 'totalBags',
			title: 'Total Bags',
			cell: 'C13',
			decimals: 0,
		},

		{
			key: 'totalRounds',
			title: 'Total Rounds',
			cell: 'E13',
			decimals: 0,
		},

		{
			key: 'activeMachines',
			title: 'Active Machines',
			cell: 'A17',
			decimals: 0,
		},

		{
			key: 'avgPiecesPerBag',
			title: 'Pieces / Bag',
			cell: 'C17',
			decimals: 2,
		},

		{
			key: 'avgPiecesPerRound',
			title: 'Pieces / Round',
			cell: 'E17',
			decimals: 2,
		},
	],
	SUMMARY_TABLES: {
		PRODUCT: {
			title: 'Product Summary',
			startRow: 22,
			startColumn: 1,
			nameField: 'productName',
			valueField: 'pieces',
			nameHeader: 'Product',
			valueHeader: 'Pieces',
		},

		MACHINE: {
			title: 'Machine Summary',
			startRow: 22,
			startColumn: 4,
			nameField: 'machineName',
			valueField: 'pieces',
			nameHeader: 'Machine',
			valueHeader: 'Pieces',
		},

		OPERATOR: {
			title: 'Operator Summary',
			startRow: 22,
			startColumn: 7,
			nameField: 'operatorName',
			valueField: 'pieces',
			nameHeader: 'Operator',
			valueHeader: 'Pieces',
		},

		DAILY: {
			title: 'Daily Summary',
			startRow: 22,
			startColumn: 10,
			nameField: 'productionDateOnly',
			valueField: 'pieces',
			nameHeader: 'Date',
			valueHeader: 'Pieces',
		},
	},
};
