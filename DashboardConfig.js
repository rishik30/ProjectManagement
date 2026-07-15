/**
 * ===========================================================
 * Dashboard Configuration
 * ===========================================================
 */

const DASHBOARD_LAYOUT = {
	TITLE: {
		ROW: 1,
		COLUMN: 1,
		WIDTH: 10,
	},

	STATUS: {
		ROW: 2,
		COLUMN: 1,
		WIDTH: 10,
	},

	KPI: {
		TITLE_ROW: 11,
		TITLE_COLUMN: 1,
		TITLE_WIDTH: 10,
	},

	HIGHLIGHTS: {
		TITLE_ROW: 21,

		LEFT: {
			ROW: 22,
			COLUMN: 1,
			WIDTH: 5,
			HEIGHT: 4,
		},

		RIGHT: {
			ROW: 22,
			COLUMN: 6,
			WIDTH: 5,
			HEIGHT: 4,
		},
	},

	SUMMARIES: {
		START_ROW: 27,
	},
};

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
		{
			key: 'uniqueProducts',
			title: 'Unique Products',
			cell: 'G13',
			decimals: 0,
		},

		{
			key: 'uniqueDies',
			title: 'Unique Dies',
			cell: 'I13',
			decimals: 0,
		},
	],
	SUMMARY_TABLES: {
		PRODUCT: {
			title: 'Product Summary',
			startRow: DASHBOARD_LAYOUT.SUMMARIES.START_ROW,
			startColumn: 1,
			nameField: 'productName',
			valueField: 'pieces',
			nameHeader: 'Product',
			valueHeader: 'Pieces',
		},

		MACHINE: {
			title: 'Machine Summary',
			startRow: DASHBOARD_LAYOUT.SUMMARIES.START_ROW,
			startColumn: 4,
			nameField: 'machineName',
			valueField: 'pieces',
			nameHeader: 'Machine',
			valueHeader: 'Pieces',
		},

		OPERATOR: {
			title: 'Operator Summary',
			startRow: DASHBOARD_LAYOUT.SUMMARIES.START_ROW,
			startColumn: 7,
			nameField: 'operatorName',
			valueField: 'pieces',
			nameHeader: 'Operator',
			valueHeader: 'Pieces',
		},

		DAILY: {
			title: 'Daily Summary',
			startRow: DASHBOARD_LAYOUT.SUMMARIES.START_ROW,
			startColumn: 10,
			nameField: 'productionDateOnly',
			valueField: 'pieces',
			nameHeader: 'Date',
			valueHeader: 'Pieces',
		},
	},
};
