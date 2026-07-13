/**
 * ===========================================================
 * Dashboard Configuration
 * ===========================================================
 */

const DASHBOARD = {
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
};
