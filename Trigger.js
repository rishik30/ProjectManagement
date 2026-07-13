/**
 * ===========================================================
 * Spreadsheet Triggers
 * ===========================================================
 */

function onEdit(e) {
	if (!e) {
		return;
	}

	handleDashboardEdit(e);
}
