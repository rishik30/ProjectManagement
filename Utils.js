/**
 * ===========================================================
 * Utils.gs
 * ===========================================================
 */

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ===========================================================
 * Date & ID Utility Functions
 * ===========================================================
 */

/**
 * Returns current timestamp
 */
function getCurrentTimestamp() {
	return new Date();
}

/**
 * Returns month name
 */
function getMonthName(date) {
	const months = [
		'January',
		'February',
		'March',
		'April',
		'May',
		'June',
		'July',
		'August',
		'September',
		'October',
		'November',
		'December',
	];

	return months[date.getMonth()];
}

/**
 * Returns financial year
 * Example:
 * April 2026 -> 2026-27
 * January 2027 -> 2026-27
 */
function getFinancialYearText(date) {
	let startYear;

	if (date.getMonth() >= 3) {
		startYear = date.getFullYear();
	} else {
		startYear = date.getFullYear() - 1;
	}

	const endYear = String(startYear + 1).slice(-2);

	return `${startYear}-${endYear}`;
}

/**
 * Returns ISO Week Number
 */
function getWeekNumber(date) {
	const d = new Date(
		Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()),
	);

	d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));

	const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

	return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function getStartOfWeek(date) {
	const d = new Date(date);

	const day = d.getDay();

	const diff = day === 0 ? -6 : 1 - day;

	d.setDate(d.getDate() + diff);

	d.setHours(0, 0, 0, 0);

	return d;
}

function getFinancialYearObject(date) {
	const year =
		date.getMonth() >= 3 ? date.getFullYear() : date.getFullYear() - 1;

	return {
		from: new Date(year, 3, 1),

		to: new Date(year + 1, 2, 31),
	};
}

/**
 * ===========================================================
 * System Error Logging
 * ===========================================================
 */

const SYSTEM_LOG_SHEET = '_System_Log';

/**
 * Returns the system log sheet.
 * Creates it automatically if it doesn't exist.
 */
function getSystemLogSheet() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	let sheet = ss.getSheetByName(SYSTEM_LOG_SHEET);

	if (!sheet) {
		sheet = ss.insertSheet(SYSTEM_LOG_SHEET);

		sheet.appendRow([
			'Timestamp',
			'User',
			'Function',
			'Error Message',
			'Stack Trace',
			'Context',
		]);

		sheet.hideSheet();
	}

	return sheet;
}

/**
 * Logs unexpected application errors.
 *
 * Logging should never interrupt the main execution.
 */
function logError(functionName, error, context) {
	try {
		const sheet = getSystemLogSheet();

		sheet.appendRow([
			new Date(),
			Session.getActiveUser().getEmail(),
			functionName,
			error && error.message ? error.message : String(error),
			error && error.stack ? error.stack : '',
			context ? JSON.stringify(context) : '',
		]);
	} catch (loggingError) {
		// Never allow logging failures to interrupt application flow.
	}
}

/**
 * ===========================================================
 * System Validation
 * ===========================================================
 */

function validateSystem() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	const requiredSheets = [
		SHEETS.USERS,
		SHEETS.PRODUCT_MASTER,
		SHEETS.MACHINE_MASTER,
		SHEETS.MACHINE_CONFIGURATION,
		SHEETS.MACHINE_HISTORY,
		SHEETS.DAILY_HEADER,
		SHEETS.DAILY_DETAIL,
		SHEETS.DASHBOARD,
	];

	const missingSheets = requiredSheets.filter(
		(name) => !ss.getSheetByName(name),
	);

	if (missingSheets.length === 0) {
		SpreadsheetApp.getUi().alert(
			'✅ System Validation Successful.\n\nAll required sheets are present.',
		);

		return;
	}

	const message =
		'The following required sheets are missing:\n\n' + missingSheets.join('\n');

	SpreadsheetApp.getUi().alert(message);
}
