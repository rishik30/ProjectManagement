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
function getFinancialYear(date) {
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

function getFinancialYear(date) {
	const year =
		date.getMonth() >= 3 ? date.getFullYear() : date.getFullYear() - 1;

	return {
		from: new Date(year, 3, 1),

		to: new Date(year + 1, 2, 31),
	};
}
