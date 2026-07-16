/**
 * ===========================================================
 * Production.gs
 * ===========================================================
 */

/**
 * Returns all active machines
 */
function getMachines() {
	const sheet = getSheet(SHEETS.MACHINE_MASTER);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) return [];

	const headers = values.shift();

	const machineIdIndex = headers.indexOf('Machine ID');
	const machineNameIndex = headers.indexOf('Machine Name');
	const defaultOperatorIndex = headers.indexOf('Default Operator ID');
	const statusIndex = headers.indexOf('Status');

	return values
		.filter((r) => String(r[statusIndex]).toLowerCase() === 'active')
		.map((r) => ({
			id: r[machineIdIndex],
			name: r[machineNameIndex],
			defaultOperatorId: r[defaultOperatorIndex],
		}));
}

/**
 * Returns all active operators
 */
function getOperators() {
	const sheet = getSheet(SHEETS.USERS);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) return [];

	const headers = values.shift();

	const idIndex = headers.indexOf('User ID');
	const nameIndex = headers.indexOf('Name');
	const roleIndex = headers.indexOf('Role');
	const activeIndex = headers.indexOf('Active');

	return values
		.filter(
			(r) =>
				String(r[roleIndex]).toLowerCase() === 'operator' &&
				r[activeIndex] === true,
		)
		.map((r) => ({
			id: r[idIndex],
			name: r[nameIndex],
		}));
}

/**
 * Returns current configuration of a machine
 */
function getMachineConfiguration(machineId) {
	const configSheet = getSheet(SHEETS.MACHINE_CONFIGURATION);
	const productSheet = getSheet(SHEETS.PRODUCT_MASTER);

	const config = configSheet.getDataRange().getValues();
	const products = productSheet.getDataRange().getValues();

	if (config.length <= 1 || products.length <= 1) {
		return [];
	}

	const productMap = {};

	const productHeaders = products.shift();

	const pIdIndex = productHeaders.indexOf('Product ID');
	const pNameIndex = productHeaders.indexOf('Product Name');

	products.forEach((row) => {
		productMap[row[pIdIndex]] = row[pNameIndex];
	});

	const configHeaders = config.shift();

	const machineIndex = configHeaders.indexOf('Machine');
	const productIndex = configHeaders.indexOf('Product');
	const mouldIndex = configHeaders.indexOf('Current Mould');

	return config
		.filter((r) => r[machineIndex] === machineId)
		.map((r) => ({
			productId: r[productIndex],
			productName: productMap[r[productIndex]] || '',
			mould: r[mouldIndex],
		}));
}

/**
 * ===========================================================
 * Module 2B Part 1
 * ===========================================================
 */

/**
 * Generates Entry ID
 *
 * Format:
 * TXN-YYYYMMDD-0001
 */
function generateEntryId(productionDate) {
	const headerSheet = getSheet(SHEETS.DAILY_HEADER);

	const values = headerSheet.getDataRange().getValues();

	const targetDate = Utilities.formatDate(
		new Date(productionDate),
		Session.getScriptTimeZone(),
		'yyyyMMdd',
	);

	let count = 0;

	for (let i = 1; i < values.length; i++) {
		const entryId = values[i][0];

		if (entryId && entryId.indexOf('TXN-' + targetDate) === 0) {
			count++;
		}
	}

	const runningNumber = String(count + 1).padStart(4, '0');

	return `TXN-${targetDate}-${runningNumber}`;
}

/**
 * Checks duplicate production entry
 *
 * Duplicate =
 * Same Date
 * Same Machine
 */
function checkDuplicateEntry(productionDate, machineId) {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) {
		return false;
	}

	for (let i = 1; i < values.length; i++) {
		const existingDate = Utilities.formatDate(
			new Date(values[i][2]),
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		const existingMachine = values[i][3];

		const inputDate = Utilities.formatDate(
			new Date(productionDate),
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		if (existingDate === inputDate && existingMachine === machineId) {
			return true;
		}
	}

	return false;
}

function testModule2BPart1() {
	Logger.log(generateEntryId('2026-07-15'));

	Logger.log(checkDuplicateEntry('2026-07-15', 'M01'));

	Logger.log(getFinancialYear(new Date('2026-07-15')));

	Logger.log(getMonthName(new Date('2026-07-15')));

	Logger.log(getWeekNumber(new Date('2026-07-15')));
}

/**
 * ===========================================================
 * Writes Header Record
 * ===========================================================
 */
function writeHeader(data) {
	const sheet = getSheet(SHEETS.DAILY_HEADER);

	const productionDate = new Date(data.date);

	const entryId = generateEntryId(data.date);

	const row = [
		entryId,

		getCurrentTimestamp(),

		productionDate,

		data.machineId,

		data.machineName,

		data.operatorId,

		data.operatorName,

		Number(data.bags),

		Number(data.rounds),

		Session.getActiveUser().getEmail(),

		getFinancialYearText(productionDate),

		getMonthName(productionDate),

		getWeekNumber(productionDate),
	];

	sheet.appendRow(row);

	return entryId;
}

function testWriteHeader() {
	const data = {
		date: '2026-07-15',

		machineId: 'M01',

		machineName: 'AJAY',

		operatorId: 'U001',

		operatorName: 'Ajay',

		bags: 1.5,

		rounds: 40,
	};

	const id = writeHeader(data);

	Logger.log(id);
}

/**
 * ===========================================================
 * Writes Detail Records
 * ===========================================================
 */
function writeDetails(entryId, data) {
	const sheet = getSheet(SHEETS.DAILY_DETAIL);

	const rows = [];

	data.products.forEach((product) => {
		rows.push([
			entryId,

			product.productId,

			product.productName,

			Number(product.mould),

			Number(product.mould) * Number(data.rounds),
		]);
	});

	if (rows.length > 0) {
		sheet
			.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
			.setValues(rows);
	}
}

function testWriteDetails() {
	const data = {
		rounds: 40,

		products: [
			{
				productId: 'P001',

				productName: 'Dana Duck',

				mould: 11,
			},

			{
				productId: 'P002',

				productName: 'Big Fruit',

				mould: 43,
			},

			{
				productId: 'P003',

				productName: 'Football',

				mould: 8,
			},
		],
	};

	writeDetails(
		'TXN-20260715-0001',

		data,
	);
}

/**
 * ===========================================================
 * Main Submit Function
 * ===========================================================
 */
function submitProduction(data) {
	try {
		const lock = LockService.getDocumentLock();
		lock.waitLock(20000);

		try {
			if (checkDuplicateEntry(data.date, data.machineId)) {
				return {
					success: false,
					duplicate: true,
					message: 'Production entry already exists for this Machine and Date.',
				};
			}

			const entryId = writeHeader(data);

			writeDetails(entryId, data);

			return {
				success: true,
				duplicate: false,
				entryId: entryId,
			};
		} finally {
			lock.releaseLock();
		}
	} catch (error) {
		logError('submitProduction', error, data);
		throw error;
	}
}

function testSubmitProduction() {
	const data = {
		date: '2026-07-15',

		machineId: 'M01',

		machineName: 'AJAY',

		operatorId: 'U001',

		operatorName: 'Ajay',

		bags: 1.5,

		rounds: 40,

		products: [
			{
				productId: 'P001',
				productName: 'Dana Duck',
				mould: 11,
			},

			{
				productId: 'P002',
				productName: 'Big Fruit',
				mould: 43,
			},
		],
	};

	Logger.log(submitProduction(data));
}
