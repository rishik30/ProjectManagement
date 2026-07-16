/**
 * ===========================================================
 * MachineSetup.gs
 * ===========================================================
 */

/**
 * ===========================================================
 * Module 3A Part 1
 * ===========================================================
 */

/**
 * Returns machines for dropdown
 */
function getMachineSetupData() {
	return {
		machines: getMachines(),
		products: getAllProducts(),
	};
}

/**
 * Returns machine configuration
 */
function getMachineConfigurationForEdit(machineId) {
	return getMachineConfiguration(machineId);
}

/**
 * Returns all active products
 */
function getAllProducts() {
	const sheet = getSheet(SHEETS.PRODUCT_MASTER);

	const values = sheet.getDataRange().getValues();

	if (values.length <= 1) {
		return [];
	}

	const headers = values.shift();

	const idIndex = headers.indexOf('Product ID');

	const nameIndex = headers.indexOf('Product Name');

	const activeIndex = headers.indexOf('Active');

	return values
		.filter((r) => r[activeIndex] === true)
		.map((r) => ({
			id: r[idIndex],

			name: r[nameIndex],
		}));
}

/**
 * ===========================================================
 * Saves Machine Configuration
 * ===========================================================
 */
function saveMachineConfiguration(machineId, changeSet) {
	const lock = LockService.getDocumentLock();

	lock.waitLock(30000);

	try {
		applyRemovedProducts(machineId, changeSet.removed);

		applyUpdatedProducts(machineId, changeSet.updated);

		applyAddedProducts(machineId, changeSet.added);

		return {
			success: true,
		};
	} finally {
		lock.releaseLock();
	}
}

function getConfigurationSheet() {
	return getSheet(SHEETS.MACHINE_CONFIGURATION);
}

function getHistorySheet() {
	return getSheet(SHEETS.MACHINE_HISTORY);
}

function applyRemovedProducts(machineId, removed) {
	if (removed.length === 0) {
		return;
	}

	const sheet = getConfigurationSheet();

	const values = sheet.getDataRange().getValues();

	for (let i = values.length - 1; i >= 1; i--) {
		const row = values[i];

		removed.forEach((item) => {
			if (row[0] === machineId && row[1] === item.productId) {
				sheet.deleteRow(i + 1);

				logHistory(machineId, item.productId, 'Removed', row[2], '');
			}
		});
	}
}

function applyUpdatedProducts(machineId, updated) {
	if (updated.length === 0) {
		return;
	}

	const sheet = getConfigurationSheet();

	const values = sheet.getDataRange().getValues();

	for (let i = 1; i < values.length; i++) {
		updated.forEach((item) => {
			if (values[i][0] === machineId && values[i][1] === item.productId) {
				sheet.getRange(i + 1, 2).setValue(item.productId);

				sheet.getRange(i + 1, 3).setValue(item.newMould);

				logHistory(
					machineId,
					item.productId,
					'Updated',
					item.oldMould,
					item.newMould,
				);
			}
		});
	}
}

function applyAddedProducts(machineId, added) {
	if (added.length === 0) {
		return;
	}

	const sheet = getConfigurationSheet();

	added.forEach((item) => {
		sheet.appendRow([machineId, item.productId, item.mould]);

		logHistory(machineId, item.productId, 'Added', '', item.mould);
	});
}

function logHistory(machineId, productId, action, oldMould, newMould) {
	getHistorySheet().appendRow([
		new Date(),
		machineId,
		productId,
		action,
		oldMould,
		newMould,
		Session.getActiveUser().getEmail(),
	]);
}
