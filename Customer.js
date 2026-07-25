/**
 * ===========================================================
 * Customer Master
 * ===========================================================
 */

function initializeCustomerMaster() {
	createSheetIfMissing(SHEETS.CUSTOMER_MASTER, [
		'Customer ID',
		'Customer Name',
		'Contact Person',
		'Mobile',
		'Email',
		'GSTIN',
		'Address',
		'City',
		'State',
		'Pincode',
		'Status',
	]);
}

/**
 * Generates Customer ID
 *
 * Format:
 * CUST0001
 */

function generateCustomerId() {
	const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

	if (sheet.getLastRow() <= 1) {
		return 'CUST0001';
	}

	const ids = sheet
		.getRange(2, 1, sheet.getLastRow() - 1, 1)
		.getValues()
		.flat()
		.filter(String);

	let max = 0;

	ids.forEach((id) => {
		const n = parseInt(id.replace('CUST', ''), 10);

		if (n > max) {
			max = n;
		}
	});

	return 'CUST' + String(max + 1).padStart(4, '0');
}

function validateCustomer(customer) {
	if (!customer) {
		throw new Error('Customer data is required.');
	}

	if (!customer.name || customer.name.trim() === '') {
		throw new Error('Customer Name is required.');
	}

	if (customer.mobile && !/^[0-9]{10}$/.test(String(customer.mobile).trim())) {
		throw new Error('Invalid Mobile Number.');
	}

	if (
		customer.email &&
		!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(customer.email.trim())
	) {
		throw new Error('Invalid Email Address.');
	}

	if (
		customer.gstin &&
		!/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(
			customer.gstin.trim().toUpperCase(),
		)
	) {
		throw new Error('Invalid GSTIN.');
	}

	if (customer.pincode && !/^[0-9]{6}$/.test(String(customer.pincode).trim())) {
		throw new Error('Invalid Pincode.');
	}
}

function isDuplicateCustomer(customerName, excludeId) {
	const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

	if (sheet.getLastRow() <= 1) {
		return false;
	}

	const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();

	const name = customerName.trim().toLowerCase();

	for (const row of data) {
		const id = row[0];

		const existing = String(row[1]).trim().toLowerCase();

		if (excludeId && id === excludeId) {
			continue;
		}

		if (existing === name) {
			return true;
		}
	}

	return false;
}

function addCustomer(customer) {
	try {
		validateCustomer(customer);

		if (isDuplicateCustomer(customer.name)) {
			throw new Error('Customer already exists.');
		}

		const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

		const id = generateCustomerId();

		sheet.appendRow([
			id,
			customer.name.trim(),
			customer.contactPerson || '',
			customer.mobile || '',
			customer.email || '',
			customer.gstin ? customer.gstin.trim().toUpperCase() : '',
			customer.address || '',
			customer.city || '',
			customer.state || '',
			customer.pincode || '',
			'Active',
		]);

		return id;
	} catch (error) {
		logError('addCustomer', error, customer);
		throw error;
	}
}

function getCustomer(customerId) {
	const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

	if (sheet.getLastRow() <= 1) {
		return null;
	}

	const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();

	for (const row of data) {
		if (row[0] === customerId) {
			return {
				id: row[0],
				name: row[1],
				contactPerson: row[2],
				mobile: row[3],
				email: row[4],
				gstin: row[5],
				address: row[6],
				city: row[7],
				state: row[8],
				pincode: row[9],
				status: row[10],
			};
		}
	}

	return null;
}

function getActiveCustomers() {
	const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

	if (sheet.getLastRow() <= 1) {
		return [];
	}

	const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();

	return data
		.filter((row) => row[10] === 'Active')
		.map((row) => ({
			id: row[0],
			name: row[1],
		}))
		.sort((a, b) => a.name.localeCompare(b.name));
}

/**
 * Updates an existing customer.
 *
 * @param {string} customerId
 * @param {Object} customer
 */
function updateCustomer(customerId, customer) {
	try {
		if (!customerId) {
			throw new Error('Customer ID is required.');
		}

		validateCustomer(customer);

		if (isDuplicateCustomer(customer.name, customerId)) {
			throw new Error('Customer already exists.');
		}

		const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

		const values = sheet.getDataRange().getValues();

		for (let i = 1; i < values.length; i++) {
			if (values[i][0] === customerId) {
				sheet
					.getRange(i + 1, 2, 1, 10)
					.setValues([
						[
							customer.name.trim(),
							customer.contactPerson || '',
							customer.mobile || '',
							customer.email || '',
							customer.gstin ? customer.gstin.trim().toUpperCase() : '',
							customer.address || '',
							customer.city || '',
							customer.state || '',
							customer.pincode || '',
							values[i][10],
						],
					]);

				return {
					success: true,
					customerId: customerId,
				};
			}
		}

		throw new Error('Customer not found.');
	} catch (error) {
		logError('updateCustomer', error, {
			customerId,
			customer,
		});

		throw error;
	}
}

/**
 * Toggles Customer Status.
 *
 * Active <-> Inactive
 */

function toggleCustomerStatus(customerId) {
	try {
		if (!customerId) {
			throw new Error('Customer ID is required.');
		}

		const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

		const values = sheet.getDataRange().getValues();

		for (let i = 1; i < values.length; i++) {
			if (values[i][0] === customerId) {
				const currentStatus = values[i][10];

				const newStatus = currentStatus === 'Active' ? 'Inactive' : 'Active';

				sheet.getRange(i + 1, 11).setValue(newStatus);

				return {
					success: true,
					customerId: customerId,
					status: newStatus,
				};
			}
		}

		throw new Error('Customer not found.');
	} catch (error) {
		logError('toggleCustomerStatus', error, customerId);

		throw error;
	}
}

/**
 * Searches customers.
 *
 * Filters:
 * searchText
 * status
 */

function searchCustomers(searchText = '', status = 'All') {
	const sheet = getSheet(SHEETS.CUSTOMER_MASTER);

	if (sheet.getLastRow() <= 1) {
		return [];
	}

	searchText = searchText.toLowerCase().trim();

	const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();

	return values
		.filter((row) => {
			if (status !== 'All' && row[10] !== status) {
				return false;
			}

			if (searchText === '') {
				return true;
			}

			return (
				String(row[0]).toLowerCase().includes(searchText) ||
				String(row[1]).toLowerCase().includes(searchText) ||
				String(row[2]).toLowerCase().includes(searchText) ||
				String(row[3]).toLowerCase().includes(searchText) ||
				String(row[5]).toLowerCase().includes(searchText)
			);
		})
		.map((row) => ({
			id: row[0],
			name: row[1],
			contactPerson: row[2],
			mobile: row[3],
			email: row[4],
			gstin: row[5],
			address: row[6],
			city: row[7],
			state: row[8],
			pincode: row[9],
			status: row[10],
		}));
}

function testSearchCustomers() {
	Logger.log(searchCustomers());

	Logger.log(searchCustomers('abc'));

	Logger.log(searchCustomers('', 'Active'));

	Logger.log(searchCustomers('', 'Inactive'));
}

function testUpdateCustomer() {
	Logger.log(
		updateCustomer('CUST0001', {
			name: 'ABC Traders Pvt Ltd',
			contactPerson: 'Rajesh Sharma',
			mobile: '9999999999',
			email: 'abc@test.com',
			gstin: '',
			address: 'Industrial Area',
			city: 'Jaipur',
			state: 'Rajasthan',
			pincode: '302001',
		}),
	);
}

function testToggleCustomer() {
	Logger.log(toggleCustomerStatus('CUST0001'));
}

function testAddCustomer() {
	Logger.log(
		addCustomer({
			name: 'ABC Traders',
			contactPerson: 'Rajesh',
			mobile: '9876543210',
			email: 'abc@gmail.com',
			gstin: '',
			address: 'Industrial Area',
			city: 'Jaipur',
			state: 'Rajasthan',
			pincode: '302001',
		}),
	);
}
