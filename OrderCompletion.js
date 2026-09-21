/**
 * ===========================================================
 * ORDER COMPLETION CALCULATOR
 * ===========================================================
 *
 * Uses:
 * 1. Current Stock
 * 2. Current Machine Configuration
 * 3. Production history from the last 7 days
 *
 * Expected production/day:
 *
 * Average historical pieces per mould per production day
 * × current configured moulds
 *
 * Required days:
 *
 * CEILING(Shortfall / Expected Production Per Day)
 * ===========================================================
 */

/**
 * Returns products available for the order calculator.
 */
function getOrderCompletionProducts() {
	const products = getAllProducts();

	return products.map((product) => ({
		id: product.id,
		name: product.name,
	}));
}

/**
 * Calculates order completion estimates.
 *
 * @param {Object[]} orderItems
 * @returns {Object}
 */
function calculateOrderCompletion(orderItems) {
	if (!Array.isArray(orderItems) || orderItems.length === 0) {
		throw new Error('Please add at least one product to the order.');
	}

	const stockMap = getCurrentStockMap();
	const mouldMap = getCurrentProductMouldMap();
	const productionRates = getSevenDayProductionRates();

	const results = orderItems.map((item) => {
		const productId = String(item.productId || '').trim();
		const orderQty = Number(item.orderQty || 0);

		if (!productId) {
			throw new Error('Product is required.');
		}

		if (orderQty <= 0) {
			throw new Error('Order quantity must be greater than zero.');
		}

		const currentStock = Number(stockMap[productId] || 0);
		const shortfall = Math.max(orderQty - currentStock, 0);

		const currentMoulds = Number(mouldMap[productId]?.moulds || 0);
		const productName =
			mouldMap[productId]?.productName || getProductName(productId);

		const rate = productionRates[productId];

		let productionPerDay = 0;
		let daysRequired = 0;
		let status = '';

		if (shortfall === 0) {
			status = 'Ready from stock';
		} else if (currentMoulds <= 0) {
			status = 'No current moulds configured';
		} else if (!rate || rate.averagePiecesPerMouldPerDay <= 0) {
			status = 'No production data in last 7 days';
		} else {
			productionPerDay = rate.averagePiecesPerMouldPerDay * currentMoulds;

			if (productionPerDay > 0) {
				daysRequired = Math.ceil(shortfall / productionPerDay);
			}
		}

		return {
			productId,
			productName,
			orderQty,
			currentStock,
			shortfall,
			currentMoulds,
			historicalAveragePiecesPerMouldPerDay:
				rate?.averagePiecesPerMouldPerDay || 0,
			expectedProductionPerDay: productionPerDay,
			daysRequired,
			status,
		};
	});

	const incompleteResults = results.filter(
		(item) => item.shortfall > 0 && item.expectedProductionPerDay <= 0,
	);

	let overallDays = 0;

	if (incompleteResults.length === 0) {
		overallDays = Math.max(...results.map((item) => item.daysRequired), 0);
	}

	const estimatedCompletionDate = new Date();
	estimatedCompletionDate.setHours(0, 0, 0, 0);
	estimatedCompletionDate.setDate(
		estimatedCompletionDate.getDate() + overallDays,
	);

	return {
		success: true,
		results,
		overallDays,
		estimatedCompletionDate: estimatedCompletionDate
			.toISOString()
			.split('T')[0],
		canComplete: incompleteResults.length === 0,
	};
}

/**
 * Returns current configured mould count by product.
 */
function getCurrentProductMouldMap() {
	const mouldMap = {};

	const machines = getMachines();

	machines.forEach((machine) => {
		const configuration = getMachineConfiguration(machine.id);

		configuration.forEach((item) => {
			const productId = item.productId;
			const moulds = Number(item.mould || 0);

			if (!productId) {
				return;
			}

			if (!mouldMap[productId]) {
				mouldMap[productId] = {
					productName: item.productName || '',
					moulds: 0,
				};
			}

			mouldMap[productId].moulds += moulds;
		});
	});

	return mouldMap;
}

/**
 * Calculates average pieces per mould per production day
 * over the last 7 calendar days.
 *
 * Days with no production for a product are not treated as
 * zero-production days.
 */
function getSevenDayProductionRates() {
	const headerSheet = getSheet(SHEETS.DAILY_HEADER);
	const detailSheet = getSheet(SHEETS.DAILY_DETAIL);

	const headerValues = headerSheet.getDataRange().getValues();
	const detailValues = detailSheet.getDataRange().getValues();

	if (headerValues.length <= 1 || detailValues.length <= 1) {
		return {};
	}

	const headerMap = buildColumnMap(headerValues[0]);
	const detailMap = buildColumnMap(detailValues[0]);

	const today = new Date();
	today.setHours(0, 0, 0, 0);

	const fromDate = new Date(today);
	fromDate.setDate(fromDate.getDate() - 6);

	const validEntries = {};

	for (let i = 1; i < headerValues.length; i++) {
		const row = headerValues[i];

		const entryId = row[headerMap['Entry ID']];
		const productionDate = new Date(row[headerMap['Production Date']]);

		if (!entryId || isNaN(productionDate.getTime())) {
			continue;
		}

		productionDate.setHours(0, 0, 0, 0);

		if (productionDate < fromDate || productionDate > today) {
			continue;
		}

		validEntries[entryId] = true;
	}

	/*
	 * Aggregate each product by production day.
	 *
	 * Example:
	 * 01 Sep:
	 *   10 moulds → 3,000 pieces
	 *
	 * 02 Sep:
	 *   12 moulds → 3,600 pieces
	 *
	 * Daily rates:
	 *   300 pieces/mould
	 *   300 pieces/mould
	 *
	 * Average:
	 *   300 pieces/mould/day
	 */
	const dailyProductData = {};

	for (let i = 1; i < detailValues.length; i++) {
		const row = detailValues[i];

		const entryId = row[detailMap['Entry ID']];

		if (!validEntries[entryId]) {
			continue;
		}

		const productId = row[detailMap['Product ID']];
		const productName = row[detailMap['Product Name']];
		const moulds = Number(row[detailMap['Mould']] || 0);
		const pieces = Number(row[detailMap['Pieces']] || 0);

		if (!productId || moulds <= 0) {
			continue;
		}

		const headerRow = headerValues.find(
			(headerRow) => headerRow[headerMap['Entry ID']] === entryId,
		);

		if (!headerRow) {
			continue;
		}

		const productionDate = new Date(headerRow[headerMap['Production Date']]);

		const dateKey = Utilities.formatDate(
			productionDate,
			Session.getScriptTimeZone(),
			'yyyy-MM-dd',
		);

		if (!dailyProductData[productId]) {
			dailyProductData[productId] = {
				productName,
				days: {},
			};
		}

		if (!dailyProductData[productId].days[dateKey]) {
			dailyProductData[productId].days[dateKey] = {
				moulds: 0,
				pieces: 0,
			};
		}

		dailyProductData[productId].days[dateKey].moulds += moulds;
		dailyProductData[productId].days[dateKey].pieces += pieces;
	}

	const rates = {};

	Object.entries(dailyProductData).forEach(([productId, productData]) => {
		const dailyRates = [];

		Object.values(productData.days).forEach((day) => {
			if (day.moulds <= 0) {
				return;
			}

			dailyRates.push(day.pieces / day.moulds);
		});

		if (dailyRates.length === 0) {
			return;
		}

		const average =
			dailyRates.reduce((sum, value) => sum + value, 0) / dailyRates.length;

		rates[productId] = {
			productName: productData.productName,
			averagePiecesPerMouldPerDay: average,
			productionDays: dailyRates.length,
		};
	});

	return rates;
}

/**
 * Returns a product name from the product master.
 */
function getProductName(productId) {
	const products = getAllProducts();

	const product = products.find(
		(item) => String(item.id) === String(productId),
	);

	return product ? product.name : productId;
}

/**
 * Builds a header-name-to-index map.
 */
function buildColumnMap(headers) {
	const map = {};

	headers.forEach((header, index) => {
		map[String(header).trim()] = index;
	});

	return map;
}
