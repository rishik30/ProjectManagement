/**
 * ===========================================================
 * Lookup Engine
 * ===========================================================
 */

function buildLookupMap(list, idField, nameField) {
	const map = {};

	list.forEach((item) => {
		map[item[nameField]] = item[idField];
	});

	return map;
}

/**
 * ===========================================================
 * Dashboard Filter Resolver
 * ===========================================================
 */

function normalizeDashboardFilters(filters) {
	const resolvedRange =
		filters.dateRange === 'Custom'
			? {
					fromDate: filters.fromDate,
					toDate: filters.toDate,
				}
			: resolveDateRange(filters.dateRange);

	if (resolvedRange.fromDate > resolvedRange.toDate) {
		throw new Error('From Date cannot be after To Date.');
	}

	const machineLookup = buildLookupMap(getMachines(), 'id', 'name');

	const operatorLookup = buildLookupMap(getOperators(), 'id', 'name');

	const productLookup = buildLookupMap(getAllProducts(), 'id', 'name');

	return {
		fromDate: resolvedRange.fromDate,
		toDate: resolvedRange.toDate,
		machineId:
			filters.machineId === 'All'
				? 'All'
				: machineLookup[filters.machineId] || 'All',
		operatorId:
			filters.operatorId === 'All'
				? 'All'
				: operatorLookup[filters.operatorId] || 'All',
		productId:
			filters.productId === 'All'
				? 'All'
				: productLookup[filters.productId] || 'All',
	};
}
