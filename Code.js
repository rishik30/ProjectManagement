/**
 * ===========================================================
 * Lakshay Industries Production Management System
 * -----------------------------------------------------------
 * File        : Code.gs
 * Description : Global constants and initialization
 * ===========================================================
 */

const APP_NAME = 'Lakshay Industries Production System';

const SHEETS = {
	README: 'README',
	SETTINGS: 'Settings',
	USERS: 'Users',
	PRODUCT_MASTER: 'Product_Master',
	MACHINE_MASTER: 'Machine_Master',
	CUSTOMER_MASTER: 'Customer_Master',
	MACHINE_CONFIGURATION: 'Machine_Configuration',

	MACHINE_HISTORY: '_Machine_Item_History',
	DAILY_HEADER: '_Daily_Production_Header',
	DAILY_DETAIL: '_Daily_Production_Detail',
	SALES_HEADER: 'Sales_Header',
	SALES_DETAILS: 'Sales_Details',

	MACHINE_REPORT: 'Machine_Report',
	DASHBOARD: 'Dashboard',
	VALIDATION: 'Validation',
	REPORT: 'Report',
	STOCK_LEDGER: 'StockLedger',
	CURRENT_STOCK: 'CurrentStock',
	STAGE_STOCK: 'StageStock',
	BUNDLE_BOM: 'Bundle_BOM',
};

const SIDEBAR = {
	TITLE: APP_NAME,
	WIDTH: 420,
};

const MACHINE_MASTER_COLUMNS = {
	ID: 1,
	NAME: 2,
};

const USER_MASTER_COLUMNS = {
	ID: 1,
	NAME: 2,
};

const PRODUCT_MASTER_COLUMNS = {
	ID: 1,
	NAME: 2,
};
