/**
 * Creates or refreshes the README sheet.
 */
function generateReadme() {
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	let sheet = ss.getSheetByName('README');
	if (!sheet) {
		sheet = ss.insertSheet('README', 0);
	} else {
		sheet.clear();
		sheet.clearFormats();
	}

	sheet.setColumnWidth(1, 1000);
	sheet.setHiddenGridlines(true);

	let row = 1;

	function title(text) {
		sheet
			.getRange(row, 1)
			.setValue(text)
			.setFontSize(22)
			.setFontWeight('bold')
			.setHorizontalAlignment('center')
			.setBackground('#1F4E78')
			.setFontColor('white');

		sheet.setRowHeight(row, 35);
		row += 2;
	}

	function heading(text) {
		sheet
			.getRange(row, 1)
			.setValue(text)
			.setFontWeight('bold')
			.setFontSize(16)
			.setBackground('#D9EAD3');

		row++;
	}

	function subHeading(text) {
		sheet
			.getRange(row, 1)
			.setValue(text)
			.setFontWeight('bold')
			.setFontSize(13)
			.setFontColor('#1F4E78');

		row++;
	}

	function paragraph(text) {
		sheet
			.getRange(row, 1)
			.setValue(text)
			.setWrap(true)
			.setVerticalAlignment('top');

		row++;
	}

	function bullets(items) {
		items.forEach((item) => {
			sheet
				.getRange(row, 1)
				.setValue('• ' + item)
				.setWrap(true);

			row++;
		});

		row++;
	}

	function workflow(items) {
		items.forEach((item, index) => {
			sheet
				.getRange(row, 1)
				.setValue(item)
				.setHorizontalAlignment('center')
				.setFontWeight('bold');

			row++;

			if (index !== items.length - 1) {
				sheet
					.getRange(row, 1)
					.setValue('↓')
					.setHorizontalAlignment('center')
					.setFontSize(18);

				row++;
			}
		});

		row++;
	}

	function divider() {
		sheet.getRange(row, 1).setBorder(true, false, false, false, false, false);

		row++;
	}

	//--------------------------------------------------------------------
	// TITLE
	//--------------------------------------------------------------------

	title('PRODUCTION MANAGEMENT SYSTEM');
	paragraph('User Guide & Training Manual');

	divider();

	//--------------------------------------------------------------------
	// WELCOME
	//--------------------------------------------------------------------

	heading('Welcome');

	paragraph(
		'Welcome to the Production Management System. ' +
			'This application has been designed to manage the complete production and sales workflow of a manufacturing unit. ' +
			'It enables you to maintain products, record production, manage inventory, create sales entries, monitor stock levels and generate reports from one place.',
	);

	divider();

	//--------------------------------------------------------------------
	// WORKFLOW
	//--------------------------------------------------------------------

	heading('System Workflow');

	workflow([
		'Master Setup',
		'Opening Stock',
		'Production Entry',
		'Stock Updated Automatically',
		'Sales Entry',
		'Stock Deducted Automatically',
		'Reports & Dashboard',
	]);

	divider();

	//--------------------------------------------------------------------
	// BEFORE STARTING
	//--------------------------------------------------------------------

	heading('Before You Start');

	bullets([
		'Add Products',
		'Add Customers',
		'Verify Settings',
		'Add Opening Stock (if applicable)',
	]);

	divider();

	//--------------------------------------------------------------------
	// MODULES
	//--------------------------------------------------------------------

	heading('Modules');

	const modules = [
		{
			title: 'Dashboard',
			desc: 'Provides an overview of current production, stock, sales and recent activities. Use this page only for monitoring.',
			points: [
				'Current Stock',
				'Production Summary',
				'Sales Summary',
				'Quick Navigation',
			],
		},

		{
			title: 'Product Master',
			desc: 'Stores every finished product manufactured by the company.',
			points: [
				'Each product should be entered only once.',
				'Do not delete products already used in transactions.',
			],
		},

		{
			title: 'Customer Master',
			desc: 'Stores customer information used during sales.',
			points: ['Customer Name', 'Address', 'Contact Details', 'GST Number'],
		},

		{
			title: 'Production Module',
			desc: 'Records production completed in the factory. Saving a production entry automatically increases stock.',
			points: [
				'Select Product',
				'Enter Quantity',
				'Save Entry',
				'Stock Updated Automatically',
			],
		},

		{
			title: 'Sales Module',
			desc: 'Records sales made to customers. Saving a sale automatically reduces stock.',
			points: [
				'Select Customer',
				'Select Product',
				'Enter Quantity',
				'Save Entry',
			],
		},

		{
			title: 'Stock Management',
			desc: 'Displays current stock of every product.',
			points: ['Opening Stock', 'Production', 'Sales', 'Stock Adjustments'],
		},

		{
			title: 'Stock Adjustment',
			desc: 'Used only for correcting stock differences or entering opening stock.',
			points: [
				'Opening Stock',
				'Physical Verification',
				'Damaged Goods',
				'Stock Correction',
			],
		},

		{
			title: 'Reports',
			desc: 'Automatically generated reports for Production, Sales and Inventory.',
			points: ['Do not edit report sheets.'],
		},
	];

	modules.forEach((module) => {
		subHeading(module.title);

		paragraph(module.desc);

		bullets(module.points);
	});

	divider();

	//--------------------------------------------------------------------
	// STOCK CALCULATION
	//--------------------------------------------------------------------

	heading('Understanding Stock');

	paragraph('Current Stock =');

	sheet
		.getRange(row, 1)
		.setValue('Opening Stock + Production - Sales ± Stock Adjustments')
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBackground('#FFF2CC');

	row += 2;

	divider();

	//--------------------------------------------------------------------
	// DAILY WORKFLOW
	//--------------------------------------------------------------------

	heading('Recommended Daily Workflow');

	subHeading('Start of the Day');

	bullets(['Check Dashboard', 'Verify Stock']);

	subHeading('During Production');

	bullets(['Record Production', 'Verify Quantity']);

	subHeading('During Sales');

	bullets(['Create Sales Entry', 'Verify Customer', 'Verify Quantity']);

	subHeading('End of the Day');

	bullets(['Verify Stock', 'Review Reports', 'Create Backup']);

	divider();

	//--------------------------------------------------------------------
	// IMPORTANT RULES
	//--------------------------------------------------------------------

	heading('Important Rules');

	bullets([
		'Never delete transaction rows.',
		'Never edit formulas.',
		'Never edit hidden sheets.',
		'Never modify generated IDs.',
		'Always use application buttons.',
	]);

	divider();

	//--------------------------------------------------------------------
	// COMMON MISTAKES
	//--------------------------------------------------------------------

	heading('Common Mistakes');

	bullets([
		'Editing formulas',
		'Deleting rows',
		'Creating duplicate products',
		'Using Stock Adjustment instead of Production',
		'Incorrect quantities',
		'Editing reports manually',
	]);

	divider();

	//--------------------------------------------------------------------
	// BACKUP
	//--------------------------------------------------------------------

	heading('Backup Recommendation');

	paragraph('Create a backup regularly using:');

	sheet
		.getRange(row, 1)
		.setValue('File → Make a Copy')
		.setFontWeight('bold')
		.setHorizontalAlignment('center');

	row += 2;

	divider();

	//--------------------------------------------------------------------
	// VERSION
	//--------------------------------------------------------------------

	heading('Version Information');

	paragraph('Application : Production Management System');
	paragraph('Version : v1.0');
	paragraph('Developed By : Your Company');
	paragraph('Last Updated : 31 July 2026');

	sheet.setFrozenRows(1);
	sheet.autoResizeRows(1, sheet.getLastRow());
}
