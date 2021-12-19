# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

{
	'name': 'All in One Sales Report in Odoo',
	'version': '14.0.1.2',
	'category': 'Sales',
	'summary': 'All sale order reports sale day wise report sale day book report sale payment report product sale summary report sale details report top customer report sales excel report sale xls report sale category report all in one sale order report all sales reports',
	'description' :"""  All in One sales order reports odoo app helps users to print sales day wise reports, payment report for customer invoice/sales, product sales summary reports, user wise sales details report, highest sales products report, top customer product report, sales day book report with particular date range for particular company in XLS and PDF format. Users can also print category wise sales order pdf reports and sale order excel reports for single or multiple sale orders.   """,
	'author': 'BrowseInfo',
	'website': 'https://www.browseinfo.in',
	"price": 75,
	"currency": 'EUR',
	'depends': ['base', 'sale_management','stock'],
	'data': [
		'security/sales_reports_security.xml',
		'security/ir.model.access.csv',

		'wizard/sales_day_wise_report_view.xml',
		'wizard/user_wise_sales_detail_report_view.xml',
		'wizard/product_sales_summary_report_view.xml',
		'wizard/customer_invoice_payment_report_view.xml',
		'wizard/top_customer_product_report_view.xml',
		'wizard/highest_selling_product_report_view.xml',
		'wizard/sale_excel_report_view.xml',
		'wizard/sale_book_day_report_view.xml',

		'report/sales_day_wise_report_template.xml',
		'report/user_wise_sales_detail_report_template.xml',
		'report/product_sales_summary_report_template.xml',
		'report/customer_invoice_payment_report_template.xml',
		'report/top_customer_product_report_template.xml',
		'report/highest_selling_product_report_template.xml',
		'report/sale_order_category_report_template.xml',
		'report/sale_book_day_report_template.xml',
		'report/bi_sales_reports.xml',

	],
	'installable': True,
	'auto_install': False,
	'live_test_url':'https://youtu.be/uLxmRnjsXAA',
	"images":['static/description/Banner.png'],
}

# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
