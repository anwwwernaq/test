# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import xlwt
import base64
import pytz, tempfile
import io
from io import StringIO
from io import BytesIO
from functools import reduce
from odoo import models, fields, api, _
from odoo.tools.misc import xlwt
from datetime import date, datetime, timedelta
import json
from odoo.exceptions import ValidationError
from datetime import datetime


class ProductsalesSummaryReport(models.TransientModel):
	_name = 'product.sales.summary.report'
	_description = 'Product Sales Summary Report'

	start_date = fields.Date(string="Start Date")
	end_date = fields.Date(string="End Date")
	select_state = fields.Selection([
		('all', 'All'),
		('done', 'Done'),
	], string='Status', default='all')
	sales_channel_ids = fields.Many2many('crm.team', string='Sales Channel')
	company_ids = fields.Many2many('res.company', string='Companies')
	excel_file = fields.Binary('Sale Excel Report')
	file_name = fields.Char('Excel File', readonly=True)
	data = fields.Binary(string="File")


	def company_record(self):
		comp_name = []
		for comp in self.company_ids:
			comp_name.append(comp.name)
		listtostr = ', '.join([str(elem) for elem in comp_name])
		return listtostr


	def channel_record(self):
		channel_name = []
		for channel in self.sales_channel_ids:
			channel_name.append(channel.name)
		listtostr = ', '.join([str(elem) for elem in channel_name])
		return listtostr


	def product_sales_summary_pdf_report(self):
		start_date = self.start_date
		end_date = self.end_date
		if end_date < start_date:
			raise ValidationError('Enter End Date greater then Start Date')
		companies = self.company_ids.ids
		if len(companies) > 0:
			selected_companies = companies
		else:
			selected_companies = self.env.user.company_ids.ids

		channel = self.sales_channel_ids.ids
		if len(channel) > 0:
			selected_channel = channel
		else:
			channel_all = self.env['crm.team'].search([]).ids
			selected_channel = channel_all

		final_data = {}
		state = []
		if self.select_state == 'all':
			state.extend(['draft', 'sent', 'sale', 'done'])
		elif self.select_state == 'done':
			state.extend(['sale', 'done'])
		elif self.select_state == False:
			state.extend(['draft', 'sent', 'sale', 'done'])

		status = ('state', 'in', state)
		sale_ids = self.env['sale.order'].search([('date_order', '>=', start_date),
												  ('date_order', '<=', end_date),
												  ('company_id', 'in', selected_companies),
												  ('team_id', 'in', selected_channel), status
												  ])
		count_total = 0
		list1 = []
		total_payment = {'Bank': 0, 'Cash': 0}
		all_tax = {}
		for product in sale_ids:
			invoice_payments = self.env['account.move'].search(
				[('id', 'in', product.invoice_ids.ids)
				 ])
			for line in product.order_line:
				list1.append([product.name,line.product_id.name, line.product_uom_qty, line.price_unit])
				count_total += line.product_uom_qty * line.price_unit

				price = line.price_unit * (1 - (line.discount or 0.0) / 100.0)
				taxes = line.tax_id.compute_all(price, line.order_id.currency_id,
												line.product_uom_qty, product=line.product_id,
												partner=line.order_id.partner_shipping_id)

				for tax in taxes.get('taxes', []):
					if tax.get('name') not in all_tax:
						all_tax.update({tax.get('name'): tax.get('amount', 0)})
					else:
						all_tax[tax.get('name')] += tax.get('amount', 0)

			if invoice_payments:
				for invoice in invoice_payments:
					if invoice.invoice_payments_widget != 'false':
						bank_amount = 0
						cash_amount = 0
						res = json.loads(invoice.invoice_payments_widget)
						for i in res['content']:
							if i['journal_name'] == 'Bank':
								bank_amount += i['amount']
							if i['journal_name'] == 'Cash':
								cash_amount += i['amount']
						total_payment['Bank'] += bank_amount
						total_payment['Cash'] += cash_amount

		if self.select_state == False:
			final_data.update({'date': [self.start_date, self.end_date, self.company_record(),
										self.select_state, self.channel_record()],
							   'sale_data': list1, 'payments': total_payment, 'taxes': all_tax})
		else:
			final_data.update({'date': [self.start_date, self.end_date, self.company_record(),
										self.select_state.capitalize(), self.channel_record()],
							   'sale_data': list1, 'payments': total_payment, 'taxes': all_tax})
		data = self.env.ref(
			'bi_all_in_one_sale_reports.action_sale_summary_report').report_action(self,
																				   data=final_data)
		return data


	def product_sales_summary_xls_report(self):
		workbook = xlwt.Workbook()
		stylePC = xlwt.XFStyle()
		worksheet = workbook.add_sheet('Product Sales Summary Report')
		bold = xlwt.easyxf("font: bold on; pattern: pattern solid, fore_colour gray25;")
		alignment = xlwt.Alignment()
		alignment.horz = xlwt.Alignment.HORZ_CENTER
		stylePC.alignment = alignment
		alignment = xlwt.Alignment()
		alignment.horz = xlwt.Alignment.HORZ_CENTER
		alignment_num = xlwt.Alignment()
		alignment_num.horz = xlwt.Alignment.HORZ_RIGHT
		horz_style = xlwt.XFStyle()
		horz_style.alignment = alignment_num
		align_num = xlwt.Alignment()
		align_num.horz = xlwt.Alignment.HORZ_RIGHT
		horz_style_pc = xlwt.XFStyle()
		horz_style_pc.alignment = alignment_num
		style1 = horz_style
		font = xlwt.Font()
		font1 = xlwt.Font()
		borders = xlwt.Borders()
		borders.bottom = xlwt.Borders.THIN
		font.bold = True
		font1.bold = True
		font.height = 400
		stylePC.font = font
		style1.font = font1
		stylePC.alignment = alignment
		pattern = xlwt.Pattern()
		pattern1 = xlwt.Pattern()
		pattern.pattern = xlwt.Pattern.SOLID_PATTERN
		pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
		pattern.pattern_fore_colour = xlwt.Style.colour_map['gray25']
		pattern1.pattern_fore_colour = xlwt.Style.colour_map['gray25']
		stylePC.pattern = pattern
		style1.pattern = pattern
		style_header = xlwt.easyxf(
			"font:height 300; font: name Liberation Sans, bold on,color black; align: vert centre, horiz center;pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading = xlwt.easyxf(
			"font: name Liberation Sans, bold on;align: horiz centre; pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading_left = xlwt.easyxf(
			"font: name Liberation Sans, bold on;align: horiz left; pattern: pattern solid, pattern_fore_colour gray25;")

		worksheet.write_merge(0, 1, 0, 3, 'Product Sales Summary Report', style=stylePC)
		worksheet.col(2).width = 5600
		worksheet.write_merge(2, 2, 0, 3, 'Companies: ' + str(self.company_record()),
							  style=xlwt.easyxf(
								  "font: name Liberation Sans, bold on; align: horiz center;"))
		worksheet.write(3, 0, 'Start Date: ' + str(self.start_date.strftime('%d-%m-%Y')),
						style=xlwt.easyxf(
							"font: name Liberation Sans, bold on;"))
		worksheet.write(3, 3, 'End Date: ' + str(
			self.end_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz left;"))
		if self.select_state == False:
			worksheet.write(4, 0, 'Status: ', style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
		else:
			worksheet.write(4, 0, 'Status: ' + str(self.select_state).capitalize(),
							style=xlwt.easyxf(
								"font: name Liberation Sans, bold on;"))
		worksheet.write(4, 3, 'Sales Channel: ' + str(self.channel_record()), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz left;"))

		row = 7
		worksheet.write_merge(6, 6, 0, 3, 'Products', style=style_line_heading)
		list1 = ['Order ref.','Product', 'Quantity', 'Price Unit']
		worksheet.col(0).width = 4000
		worksheet.write(row, 0, list1[0], style=style_line_heading_left)
		worksheet.col(1).width = 5000
		worksheet.write(row, 1, list1[1], style=style_line_heading_left)
		worksheet.col(2).width = 4000
		worksheet.write(row, 2, list1[2], style1)
		worksheet.col(3).width = 4000
		worksheet.write(row, 3, list1[3], style1)
		row += 1

		sale_records = self.product_sales_summary_pdf_report()
		if sale_records['context'].get('report_action') == None:
			sale_datas = sale_records['data']['sale_data']
		else:
			sale_datas = sale_records['context']['report_action']['data']['sale_data']
		count_total = 0
		for product in sale_datas:
			worksheet.write(row, 0, product[0])
			worksheet.write(row, 1, product[1])
			worksheet.write(row, 2, product[2])
			worksheet.write(row, 3, product[3])
			count_total += (product[2] * product[3])
			row = row + 1
		row += 1
		list2 = ['Name', 'Total']
		worksheet.write_merge(row, row, 0, 2, 'Payments', style=style_line_heading)
		row += 1
		worksheet.col(0).width = 5000
		worksheet.write(row, 0, list2[0], style=style_line_heading_left)
		worksheet.col(1).width = 5000
		worksheet.write(row, 1, '', style1)
		worksheet.col(2).width = 5000
		worksheet.write(row, 2, list2[1], style1)
		row += 1
		if sale_records['context'].get('report_action')==None:
			payment_record = sale_records['data']['payments']
		else:
			payment_record = sale_records['context']['report_action']['data']['payments']
		for pay in payment_record.items():
			worksheet.write(row, 0, pay[0])
			worksheet.write(row, 2, pay[1])
			row = row + 1
		row += 1
		worksheet.write_merge(row, row, 0, 2, 'Taxes', style=style_line_heading)
		row += 1
		worksheet.col(0).width = 5000
		worksheet.write(row, 0, list2[0], style=style_line_heading_left)
		worksheet.col(1).width = 5000
		worksheet.write(row, 1, '', style1)
		worksheet.col(2).width = 5000
		worksheet.write(row, 2, list2[1], style1)
		row += 1
		if sale_records['context'].get('report_action')==None:
			tax_data = sale_records['data']['taxes']
		else:
			tax_data = sale_records['context']['report_action']['data']['taxes']
		for data in tax_data.items():
			worksheet.write(row, 0, data[0])
			worksheet.write(row, 2, data[1])
			row = row + 1
		row += 2
		file_data = BytesIO()
		workbook.save(file_data)
		self.write({
			'data': base64.encodestring(file_data.getvalue()),
			'file_name': 'Product Sales Summary Report.xls'
		})
		action = {
			'type': 'ir.actions.act_url',
			'name': 'contract',
			'url': '/web/content/product.sales.summary.report/%s/data/Product Sales Summary Report.xls?download=true' % (
				self.id),
			'target': 'self',
		}
		return action


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4: