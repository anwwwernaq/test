# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import xlwt
import base64
import pytz, tempfile
from io import BytesIO
from odoo import models, fields, api, _
from odoo.tools.misc import xlwt
from datetime import date, datetime, timedelta
from datetime import datetime
from odoo.exceptions import ValidationError
from odoo.addons import decimal_precision as dp
import calendar


class SalesDayWiseReport(models.TransientModel):
	_name = 'sales.day.wise.report'
	_description = 'Sales Day Wise Report'

	start_date = fields.Date(string="Start Date")
	end_date = fields.Date(string="End Date")
	company_ids = fields.Many2many('res.company', string='Companies')
	file_name = fields.Char('Excel File', readonly=True)
	data = fields.Binary(string="File")


	def sales_day_wise_pdf_report(self):
		start_date = self.start_date
		end_date = self.end_date
		if start_date > end_date:
			raise ValidationError(_("Please enter valid date range"))
		companies = self.company_ids.ids
		if len(companies) > 0:
			selected_companies = companies
		else:
			selected_companies = self.env.user.company_ids.ids
		sale_order = self.env['sale.order'].search([('date_order', '>=', start_date),
													('date_order', '<=', end_date),
													('company_id', 'in', selected_companies),
													])

		sale_order = sale_order.filtered(lambda s: s.state == "sale")
		data = {}
		for order in sale_order:
			day = calendar.weekday(order.date_order.date().year, order.date_order.date().month,
								   order.date_order.date().day)
			for line in order.order_line:
				if line.product_id.name in data:
					data[line.product_id.name][day] += int(line.product_uom_qty)
					data[line.product_id.name][7] += int(line.product_uom_qty)
				else:
					data[line.product_id.name] = [0, 0, 0, 0, 0, 0, 0, 0]
					data[line.product_id.name][day] = int(line.product_uom_qty)
					data[line.product_id.name][7] = int(line.product_uom_qty)

		day_total = [0, 0, 0, 0, 0, 0, 0]
		for i in range(0, 7):
			for product in data.keys():
				day_total[i] += data[product][i]

		data_all = {'data': [data, day_total, [start_date, end_date],self.company_record()],}
		record = self.env.ref('bi_all_in_one_sale_reports.sales_day_wise_report_action').report_action(self, data=data_all)
		return record


	def company_record(self):
		comp_name = []
		for comp in self.company_ids:
			comp_name.append(comp.name)
		listtostr = ', '.join([str(elem) for elem in comp_name])
		return listtostr


	def sales_day_wise_xls_report(self):
		workbook = xlwt.Workbook()
		worksheet = workbook.add_sheet('Sales Day Wise Report')
		worksheet.col(0).width = 8000
		style_header = xlwt.easyxf(
			"font:height 400; font: name Liberation Sans, bold on,color black; align: vert centre, horiz center;pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading = xlwt.easyxf(
			"font: name Liberation Sans, bold on; pattern: pattern solid, pattern_fore_colour gray25;")
		style_bold = xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz right;")
		worksheet.write_merge(0, 1, 0, 8, 'Sales Day Wise Report', style=style_header)
		worksheet.write_merge(2, 2, 0, 8, 'Companies: '+str(self.company_record()), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz center;"))
		worksheet.col(2).width = 5600
		row = 4
		list1 = ['Product Name', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday',
				 'Saturday', 'Sunday', 'Total', 'Start Date', 'End Date']
		worksheet.col(0).width = 5000
		worksheet.write(row, 0, 'Start Date: '+str(self.start_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on;"))
		worksheet.col(6).width = 5000
		worksheet.write(row, 8, 'End Date: '+str(
			self.end_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz left;"))
		row += 2
		worksheet.col(0).width = 5000
		worksheet.write(row, 0, list1[0], style=style_line_heading)
		worksheet.col(1).width = 5000
		worksheet.write(row, 1, list1[1], style=style_line_heading)
		worksheet.col(2).width = 5000
		worksheet.write(row, 2, list1[2], style=style_line_heading)
		worksheet.col(3).width = 5000
		worksheet.write(row, 3, list1[3], style=style_line_heading)
		worksheet.col(4).width = 5000
		worksheet.write(row, 4, list1[4], style=style_line_heading)
		worksheet.col(5).width = 5000
		worksheet.write(row, 5, list1[5], style=style_line_heading)
		worksheet.col(6).width = 5000
		worksheet.write(row, 6, list1[6], style=style_line_heading)
		worksheet.col(7).width = 5000
		worksheet.write(row, 7, list1[7], style=style_line_heading)
		worksheet.col(8).width = 5000
		worksheet.write(row, 8, list1[8], style=style_line_heading)
		row = row + 1
		sale_records = self.sales_day_wise_pdf_report()
		if sale_records['context'].get('report_action') == None:
			data = sale_records['data']
		else:
			data = sale_records['context']['report_action']['data']

		for i in data.values():
			order = i[0]
			total = [0, 0, 0, 0, 0, 0, 0, 0]
			for product in order:
				worksheet.write(row, 0, product)
				worksheet.write(row, 1, order[product][0])
				total[0] += order[product][0]
				worksheet.write(row, 2, order[product][1])
				total[1] += order[product][1]
				worksheet.write(row, 3, order[product][2])
				total[2] += order[product][2]
				worksheet.write(row, 4, order[product][3])
				total[3] += order[product][3]
				worksheet.write(row, 5, order[product][4])
				total[4] += order[product][4]
				worksheet.write(row, 6, order[product][5])
				total[5] += order[product][5]
				worksheet.write(row, 7, order[product][6])
				total[6] += order[product][6]
				worksheet.write(row, 8, order[product][7])
				total[7] += order[product][7]
				row = row + 1
			row += 1
			worksheet.write(row, 0, 'Total', style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			worksheet.write(row, 1, total[0], style=style_bold)
			worksheet.write(row, 2, total[1], style=style_bold)
			worksheet.write(row, 3, total[2], style=style_bold)
			worksheet.write(row, 4, total[3], style=style_bold)
			worksheet.write(row, 5, total[4], style=style_bold)
			worksheet.write(row, 6, total[5], style=style_bold)
			worksheet.write(row, 7, total[6], style=style_bold)
			worksheet.write(row, 8, total[7], style=style_bold)
		tz = pytz.timezone('Asia/Kolkata')
		file_data = BytesIO()
		workbook.save(file_data)
		self.write({
			'data': base64.encodestring(file_data.getvalue()),
			'file_name': 'Sales Day Wise Report.xls'
		})
		action = {
			'type': 'ir.actions.act_url',
			'name': 'contract',
			'url': '/web/content/sales.day.wise.report/%s/data/Sales Day Wise Report.xls?download=true' % (self.id),
			'target': 'self',
		}
		return action


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
