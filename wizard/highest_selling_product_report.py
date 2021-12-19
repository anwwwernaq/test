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
import operator
from odoo.exceptions import ValidationError


class HighestSellingProductReport(models.TransientModel):
	_name = 'highest.selling.product.report'
	_description = 'Highest Selling Product Report'

	report_type = fields.Selection([('basic', 'Basic'), ('compare', 'Compare')],
								   string='Report Type', default='basic')
	from_date = fields.Date(string="From Date")
	compare_from_date = fields.Date(string="Compare From Date")
	to_date = fields.Date(string="To Date")
	compare_to_date = fields.Date(string="Compare To Date")
	no_item = fields.Integer(string="No Of Item", required=True, default=10)
	total_qty_sold = fields.Float(string="Total Qty.Sold")
	sales_channel_ids = fields.Many2one('crm.team', string='Sales Channel')
	company_ids = fields.Many2many('res.company', string='Companies')
	file_name = fields.Char('Excel File', readonly=True)
	data = fields.Binary(string="File")
	basic_purchase_orders = fields.Many2many('sale.order','product_basic_sale_orders')
	compare_purchase_orders = fields.Many2many('sale.order', 'product_compare_sale_orders')

	@api.onchange('report_type')
	def report_type_selected(self):
		if self.report_type != 'compare':
			self.compare_from_date = False
			self.compare_to_date = False

	@api.onchange('report_type')
	def onchange_partner_id(self):
		for rec in self:
			return {'domain': {'company_ids': [('id', 'in', self.env.user.company_ids.ids)]}}

	def update_top_selling_pdf_report(self):
		from_date = self.from_date
		to_date = self.to_date
		if to_date < from_date:
			raise ValidationError('End Date should be greater then Start Date')
		if self.report_type == 'compare':
			compare_from_date = self.compare_from_date
			compare_to_date = self.compare_to_date
			if compare_to_date < compare_from_date:
				raise ValidationError('End Date should be greater then Start Date')

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
		basic_purchase_orders = self.env['sale.order'].search([('date_order', '>=', self.from_date), ('date_order', '<=', self.to_date), ('company_id', 'in', selected_companies),('team_id','in',selected_channel),('state', '=', 'sale')])
		self.basic_purchase_orders = [(6, 0, basic_purchase_orders.ids)]

		if self.report_type == 'compare':
			compare_purchase_orders = self.env['sale.order'].search([('date_order', '>=', self.compare_from_date), ('date_order', '<=', self.compare_to_date), ('company_id', 'in', selected_companies),('team_id','in',selected_channel),('state', '=', 'sale')])
			self.compare_purchase_orders = [(6, 0, compare_purchase_orders.ids)]

		return self.env.ref('bi_all_in_one_sale_reports.highest_selling_product_report_action').report_action(self.id)

	def Sort(self,sub_li):
		l = len(sub_li)
		for i in range(0, l):
			for j in range(0, l-i-1):
				if (sub_li[j][1] > sub_li[j + 1][1]):
					tempo = sub_li[j]
					sub_li[j]= sub_li[j + 1]
					sub_li[j + 1]= tempo
		sub_li.sort(key=lambda element:element[1], reverse=True)
		return sub_li

	def set_table_values(self):
		basic_product_purchase = []
		compare_product_purchase = []
		new_products = []
		lost_products = []
		companies = self.company_ids.ids
		if len(companies) > 0:
			selected_companies = companies
		else:
			selected_companies = [self.env.company.id]
		channel = self.sales_channel_ids.ids
		if len(channel) > 0:
			selected_channel = channel
		else:
			channel_all = self.env['crm.team'].search([]).ids
			selected_channel = channel_all
		basic_purchase_orders = self.env['sale.order'].search(
			[('date_order', '>=', self.from_date), ('date_order', '<=', self.to_date),
			 ('company_id', 'in', selected_companies), ('team_id', 'in', selected_channel), ('state', '=', 'sale')])
		self.basic_purchase_orders = [(6, 0, basic_purchase_orders.ids)]

		basic_product_purchase = self.Sort([i for i in self.get_product_data(self.basic_purchase_orders) if i[1] >= self.total_qty_sold])[0:self.no_item]

		if self.report_type == 'compare':
			compare_purchase_orders = self.env['sale.order'].search(
				[('date_order', '>=', self.compare_from_date), ('date_order', '<=', self.compare_to_date),
				 ('company_id', 'in', selected_companies), ('team_id', 'in', selected_channel), ('state', '=', 'sale')])
			self.compare_purchase_orders = [(6, 0, compare_purchase_orders.ids)]

			compare_product_purchase = self.Sort([i for i in self.get_product_data(self.compare_purchase_orders) if i[1] >= self.total_qty_sold])[0:self.no_item]

			basic_purchase_list = [i[0] for i in basic_product_purchase]
			compare_purchase_list = [i[0] for i in compare_product_purchase]

			for i in compare_product_purchase:
				if i[0] not in basic_purchase_list:
					lost_products.append(i)

			for i in basic_product_purchase:
				if i[0] not in compare_purchase_list:
					new_products.append(i[0])

		return {'basic':basic_product_purchase,'compare':compare_product_purchase,'new':new_products,'lost':lost_products}

	def get_product_data(self, purchase_orders):
		product_list = list()
		total_list = ['Total']
		products = list()

		for rec in purchase_orders:
			for product in rec.order_line:
				if product.product_id.name_get()[0][1] not in [i[0] for i in product_list] and product.qty_delivered>=1:
					product_list.append([product.product_id.name_get()[0][1], product.product_uom_qty,product.product_id])
				elif product.product_id.name_get()[0][1] in [i[0] for i in product_list] and product.qty_delivered>=1:
					for i in product_list:
						if product.product_id.name_get()[0][1] == i[0]:
							product_list[product_list.index(i)][1] += product.product_uom_qty

		return product_list

	def update_top_selling_xls_report(self):
		if self.to_date < self.from_date:
			raise ValidationError('End Date should be greater then Start Date')
		if self.report_type == 'compare':
			if self.compare_to_date < self.compare_from_date:
				raise ValidationError('End Date should be greater then Start Date')
		workbook = xlwt.Workbook()
		worksheet = workbook.add_sheet('Highest Selling Product Report')
		worksheet.col(0).width = 4000
		worksheet.col(1).width = 8000
		worksheet.col(2).width = 4000
		worksheet.col(4).width = 4000
		worksheet.col(5).width = 8000
		worksheet.col(6).width = 4000

		style_header = xlwt.easyxf(
			"font:height 400; font: name Liberation Sans, bold on,color black; align: vert centre, horiz center;pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading = xlwt.easyxf(
			"font: name Liberation Sans, bold on; pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading_right = xlwt.easyxf(
			"font: name Liberation Sans, bold on;align: horiz right; pattern: pattern solid, pattern_fore_colour gray25;")

		style_line_left = xlwt.easyxf("align: horiz left")
		style_line_right = xlwt.easyxf("align: horiz right")
		data = self.set_table_values()
		if self.report_type == 'basic':
			row = 2
			worksheet.write_merge(0, 1, 0, 2, "Highest Selling Product Report", style=style_header)
			row+=1
			worksheet.write_merge(row, row, 0, 2, 'Companies: '+str(self.company_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			row += 1
			worksheet.write(4, 0, 'Start Date: '+str(self.from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			row += 1
			worksheet.write(4, 2, 'End Date: '+str(
				self.to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz left;"))
			worksheet.write(5,2, 'Sales Channel: '+str(self.channel_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))

			row += 2
			worksheet.write(row, 0, '#', style=style_line_heading)
			worksheet.write(row, 1, 'Product', style=style_line_heading)
			worksheet.write(row, 2, 'Qty Sold', style=style_line_heading_right)
			row += 1
			count = 0
			for value in data['basic']:
				count += 1
				worksheet.write(row, 0, count, style=style_line_left)
				worksheet.write(row, 1, value[0])
				worksheet.write(row, 2, value[1], style=style_line_right)
				row += 1
			row += 1
		if self.report_type == 'compare':
			row = 2
			worksheet.write_merge(0, 1, 0, 6, "Highest Selling Product Report", style=style_header)
			row += 1
			worksheet.write_merge(row, row, 0, 6, 'Companies: '+str(self.company_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			row += 1
			row += 1
			worksheet.write(row, 0, 'From Date: '+str(self.from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(row, 6, 'Compare From Date: '+str(self.compare_from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			row += 1
			worksheet.write(row, 0, 'To Date: '+str(self.to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(row, 6, 'Compare To Date: '+str(self.compare_to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			row += 1
			worksheet.write(row, 6, 'Sales Channel: '+str(self.channel_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			row += 2
			worksheet.write(row, 0, '#', style=style_line_heading)
			worksheet.write(row, 1, 'Product', style=style_line_heading)
			worksheet.write(row, 2, 'Qty Sold', style=style_line_heading_right)
			worksheet.write(row, 4, '#', style=style_line_heading)
			worksheet.write(row, 5, 'Product', style=style_line_heading)
			worksheet.write(row, 6, 'Qty Sold', style=style_line_heading_right)

			row += 1
			index = row
			count = 1
			for i in data['basic']:
				worksheet.write(row,0, count)
				worksheet.write(row,1, i[0])
				worksheet.write(row,2, i[1])
				count += 1
				row += 1
			count= 1

			for i in data['compare']:
				worksheet.write(index,4, count)
				worksheet.write(index,5, i[0])
				worksheet.write(index,6, i[1])
				count += 1
				index += 1

			if (len(data['compare'])) > len(data['basic']):
				row = index
			else:
				row = row
			row+=2
			count=0
			index = row
			worksheet.write_merge(row,row,0, 1, 'New Product', style=style_line_heading)
			worksheet.write_merge(row,row,4, 5, 'Lost Product', style=style_line_heading)
			row+=1
			for i in data['new']:
				worksheet.write_merge(row, row, 0, 1,i)
				count += 1
				row += 1
			index+=1
			for i in data['lost']:
				worksheet.write_merge(index, index, 4, 5,i[0])
				index +=1

		tz = pytz.timezone('Asia/Kolkata')
		file_data = BytesIO()
		workbook.save(file_data)

		self.write({
			'data': base64.encodestring(file_data.getvalue()),
			'file_name': 'Highest Selling Product Report.xls'
		})
		action = {
			'type': 'ir.actions.act_url',
			'name': 'contract',
			'url': '/web/content/highest.selling.product.report/%s/data/Highest Selling Product Report.xls?download=true' % (self.id),
			'target': 'self',
		}
		return action

	def company_record(self):
		comp_name = []
		if self.company_ids:
			for comp in self.company_ids:
				comp_name.append(comp.name)
		else:
			self.company_ids = [self.env.company.id]
			comp_name.append(self.company_ids.name)

		listtostr = ', '.join([str(elem) for elem in comp_name])
		return listtostr

	def channel_record(self):
		channel_name = []
		for channel in self.sales_channel_ids:
			channel_name.append(channel.name)
		listtostr = ', '.join([str(elem) for elem in channel_name])
		return listtostr


class HighestSellingProduct(models.TransientModel):
	_name = "highest.selling.product"
	_description = "Highest Selling Product"


	product_id = fields.Many2one('product.product', 'Product')
	product_uom_qty = fields.Float('Qty Sold')

	def filter_all_product_record(self):
		all_sale_orders = self.env['sale.order'].search(
			[('state', '=', 'sale')])
		product_list_data = list()
		for record in all_sale_orders:
			for product in record.order_line:
				if product.product_id.name_get()[0][1] not in [i[0] for i in product_list_data]:
					product_list_data.append(
						[product.product_id.name_get()[0][1], product.product_uom_qty,
						 product.product_id])
				elif product.product_id.name_get()[0][1] in [i[0] for i in product_list_data]:
					for i in product_list_data:
						if product.product_id.name_get()[0][1] == i[0]:
							product_list_data[product_list_data.index(i)][
								1] += product.product_uom_qty
		product_list_data.sort(key=lambda product_list_data: product_list_data[1], reverse=True)

		return product_list_data

	def product_details(self):
		product_ids = []
		sale_records = self.filter_all_product_record()
		for entry in sale_records:
			record_id = self.env['highest.selling.product'].create(
				{'product_id': entry[2].id, 'product_uom_qty': entry[1]})
			product_ids.append(record_id.id)

		domain = [('id', 'in', product_ids)]
		return {
			'name': 'Highest Selling Products',
			'type': 'ir.actions.act_window',
			'view_type': 'tree',
			'view_mode': 'list,form',
			'domain': domain,
			'res_model': 'highest.selling.product',
			'target': 'current'
		}


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
