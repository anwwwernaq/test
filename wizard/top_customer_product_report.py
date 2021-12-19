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
from operator import itemgetter
from odoo.exceptions import ValidationError


class TopCustomerProductReport(models.TransientModel):
	_name = 'top.customer.product.report'
	_description = 'Top Customer Product Report'


	report_type = fields.Selection([('basic', 'Basic'), ('compare', 'Compare')],
								   string='Report Type', default='basic')
	from_date = fields.Date(string="From Date")
	compare_from_date = fields.Date(string="Compare From Date")
	to_date = fields.Date(string="To Date")
	compare_to_date = fields.Date(string="Compare To Date")
	no_item = fields.Integer(string="No Of Item", required=True, default=10)
	total_amount = fields.Float(string="Total Sale Amount")
	company_ids = fields.Many2many('res.company', string='Companies')
	sales_channel_ids = fields.Many2one('crm.team', string='Sales Channel')
	file_name = fields.Char('Excel File', readonly=True)
	data = fields.Binary(string="File")
	basic_sales_orders = fields.Many2many('sale.order','product_basic_sale_orders_vendors')
	compare_sales_orders = fields.Many2many('sale.order', 'product_compare_sale_orders_vendors')


	@api.onchange('report_type')
	def report_type_selected(self):
		if self.report_type != 'compare':
			self.compare_from_date = False
			self.compare_to_date = False


	@api.onchange('report_type')
	def onchange_partner_id(self):
		for rec in self:
			return {'domain': {'company_ids': [('id', 'in', self.env.user.company_ids.ids)]}}


	def top_customer_product_pdf_report(self):
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

		basic_sales_orders = self.env['sale.order'].search([('date_order', '>=', self.from_date), ('date_order', '<=', self.to_date), ('company_id', 'in', selected_companies),('team_id','in',selected_channel),])
		self.basic_sales_orders = [(6, 0, basic_sales_orders.ids)]

		if self.report_type == 'compare':
			compare_sales_orders = self.env['sale.order'].search([('date_order', '>=', self.compare_from_date), ('date_order', '<=', self.compare_to_date), ('company_id', 'in', selected_companies),('team_id','in',selected_channel),])
			self.compare_sales_orders = [(6, 0, compare_sales_orders.ids)]

		data = self.env.ref('bi_all_in_one_sale_reports.top_customer_product_report_action').report_action(self.id)
		return data


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


	def top_customer_product_xls_report(self):
		if self.to_date < self.from_date:
			raise ValidationError('End Date should be greater then Start Date')
		if self.report_type == 'compare':
			if self.compare_to_date < self.compare_from_date:
				raise ValidationError('End Date should be greater then Start Date')
		data = self.set_table_values()
		workbook = xlwt.Workbook()
		worksheet = workbook.add_sheet('Top Customer')
		worksheet.col(0).width = 5000
		worksheet.col(1).width = 7000
		worksheet.col(2).width = 4000
		worksheet.col(4).width = 5000
		worksheet.col(5).width = 7000
		worksheet.col(6).width = 4000
		style_header = xlwt.easyxf(
			"font:height 400; font: name Liberation Sans, bold on,color black; align: vert centre, horiz center;pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_heading = xlwt.easyxf(
			"font: name Liberation Sans, bold on; pattern: pattern solid, pattern_fore_colour gray25;")
		style_line_left = xlwt.easyxf("align: horiz left")

		if self.report_type == 'basic':
			row = 2
			worksheet.write_merge(0, 1, 0, 2, "Top Customer Product Report", style=style_header)
			worksheet.write_merge(3, 3, 0, 2, 'Companies: '+str(self.company_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			worksheet.write(4, 0, 'Start Date: '+str(self.from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(4, 2, 'End Date: '+str(
			self.to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
			"font: name Liberation Sans, bold on; align: horiz left;"))
			worksheet.write(5, 2, 'Sales Channel: '+str(self.channel_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz left;"))
			row = 7
			worksheet.write(row, 0, '#', style=style_line_heading)
			worksheet.write(row, 1, 'Customer', style=style_line_heading)
			worksheet.write(row, 2, 'Sales Amount', style=style_line_heading)
			row += 1
			count = 0
			for value in data['basic']:
				count += 1
				worksheet.write(row, 0, count,style=style_line_left)
				worksheet.write(row, 1, value[0])
				worksheet.write(row, 2, round(value[1],4))
				row += 1
		if self.report_type == 'compare':
			row = 2
			worksheet.write_merge(0, 1, 0, 6, "Top Customer Product Report", style=style_header)
			worksheet.write_merge(3, 3, 0, 6, 'Companies: '+str(self.company_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			worksheet.write(4, 0, 'From Date: '+str(self.from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(4, 6, 'Compare From Date: '+str(self.compare_from_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(5, 0, 'To Date: '+str(self.to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(5, 6, 'Compare To Date: '+str(self.compare_to_date.strftime('%d-%m-%Y')), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on;"))
			worksheet.write(6, 6, 'Sales Channel: '+str(self.channel_record()), style=xlwt.easyxf(
				"font: name Liberation Sans, bold on; align: horiz center;"))
			row = 8
			worksheet.write(row, 0, '#', style=style_line_heading)
			worksheet.write(row, 1, 'Customer', style=style_line_heading)
			worksheet.write(row, 2, 'Sales Amount', style=style_line_heading)
			worksheet.write(row, 4, '#', style=style_line_heading)
			worksheet.write(row, 5, 'Customer', style=style_line_heading)
			worksheet.write(row, 6, 'Sales Amount', style=style_line_heading)
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
			worksheet.write(row, 0, 'New Customer', style=style_line_heading)
			worksheet.write(row, 4, 'Lost Customer', style=style_line_heading)
			row+=1
			for i in data['new']:
				worksheet.write_merge(row, row, 0, 1,i)
				count += 1
				row += 1
			index+=1
			for i in data['lost']:
				worksheet.write_merge(index, index, 4, 5,i)
				index +=1

		file_data = BytesIO()
		workbook.save(file_data)

		self.write({
			'data': base64.encodebytes(file_data.getvalue()),
			'file_name': 'Top Customer Product Report.xls'
		})
		action = {
			'type': 'ir.actions.act_url',
			'name': 'contract',
			'url': '/web/content/top.customer.product.report/%s/data/Top Customer Product Report.xls?download=true' % (self.id),
			'target': 'self',
		}
		return action


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
		basic_vendors = []
		compare_vendors = []
		new_vendors = []
		lost_vendors = []
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
		basic_sales_orders = self.env['sale.order'].search(
			[('date_order', '>=', self.from_date), ('date_order', '<=', self.to_date),
			 ('company_id', 'in', selected_companies), ('team_id', 'in', selected_channel), ])
		self.basic_sales_orders = [(6, 0, basic_sales_orders.ids)]

		basic_vendors = self.Sort([i for i in self.get_product_data(self.basic_sales_orders) if i[1] >= self.total_amount])[0:self.no_item]

		if self.report_type == 'compare':
			compare_sales_orders = self.env['sale.order'].search(
				[('date_order', '>=', self.compare_from_date), ('date_order', '<=', self.compare_to_date),
				 ('company_id', 'in', selected_companies), ('team_id', 'in', selected_channel), ])
			self.compare_sales_orders = [(6, 0, compare_sales_orders.ids)]

			compare_vendors = self.Sort([i for i in self.get_product_data(self.compare_sales_orders) if i[1] >= self.total_amount])[0:self.no_item]

			basic_vendors_list = [i[0] for i in basic_vendors]
			compare_vendors_list = [i[0] for i in compare_vendors]

			for i in compare_vendors:
				if i[0] not in basic_vendors_list:
					lost_vendors.append(i[0])

			for i in basic_vendors:
				if i[0] not in compare_vendors_list:
					new_vendors.append(i[0])

		return {'basic':basic_vendors,'compare':compare_vendors,'new':new_vendors,'lost':lost_vendors}


	def get_product_data(self, purchase_orders):
		vendor_list = list()
		total_list = ['Total']
		products = list()

		for rec in purchase_orders:
			if rec.partner_id.name_get()[0][1] not in [i[0] for i in vendor_list]:
				vendor_list.append([rec.partner_id.name_get()[0][1], rec.amount_total])

			elif rec.partner_id.name_get()[0][1] in [i[0] for i in vendor_list]:
				for i in vendor_list:
					if rec.partner_id.name_get()[0][1] == i[0]:
						vendor_list[vendor_list.index(i)][1] += rec.amount_total

		return vendor_list


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
