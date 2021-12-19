# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

import io
from io import BytesIO, StringIO
from functools import reduce
from odoo import fields, models, api, _
from datetime import datetime
import xlsxwriter
import base64
import xlwt
import csv
from PIL import Image
from odoo.exceptions import UserError, ValidationError


class SaleExcelReport(models.TransientModel):

	_name = "sale.excel.report"
	_description = "Sales Excel Reports"


	def sale_xls_report(self):
		sale_order = self.env['sale.order'].browse(self._context.get('active_ids',[]))
		filename = 'Sale Excel Report.xls'
		workbook = xlwt.Workbook()

		for sale in sale_order:			
			worksheet = workbook.add_sheet(sale.name)
			worksheet.col(0).width = 8000
			for i in range(1,9):
				worksheet.col(i).width = 3000
			date_format = xlwt.XFStyle()
			date_format.num_format_str = 'dd/mm/yyyy'
			heading_style = xlwt.easyxf("font:height 300; font: name Liberation Sans, bold on,color black; align: vert centre, horiz center;pattern: pattern solid, fore_colour gray50;")
			po_style = xlwt.easyxf("font: bold 1, height 280, color gray80; align: horiz center")
			heading_xf = xlwt.easyxf('font: bold on, color white; align: wrap on, vert centre, horiz center; pattern: pattern solid, pattern_fore_colour black;')
			tbl_header_style = xlwt.easyxf("font: bold 1, height 220, color white; align: horiz center; pattern: pattern solid, pattern_fore_colour black;")
			section_style = xlwt.easyxf("font: bold 1, height 220, color gray80; align: horiz center;")
			footer_style = xlwt.easyxf("font: bold 1, height 250, color gray80; align: horiz left; borders: left thin, right thin, top thin, bottom thin;")
			foo_style = xlwt.easyxf("font: bold 1, height 250, color gray80; align: horiz left; borders: left thin, right thin, top thin, bottom thin;")

			first_col = worksheet.col(0)
			one_col = worksheet.col(1)
			two_col = worksheet.col(2)
			three_col = worksheet.col(3)
			four_col = worksheet.col(4)
			five_col = worksheet.col(5)
			six_col = worksheet.col(6)

			first_col.width = 320*20
			one_col.width = 320*20
			two_col.width = 180*20
			three_col.width = 280*20
			four_col.width = 180*20
			five_col.width = 250*20
			six_col.width = 200*20

			worksheet.row(0).height_mismatch = True
			worksheet.row(0).height = 500
			worksheet.write_merge(0, 1, 0, 6, "Sale Excel Report", style= heading_style)

			name = sale.partner_id.name 
			street = sale.partner_id.street or ''
			street2 = sale.partner_id.street2 or ''
			city = sale.partner_id.city or ''
			state = sale.partner_id.state_id.name or ''
			country = sale.partner_id.country_id.name or ''
			zip_code = str(sale.partner_id.zip) or ''
			shipping_address = name +"\n"+ street + street2 +"\n"+ city +" "+ zip_code +"\n"+ state +"\n"+ country 

			worksheet.write(4,0, 'Delivery Address', style = heading_xf)
			worksheet.write_merge(5, 10, 0, 0,shipping_address)

			user_id = self.env['res.users'].browse(self._uid)
			name = user_id.company_id.name 
			street = user_id.company_id.street or ''
			street2 = user_id.company_id.street2 or ''
			city = user_id.company_id.city or ''
			state = user_id.company_id.state_id.name or ''
			country = user_id.company_id.country_id.name or ''
			zip_code = str(user_id.company_id.zip) or ''
			company_address = name +"\n"+ street + street2 +"\n"+ city +" "+ zip_code +"\n"+ state +"\n"+ country 
			worksheet.write_merge(4, 4, 5, 6, 'Company Address', style = heading_xf)
			worksheet.write_merge(5, 10, 5, 6,company_address)

			sale_name = ''
			if sale.state not in ['draft','sent']:
				sale_name = 'Sale Order #' + sale.name
			else:
				sale_name = 'Request for Quotation #' + sale.name

			worksheet.write_merge(14,15,1,5,sale_name, style = po_style)

			worksheet.write(18,0, 'Salesperson', style = section_style)
			worksheet.write(19,0, sale.user_id.name)
			worksheet.write(18,3, 'Order Date', style = section_style)
			worksheet.write(19,3, sale.date_order,date_format)
			worksheet.write(18,6, 'Order State', style = section_style)

			state = ''
			if sale.state == 'draft':
				state = 'Quotation'
			elif sale.state == 'sent':
				state = 'Quotation Sent'
			elif sale.state == 'sale':
				state = 'Sales Order'
			elif sale.state == 'done':
				state = 'Sales Order'
			elif sale.state == 'cancel':
				state = 'Cancelled'
			else:
   				state = ' '
			worksheet.write(19,6, state)

			worksheet.write(23,0, 'Product', style = tbl_header_style)
			worksheet.write(23,1, 'Description', style = tbl_header_style)
			worksheet.write(23,2, 'Quantity', style = tbl_header_style)
			worksheet.write(23,3, 'UOM', style = tbl_header_style)
			worksheet.write(23,4, 'Unit Price', style = tbl_header_style)
			worksheet.write(23,5, 'Taxes', style = tbl_header_style)
			worksheet.write(23,6, 'Subtotal', style = tbl_header_style)

			row = 24
			line_end_row = 0
			currency = str(sale.currency_id.symbol)
			for line in sale.order_line:
				tax_id = line.tax_id
				worksheet.write(row,0, line.product_id.name)
				worksheet.write(row,1, line.name)
				worksheet.write(row,2, line.product_uom_qty)
				worksheet.write(row,3, line.product_uom.name)
				worksheet.write(row,4, line.price_unit)
				taxes = []
				if not line.tax_id:
					' '
				else:
					for tax in line.tax_id:
						taxes.append(tax.name + ', ') 
				worksheet.write(row,5, taxes)
				worksheet.write(row,6, currency+ " " + str(line.price_subtotal))
				row += 1
				line_end_row = row

			line_end_row +=2
			line_end_row1 = line_end_row + 1
			line_end_row2  = line_end_row1 + 1
			worksheet.write(line_end_row,5, 'Untaxed Amount :', style = footer_style)
			worksheet.write(line_end_row1,5, 'Taxes :', style = footer_style)
			worksheet.write(line_end_row2,5, 'Total :', style = footer_style)

			amount_untaxed = str(sale.amount_untaxed)
			amount_tax = str(sale.amount_tax)
			amount_total = str(sale.amount_total)

			worksheet.write(line_end_row,6, currency +" "+ amount_untaxed, style = foo_style)
			worksheet.write(line_end_row1,6, currency +" "+ amount_tax, style = foo_style)
			worksheet.write(line_end_row2,6, currency +" "+ amount_total, style = foo_style)

		fp = io.BytesIO()
		workbook.save(fp)

		export_id = self.env['excel.report'].create({'excel_file': base64.encodebytes(fp.getvalue()), 'file_name': filename})
		res = {
			'view_mode': 'form',
			'res_id': export_id.id,
			'res_model': 'excel.report',
			'view_type': 'form',
			'type': 'ir.actions.act_window',
			'target': 'new'
		}
		return res


class sale_report_excel(models.TransientModel):
	_name = "excel.report"
	_description = "Excel File"

	excel_file = fields.Binary('Excel Report')
	file_name = fields.Char('Excel File', size=64)


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4:
