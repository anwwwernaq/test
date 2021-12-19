# -*- coding: utf-8 -*-
# Part of BrowseInfo. See LICENSE file for full copyright and licensing details.

from odoo import api, fields, models, tools, _
from odoo.exceptions import UserError
import datetime


class ReportSaleWizard(models.AbstractModel):
	_name = 'report.bi_all_in_one_sale_reports.user_wise_sales_detail_doc'
	_description = 'User Wise Sale Detail Report'

	@api.model
	def _get_report_values(self, docids, data=None):
		report = self.env['ir.actions.report']._get_report_from_name(
			'bi_all_in_one_sale_reports.user_wise_sales_detail_doc')
		record = {
			'doc_ids': self.env['user.wise.sales.detail.report'].search([('id', 'in', list(data["ids"]))]),
			'doc_model': report.model,
			'docs': self,
			'data': data,
			}

		return record


# vim:expandtab:smartindent:tabstop=4:softtabstop=4:shiftwidth=4: