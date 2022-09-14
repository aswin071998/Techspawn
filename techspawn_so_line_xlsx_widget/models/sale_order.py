from odoo import fields, models, _, api
import xlsxwriter

from odoo.exceptions import UserError
from odoo.tools.misc import xlwt
import io
from datetime import date, datetime, timedelta
import base64
import os
import tempfile
from PIL import Image
from io import BytesIO

# from cStringIO import StringIO
from io import StringIO
import base64


class SaleOrder(models.Model):
    _inherit = 'sale.order'

    xl_report = fields.Binary(string='XL Report')
    xl_report_name = fields.Char()

    example_integer_widget = fields.Integer(string="Demo Integer(Rupee Widget)")
    example_float_widget = fields.Float(string="Demo Float(Rupee Widget)")

    def send_excel_report_by_email(self):
        self.ensure_one()
        self.generate_excel_report()
        attachment = self.env['ir.attachment'].create({
            'name': self.name,
            'type': 'binary',
            'datas': self.xl_report,
            'res_model': 'sale.order',
            'res_id': self.id
        })
        ir_model_data = self.env['ir.model.data']
        try:
            template_id = \
                ir_model_data.get_object_reference('techspawn_so_line_xlsx_widget', 'email_sale_order_xlsx')[1]
        except ValueError:
            template_id = False
        try:
            compose_form_id = ir_model_data.get_object_reference('mail', 'email_compose_message_wizard_form')[1]
        except ValueError:
            compose_form_id = False

        template = self.env['mail.template'].browse(template_id)
        template.attachment_ids = [(6, 0, [attachment.id])]
        ctx = {
            'default_model': 'sale.order',
            'default_res_id': self.ids[0],
            'default_use_template': bool(template_id),
            'default_template_id': template_id,
            'default_composition_mode': 'comment',
        }

        return {
            'name': _('Compose Email'),
            'type': 'ir.actions.act_window',
            'view_mode': 'form',
            'res_model': 'mail.compose.message',
            'views': [(compose_form_id, 'form')],
            'view_id': compose_form_id,
            'target': 'new',
            'context': ctx,
        }

    def _format_amount(self, amount, currency):
        fmt = "%.{0}f".format(currency.decimal_places)
        lang = self.env['res.lang']._lang_get(self.env.context.get('lang') or 'en_US')
        res = lang.format(fmt, currency.round(amount), grouping=True, monetary=True) \
            .replace(r' ', u'\N{NO-BREAK SPACE}').replace(r'-', u'-\N{ZERO WIDTH NO-BREAK SPACE}')

        if currency and currency.symbol:
            if currency.position == 'after':
                res = '%s %s' % (res, currency.symbol)
            elif currency and currency.position == 'before':
                res = '%s %s' % (currency.symbol, res)
        return res

    def get_open_order(self):
        # Returns dictionary of sale order data
        data_list = []
        # order_ids = self.search([('state', '=', 'sale')])
        order = self
        # for order in self:
        report_data_list = []
        for line in order.order_line:
            open_qty = line.product_uom_qty - line.qty_delivered
            expected_date = ''
            report_data_list.append(
                {'order': line.order_id, 'name': line.product_id,
                 'description': line.name, 'req_date': expected_date,
                 'order_date': False,
                 'unit_price': self._format_amount(
                     line.price_unit,
                     line.order_id.company_id.currency_id),
                 'discount': line.discount,
                 'product_uom': line.product_uom.name,
                 'order_qty': line.product_uom_qty,
                 'ship_qty': line.qty_delivered,
                 'on_hand': 0.00,  # line.product_id.qty_available,
                 'open_qty': open_qty,
                 'rate': 0.00,  # line.order_id.currency_rate,
                 'total': self._format_amount(
                     line.price_subtotal,
                     line.order_id.company_id.currency_id)})
        data_list.append({'order': order, 'lines': report_data_list})
        return data_list

    def print_excel_report(self):
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        obj = self
        customer_data = ''
        company_format = workbook.add_format(
            {'bg_color': 'gray', 'align': 'center', 'font_size': 25,
                'font_color': 'white'})
        order_format = workbook.add_format(
            {'bg_color': 'gray', 'align': 'center', 'font_size': 14,
                'font_color': 'white', 'border': 1})
        table_header_left = workbook.add_format(
            {'bg_color': 'gray', 'align': 'left', 'font_size': 12,
                'font_color': 'white'})
        table_row_left = workbook.add_format(
            {'align': 'left', 'font_size': 12, 'border': 1})
        table_header_right = workbook.add_format(
            {'bg_color': 'gray', 'align': 'right', 'font_size': 12,
                'font_color': 'white', 'border': 1})
        table_row_right = workbook.add_format(
            {'align': 'right', 'font_size': 12, 'border': 1})
        customer_header_format = workbook.add_format({
            'align': 'center', 'font_size': 13, 'bold': True, 'border': 1})
        customer_format = workbook.add_format({
            'align': 'center', 'font_size': 13, 'border': 1})
        table_left = workbook.add_format(
            {'align': 'left', 'bold': True, 'border': 1})
        table_right = workbook.add_format(
            {'align': 'right', 'bold': True, 'border': 1})
        if obj.partner_id.name:
            customer_data += obj.partner_id.name + '\n'
        if obj.partner_id.street:
            customer_data += obj.partner_id.street + '\n'
        if obj.partner_id.street2:
            customer_data += obj.partner_id.street2 + '\n'
        if obj.partner_id.city:
            customer_data += obj.partner_id.city + ' '
        if obj.partner_id.state_id:
            customer_data += str(obj.partner_id.state_id.name + ' ')
        if obj.partner_id.zip:
            customer_data += obj.partner_id.zip + ' '
        if obj.partner_id.country_id:
            customer_data += '\n' + str(obj.partner_id.country_id.name)
        worksheet = workbook.add_worksheet(obj.name)
        worksheet.merge_range('A2:F3', obj.company_id.name, company_format)
        worksheet.merge_range('A4:F4', '')
        if obj.state not in ['draft', 'sent']:
            worksheet.merge_range(
                'A5:F5', 'Order :- ' + obj.name, order_format)
            worksheet.merge_range(
                'C7:D7', 'Order Date', customer_header_format)
            worksheet.merge_range(
                'E7:F7', str(obj.date_order.date()), customer_format)
        elif obj.state in ['draft', 'sent']:
            worksheet.merge_range(
                'A5:F5', 'Quotation :- ' + obj.name, order_format)
            worksheet.merge_range(
                'C7:D7', 'Quotation Date', customer_header_format)
            worksheet.merge_range(
                'E7:F7', str(obj.date_order.date()), customer_format)
        worksheet.merge_range('A6:F6', '')
        worksheet.merge_range(
            'A7:B7', 'Customer', customer_header_format)
        worksheet.merge_range(
            'A8:B12', customer_data, customer_format)
        worksheet.merge_range(
            'C8:D8', 'Salesperson', customer_header_format)
        worksheet.merge_range(
            'E8:F8', obj.user_id.name, customer_format)
        if obj.client_order_ref:
            worksheet.merge_range(
                'C9:D9', 'Your Reference', customer_header_format)
            worksheet.merge_range(
                'E9:F9', obj.client_order_ref, customer_format)
            if obj.payment_term_id:
                worksheet.merge_range(
                    'C10:D10', 'Payment Terms', customer_header_format)
                worksheet.merge_range(
                    'E10:F10', obj.payment_term_id.name, customer_format)
        elif obj.payment_term_id:
            worksheet.merge_range(
                'C9:D9', 'Payment Terms', customer_header_format)
            worksheet.merge_range(
                'E9:F9', obj.payment_term_id.name, customer_format)
        worksheet.merge_range('A13:I13', '')

        row = 14
        worksheet.set_column('A:A', 40)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 15)

        group = self.env.user.has_group(
            'product.group_discount_per_so_line')
        display_discount = any([l.discount for l in obj.order_line])
        display_tax = any([l.tax_id for l in obj.order_line])
        worksheet.write(row, 0, 'Product', table_header_left)
        worksheet.write(row, 1, 'Quantity', table_header_right)
        worksheet.write(row, 2, 'Unit Price', table_header_right)
        if display_discount and group:
            worksheet.write(row, 3, 'Disc.%', table_header_right)
            if display_tax:
                worksheet.write(row, 4, 'Taxes', table_header_right)
                worksheet.write(row, 5, 'Amount', table_header_right)
            else:
                worksheet.write(row, 4, 'Amount', table_header_right)
        elif display_tax:
            worksheet.write(row, 3, 'Taxes', table_header_right)
            worksheet.write(row, 4, 'Amount', table_header_right)
        else:
            worksheet.write(row, 3, 'Amount', table_header_right)
        row += 1

        for line in obj.order_line:
            worksheet.write(row, 0, line.name, table_row_left)
            worksheet.write(row, 1, line.product_uom_qty, table_row_right)
            worksheet.write(row, 2, line.price_unit, table_row_right)
            if display_discount and group:
                worksheet.write(row, 3, line.discount, table_row_right)
                if display_tax and line.tax_id:
                    worksheet.write(
                        row, 4, line.tax_id.name, table_row_right)
                    worksheet.write(
                        row, 5, line.price_subtotal, table_row_right)
                    row += 1
                elif not line.tax_id and display_tax:
                    worksheet.write(row, 4, '0', table_row_right)
                    worksheet.write(
                        row, 5, line.price_subtotal, table_row_right)
                    row += 1
                else:
                    worksheet.write(
                        row, 4, line.price_subtotal, table_row_right)
                    row += 1
            elif display_tax:
                if display_tax and line.tax_id:
                    worksheet.write(
                        row, 3, line.tax_id.name, table_row_right)
                    worksheet.write(
                        row, 4, line.price_subtotal, table_row_right)
                    row += 1
                elif not line.tax_id:
                    worksheet.write(row, 3, '0', table_row_right)
                    worksheet.write(
                        row, 4, line.price_subtotal, table_row_right)
                    row += 1
                else:
                    worksheet.write(
                        row, 3, line.price_subtotal, table_row_right)
                    row += 1
            else:
                worksheet.write(
                    row, 3, line.price_subtotal, table_row_right)
                row += 1
        if display_discount and group and display_tax:
            worksheet.merge_range(row, 0, row, 5, '')
            worksheet.write(row + 1, 4, 'Untaxed Amount', table_left)
            worksheet.write(row + 1, 5, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 4, 'Taxes', table_left)
            worksheet.write(row + 2, 5, obj.amount_tax, table_right)
            worksheet.write(row + 3, 4, 'Total', table_left)
            worksheet.write(row + 3, 5, obj.amount_total, table_right)
        elif not group and not display_tax and not display_discount:
            worksheet.merge_range(row, 0, row, 3, '')
            worksheet.write(row + 1, 2, 'Subtotal', table_left)
            worksheet.write(row + 1, 3, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 2, 'Total', table_left)
            worksheet.write(row + 2, 3, obj.amount_total, table_right)
        elif not group and not display_tax:
            worksheet.merge_range(row, 0, row, 3, '')
            worksheet.write(row + 1, 2, 'Subtotal', table_left)
            worksheet.write(row + 1, 3, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 2, 'Total', table_left)
            worksheet.write(row + 2, 3, obj.amount_total, table_right)
        elif not display_tax and not display_discount:
            worksheet.merge_range(row, 0, row, 3, '')
            worksheet.write(row + 1, 2, 'Subtotal', table_left)
            worksheet.write(row + 1, 3, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 2, 'Total', table_left)
            worksheet.write(row + 2, 3, obj.amount_total, table_right)
        elif group and display_discount:
            worksheet.merge_range(row, 0, row, 4, '')
            worksheet.write(row + 1, 3, 'Subtotal', table_left)
            worksheet.write(row + 1, 4, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 3, 'Total', table_left)
            worksheet.write(row + 2, 4, obj.amount_total, table_right)
        elif display_tax:
            worksheet.merge_range(row, 0, row, 4, '')
            worksheet.write(row + 1, 3, 'Subtotal', table_left)
            worksheet.write(row + 1, 4, obj.amount_untaxed, table_right)
            worksheet.write(row + 2, 3, 'Taxes', table_left)
            worksheet.write(row + 2, 4, obj.amount_tax, table_right)
            worksheet.write(row + 3, 3, 'Total', table_left)
            worksheet.write(row + 3, 4, obj.amount_total, table_right)
        workbook.close()
        fp.seek(0)
        result = fp.read()
        return result

    def generate_excel_report(self):
        data = base64.encodestring(self.print_excel_report())
        self.xl_report = data
        self.xl_report_name = 'sale report.xls'

    def print_line_item_excel_report(self):
        # Method to print sale order line item excel report
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        title_format = workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'align': 'center'})
        row_header_format = workbook.add_format(
            {'font_name': 'Calibri', 'font_size': 11, 'bold': 1,
             'align': 'center'})
        align_right = workbook.add_format(
            {'align': 'right'})
        row_format = workbook.add_format(
            {'font_size': 11})
        row_format.set_text_wrap()

        worksheet = workbook.add_worksheet('SO - Line Item')
        worksheet.merge_range(
            0, 0, 0, 8, 'Sale Order Line Item Report',
            title_format)
        header_str = [
            'SO#', 'Customer', 'Order Date', 'Item', 'UOM', 'Ordered Qty',
            'Unit Price', 'Discount', 'Net Price']

        worksheet.set_column('A:A', 20)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:C', 15)
        worksheet.set_column('D:D', 15)
        worksheet.set_column('G:G', 15)
        worksheet.set_column('I:I', 15)
        worksheet.set_column('F:F', 15)
        row = 0
        col = 0

        data_list = self.get_open_order()
        if data_list:
            row += 2
            for index, header in enumerate(header_str, start=0):
                worksheet.write(row, index, header, row_header_format)
            for data in data_list:
                for lines in data['lines']:
                    row += 1
                    worksheet.set_row(row, 30)
                    worksheet.write(row, col, lines.get('order').name)
                    worksheet.write(row, col + 1,
                                    lines.get('order').partner_id.name)
                    worksheet.write(row, col + 2, lines.get('order_date'),
                                    align_right)
                    worksheet.write(row, col + 3, lines.get('name').name,
                                    row_format)
                    worksheet.write(row, col + 4, lines.get('product_uom'))
                    worksheet.write(row, col + 5, lines.get('order_qty'),
                                    align_right)
                    worksheet.write(row, col + 6, lines.get('unit_price'),
                                    align_right)
                    worksheet.write(row, col + 7, lines.get('discount'),
                                    align_right)
                    worksheet.write(row, col + 8, lines.get('total'),
                                    align_right)

            workbook.close()
            fp.seek(0)
            result = base64.b64encode(fp.read())

            report_id = self.env['sale.report.out'].sudo().create(
                {'filedata': result, 'filename': 'sale report.xls'})
            return {
                'type': 'ir.actions.act_url',
                'url': '/web/binary/download_document?model=sale.report.out&field=filedata&id=%s&filename=%s.xls' % (
                report_id.id, 'sale report.xls'),
                'target': 'new',
            }
        else:
            raise UserError(_('No records found'))

# class SaleReportOut(models.TransientModel):
#
#     _name = 'sale.report.out'
#
#     filedata = fields.Binary('Download file', readonly=True)
#     filename = fields.Char('Filename', size=64, readonly=True)


class SaleOrderLine(models.Model):
    _inherit = 'sale.order.line'

    product_tmpl_id = fields.Many2one('product.template', string='Product tmpl')

    # This is the function to update other fields as normal.(eg: Description)
    @api.onchange('product_tmpl_id')
    def product_tmpl_id_onchange(self):
        if not self.product_tmpl_id:
            return
        product_id = self.env['product.product'].search([('product_tmpl_id', '=', self.product_tmpl_id.id)], limit=1)
        self.product_id = product_id.id
