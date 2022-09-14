from odoo import fields, models, api


class ResConfigSettings(models.TransientModel):
    _inherit = 'res.config.settings'

    group_product_template_so_line = fields.Boolean("Product Template in SO Line",
                                                    implied_group='techspawn_so_line_xlsx_widget.group_product_template_so_line')

    def set_values(self):
        super(ResConfigSettings, self).set_values()
        self.env['ir.config_parameter'].sudo().set_param('techspawn_so_line_xlsx_widget.group_product_template_so_line', self.group_product_template_so_line)
