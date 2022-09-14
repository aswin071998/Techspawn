# -*- coding: utf-8 -*-

{
    'name': 'Techspawn SO Line, SO xlsx and Custom Widget',
    'version': '14.0.0.1',
    'category': 'Sales',
    'description': """
    Custom SO Line Products, SO xlsx and Widget
    """,
    'summary': 'Custom SO Line Products, SO xlsx and Widget',
    'author': 'Aswin',
    'website': '',
    'license': '',
    'depends': ['sale_management'],
    'data': [
        'security/res_groups.xml',
        'data/mail_template_xlsx.xml',
        'views/sale_order_views.xml',
        'views/res_config_settings.xml',
        'views/rupee_widget_format_view.xml',
    ],
    'qweb': [
        'static/src/xml/rupee_widget.xml',
    ],
    'installable': True,
    'auto_install': False,
    'application': True,
}
