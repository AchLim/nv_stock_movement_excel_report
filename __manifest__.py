# -*- coding: utf-8 -*-
{
    'name': 'Stock Movement Excel Report',
    'version': '19.0.1.0.0',
    'category': 'Inventory/Reporting',
    'summary': 'Generate Excel reports for stock movements by product variant with monthly breakdown',
    'description': """
        Stock Movement Excel Report
        ===========================
        This module generates comprehensive Excel reports showing:
        - All product variants with their attributes
        - Monthly stock movements (purchases, sales, opening/closing quantities)
        - Yearly totals
        - Price information from POS, Sales, and Purchase orders
    """,
    'author': 'Vincent Lim',
    'website': 'https://github.com/AchLim',
    'license': 'LGPL-3',
    'depends': [
        'base',
        'product',
        'stock',
        'purchase',
        'sale',
        'point_of_sale',
        'mrp',
    ],
    'data': [
        'security/ir.model.access.csv',
        'views/stock_movement_report_wizard_view.xml',
    ],
    'external_dependencies': {
        'python': ['xlsxwriter'],
    },
    'installable': True,
    'application': False,
    'auto_install': False,
}
