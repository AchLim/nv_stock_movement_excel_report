# -*- coding: utf-8 -*-
import base64
import io
from datetime import timedelta
from dateutil.relativedelta import relativedelta

from odoo import models, fields, api, _
from odoo.exceptions import UserError

try:
    import xlsxwriter
except ImportError:
    xlsxwriter = None


class StockMovementReportWizard(models.TransientModel):
    _name = 'stock.movement.report.wizard'
    _description = 'Stock Movement Excel Report Wizard'

    date_from = fields.Date(
        string='Start Date',
        required=True,
        default=lambda self: fields.Date.today().replace(month=1, day=1)
    )
    date_to = fields.Date(
        string='End Date',
        required=True,
        default=fields.Date.today
    )
    product_ids = fields.Many2many(
        'product.product',
        string='Products',
        help='Leave empty to include all products'
    )
    category_ids = fields.Many2many(
        'product.category',
        string='Product Categories',
        help='Filter by product categories'
    )
    warehouse_ids = fields.Many2many(
        'stock.warehouse',
        string='Warehouses',
        help='Leave empty to include all warehouses'
    )
    include_pos = fields.Boolean(
        string='Include POS Sales',
        default=True
    )
    include_sales = fields.Boolean(
        string='Include Sales Orders',
        default=True
    )
    include_purchases = fields.Boolean(
        string='Include Purchases',
        default=True
    )
    excel_file = fields.Binary(string='Excel Report')
    file_name = fields.Char(string='File Name')

    @api.constrains('date_from', 'date_to')
    def _check_dates(self):
        for wizard in self:
            if wizard.date_from > wizard.date_to:
                raise UserError(_('Start Date must be before End Date.'))

    def _get_months_in_range(self):
        """Generate list of months between date_from and date_to"""
        months = []
        current = self.date_from.replace(day=1)
        end = self.date_to.replace(day=1)

        while current <= end:
            months.append({
                'date': current,
                'name': current.strftime('%B %Y'),
                'year': current.year,
                'month': current.month,
                'start': current,
                'end': (current + relativedelta(months=1)) - timedelta(days=1)
            })
            current = current + relativedelta(months=1)

        return months

    def _get_products(self):
        """Get product variants to include in report, excluding phantom BoM (kit) products"""
        domain = [('type', '=', 'consu')]

        if self.product_ids:
            domain.append(('id', 'in', self.product_ids.ids))

        if self.category_ids:
            domain.append(('categ_id', 'child_of', self.category_ids.ids))

        products = self.env['product.product'].search(domain, order='name')

        # Exclude products that have phantom BoM (kits) to avoid double counting
        # Kit products' stock movements are already reflected in their components
        if products:
            phantom_bom_product_tmpl_ids = self.env['mrp.bom'].search([
                ('type', '=', 'phantom'),
                ('product_tmpl_id', 'in', products.product_tmpl_id.ids)
            ]).mapped('product_tmpl_id').ids

            if phantom_bom_product_tmpl_ids:
                products = products.filtered(
                    lambda p: p.product_tmpl_id.id not in phantom_bom_product_tmpl_ids
                )

        return products

    def _get_product_display_name(self, product):
        """Get product name with all variant attributes in one cell"""
        name = product.product_tmpl_id.name

        if product.product_template_attribute_value_ids:
            attributes = []
            for attr_value in product.product_template_attribute_value_ids:
                attr_name = attr_value.attribute_id.name
                value_name = attr_value.name
                attributes.append(f"{attr_name}: {value_name}")

            if attributes:
                name = f"{name} ({', '.join(attributes)})"

        return name

    def _get_location_ids(self):
        """Get internal location IDs for stock calculation"""
        if self.warehouse_ids:
            locations = self.env['stock.location'].search([
                ('usage', '=', 'internal'),
                ('warehouse_id', 'in', self.warehouse_ids.ids)
            ])
        else:
            locations = self.env['stock.location'].search([
                ('usage', '=', 'internal')
            ])
        return locations.ids

    def _get_stock_at_date(self, product_id, date, location_ids):
        """Calculate stock quantity at a specific date using product_qty (already in product's base UoM)"""
        self.env.cr.execute("""
                            SELECT COALESCE(SUM(
                                                    CASE
                                                        WHEN sm.location_dest_id IN %s THEN sm.product_qty
                                                        WHEN sm.location_id IN %s THEN -sm.product_qty
                                                        ELSE 0
                                                        END
                                            ), 0) as qty
                            FROM stock_move sm
                            WHERE sm.product_id = %s
                              AND sm.state = 'done'
                              AND sm.date::date <= %s
                              AND (sm.location_id IN %s
                               OR sm.location_dest_id IN %s)
                            """, (
                                tuple(location_ids), tuple(location_ids),
                                product_id, date,
                                tuple(location_ids), tuple(location_ids)
                            ))
        result = self.env.cr.fetchone()
        return result[0] if result else 0

    def _get_stock_moves_data(self, product_id, date_start, date_end, location_ids):
        """Get stock move data for a product in a date range"""
        data = {
            'qty_in': 0,
            'qty_out': 0,
            'value_in': 0,
            'value_out': 0,
            'purchase_qty': 0,
            'purchase_value': 0,
            'sale_qty': 0,
            'sale_value': 0,
            'pos_qty': 0,
            'pos_value': 0,
        }

        # Find phantom BoMs that contain this product as a component
        phantom_bom_data = self._get_phantom_bom_components(product_id)

        # Get incoming moves using product_qty (already in product's base UoM)
        self.env.cr.execute("""
                            SELECT COALESCE(SUM(sm.product_qty), 0) as qty,
                                   COALESCE(SUM(sm.product_qty * COALESCE(
                                           (SELECT pol.price_unit
                                            FROM purchase_order_line pol
                                                     JOIN stock_move sm2 ON sm2.purchase_line_id = pol.id
                                            WHERE sm2.id = sm.id LIMIT 1),
                                           sm.price_unit
                                   )), 0) as value
                            FROM stock_move sm
                            WHERE sm.product_id = %s
                              AND sm.state = 'done'
                              AND sm.date:: date >= %s
                              AND sm.date:: date <= %s
                              AND sm.location_dest_id IN %s
                              AND sm.location_id NOT IN %s
                            """, (product_id, date_start, date_end, tuple(location_ids), tuple(location_ids)))

        result = self.env.cr.fetchone()
        if result:
            data['qty_in'] = result[0] or 0
            data['value_in'] = result[1] or 0

        # Get outgoing moves using product_qty (already in product's base UoM)
        self.env.cr.execute("""
                            SELECT COALESCE(SUM(sm.product_qty), 0) as qty,
                                   COALESCE(SUM(sm.product_qty * sm.price_unit), 0) as value
                            FROM stock_move sm
                            WHERE sm.product_id = %s
                              AND sm.state = 'done'
                              AND sm.date:: date >= %s
                              AND sm.date:: date <= %s
                              AND sm.location_id IN %s
                              AND sm.location_dest_id NOT IN %s
                            """, (product_id, date_start, date_end, tuple(location_ids), tuple(location_ids)))

        result = self.env.cr.fetchone()
        if result:
            data['qty_out'] = result[0] or 0
            data['value_out'] = result[1] or 0

        # Get purchase-specific data with UoM conversion
        if self.include_purchases:
            self.env.cr.execute("""
                                SELECT COALESCE(SUM(pol.qty_received / pol_uom.factor * prod_uom.factor), 0) as qty,
                                       COALESCE(SUM(pol.qty_received * pol.price_unit), 0) as value
                                FROM purchase_order_line pol
                                    JOIN purchase_order po
                                ON pol.order_id = po.id
                                    JOIN product_product pp ON pol.product_id = pp.id
                                    JOIN product_template pt ON pp.product_tmpl_id = pt.id
                                    JOIN uom_uom prod_uom ON pt.uom_id = prod_uom.id
                                    JOIN uom_uom pol_uom ON pol.product_uom_id = pol_uom.id
                                WHERE pol.product_id = %s
                                  AND po.state IN ('purchase'
                                    , 'done')
                                  AND po.date_approve:: date >= %s
                                  AND po.date_approve:: date <= %s
                                """, (product_id, date_start, date_end))

            result = self.env.cr.fetchone()
            if result:
                data['purchase_qty'] = result[0] or 0
                data['purchase_value'] = result[1] or 0

        # Get sales order data with UoM conversion
        if self.include_sales:
            # Direct sales of the product
            self.env.cr.execute("""
                                SELECT COALESCE(SUM(sol.qty_delivered / sol_uom.factor * prod_uom.factor), 0) as qty,
                                       COALESCE(SUM(sol.qty_delivered * sol.price_unit), 0) as value
                                FROM sale_order_line sol
                                    JOIN sale_order so
                                ON sol.order_id = so.id
                                    JOIN product_product pp ON sol.product_id = pp.id
                                    JOIN product_template pt ON pp.product_tmpl_id = pt.id
                                    JOIN uom_uom prod_uom ON pt.uom_id = prod_uom.id
                                    JOIN uom_uom sol_uom ON sol.product_uom_id = sol_uom.id
                                WHERE sol.product_id = %s
                                  AND so.state IN ('sale'
                                    , 'done')
                                  AND so.date_order:: date >= %s
                                  AND so.date_order:: date <= %s
                                """, (product_id, date_start, date_end))

            result = self.env.cr.fetchone()
            if result:
                data['sale_qty'] = result[0] or 0
                data['sale_value'] = result[1] or 0

            # Add sales from phantom BoM (kit) products
            for kit_product_id, bom_qty in phantom_bom_data.items():
                self.env.cr.execute("""
                                    SELECT COALESCE(SUM(sol.qty_delivered / sol_uom.factor * prod_uom.factor),
                                                    0) as qty,
                                           COALESCE(SUM(sol.qty_delivered * sol.price_unit), 0) as value
                                    FROM sale_order_line sol
                                        JOIN sale_order so
                                    ON sol.order_id = so.id
                                        JOIN product_product pp ON sol.product_id = pp.id
                                        JOIN product_template pt ON pp.product_tmpl_id = pt.id
                                        JOIN uom_uom prod_uom ON pt.uom_id = prod_uom.id
                                        JOIN uom_uom sol_uom ON sol.product_uom_id = sol_uom.id
                                    WHERE sol.product_id = %s
                                      AND so.state IN ('sale'
                                        , 'done')
                                      AND so.date_order:: date >= %s
                                      AND so.date_order:: date <= %s
                                    """, (kit_product_id, date_start, date_end))

                result = self.env.cr.fetchone()
                if result and result[0]:
                    data['sale_qty'] += (result[0] or 0) * bom_qty
                    data['sale_value'] += (result[1] or 0) * bom_qty

        # Get POS sales data (POS uses product's default UoM)
        if self.include_pos:
            # Direct POS sales of the product
            self.env.cr.execute("""
                                SELECT COALESCE(SUM(pol.qty), 0) as qty,
                                       COALESCE(SUM(pol.price_subtotal_incl), 0) as value
                                FROM pos_order_line pol
                                    JOIN pos_order po
                                ON pol.order_id = po.id
                                WHERE pol.product_id = %s
                                  AND po.state IN ('paid'
                                    , 'done'
                                    , 'invoiced')
                                  AND po.date_order:: date >= %s
                                  AND po.date_order:: date <= %s
                                """, (product_id, date_start, date_end))

            result = self.env.cr.fetchone()
            if result:
                data['pos_qty'] = result[0] or 0
                data['pos_value'] = result[1] or 0

            # Add POS sales from phantom BoM (kit) products
            for kit_product_id, bom_qty in phantom_bom_data.items():
                self.env.cr.execute("""
                                    SELECT COALESCE(SUM(pol.qty), 0) as qty,
                                           COALESCE(SUM(pol.price_subtotal_incl), 0) as value
                                    FROM pos_order_line pol
                                        JOIN pos_order po
                                    ON pol.order_id = po.id
                                    WHERE pol.product_id = %s
                                      AND po.state IN ('paid'
                                        , 'done'
                                        , 'invoiced')
                                      AND po.date_order:: date >= %s
                                      AND po.date_order:: date <= %s
                                    """, (kit_product_id, date_start, date_end))

                result = self.env.cr.fetchone()
                if result and result[0]:
                    data['pos_qty'] += (result[0] or 0) * bom_qty
                    data['pos_value'] += (result[1] or 0) * bom_qty

        return data

    def _get_phantom_bom_components(self, product_id):
        """
        Find all phantom BoM products that contain this product as a component.
        Returns a dict: {kit_product_id: quantity_of_component_in_kit}
        """
        result = {}

        # Find BoM lines where this product is a component
        bom_lines = self.env['mrp.bom.line'].search([
            ('product_id', '=', product_id)
        ])

        for line in bom_lines:
            bom = line.bom_id
            # Only consider phantom (kit) BoMs
            if bom.type == 'phantom':
                # Get the kit product(s)
                if bom.product_id:
                    # BoM is for a specific variant
                    kit_product_id = bom.product_id.id
                else:
                    # BoM is for template, get all variants
                    kit_products = bom.product_tmpl_id.product_variant_ids
                    for kit_product in kit_products:
                        kit_product_id = kit_product.id
                        # Calculate the quantity considering UoM
                        component_qty = line.product_qty
                        # Convert to product UoM if different
                        if line.product_uom_id != line.product_id.uom_id:
                            component_qty = line.product_uom_id._compute_quantity(
                                line.product_qty,
                                line.product_id.uom_id
                            )
                        result[kit_product_id] = component_qty
                    continue

                # Calculate the quantity considering UoM
                component_qty = line.product_qty
                # Convert to product UoM if different
                if line.product_uom_id != line.product_id.uom_id:
                    component_qty = line.product_uom_id._compute_quantity(
                        line.product_qty,
                        line.product_id.uom_id
                    )
                result[kit_product_id] = component_qty

        return result

    def action_generate_report(self):
        """Generate the Excel report"""
        self.ensure_one()

        if not xlsxwriter:
            raise UserError(_('xlsxwriter library is required. Please install it using: pip install xlsxwriter'))

        products = self._get_products()
        if not products:
            raise UserError(_('No products found with the given criteria.'))

        months = self._get_months_in_range()
        if not months:
            raise UserError(_('No months found in the selected date range.'))

        location_ids = self._get_location_ids()
        if not location_ids:
            raise UserError(_('No stock locations found.'))

        # Create Excel file
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Define formats
        formats = self._create_excel_formats(workbook)

        # Create worksheet
        worksheet = workbook.add_worksheet('Stock Movement Report')

        # Write headers and data
        self._write_excel_content(
            worksheet, formats, products, months, location_ids
        )

        workbook.close()
        output.seek(0)

        # Save file
        file_data = base64.b64encode(output.read())
        file_name = f'Stock_Movement_Report_{self.date_from}_{self.date_to}.xlsx'

        self.write({
            'excel_file': file_data,
            'file_name': file_name,
        })

        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/?model={self._name}&id={self.id}&field=excel_file&filename_field=file_name&download=true',
            'target': 'self',
        }

    def _create_excel_formats(self, workbook):
        """Create Excel cell formats"""
        formats = {}

        # Title format
        formats['title'] = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#2E7D32',
            'border': 1,
        })

        # Header level 1 (Month names)
        formats['header_month'] = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#1565C0',
            'border': 1,
            'text_wrap': True,
        })

        # Header level 2 (Column names)
        formats['header_col'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#42A5F5',
            'border': 1,
            'text_wrap': True,
        })

        # Product name format
        formats['product'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'align': 'left',
            'valign': 'vcenter',
            'bg_color': '#E3F2FD',
            'border': 1,
            'text_wrap': True,
        })

        # Number format
        formats['number'] = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00',
            'border': 1,
        })

        # Currency format
        formats['currency'] = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'Rp #,##0.00',
            'border': 1,
        })

        # Integer format
        formats['integer'] = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0',
            'border': 1,
        })

        # Yearly total header
        formats['header_year'] = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#FF6F00',
            'border': 1,
            'text_wrap': True,
        })

        # Yearly column header
        formats['header_year_col'] = workbook.add_format({
            'bold': True,
            'font_size': 10,
            'align': 'center',
            'valign': 'vcenter',
            'font_color': '#FFFFFF',
            'bg_color': '#FFB300',
            'border': 1,
            'text_wrap': True,
        })

        # Yearly data format
        formats['year_number'] = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': '#,##0.00',
            'border': 1,
            'bg_color': '#FFF3E0',
        })

        formats['year_currency'] = workbook.add_format({
            'font_size': 10,
            'align': 'right',
            'valign': 'vcenter',
            'num_format': 'Rp #,##0.00',
            'border': 1,
            'bg_color': '#FFF3E0',
        })

        return formats

    def _write_excel_content(self, worksheet, formats, products, months, location_ids):
        """Write content to Excel worksheet"""

        # Define columns per month
        month_columns = [
            ('opening_qty', 'Opening\nStock', 'integer'),
            ('purchase_qty', 'Purchased\nQty', 'integer'),
            ('purchase_value', 'Purchase\nValue', 'currency'),
            ('sale_qty', 'Sold Qty\n(Sales)', 'integer'),
            ('sale_value', 'Sales\nValue', 'currency'),
            ('pos_qty', 'Sold Qty\n(POS)', 'integer'),
            ('pos_value', 'POS\nValue', 'currency'),
            ('closing_qty', 'Closing\nStock', 'integer'),
        ]

        cols_per_month = len(month_columns)

        # Yearly summary columns
        year_columns = [
            ('total_purchase_qty', 'Total\nPurchased', 'integer'),
            ('total_purchase_value', 'Total\nPurchase Value', 'currency'),
            ('total_sale_qty', 'Total\nSold (Sales)', 'integer'),
            ('total_sale_value', 'Total\nSales Value', 'currency'),
            ('total_pos_qty', 'Total\nSold (POS)', 'integer'),
            ('total_pos_value', 'Total\nPOS Value', 'currency'),
        ]

        # Group months by year
        years_in_range = sorted(set(m['year'] for m in months))

        # Calculate total columns needed
        total_month_cols = len(months) * cols_per_month
        total_year_cols = len(years_in_range) * len(year_columns)
        total_cols = 1 + total_month_cols + total_year_cols  # 1 for product column

        # Row 0: Title
        worksheet.merge_range(
            0, 0, 0, total_cols - 1,
            f'Stock Movement Report ({self.date_from} to {self.date_to})',
            formats['title']
        )

        # Row 1: Month headers (merged across their columns)
        row = 1
        col = 1  # Start after product column

        worksheet.write(row, 0, 'Product', formats['header_month'])
        worksheet.write(row + 1, 0, 'Variant', formats['header_col'])

        for month in months:
            end_col = col + cols_per_month - 1
            worksheet.merge_range(row, col, row, end_col, month['name'], formats['header_month'])
            col = end_col + 1

        # Year summary headers
        for year in years_in_range:
            end_col = col + len(year_columns) - 1
            worksheet.merge_range(row, col, row, end_col, f'Year {year} Total', formats['header_year'])
            col = end_col + 1

        # Row 2: Column sub-headers
        row = 2
        col = 1

        for month in months:
            for col_def in month_columns:
                worksheet.write(row, col, col_def[1], formats['header_col'])
                col += 1

        for year in years_in_range:
            for col_def in year_columns:
                worksheet.write(row, col, col_def[1], formats['header_year_col'])
                col += 1

        # Set column widths
        worksheet.set_column(0, 0, 45)  # Product column
        worksheet.set_column(1, total_cols - 1, 12)  # Data columns

        # Set row heights
        worksheet.set_row(0, 30)  # Title row
        worksheet.set_row(1, 25)  # Month header row
        worksheet.set_row(2, 40)  # Column header row

        # Write product data
        data_row = 3

        for product in products:
            product_name = self._get_product_display_name(product)
            worksheet.write(data_row, 0, product_name, formats['product'])

            col = 1
            yearly_totals = {year: {
                'total_purchase_qty': 0,
                'total_purchase_value': 0,
                'total_sale_qty': 0,
                'total_sale_value': 0,
                'total_pos_qty': 0,
                'total_pos_value': 0,
            } for year in years_in_range}

            for month in months:
                # Get opening stock (stock at start of month)
                opening_qty = self._get_stock_at_date(
                    product.id,
                    month['start'] - timedelta(days=1),
                    location_ids
                )

                # Get movement data for the month
                move_data = self._get_stock_moves_data(
                    product.id,
                    month['start'],
                    month['end'],
                    location_ids
                )

                # Get closing stock (stock at end of month)
                closing_qty = self._get_stock_at_date(
                    product.id,
                    month['end'],
                    location_ids
                )

                # Write monthly data
                month_data = {
                    'opening_qty': opening_qty,
                    'purchase_qty': move_data['purchase_qty'],
                    'purchase_value': move_data['purchase_value'],
                    'sale_qty': move_data['sale_qty'],
                    'sale_value': move_data['sale_value'],
                    'pos_qty': move_data['pos_qty'],
                    'pos_value': move_data['pos_value'],
                    'closing_qty': closing_qty,
                }

                for col_def in month_columns:
                    key, label, fmt_type = col_def
                    value = month_data.get(key, 0)

                    if fmt_type == 'currency':
                        worksheet.write(data_row, col, value, formats['currency'])
                    elif fmt_type == 'integer':
                        worksheet.write(data_row, col, value, formats['integer'])
                    else:
                        worksheet.write(data_row, col, value, formats['number'])
                    col += 1

                # Accumulate yearly totals
                year = month['year']
                yearly_totals[year]['total_purchase_qty'] += move_data['purchase_qty']
                yearly_totals[year]['total_purchase_value'] += move_data['purchase_value']
                yearly_totals[year]['total_sale_qty'] += move_data['sale_qty']
                yearly_totals[year]['total_sale_value'] += move_data['sale_value']
                yearly_totals[year]['total_pos_qty'] += move_data['pos_qty']
                yearly_totals[year]['total_pos_value'] += move_data['pos_value']

            # Write yearly totals
            for year in years_in_range:
                for col_def in year_columns:
                    key, label, fmt_type = col_def
                    value = yearly_totals[year].get(key, 0)

                    if fmt_type == 'currency':
                        worksheet.write(data_row, col, value, formats['year_currency'])
                    else:
                        worksheet.write(data_row, col, value, formats['year_number'])
                    col += 1

            data_row += 1

        # Freeze panes (freeze product column and header rows)
        worksheet.freeze_panes(3, 1)

        # Add autofilter
        worksheet.autofilter(2, 0, data_row - 1, total_cols - 1)
