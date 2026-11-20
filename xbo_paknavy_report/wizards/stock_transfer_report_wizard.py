import base64
import io
from collections import OrderedDict
from datetime import datetime, time

import pytz
import xlsxwriter

from odoo import _, fields, models
from odoo.exceptions import UserError


class StockTransferReportWizard(models.TransientModel):
    _name = 'stock.transfer.report.wizard'
    _description = 'Stock Transfer Report Wizard'

    date_from = fields.Date(string='Date From', required=True)
    date_to = fields.Date(string='Date To', required=True)
    location_id = fields.Many2one('stock.location', string='From Location')
    partner_id = fields.Many2one(
        'res.partner', string='To Location',
        domain=[('is_branch', '=', True)],
    )

    # ---------------- Pickings dhundo (koi change nahi) ----------------
    def _get_pickings(self):
        self.ensure_one()
        tz = pytz.timezone(self.env.user.tz or 'UTC')

        start = tz.localize(datetime.combine(self.date_from, time.min)).astimezone(pytz.utc).replace(tzinfo=None)
        end = tz.localize(datetime.combine(self.date_to, time.max)).astimezone(pytz.utc).replace(tzinfo=None)

        domain = [
            ('state', '=', 'done'),
            ('picking_type_code', '=', 'outgoing'),
            ('date_done', '>=', start),
            ('date_done', '<=', end),
            ('sale_id', '!=', False),
        ]
        if self.partner_id:
            domain.append(('partner_id', '=', self.partner_id.id))
        else:
            domain.append(('partner_id.is_branch', '=', True))
        if self.location_id:
            domain.append(('location_id', '=', self.location_id.id))
        return self.env['stock.picking'].search(domain, order='date_done, id')

    # ---------------- Data tayyar karo: ab PICKING WISE groups ----------------
    def _prepare_lines(self, pickings):
        groups = []                                  # har picking ka aik group (dictionary)
        sale_orders = self.env['sale.order']         # saare sale orders (footer ke liye)
        for picking in pickings:
            sale_orders |= picking.sale_id           # is picking ka sale order jama karo
            rows = []                                # is picking ki product/lot wali rows
            for move in picking.move_ids.filtered(lambda m: m.state == 'done'):
                qty_by_lot = OrderedDict()           # lot wise qty jama karne ke liye
                for ml in move.move_line_ids:
                    entry = qty_by_lot.setdefault(ml.lot_id.id, [ml.lot_id, 0.0])
                    entry[1] += ml.quantity
                unit_price = move.sale_line_id.price_unit if move.sale_line_id else 0.0
                for lot, qty in qty_by_lot.values():
                    rows.append({
                        'product': move.product_id.display_name,
                        'lot': lot.name or '',
                        'qty': qty,
                        'unit_price': unit_price,
                        'total': unit_price * qty,
                    })
            if not rows:
                continue                             # agar picking me koi done line nahi to chor do

            # is picking ke sale order ki posted invoices
            invoices = picking.sale_id.invoice_ids.filtered(
                lambda inv: inv.move_type == 'out_invoice' and inv.state == 'posted')
            invoiced = sum(invoices.mapped('amount_total'))     # invoice ka total amount
            due = sum(invoices.mapped('amount_residual'))       # abhi baqi amount
            groups.append({
                'picking': picking.name,
                'location': picking.location_id.complete_name or '',
                'customer': picking.partner_id.name or '',
                'invoiced': invoiced,
                'paid': invoiced - due,              # ada ho chuka = total - baqi
                'due': due,
                'rows': rows,
            })
        return groups, sale_orders

    # ---------------- Excel banao ----------------
    def _build_excel(self, groups, sale_orders):
        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        sheet = workbook.add_worksheet('Stock Transfer Report')

        title = workbook.add_format({'bold': True, 'font_size': 16, 'align': 'center'})
        bold = workbook.add_format({'bold': True})
        head = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D9D9D9',
                                    'align': 'center', 'valign': 'vcenter'})
        # valign vcenter: merge hui cell me text beech me nazar aaye
        cell = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        num = workbook.add_format({'border': 1, 'valign': 'vcenter', 'num_format': '#,##0.00'})
        bold_num = workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00'})
        bold_cell = workbook.add_format({'bold': True, 'border': 1})

        # Title aur filters (ab 12 columns hain: A se L tak)
        sheet.merge_range(0, 0, 0, 11, 'Stock Transfer Report', title)
        sheet.write(2, 0, 'Date From', bold)
        sheet.write(2, 1, self.date_from.strftime('%d-%m-%Y'))
        sheet.write(2, 10, 'From Location', bold)
        sheet.write(2, 11, self.location_id.complete_name or '')
        sheet.write(3, 0, 'Date To', bold)
        sheet.write(3, 1, self.date_to.strftime('%d-%m-%Y'))
        sheet.write(3, 10, 'To Customer', bold)
        sheet.write(3, 11, self.partner_id.name or '')

        # Header: To Customer ke baad 3 naye columns (Total Amount, Amount Paid, Amount Due)
        headers = ['Sr No', 'Name', 'From Location', 'To Customer',
                   'Total Amount', 'Amount Paid', 'Amount Due',
                   'Product', 'Lot Number', 'Qty', 'Unit Price', 'Total Price']
        for col, title_text in enumerate(headers):
            sheet.write(5, col, title_text, head)

        row = 6
        total_qty = 0.0
        total_amount = 0.0
        # enumerate start=1: har PICKING ko aik Sr No (1, 2, 3...)
        for sr, grp in enumerate(groups, start=1):
            first = row                              # is picking ki pehli row
            last = row + len(grp['rows']) - 1        # is picking ki aakhri row

            # Picking level ki values: ye saari lines me aik hi hain, isliye merge karenge
            common = [
                (sr, cell),
                (grp['picking'], cell),
                (grp['location'], cell),
                (grp['customer'], cell),
                (grp['invoiced'], num),
                (grp['paid'], num),
                (grp['due'], num),
            ]
            for col, (val, fmt) in enumerate(common):
                if last > first:
                    sheet.merge_range(first, col, last, col, val, fmt)   # 1 se zyada lines: merge
                else:
                    sheet.write(first, col, val, fmt)                    # sirf 1 line: seedha likho

            # Har product/lot ki apni row (ye repeat nahi hota)
            for r in grp['rows']:
                sheet.write(row, 7, r['product'], cell)
                sheet.write(row, 8, r['lot'], cell)
                sheet.write(row, 9, r['qty'], num)
                sheet.write(row, 10, r['unit_price'], num)
                sheet.write(row, 11, r['total'], num)
                total_qty += r['qty']
                total_amount += r['total']
                row += 1

        # Neeche ka total (pehle jaisa): saare sale orders ki invoices ek baar gino
        invoices = sale_orders.invoice_ids.filtered(
            lambda inv: inv.move_type == 'out_invoice' and inv.state == 'posted')
        invoiced = sum(invoices.mapped('amount_total'))
        due = sum(invoices.mapped('amount_residual'))
        paid = invoiced - due

        row += 1
        footer = [('Total Qty', total_qty), ('Total Amount', total_amount),
                  ('Amount Paid', paid), ('Amount Due', due)]
        for label, value in footer:
            sheet.write(row, 10, label, bold_cell)
            sheet.write(row, 11, value, bold_num)
            row += 1

        # Column widths
        sheet.set_column(0, 0, 7)     # Sr No
        sheet.set_column(1, 2, 18)    # Name, From Location
        sheet.set_column(3, 3, 22)    # To Customer
        sheet.set_column(4, 6, 15)    # 3 amount columns
        sheet.set_column(7, 7, 28)    # Product
        sheet.set_column(8, 8, 14)    # Lot Number
        sheet.set_column(9, 9, 8)     # Qty
        sheet.set_column(10, 11, 15)  # Unit Price, Total Price
        workbook.close()
        return output.getvalue()

    def action_print_excel(self):
        self.ensure_one()
        if self.date_from > self.date_to:
            raise UserError(_('Date From must be before Date To.'))
        pickings = self._get_pickings()
        if not pickings:
            raise UserError(_('No transfers found for the selected filters.'))
        groups, sale_orders = self._prepare_lines(pickings)
        data = self._build_excel(groups, sale_orders)
        attachment = self.env['ir.attachment'].create({
            'name': 'Stock_Transfer_%s_%s.xlsx' % (self.date_from, self.date_to),
            'type': 'binary',
            'datas': base64.b64encode(data),
            'mimetype': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        return {
            'type': 'ir.actions.act_url',
            'url': f'/web/content/{attachment.id}?download=true',
            'target': 'self',
        }