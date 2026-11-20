# -*- coding: utf-8 -*-
import base64
import io
from datetime import datetime, time

import pytz
import xlsxwriter

from odoo import fields, models, _
from odoo.exceptions import ValidationError


class BarcodeReportWizard(models.TransientModel):
    _name = "barcode.report.wizard"
    _description = "Barcode Report Wizard"

    # ---- Fields ----
    date_from = fields.Date(string="Date From", required=True, default=fields.Date.context_today)
    date_to = fields.Date(string="Date To", required=True, default=fields.Date.context_today)
    pos_config_ids = fields.Many2many("pos.config", string="Point of Sale")
    # Excel file yahan save hogi taake user download kar sake
    file_data = fields.Binary(string="File", readonly=True)
    # Download hone wali file ka naam
    file_name = fields.Char(string="File Name")

    def action_print(self):
        self.ensure_one()  # sirf aik record par chale, warna error de

        if self.date_from > self.date_to:
            raise ValidationError(_("Date From cannot be greater than Date To."))

        # ---------- STEP 1: Date range ko UTC me badlo ----------
        # Odoo database me date_order UTC me hota hai, user ka timezone alag ho sakta hai
        tz = pytz.timezone(self.env.user.tz or "UTC")
        start_local = tz.localize(datetime.combine(self.date_from, time.min))
        end_local = tz.localize(datetime.combine(self.date_to, time.max))
        start_utc = start_local.astimezone(pytz.utc).replace(tzinfo=None)
        end_utc = end_local.astimezone(pytz.utc).replace(tzinfo=None)

        # ---------- STEP 2: POS Orders dhundo ----------
        order_domain = [
            ("date_order", ">=", start_utc),
            ("date_order", "<=", end_utc),
            ("state", "in", ["paid", "done"]),
        ]
        if self.pos_config_ids:
            # agar POS select kiye hain to sirf unhi ka data, warna sab POS ka
            order_domain.append(("config_id", "in", self.pos_config_ids.ids))
        orders = self.env["pos.order"].search(order_domain, order="date_order asc")

        # ---------- STEP 3: Har order ki sirf LOT wali lines nikalo ----------
        data = []  # har item = (order, [lines])
        for order in orders:
            lines = order.lines.filtered(lambda l: l.product_id.tracking == "lot")
            if lines:  # agar order me lot wali koi line hai tabhi report me lo
                data.append((order, lines))

        # ---------- STEP 4: Excel file banao ----------
        output = io.BytesIO()  # memory me khali file
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        sheet = workbook.add_worksheet("Barcode Report")

        navy = "#1F3864"  # image wala dark blue colour
        title_fmt = workbook.add_format({"bold": True, "font_size": 20, "font_color": "white",
                                         "bg_color": navy, "align": "center", "valign": "vcenter"})
        label_fmt = workbook.add_format({"bold": True})
        head_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": navy,
                                        "align": "center", "valign": "vcenter", "border": 1})
        cell_fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1})
        left_fmt = workbook.add_format({"align": "left", "valign": "vcenter", "border": 1})
        money_fmt = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1,
                                         "num_format": "#,##0.00"})
        tot_lbl_fmt = workbook.add_format({"bold": True, "align": "right", "bg_color": "#F0F0F0",
                                           "border": 1, "font_color": navy})
        tot_val_fmt = workbook.add_format({"bold": True, "align": "center", "bg_color": "#F0F0F0",
                                           "border": 1, "num_format": "#,##0.00"})

        # column widths: Sr, Date, Customer, Product, Barcode, Qty, Unit Price, Total Price
        for col, width in enumerate([8, 16, 24, 30, 18, 8, 14, 14]):
            sheet.set_column(col, col, width)

        # Title (A se H tak merge)
        sheet.merge_range(0, 0, 0, 7, "BARCODE REPORT", title_fmt)
        sheet.set_row(0, 35)

        # Upar ki filter info: Date From, Date To, POS
        pos_names = ", ".join(self.pos_config_ids.mapped("name")) if self.pos_config_ids else "All POS"
        sheet.write(2, 0, "Date From:", label_fmt)
        sheet.write(2, 1, str(self.date_from))
        sheet.write(3, 0, "Date To:", label_fmt)
        sheet.write(3, 1, str(self.date_to))
        sheet.write(4, 0, "POS:", label_fmt)
        sheet.write(4, 1, pos_names)

        # Table header (row 6 par)
        row = 6
        headers = ["Sr No", "Date", "Customer", "Product", "Barcode", "Qty", "Unit Price", "Total Price"]
        for col, h in enumerate(headers):
            sheet.write(row, col, h, head_fmt)
        row += 1

        total_qty = 0.0
        total_price = 0.0

        # Data rows
        for sr, (order, lines) in enumerate(data, start=1):
            first = row                      # is order ki pehli row
            last = row + len(lines) - 1      # is order ki aakhri row
            # order ki date user ke timezone me
            order_date = fields.Datetime.context_timestamp(self, order.date_order).strftime("%d-%b-%Y")
            customer = order.partner_id.name or ""

            # Sr No, Date, Customer: agar 1 se zyada lines hain to cells merge karo (image jaisa)
            for col, val in enumerate([sr, order_date, customer]):
                if last > first:
                    sheet.merge_range(first, col, last, col, val, cell_fmt)
                else:
                    sheet.write(first, col, val, cell_fmt)

            # Har line ki apni row
            for line in lines:
                line_total = line.price_subtotal_incl  # tax samet total
                lot_names = ", ".join(line.pack_lot_ids.mapped("lot_name"))  # pehle POS line ke lots lo
                if not lot_names:  # agar POS line par lot nahi mila to
                    move_lines = line.order_id.picking_ids.move_line_ids.filtered(  # order ki delivery ki stock lines
                        lambda ml: ml.product_id == line.product_id and ml.lot_id
                        # isi product ki wo lines jin par lot ho
                    )
                    lot_names = ", ".join(move_lines.mapped("lot_id.name"))
                sheet.write(row, 3, line.product_id.display_name, left_fmt)
                sheet.write(row, 4, lot_names, cell_fmt)
                sheet.write(row, 5, line.qty, cell_fmt)
                sheet.write(row, 6, line.price_unit, money_fmt)
                sheet.write(row, 7, line_total, money_fmt)
                total_qty += line.qty
                total_price += line_total
                row += 1

        # Totals (neeche, Unit Price aur Total Price columns me)
        row += 1
        sheet.write(row, 6, "Total Qty", tot_lbl_fmt)
        sheet.write(row, 7, total_qty, tot_val_fmt)
        sheet.write(row + 1, 6, "Total Price", tot_lbl_fmt)
        sheet.write(row + 1, 7, total_price, tot_val_fmt)

        workbook.close()  # file complete

        # ---------- STEP 5: File wizard me save karke download karwao ----------
        self.write({
            "file_data": base64.b64encode(output.getvalue()),
            "file_name": "Barcode_Report_%s_to_%s.xlsx" % (self.date_from, self.date_to),
        })
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/?model=%s&id=%s&field=file_data&filename_field=file_name&download=true"
                   % (self._name, self.id),
            "target": "self",
        }