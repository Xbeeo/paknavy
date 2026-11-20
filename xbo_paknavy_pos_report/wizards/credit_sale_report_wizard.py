# -*- coding: utf-8 -*-
import base64          # file ko base64 me badalne ke liye (attachment ko yehi chahiye)
import io              # memory me Excel file banane ke liye
import xlsxwriter      # Excel banane ki library (Odoo ke sath aati hai)

from odoo import fields, models, _
from odoo.exceptions import ValidationError


class CreditSaleReportWizard(models.TransientModel):
    _name = "credit.sale.report.wizard"
    _description = "Credit Sale Report Wizard"

    date_from = fields.Date(string="Date From", required=True, default=fields.Date.context_today)
    date_to = fields.Date(string="Date To", required=True, default=fields.Date.context_today)
    pos_config_ids = fields.Many2many("pos.config", string="Point of Sale")

    # ---------------- PDF (pehle jaisa, koi change nahi) ----------------
    def action_print(self):
        self.ensure_one()

        if self.date_from > self.date_to:
            raise ValidationError(_("Date From cannot be greater than Date To."))

        return self.env.ref("xbo_paknavy_pos_report.action_report_credit_sale").report_action(self)

    # ---------------- Lot nikalne ka helper ----------------
    def _get_lot_pool(self, move, product):
        """Invoice ke is product ke SAARE lots ki list wapis karta hai (har lot aik baar)."""
        # invoice se juray POS orders ki wo lines jin ka product yehi ho
        pos_lines = move.pos_order_ids.lines.filtered(lambda l: l.product_id == product)
        # in POS lines par jo lot likhe hain unke naam (list)
        names = pos_lines.pack_lot_ids.mapped("lot_name")
        if not names:
            # agar POS line par lot nahi mila to delivery ki stock move lines se lo
            move_lines = move.pos_order_ids.picking_ids.move_line_ids.filtered(
                lambda ml: ml.product_id == product and ml.lot_id
            )
            names = move_lines.mapped("lot_id.name")
        return list(names)  # list wapis, taake aage lines me baant sakein

    # ---------------- Excel ----------------
    def action_print_excel(self):
        self.ensure_one()

        if self.date_from > self.date_to:
            raise ValidationError(_("Date From cannot be greater than Date To."))

        # PDF report wala data (sections) dobara use karo, taake logic dohra na likhna paray
        values = self.env["report.xbo_paknavy_pos_report.credit_sale_report_template"]._get_report_values(self.ids)
        sections = values["sections"]
        if not sections:
            raise ValidationError(_("No credit invoices found in this date range."))

        output = io.BytesIO()                                              # memory me khali file
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})        # nayi Excel file
        sheet = workbook.add_worksheet("Credit Sale Report")               # sheet ka naam

        # ---- formats (styles) ----
        title = workbook.add_format({"bold": True, "font_size": 16, "align": "center", "valign": "vcenter"})
        bold = workbook.add_format({"bold": True})
        head = workbook.add_format({"bold": True, "border": 1, "bg_color": "#E9ECEF",
                                    "align": "center", "valign": "vcenter"})
        sec_fmt = workbook.add_format({"bold": True, "font_size": 12})
        cell = workbook.add_format({"border": 1, "align": "center", "valign": "vcenter"})
        left = workbook.add_format({"border": 1, "align": "left", "valign": "vcenter"})
        num = workbook.add_format({"border": 1, "align": "right", "valign": "vcenter", "num_format": "#,##0.00"})
        tot_lbl = workbook.add_format({"bold": True, "border": 1, "bg_color": "#DEE2E6"})
        tot_num = workbook.add_format({"bold": True, "border": 1, "bg_color": "#DEE2E6", "num_format": "#,##0.00"})

        # ---- column widths: Sr, Date, Name, Job, Paid, Due, Item, Lot, Unit, Total ----
        for col, width in enumerate([7, 13, 24, 18, 14, 14, 30, 16, 14, 14]):
            sheet.set_column(col, col, width)

        # ---- title aur date range ----
        sheet.merge_range(0, 0, 0, 9, "Credit Sale Report", title)
        sheet.write(2, 0, "Date From:", bold)
        sheet.write(2, 1, self.date_from.strftime("%d/%m/%Y"))   # d/m/Y format
        sheet.write(2, 8, "Date To:", bold)
        sheet.write(2, 9, self.date_to.strftime("%d/%m/%Y"))

        headers = ["Sr No.", "Date", "Name", "Job Position", "Total Paid", "Amount Due",
                   "Item", "Lot No", "Unit Price", "Total Price"]

        row = 4  # yahan se sections shuru
        for sec in sections:
            # POS ka naam (agar POS select nahi kiya to "All POS")
            sheet.write(row, 0, sec["config"].name or "All POS", sec_fmt)
            row += 1
            # table ka header
            for col, h in enumerate(headers):
                sheet.write(row, col, h, head)
            row += 1

            for inv in sec["invoices"]:
                move = inv["move"]
                lines = inv["lines"]
                if not lines:
                    continue  # agar invoice me product line hi nahi to chor do
                first = row                          # is invoice ki pehli row
                last = row + len(lines) - 1          # is invoice ki aakhri row
                paid = move.amount_total - move.amount_residual   # ab tak ki payment
                # invoice level ki values jo saari lines ke liye aik hi hain
                common = [
                    (inv["sr"], cell),
                    (move.invoice_date.strftime("%d/%m/%Y") if move.invoice_date else "", cell),
                    (move.partner_id.name or "", left),
                    (inv["job"], left),
                    (paid, num),
                    (move.amount_residual, num),
                ]
                for col, (val, fmt) in enumerate(common):
                    if last > first:
                        sheet.merge_range(first, col, last, col, val, fmt)  # lines zyada hon to merge
                    else:
                        sheet.write(first, col, val, fmt)

                # is invoice ke har product ke lots ki list yahan jama hongi {product_id: [lots]}
                pools = {}
                # har product line ki apni row
                for line in lines:
                    lot_text = ""  # default: lot nahi
                    if line.product_id.tracking == "lot":  # sirf lot wale product ke liye
                        # is product ke lots ki list (pehli dafa nikalo, phir wahi use karo)
                        if line.product_id.id not in pools:
                            pools[line.product_id.id] = self._get_lot_pool(move, line.product_id)
                        pool = pools[line.product_id.id]
                        qty = max(int(round(line.quantity)), 1)  # is line me kitni qty hai
                        lot_text = ", ".join(pool[:qty])          # utne hi lots is line ko do
                        del pool[:qty]                            # wo lots list se nikal do, agli line ko na milein
                    sheet.write(row, 6, line.product_id.display_name or line.name, left)
                    sheet.write(row, 7, lot_text, cell)   # lot number
                    sheet.write(row, 8, line.price_unit, num)
                    sheet.write(row, 9, line.price_total, num)
                    row += 1

            # section ka total
            sheet.merge_range(row, 0, row, 8, "Total (%s)" % (sec["config"].name or "All POS"), tot_lbl)
            sheet.write(row, 9, sec["total"], tot_num)
            row += 2  # agle section se pehle aik khali row

        workbook.close()  # file mukammal

        # ---- download ----
        attachment = self.env["ir.attachment"].create({
            "name": "Credit_Sale_Report_%s_to_%s.xlsx" % (self.date_from, self.date_to),
            "type": "binary",
            "datas": base64.b64encode(output.getvalue()),
            "mimetype": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        })
        return {
            "type": "ir.actions.act_url",
            "url": "/web/content/%s?download=true" % attachment.id,
            "target": "self",
        }