# -*- coding: utf-8 -*-
from odoo import api, models


class ReportCreditSale(models.AbstractModel):
    _name = "report.xbo_paknavy_pos_report.credit_sale_report_template"
    _description = "Credit Sale Report"

    @api.model
    def _get_report_values(self, docids, data=None):
        wizard = self.env["credit.sale.report.wizard"].browse(docids)

        configs = wizard.pos_config_ids

        domain = [
            ("move_type", "=", "out_invoice"),
            ("state", "=", "posted"),
            ("invoice_date", ">=", wizard.date_from),
            ("invoice_date", "<=", wizard.date_to),
            ("payment_state", "in", ["partial", "not_paid"]),
        ]

        if configs:
            domain.append(("pos_order_ids.config_id", "in", configs.ids))

        moves = self.env["account.move"].search(domain, order="invoice_date, id")

        grouped = {}
        for move in moves:
            config = (move.pos_order_ids.config_id & configs)[:1]
            grouped.setdefault(config, self.env["account.move"])
            grouped[config] |= move
        Employee = self.env["hr.employee"].sudo()
        sections = []
        for config in sorted(grouped, key=lambda c: c.name):
            invoices = []
            for sr, move in enumerate(grouped[config], start=1):
                employee = Employee.search(
                    [("work_contact_id", "=", move.partner_id.id)],
                    limit=1,
                )
                invoices.append({
                    "sr": sr,
                    "move": move,
                    "job": employee.job_id.name or "",
                    "lines": move.invoice_line_ids.filtered(
                        lambda l: l.display_type == "product"
                    ),
                })
            sections.append({
                "config": config,
                "invoices": invoices,
                "total": sum(grouped[config].mapped("amount_total")),
                "balance": sum(grouped[config].mapped("amount_residual")),
            })

        return {
            "doc_ids": docids,
            "doc_model": "credit.sale.report.wizard",
            "docs": wizard,
            "date_from": wizard.date_from,
            "date_to": wizard.date_to,
            "sections": sections,
        }