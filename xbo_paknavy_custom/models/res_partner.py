from odoo import models , api ,fields,_
from odoo.exceptions import ValidationError

class ResPartnerInherit(models.Model):
    _inherit = "res.partner"

    is_branch = fields.Boolean(string="Is Branch")

