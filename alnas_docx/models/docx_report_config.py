import re

from odoo import models, fields, api
from odoo.exceptions import UserError
from odoo.tools.safe_eval import safe_eval, time


# Maps a {placeholder} in the file name pattern to a small piece of Python.
# {field} and {model} are filled in per record when the expression is built.
NAME_PLACEHOLDERS = {
    "record_name": "(object.{field} or '')",
    "model_name": "'{model}'",
    "year": "(object.create_date and object.create_date.year or '')",
    "quarter": "(object.create_date and 'Q%d' % ((object.create_date.month - 1) // 3 + 1) or '')",
    "month": "(object.create_date and '%02d' % object.create_date.month or '')",
    "day": "(object.create_date and '%02d' % object.create_date.day or '')",
}


class DocxReportConfig(models.Model):
    _name = "docx.report.config"
    _description = "DOCX Report Configuration"

    _inherit = ["mail.thread", "mail.activity.mixin"]
    
    _sql_constraints = [
        ('report_code_name', 'UNIQUE(report_name)', 'Report code name must be unique!.')
    ] 

    name = fields.Char(
        string="Report Name",
        required=True,
        readonly=True,
        help="Name of the report",
    )
    report_name = fields.Char(
        string="Report Code",
        required=True,
        help="Report Unique Code use for Technical Purpose",
        copy=False
    )
    
    model_id = fields.Many2one(
        "ir.model",
        string="Model",
        required=True,
        ondelete="cascade",
        readonly=True,
        help="Model to which this report will be attached",
    )
    field_id = fields.Many2one(
        "ir.model.fields",
        string="Field Name",
        required=True,
        ondelete="cascade",
        domain="[('model_id', '=', model_id),('ttype', '=', 'char')]",
        readonly=True,
        help="Field to be used as the report name",
    )
    report_docx_template = fields.Binary(
        string="Report DOCX Template",
        required=True,
        readonly=True,
        help="DOCX template to be used for the report",
    )
    report_docx_template_filename = fields.Char(
        string="Report DOCX Template Name",
        required=True,
        readonly=True,
    )
    prefix = fields.Char(
        string="Prefix",
        readonly=True,
        help="Prefix to be used in the report name",
    )
    state = fields.Selection(
        [("draft", "Draft"), ("published", "Published")],
        string="State",
        default="draft",
        tracking=True,
        copy=False,
        readonly=True,
    )
    action_report_id = fields.Many2one(
        "ir.actions.report", string="Related Report Action", readonly=True, copy=False
    )
    docx_merge_mode = fields.Selection(
        [("composer", "Composer"), ("zip", "Zip"), ("pdf", "PDF")],
        string="DOCX Merge Mode",
        default="composer",
        required=True,
        readonly=True,
        help="Mode to be used for merging the DOCX template with the data, \n \
            if 'Composer' is selected, the report will be generated as a single DOCX file, \n \
            if 'Zip' is selected, the report will be generated as a ZIP file containing multiple DOCX files, \n \
            if 'PDF' is selected, the report will be converted to PDF file.",
    )

    name_pattern = fields.Char(
        string="File Name Pattern",
        readonly=True,
        help="Text used to build the report file name. "
        "Placeholders you can use: {record_name}, {year}, {quarter}, {month}, "
        "{day}, {model_name}. Example: Contract {record_name} {year}. "
        "A change here only takes effect after the report has been "
        "unpublished and published again.",
    )
    print_report_name = fields.Char(
        string="Print Report Name",
        compute="_compute_print_report_name",
        store=True,
        help="Python expression used to build the file name. "
        "Generated from the File Name Pattern.",
    )
    autoescape = fields.Boolean(
        string="Autoescape",
        default=False,
        help="Enable autoescape for special character like <, > and &.",
    )

    @api.depends("name_pattern", "model_id", "field_id", "prefix")
    def _compute_print_report_name(self):
        for rec in self:
            rec.print_report_name = rec._build_name_expression()

    def _build_name_expression(self):
        """Turn the friendly File Name Pattern into a Python expression string.

        The result is later evaluated with ``object`` (the record) and ``time``
        in scope, in the report controller and in ``_render_zip_mode``.
        """
        self.ensure_one()
        field_name = self.field_id.name or "id"
        model_name = self.model_id.name or "Report"

        # No pattern typed: keep the previous behaviour (prefix + record field).
        pattern = self.name_pattern
        if not pattern:
            prefix = self.prefix or model_name
            pattern = prefix + " {record_name}"

        pieces = []
        # Split into plain text and {placeholder} tokens, keeping both.
        for token in re.split(r"(\{[a-z_]+\})", pattern):
            if not token:
                continue
            if token.startswith("{") and token.endswith("}"):
                key = token[1:-1]
                snippet = NAME_PLACEHOLDERS.get(key)
                if snippet is None:
                    raise UserError("Unknown file name placeholder: %s" % token)
                snippet = snippet.format(field=field_name, model=model_name)
                pieces.append("str(%s)" % snippet)
            else:
                # Plain text: repr() wraps it safely in quotes so that spaces,
                # quotes or % signs cannot break the expression.
                pieces.append(repr(token))

        return " + ".join(pieces) if pieces else "''"

    @api.constrains("report_docx_template_filename")
    def _check_report_docx_template_filename(self):
        for rec in self:
            if not rec.report_docx_template_filename.endswith(".docx"):
                raise UserError("Please upload a DOCX template.")

    @api.constrains("print_report_name", "model_id")
    def _check_print_report_name(self):
        for rec in self:
            if not rec.print_report_name or not rec.model_id:
                continue
            try:
                dummy = rec.env[rec.model_id.model].new({})
                safe_eval(rec.print_report_name, {"object": dummy, "time": time})
            except Exception as e:
                raise UserError("The file name pattern is not valid: %s" % e)

    def _action_publish(self):
        for record in self:
            if record.state == "draft":
                val = record._prepare_action_val()
                if not record.action_report_id:
                    action_report = self.env["ir.actions.report"].sudo().create(val)
                else:
                    action_report = record.action_report_id
                    action_report.sudo().write(val)

                action_report.create_action()
                record.action_report_id = action_report
                record.state = "published"
            else:
                raise UserError("Report already published")

        return True

    def action_publish(self):
        self._action_publish()
        return self._refresh_page()

    def _action_unpublish(self):
        for record in self:
            if record.state == "published":
                record.action_report_id.unlink_action()
                record.state = "draft"
            else:
                raise UserError("Report already unpublished")
        return True

    def action_unpublish(self):
        self._action_unpublish()
        return self._refresh_page()

    def _prepare_action_val(self):
        return {
            "name": self.name,
            "model": self.model_id.model,
            "report_type": "docx",
            "report_docx_template": self.report_docx_template,
            "report_docx_template_name": self.report_docx_template_filename,
            "report_name": self.report_name,
            "docx_merge_mode": self.docx_merge_mode,
            'docx_autoescape': self.autoescape,
            "print_report_name": self.print_report_name,
        }

    @api.ondelete(at_uninstall=False)
    def _unlink_docx_report(self):
        for rec in self:
            if rec.state == "published":
                rec.action_unpublish()
            if rec.action_report_id:
                rec.action_report_id.unlink()

    def _refresh_page(self):
        return {
            "type": "ir.actions.client",
            "tag": "reload",
        }
