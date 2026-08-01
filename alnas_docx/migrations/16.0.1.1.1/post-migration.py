from hashlib import sha256

from odoo import SUPERUSER_ID, api


def migrate(cr, version):
    env = api.Environment(cr, SUPERUSER_ID, {})
    records = env["docx.report.config"].search([])

    for record in records.filtered(lambda rec: not rec.report_name):
        record.report_name = record.action_report_id.report_name or (
            f"alnas_docx.{sha256(str(record.id).encode()).hexdigest()}"
        )

    records.filtered(lambda rec: not rec.print_report_name)._compute_print_report_name()
    records.flush_model(["report_name", "print_report_name"])
