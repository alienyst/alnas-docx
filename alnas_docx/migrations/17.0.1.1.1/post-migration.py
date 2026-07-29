from odoo import api, SUPERUSER_ID


def migrate(cr, version):
    env = api.Environment(cr, SUPERUSER_ID, {})

    records = env["docx.report.config"].search([
        "|",
        ("print_report_name", "=", False),
        ("print_report_name", "=", ""),
    ])

    if records:
        for rec in records:
            rec.print_report_name = False

        records._compute_print_report_name()
        env["docx.report.config"].flush_model(["print_report_name"])