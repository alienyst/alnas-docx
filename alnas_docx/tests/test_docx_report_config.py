from base64 import b64encode

from odoo.tests.common import TransactionCase


class TestDocxReportConfig(TransactionCase):
    def test_odoo19_domain_copy_and_model_onchange(self):
        model = self.env["ir.model"]._get("res.partner")
        field = self.env["ir.model.fields"].search(
            [("model_id", "=", model.id), ("name", "=", "name")], limit=1
        )
        config = self.env["docx.report.config"].create(
            {
                "name": "Partner Report",
                "report_name": "partner_report",
                "model_id": model.id,
                "field_id": field.id,
                "report_docx_template": b64encode(b"docx"),
                "report_docx_template_filename": "partner.docx",
                "domain": "[('active', '=', True)]",
            }
        )

        copied = config.copy_data()[0]
        self.assertEqual(copied["name"], "Partner Report (copy)")
        self.assertEqual(copied["report_name"], "partner_report (copy)")
        config._action_publish()
        self.assertEqual(config.action_report_id.domain, config.domain)
        config._action_unpublish()

        onchange = self.env["docx.report.config"].new({"model_id": model.id})
        onchange._onchange_model_id()
        self.assertEqual(onchange.field_id.model_id, model)
        self.assertEqual(onchange.field_id.ttype, "char")
