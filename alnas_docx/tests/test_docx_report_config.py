from base64 import b64encode

from odoo.tests.common import TransactionCase


class TestDocxReportConfig(TransactionCase):
    def test_copy_model_onchange_and_idempotent_actions(self):
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
            }
        )

        copied = config.copy_data()[0]
        self.assertEqual(copied["name"], "Partner Report (copy)")
        self.assertEqual(copied["report_name"], "partner_report (copy)")

        onchange = self.env["docx.report.config"].new({"model_id": model.id})
        onchange._onchange_model_id()
        self.assertEqual(onchange.field_id.model_id, model)
        self.assertEqual(onchange.field_id.ttype, "char")

        config._action_publish()
        config._action_publish()
        config._action_unpublish()
        config._action_unpublish()
