from base64 import b64encode
from io import BytesIO

from docx import Document
from odoo.tests.common import TransactionCase

from odoo.addons.alnas_docx.tools import misc


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

    def test_subdoc_accepts_base64_binary_values(self):
        source = BytesIO()
        Document().save(source)
        raw = source.getvalue()

        class Template:
            def new_subdoc(self, stream):
                return stream.read()

        self.assertEqual(misc.add_new_subdoc(Template(), b64encode(raw)), raw)
        self.assertEqual(misc.add_new_subdoc(Template(), raw), raw)

    def test_barcode_failure_returns_empty_value(self):
        class FailingReport:
            def barcode(self, *args, **kwargs):
                raise RuntimeError("barcode unavailable")

        class FailingEnv:
            def __getitem__(self, model):
                return FailingReport()

        self.assertEqual(misc.render_barcode(FailingEnv(), None, "123"), "")
