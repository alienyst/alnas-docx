from base64 import b64encode
from io import BytesIO
from zipfile import ZipFile
import base64

from odoo.tests.common import TransactionCase
from odoo.addons.alnas_docx.tools.misc import _ensure_supported_image

class TestWebPRendering(TransactionCase):
    def test_webp_image_converts_to_png(self):
        webp_b64 = b"UklGRhoAAABXRUJQVlA4TA0AAAAvAAAAEAcQERGIiP4HAA=="
        
        # Test helper directly
        raw_bytes = base64.b64decode(webp_b64)
        processed = _ensure_supported_image(raw_bytes)
        
        # WEBP should be converted to PNG internally
        self.assertTrue(processed.startswith(b"\x89PNG\r\n\x1a\n"), "Image was not converted to PNG format.")

        # Try rendering in Odoo via python-docx
        partner = self.env["res.partner"].create({
            "name": "WebP Partner",
            "image_1920": base64.b64encode(processed) # passing PROCESSED bytes as b64 into image_1920!
        })

        from docx import Document
        template_io = BytesIO()
        doc = Document()
        doc.add_paragraph("{{ render_image(docs.image_1920) }}")
        doc.save(template_io)

        model = self.env["ir.model"]._get("res.partner")
        field = self.env["ir.model.fields"].search(
            [("model_id", "=", model.id), ("name", "=", "name")], limit=1
        )
        config = self.env["docx.report.config"].create({
            "name": "WebP Preview",
            "report_name": "webp_preview",
            "model_id": model.id,
            "field_id": field.id,
            "report_docx_template": b64encode(template_io.getvalue()),
            "report_docx_template_filename": "webp.docx",
        })

        content = config._render_preview_docx(partner.id)
        
        with ZipFile(BytesIO(content)) as rendered:
            self.assertTrue(any(f.startswith("word/media/") for f in rendered.namelist()), "Image was not included in DOCX media folder.")
