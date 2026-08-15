from base64 import b64encode
from io import BytesIO
import base64

from odoo.tests.common import TransactionCase
from odoo.addons.alnas_docx.tools.misc import _ensure_supported_image

class TestWebPRendering(TransactionCase):
    def test_webp_image_converts_to_png(self):
        webp_b64 = b"UklGRhoAAABXRUJQVlA4TA0AAAAvAAAAEAcQERGIiP4HAA=="
        raw_bytes = base64.b64decode(webp_b64)
        
        from PIL import Image
        try:
            from PIL import WebPImagePlugin
            Image.register_open(WebPImagePlugin.WebPImageFile.format, WebPImagePlugin.WebPImageFile, WebPImagePlugin._accept)
            Image.register_extension(WebPImagePlugin.WebPImageFile.format, ".webp")
            Image.register_mime(WebPImagePlugin.WebPImageFile.format, "image/webp")
        except Exception:
            pass
            
        try:
            stream = BytesIO(raw_bytes)
            with Image.open(stream) as img:
                out = BytesIO()
                img.save(out, format='PNG')
        except Exception as e:
            self.fail(f"PIL conversion failed inside Odoo: {type(e).__name__} - {str(e)}")
            
        processed = _ensure_supported_image(raw_bytes)
        self.assertTrue(processed.startswith(b"\x89PNG\r\n\x1a\n"), "Image was not converted to PNG format.")
