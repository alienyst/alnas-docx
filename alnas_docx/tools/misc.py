from io import BytesIO
from base64 import b64decode
from datetime import date, datetime
from zoneinfo import ZoneInfo
from docx import Document
from docx.shared import Mm
from docxtpl import InlineImage, RichText
from bs4 import BeautifulSoup
from num2words import num2words
from babel.dates import format_date
from babel.numbers import format_currency
from htmldocx import HtmlToDocx

# Partial Function
def render_image(tpl, imgb64, width=None, height=None):
    width = Mm(width) if width else None
    height = Mm(height) if height else None
        
    if not imgb64:
        return ''

    image_stream = BytesIO(b64decode(imgb64))
    return InlineImage(
        tpl, image_descriptor=image_stream, width=width, height=height
    )
    
def render_html_as_subdoc(tpl, html_code=None):
    if not (
        isinstance(html_code, str)
        and bool(BeautifulSoup(html_code, "html.parser").find())
    ):
        return ""

    temp = BytesIO()
    desc_document = Document()
    new_parser = HtmlToDocx()
    new_parser.add_html_to_document(html_code, desc_document)
    desc_document.save(temp)
    temp.seek(0)
    return tpl.new_subdoc(temp)


def add_new_subdoc(tpl, docx_file):
    if not docx_file:
        return ""
    try:
        # Odoo Binary fields are base64 str; callers may also pass raw file bytes.
        if isinstance(docx_file, str):
            raw = b64decode(docx_file)
        else:
            raw = bytes(docx_file)
        # skip if not a valid docx file
        if not raw.startswith(b"PK\x03\x04"):
            return ""
        return tpl.new_subdoc(BytesIO(raw))
    except Exception:
        return ""

def linked_attachments_for_record(env, record):
    """Attachments linked to ``record`` via ``res_model`` / ``res_id`` (binary only)."""
    if not record or not record.ids:
        return env["ir.attachment"].browse()
    try:
        res_id = int(record.ids[0])
    except (TypeError, ValueError):
        return env["ir.attachment"].browse()
    return env["ir.attachment"].search(
        [
            ("res_model", "=", record._name),
            ("res_id", "=", res_id),
            ("type", "=", "binary"),
        ],
        order="id",
    )

def replace_image(tpl, dummy_pic, imgb64):
    if not imgb64:
        return ''

    tpl.replace_pic(dummy_pic, BytesIO(b64decode(imgb64)))
    return ''

def replace_media(tpl, dummy_pic, imgb64):
    if not imgb64:
        return ''

    tpl.replace_media(dummy_pic, BytesIO(b64decode(imgb64)))
    return ''

def replace_embedded(tpl, dummy_embeed, file_b64):
    if not file_b64:
        return ''

    tpl.replace_embedded(dummy_embeed, BytesIO(b64decode(file_b64)))
    return ''

def replace_zipname(tpl, embedded_object_path, file_b64):
    if not file_b64:
        return ''

    tpl.replace_zipname(embedded_object_path, BytesIO(b64decode(file_b64)))
    return ''

# Formatting Function
def parse_html(html):
    if not html:
        return ""
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text()

def formatdate(date_required=datetime.today(), format="full", lang="id_ID", **kwargs):
    return format_date(date_required, format=format, locale=lang, **kwargs)

def format_datetime(dt, tz=None, format="%Y-%m-%d %H:%M:%S", **kwargs):
    if not dt:
        return ""
    if tz is None:
        tz_target = ZoneInfo("UTC")
    elif isinstance(tz, ZoneInfo):
        tz_target = tz
    else:
        try:
            tz_target = ZoneInfo(str(tz))
        except Exception:
            tz_target = ZoneInfo("UTC")
    if isinstance(dt, date) and not isinstance(dt, datetime):
        return dt.strftime(format)
    if isinstance(dt, datetime):
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=ZoneInfo("UTC"))
        local_dt = dt.astimezone(tz_target)
        return local_dt.strftime(format)
    return str(dt)

def spelled_out(number, lang="id_ID", to="cardinal", **kwargs):
    return num2words(number, lang=lang, to=to, **kwargs)

def convert_currency(number, currency_field, locale='id_ID', **kwargs):
    return format_currency(number, currency_field.name, locale=locale, **kwargs)

def format_abs(number):
    return abs(number)

def rich_text(text, **kwargs):
    if not text:
        return ""
    
    return RichText(text, **kwargs)