from base64 import b64decode
from io import BytesIO
from zoneinfo import ZoneInfo

from babel.dates import format_date, format_datetime
from babel.numbers import format_currency
from docx import Document
from docx.shared import Mm
from docxtpl import InlineImage, RichText
from htmldocx import HtmlToDocx
from num2words import num2words
from odoo import _
from odoo.exceptions import UserError
from odoo.tools import is_html_empty
from odoo.tools.pdf import PdfReader, merge_pdf
from pytz import timezone


# Partial Function
def render_image(tpl, imgb64, width=None, height=None):
    width = Mm(width) if width else None
    height = Mm(height) if height else None

    if not imgb64:
        return ""

    image_stream = BytesIO(b64decode(imgb64))
    return InlineImage(tpl, image_descriptor=image_stream, width=width, height=height)


def render_html_as_subdoc(tpl, html_code=None):
    if not isinstance(html_code, str) or is_html_empty(html_code):
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


def _pdf_bytes_from_source(source, label=None):
    """Return (raw_pdf_bytes, error_label) or None to skip (empty / falsy)."""
    if source is None or source is False:
        return None
    if hasattr(source, "ids"):
        if len(source) == 0:
            return None
        if len(source) != 1:
            raise UserError(
                _("add_pdf: expected a single record, got %(count)d.")
                % {"count": len(source)}
            )
        source = source[0]
    if hasattr(source, "_name") and source._name == "ir.attachment":
        att = source.with_context(bin_size=False)
        data = None
        if getattr(att, "raw", None):
            data = att.raw
        if not data and att.datas:
            data = b64decode(att.datas)
        if not data:
            return None
        lbl = label or att.display_name or att.name or _("attachment")
        return data, lbl
    if isinstance(source, (bytes, bytearray, memoryview)):
        data = bytes(source)
        if not data:
            return None
        return data, label or _("PDF")
    if isinstance(source, str):
        if not source.strip():
            return None
        try:
            data = b64decode(source, validate=True)
        except TypeError:
            data = b64decode(source)
        if not data:
            return None
        return data, label or _("PDF")
    raise UserError(
        _("add_pdf: unsupported type %(typ)s. Use binary data or ir.attachment.")
        % {"typ": type(source).__name__}
    )


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


def _coerce_pdf_bytes(data):
    """Return raw PDF bytes: decode when the buffer is still base64 text (e.g. ``JVBERi...``)."""
    if not data:
        return data
    if isinstance(data, str):
        try:
            chunk = data.strip().encode("ascii")
        except UnicodeEncodeError:
            return data
    else:
        chunk = bytes(data)
    if chunk.lstrip().startswith(b"%PDF"):
        return chunk
    try:
        decoded = b64decode(chunk.strip())
    except Exception:
        return chunk
    if decoded.lstrip().startswith(b"%PDF"):
        return decoded
    return chunk


def _make_pdf_reader(data):
    stream = BytesIO(data)
    try:
        return PdfReader(stream, strict=False, root_object_recovery_limit=None)
    except TypeError:
        return PdfReader(stream, strict=False)


def _validate_pdf_bytes(data, label):
    """Validate and return coerced PDF bytes for storage and merging."""
    data = _coerce_pdf_bytes(data)
    reader = None
    try:
        reader = _make_pdf_reader(data)
        if len(reader.pages) == 0:
            raise UserError(_("PDF has no pages (%(label)s).") % {"label": label})
    except UserError:
        raise
    except Exception as err:
        raise UserError(
            _("Not a valid PDF (%(label)s): %(msg)s")
            % {"label": label, "msg": str(err)}
        ) from err
    finally:
        if reader is not None and hasattr(reader, "close"):
            reader.close()
    return data


def add_pdf_factory(before_list, after_list):
    """Side-effect helpers for PDF output mode only (merge after main report PDF)."""

    def add_pdf(source, position="after", label=None):
        got = _pdf_bytes_from_source(source, label=label)
        if got is None:
            return ""
        data, err_label = got

        data = _validate_pdf_bytes(data, err_label)
        pos = (position or "after").lower()
        if pos not in ("before", "after"):
            raise UserError(
                _("add_pdf: position must be 'before' or 'after', not %(pos)r.")
                % {"pos": position}
            )
        if pos == "before":
            before_list.append(data)
        else:
            after_list.append(data)
        return ""

    return add_pdf


def merge_pdf_bytes(main_pdf_bytes, before_list, after_list):
    """Concatenate PDFs: before_list + main + after_list."""
    if not before_list and not after_list:
        return main_pdf_bytes

    chunks = [_coerce_pdf_bytes(x) for x in [*before_list, main_pdf_bytes, *after_list]]

    return merge_pdf(chunks)


def replace_image(tpl, dummy_pic, imgb64):
    if not imgb64:
        return ""

    tpl.replace_pic(dummy_pic, BytesIO(b64decode(imgb64)))
    return ""


def replace_media(tpl, dummy_pic, imgb64):
    if not imgb64:
        return ""

    tpl.replace_media(dummy_pic, BytesIO(b64decode(imgb64)))
    return ""


def replace_embedded(tpl, dummy_embeed, file_b64):
    if not file_b64:
        return ""

    tpl.replace_embedded(dummy_embeed, BytesIO(b64decode(file_b64)))
    return ""


def replace_zipname(tpl, embedded_object_path, file_b64):
    if not file_b64:
        return ""

    tpl.replace_zipname(embedded_object_path, BytesIO(b64decode(file_b64)))
    return ""

def format_selection(record, field_name):
    """Return the translated label of a selection field for the given record."""
    if not record or not field_name or field_name not in record._fields:
        return ""
    try:
        field_info = record.fields_get([field_name]).get(field_name, {})
        selection = dict(field_info.get("selection", []))
        value = record[field_name]
        return selection.get(value, value or "")
    except Exception:
        return record[field_name] or ""

def render_qrcode(env, tpl, value, width=15, height=15):
    """Generate a QR code using Odoo's native barcode generator."""
    return render_barcode(env, tpl, str(value), barcode_type="QR", width=width, height=height)

def render_barcode(env, tpl, value, barcode_type="Code128", width=None, height=None, **kwargs):
    """Generate a barcode/QR code using Odoo's native barcode generator."""
    if not isinstance(value, str) or not value:
        return ""
    try:
        bar_bytes = env['ir.actions.report'].barcode(barcode_type, value, **kwargs)
        w = Mm(width) if width else None
        h = Mm(height) if height else None
        return InlineImage(tpl, BytesIO(bar_bytes), width=w, height=h)
    except Exception as e:
        _logger.warning("Failed to generate %s: %s", barcode_type, str(e))
        return ""


# Formatting Function
def formatdate(date_required=None, format="full", lang="id_ID", **kwargs):
    if not date_required:
        return ""
    return format_date(date_required, format=format, locale=lang, **kwargs)


def formatdatetime(
    dt, user_tz="UTC", format="dd/MM/yyyy HH:mm", lang="id_ID", **kwargs
):
    if not dt:
        return ""
    return format_datetime(
        dt,
        format=format,
        tzinfo=timezone(user_tz),
        locale=lang,
        **kwargs,
    )


def spelled_out(number, lang="id_ID", to="cardinal", **kwargs):
    return num2words(number, lang=lang, to=to, **kwargs)


def convert_currency(number, currency_field, locale="id_ID", **kwargs):
    return format_currency(number, currency_field.name, locale=locale, **kwargs)


def format_abs(number):
    return abs(number)


def rich_text(text, **kwargs):
    if not text:
        return ""

    return RichText(text, **kwargs)
