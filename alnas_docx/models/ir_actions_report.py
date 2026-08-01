import base64
import zipfile
import os
import subprocess
import tempfile
import shutil
import logging
from io import BytesIO
from functools import partial
from docx import Document
from docxtpl import DocxTemplate
from docxcompose.composer import Composer

from odoo import _, api, fields, models, tools
from odoo.tools.safe_eval import safe_eval, time
from odoo.exceptions import ValidationError, MissingError, UserError

from ..tools import misc as misc_tools

_logger = logging.getLogger(__name__)


class IrActionsReport(models.Model):
    _inherit = "ir.actions.report"

    report_type = fields.Selection(
        selection_add=[("docx", "DOCX")], ondelete={"docx": "cascade"}
    ) # add docx type
    report_docx_template = fields.Binary(string="Report DOCX Template")
    report_docx_template_name = fields.Char(string="Report DOCX Template Name")
    docx_merge_mode = fields.Selection(
        [("composer", "Composer"), ("zip", "Zip"), ("pdf", "PDF")],
        string="DOCX Mode",
        default="composer",
    )
    docx_autoescape = fields.Boolean(
        string="Autoescape (Docx)",
        default=False,
        help="Enable autoescape for special character like <, > and &.",
    )

    @api.constrains("report_type", "report_docx_template", "report_docx_template_name")
    def _check_report_type(self):
        for rec in self:
            if rec.report_type == "docx":
                if not rec.report_docx_template or not (rec.report_docx_template_name or "").lower().endswith(".docx"):
                    raise ValidationError(_("Please upload a valid .docx template."))

    def _get_rendering_context_docx(self, doc_template, extra_pdfs=None):
        context = self.env["mail.render.mixin"]._render_eval_context()
        context.update({
            "company": self.env.company,
            "lang": self._context.get("lang", "id_ID"),
            "sysdate": fields.Datetime.now(),
            "html2plaintext": tools.html2plaintext,
            "spelled_out": misc_tools.spelled_out,
            "format_selection": misc_tools.format_selection,
            "parsehtml": tools.html2plaintext, # Aliased for backward compatibility
            "formatdate": misc_tools.formatdate, # Deprecated: use native format_date
            "formatdatetime": partial(misc_tools.formatdatetime, user_tz=self.env.user.tz or 'UTC'), # Deprecated: use native format_datetime
            "convert_currency": misc_tools.convert_currency, # Deprecated: use native format_amount
            "formatabs": misc_tools.format_abs,
            "rich_text": misc_tools.rich_text,
            "render_image": partial(misc_tools.render_image, doc_template),
            "render_qrcode": partial(misc_tools.render_qrcode, doc_template),
            "render_barcode": partial(misc_tools.render_barcode, self.env, doc_template),
            "html2docx": partial(misc_tools.render_html_as_subdoc, doc_template),
            "add_subdoc": partial(misc_tools.add_new_subdoc, doc_template),
            "replace_image": partial(misc_tools.replace_image, doc_template),
            "replace_media": partial(misc_tools.replace_media, doc_template),
            "replace_embedded": partial(misc_tools.replace_embedded, doc_template),
            "replace_zipname": partial(misc_tools.replace_zipname, doc_template),
            "linked_attachments": lambda record: misc_tools.linked_attachments_for_record(
                self.env, record
            ),
            "add_pdf": (
                misc_tools.add_pdf_factory(extra_pdfs["before"], extra_pdfs["after"])
                if extra_pdfs is not None
                else (lambda *args, **kwargs: "")
            ),
        })
        return context
    
    def _render_docx(self, report_ref, docids, data):
        report = self._get_report(report_ref)
        template = report.report_docx_template

        if not template:
            raise MissingError("No DOCX template found.")

        doc_template = DocxTemplate(BytesIO(base64.b64decode(template)))
        doc_obj = self.env[report.model].browse(docids).with_context(
            bin_size=False
        )
        extra_pdfs = (
            {"before": [], "after": []}
            if report.docx_merge_mode == "pdf"
            else None
        )
        context = self._get_rendering_context_docx(
            doc_template=doc_template, extra_pdfs=extra_pdfs
        )
        autoescape = report.docx_autoescape
        
        if report.docx_merge_mode == "composer":
            return self._render_composer_mode(doc_template, doc_obj, data, context, autoescape=autoescape)
        elif report.docx_merge_mode == "zip":
            return self._render_zip_mode(
                doc_template,
                doc_obj,
                data,
                context,
                report_name=report.print_report_name,
                autoescape=autoescape,
            )
        else:
            return self._render_docx_to_pdf_mode(
                doc_template, doc_obj, data, context, extra_pdfs, autoescape=autoescape
            )

    def _render_composer_mode(self, doc_template, doc_obj, data, context, autoescape=False):
        if not doc_obj:
            temp = BytesIO()
            doc_template.render({**context, "docs": doc_obj, "data": data}, autoescape=autoescape)
            doc_template.save(temp)
            temp.seek(0)
            return temp.read(), 'docx'

        for idx, obj in enumerate(doc_obj):
            ctx = {
                **context,
                "docs": obj,
                "data": data
            }

            temp = BytesIO()
            doc_template.render(ctx, autoescape=autoescape)
            doc_template.save(temp)
            temp.seek(0)

            if len(doc_obj) == 1:
                return temp.read(), 'docx'
            else:
                if idx == 0:
                    master_doc = Document(temp)
                    composer = Composer(master_doc)
                else:
                    doc_to_append = Document(temp)
                    master_doc.add_page_break()
                    composer.append(doc_to_append)

        temp_output = BytesIO()
        composer.save(temp_output)
        temp_output.seek(0)

        return temp_output.read(), 'docx'

    def _render_zip_mode(
        self, doc_template, doc_obj, data, context, report_name="report", autoescape=False
    ):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for idx, obj in enumerate(doc_obj):
                ctx = {
                    **context,
                    "docs": obj,
                    "data": data
                }

                temp = BytesIO()
                doc_template.render(ctx, autoescape=autoescape)
                doc_template.save(temp)
                
                name = safe_eval(report_name, {"object": obj, "time": time}) if report_name else f"report_{idx+1}"
                safe_filename = str(name).replace("/", "_").replace("\\", "_") + ".docx"
                zip_file.writestr(safe_filename, temp.getvalue())

        return zip_buffer.getvalue(), 'zip'

    def _render_docx_to_pdf_mode(
        self, doc_template, doc_obj, data, context, extra_pdfs, autoescape=False
    ):
        docx_file, _ = self._render_composer_mode(
            doc_template, doc_obj, data, context, autoescape=autoescape
        )
        temp_dir = tempfile.mkdtemp()

        try:
            docx_file_path = os.path.join(temp_dir, 'document.docx')
            with open(docx_file_path, 'wb') as f:
                f.write(docx_file)

            pdf_file_path = self.convert_file_to_pdf(docx_file_path, temp_dir)

            if not pdf_file_path:
                raise UserError('PDF conversion failed.')

            with open(pdf_file_path, 'rb') as pdf_file:
                main_pdf = pdf_file.read()

        finally:
            shutil.rmtree(temp_dir)

        if extra_pdfs and (extra_pdfs["before"] or extra_pdfs["after"]):
            main_pdf = misc_tools.merge_pdf_bytes(
                main_pdf, extra_pdfs["before"], extra_pdfs["after"]
            )

        return main_pdf, 'pdf'

    def convert_file_to_pdf(self, file_path, output_dir):
        librepath = self._get_libreoffice_path()
        profile_dir = os.path.join(output_dir, "lo_profile")
        
        command = [
            librepath,
            '--headless',
            '--invisible',
            '--nologo',
            '--nodefault',
            '--norestore',
            '--nolockcheck',
            '--nofirststartwizard',
            f'-env:UserInstallation=file://{profile_dir}',
            '--convert-to', 'pdf',
            '--outdir', output_dir,
            file_path
        ]
        
        env = os.environ.copy()
        env['SAL_DISABLE_OPENCL'] = '1'
        
        try:
            result = subprocess.run(command, env=env, timeout=120, capture_output=True, text=True)
            if result.returncode != 0:
                _logger.error("LibreOffice PDF conversion failed.\nSTDOUT: %s\nSTDERR: %s", result.stdout, result.stderr)
                raise UserError(f"PDF conversion failed (exit code {result.returncode}). Check logs for details.\nSTDERR: {result.stderr.strip()[-200:]}")
        except subprocess.TimeoutExpired as e:
            _logger.error("LibreOffice PDF conversion timed out.\nSTDOUT: %s\nSTDERR: %s", e.stdout, e.stderr)
            raise UserError('PDF conversion timed out.')
        except Exception as e:
            if isinstance(e, UserError):
                raise
            _logger.exception("LibreOffice PDF conversion exception.")
            raise UserError(f'PDF conversion error: {str(e)}')
        
        pdf_file_name = os.path.splitext(os.path.basename(file_path))[0] + '.pdf' 
        pdf_file_path = os.path.join(output_dir, pdf_file_name)        
        if os.path.exists(pdf_file_path):
            return pdf_file_path
        else:
            return None

    def _get_libreoffice_path(self):
        libreoffice = self.env.ref('alnas_docx.default_libreoffice_path')
        if not libreoffice and not libreoffice.value:
            raise ValidationError('Libreoffice path doesnt exits, \n \
                please set in Settings => Technical => Parameters => System Parameters => default_libreoffice_path')
            
        return libreoffice.value
