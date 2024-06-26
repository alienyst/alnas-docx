import json
from werkzeug.urls import url_decode

from odoo.http import (
    content_disposition,
    request,
    route,
    serialize_exception as _serialize_exception,
)

from odoo.tools import html_escape
from odoo.tools.safe_eval import safe_eval, time

from odoo.addons.web.controllers.report import ReportController


class DocxReportController(ReportController):
    @route()
    def report_routes(self, reportname, docids=None, converter=None, **data):
        if converter == "docx":
            
            report = request.env["ir.actions.report"]._get_report_from_name(reportname)
            context = dict(request.env.context)
            
            if docids:
                docids = [int(i) for i in docids.split(",")]
            if data.get("options"):
                data.update(json.loads(data.pop("options")))
            if data.get("context"):
                data["context"] = json.loads(data["context"])
                context.update(data["context"])
                
            docx_files = report.with_context(**context)._render_docx(reportname, docids, data=data)
            
            if report.docx_merge_mode == 'composer':
                httpheaders = [
                    ('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
                ]
            elif report.docx_merge_mode == 'zip':
                httpheaders = [
                    ('Content-Type', 'application/zip'),
                ]
            else:
                httpheaders = [
                    ('Content-Type', 'application/pdf'),
                ]
            return request.make_response(docx_files, headers=httpheaders)
        
        return super().report_routes(reportname, docids, converter, **data)

    @route()
    def report_download(self, data, context=None, token=None):
        requestcontent = json.loads(data)
        url, report_type = requestcontent[0], requestcontent[1]
        try:
            if report_type == "docx":
                reportname = url.split("/report/docx/")[1].split("?")[0]
                docids = None
                if "/" in reportname:
                    reportname, docids = reportname.split("/")
                if docids:
                    # Generic report:
                    response = self.report_routes(
                        reportname, docids=docids, converter="docx", context=context
                    )
                else:
                    # Particular report:
                    data = dict(
                        url_decode(url.split("?")[1]).items()
                    )  # decoding the args represented in JSON
                    if "context" in data:
                        context, data_context = json.loads(context or "{}"), json.loads(
                            data.pop("context")
                        )
                        context = json.dumps({**context, **data_context})
                    response = self.report_routes(
                        reportname, docids=docids, converter="docx", context=context, **data
                    )

                report = request.env["ir.actions.report"]._get_report_from_name(
                    reportname
                )
                
                filename = self._get_filename_by_report_type(report, report.name)
                    
                if docids:
                    ids = [int(x) for x in docids.split(",")]
                    obj = request.env[report.model].browse(ids)
                    if report.print_report_name and not len(obj) > 1:
                        report_name = safe_eval(
                            report.print_report_name, {"object": obj, "time": time}
                        )
                        filename = self._get_filename_by_report_type(report, report_name)
                if not response.headers.get("Content-Disposition"):
                    response.headers.add(
                        "Content-Disposition", content_disposition(filename)
                    )
                return response
            else:
                return super().report_download(data, context=context, token=token)
        except Exception as e:
            se = _serialize_exception(e)
            error = {"code": 200, "message": "Odoo Server Error", "data": se}
            return request.make_response(html_escape(json.dumps(error)))
        
    def _get_filename_by_report_type(self, report, name):
        if report.docx_merge_mode == 'composer':
            filename = "%s.%s" % (name, "docx")
        elif report.docx_merge_mode == 'zip':
            filename = "%s.%s" % (name, "zip")
        else:
            filename = "%s.%s" % (name, "pdf")
        return filename

