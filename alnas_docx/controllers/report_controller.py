import json

from odoo.http import content_disposition, request, route, serialize_exception
from odoo.tools import html_escape
from odoo.tools.safe_eval import safe_eval

from odoo.addons.web.controllers import main as report


class ReportController(report.ReportController):
    @route()
    def report_routes(self, reportname, docids=None, converter=None, **data):
        if converter == "docxtpl":
            return self._report_routes_docxtpl(reportname, docids, converter, **data)
        return super(ReportController, self).report_routes(
            reportname, docids, converter, **data
        )

    def _report_routes_docxtpl(self, reportname, docids=None, converter=None, **data):
        try:
            report = request.env["ir.actions.report"]._get_report_from_name(reportname)
            filename = self._get_filename_by_report_type(report, report.name)

            context = dict(request.env.context)
            if docids:
                docids = [int(i) for i in docids.split(",")]
            if data.get("options"):
                data.update(json.loads(data.pop("options")))
            if data.get("context"):
                data["context"] = json.loads(data["context"])
                if data["context"].get("lang"):
                    del data["context"]["lang"]
                context.update(data["context"])

            docxtpl_files, _ = report.with_context(**context)._render_docxtpl(res_ids=docids, data=data)

            report_name = report.name
            if report.print_report_name and not len(docids) > 1:
                obj = request.env[report.model].browse(docids[0])
                report_name = safe_eval(report.print_report_name, {"object": obj})
                filename = self._get_filename_by_report_type(report, report_name)
            
            httpheaders = [
                ("Content-Length", len(docxtpl_files)),
                ("Content-Disposition", content_disposition(filename))
            ]
            
            if report.docx_merge_mode == 'composer':
                httpheaders.append(
                    ('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
                )
            elif report.docxtpl_merge_mode == 'zip':
                httpheaders.append(
                    ('Content-Type', 'application/zip'),
                )
            else:
                httpheaders.append(
                    ('Content-Type', 'application/pdf'),
                )

            return request.make_response(docxtpl_files, headers=httpheaders)

        except Exception as e:
            se = serialize_exception(e)
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

