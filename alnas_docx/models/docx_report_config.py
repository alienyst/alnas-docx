from odoo import models, fields, api
from odoo.exceptions import UserError



class DocxReportConfig(models.Model):
    _name = 'docx.report.config'
    _description = 'DOCX Report Configuration'
    
    _inherit = ["mail.thread", "mail.activity.mixin"]

    name = fields.Char(string='Report Name', required=True, readonly=True, states={"draft": [("readonly", False)]}, help="Name of the report")
    model_id = fields.Many2one('ir.model', string='Model', required=True, ondelete='cascade', readonly=True, states={"draft": [("readonly", False)]}, help="Model to which this report will be attached")
    field_id = fields.Many2one('ir.model.fields', string='Field Name', required=True, ondelete='cascade', domain="[('model_id', '=', model_id),('ttype', '=', 'char')]", readonly=True, states={"draft": [("readonly", False)]}, help="Field to be used as the report name")
    report_docx_template = fields.Binary(string='Report DOCX Template', required=True, readonly=True, states={"draft": [("readonly", False)]}, help="DOCX template to be used for the report")
    report_docx_template_filename = fields.Char(string='Report DOCX Template Name', required=True, readonly=True, states={"draft": [("readonly", False)]})
    prefix = fields.Char(string='Prefix', readonly=True, states={"draft": [("readonly", False)]}, help="Prefix to be used in the report name")
    state = fields.Selection([('draft', 'Draft'),('published', 'Published')],
        string='State',
        default='draft',
        tracking=True,
        copy=False,
        readonly=True, states={"draft": [("readonly", False)]}
    )
    action_report_id = fields.Many2one('ir.actions.report', string='Related Report Action', readonly=True, copy=False)
    docx_merge_mode = fields.Selection(
        [
            ("composer", "Composer"),
            ("zip", "Zip")
        ],
        string="DOCX Merge Mode",
        default="composer",
        required=True,
        readonly=True, states={"draft": [("readonly", False)]},
        help="Mode to be used for merging the DOCX template with the data, if 'Composer' is selected, the report will be generated as a single DOCX file, if 'Zip' is selected, the report will be generated as a ZIP file containing multiple DOCX files."
    )
    
    print_report_name = fields.Char(string='Print Report Name', compute='_compute_print_report_name', help="Filename generated for the report")
            
    @api.depends('model_id', 'field_id', 'prefix')
    def _compute_print_report_name(self):
        for rec in self:
            if rec.prefix:
                rec.print_report_name = f"'{rec.prefix} %s' % object.{rec.field_id.name} if object.{rec.field_id.name} else ''"
            else:
                rec.print_report_name = f"'{rec.model_id.name} %s' % object.{rec.field_id.name} if object.{rec.field_id.name} else ''"
    
    @api.constrains('report_docx_template_filename')
    def _check_report_docx_template_filename(self):
        for rec in self:
            if not rec.report_docx_template_filename.endswith('.docx'):
                raise UserError('Please upload a DOCX template.')
    
    def action_publish(self):
        for record in self:
            if record.state == 'draft':
                if not record.action_report_id:                
                    action_report = self.env['ir.actions.report'].sudo().create({
                        'name': record.name,
                        'model': record.model_id.model,
                        'report_type': 'docx',
                        'report_docx_template': record.report_docx_template,
                        'report_docx_template_name': record.report_docx_template_filename,
                        'report_name': record.report_docx_template_filename,
                        'docx_merge_mode': record.docx_merge_mode,
                        'print_report_name': record.print_report_name,
                    })
                else:
                    action_report = record.action_report_id
                    action_report.sudo().write({
                        'name': record.name,
                        'model': record.model_id.model,
                        'report_type': 'docx',
                        'report_docx_template': record.report_docx_template,
                        'report_docx_template_name': record.report_docx_template_filename,
                        'report_name': record.report_docx_template_filename,
                        'docx_merge_mode': record.docx_merge_mode,
                        'print_report_name': record.print_report_name,
                    })
                
                action_report.create_action()
                record.action_report_id = action_report
                record.state = 'published'
                
                return self._refresh_page()
                
            else:
                raise UserError('Report already published')
            
    def action_unpublish(self):
        for record in self:
            if record.state == 'published':
                record.action_report_id.unlink_action()
                record.state = 'draft'
                
                return self._refresh_page()
                
            else:
                raise UserError('Report already unpublished')
            
    @api.ondelete(at_uninstall=False)
    def _unlink_docx_report(self):
        for rec in self:
            if rec.state == 'published':
                rec.action_unpublish()
            if rec.action_report_id:
                rec.action_report_id.unlink()
                
    def _refresh_page(self):
        return {
            'type': 'ir.actions.client',
            'tag': 'reload',
        }