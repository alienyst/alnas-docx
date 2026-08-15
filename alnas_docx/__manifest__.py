{
    "name": "Docx Report Generator",
    "summary": """
        Generate your Report with DOCX template""",
    "description": """
        Simple module to generate report with DOCX template
    """,
    "author": "Ali Ns",
    "maintainers": ["salvorapi", "joachimnasution", "jankkm"],
    "website": "https://github.com/alienyst",
    "images": ["static/description/banner.png"],
    "category": "Technical",
    "version": "18.0.1.4.0",
    "application": True,
    "installable": True,
    "depends": ["mail"],
    "data": [
        "security/security.xml",
        "security/ir.model.access.csv",
        "data/ir_config_data.xml",
        "views/docx_report_config_view.xml",
        "views/ir_action_report_view.xml",
    ],
    "assets": {
        "web.assets_backend": [
            "alnas_docx/static/src/js/report/action_manager_report.esm.js",
            "alnas_docx/static/src/js/field/docx_preview.js",
            "alnas_docx/static/src/xml/docx_preview.xml",
            "alnas_docx/static/src/scss/docx_preview.scss",
        ],
        "alnas_docx.docx_preview": [
            "alnas_docx/static/vendor/jszip/jszip.min.js",
            "alnas_docx/static/vendor/docx-preview/docx-preview.min.js",
        ],
    },
    "license": "LGPL-3",
    "external_dependencies": {
        "python": ["docxtpl", "docxcompose", "htmldocx"],
    },
}
