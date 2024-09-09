{
    'name': "Docx Report Generator",

    'summary': """
        Generate your Report with DOCX template""",

    'description': """
        Simple module to generate report with DOCX template
    """,

    'author': "Ali Ns",

    'website': "https://github.com/alienyst",
    
    'images': ["static/description/banner.png"],

    'category': 'Technical',
    
    'version': '17.0.1.0.0',
        
    'application': True,
    
    'installable': True,

    'depends': ['base', 'mail'],

    'data': [
        'security/ir.model.access.csv',
        
        'data/ir_config_data.xml',
        
        'views/docx_report_config_view.xml',
        
        'views/ir_action_report_view.xml',
    ],
    
    'assets': {
        'web.assets_backend': [
            'alnas_docx/static/src/js/report/action_manager_report.esm.js'
        ]
    },

    'license': 'LGPL-3',
    
    'external_dependencies': {
        'python': ['docxtpl', 'docxcompose', 'htmldocx', 'beautifulsoup4'],
    }
    
}
