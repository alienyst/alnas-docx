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
    
    'version': '1.0',
        
    'application': True,
    
    'installable': True,

    'depends': ['base', 'mail'],

    'data': [
        'security/ir.model.access.csv',
        
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
        'python': ['docxtpl', 'docxcompose', 'htmldocx', 'num2words', 'Babel', 'beautifulsoup4'],
    }
    
}
