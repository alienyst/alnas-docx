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

    'depends': ['base', 'web', 'mail'],

    'data': [
        'security/ir.model.access.csv',
        
        'data/ir_config_data.xml',
        
        'views/docx_report_config_view.xml',
        
        'views/ir_action_report_view.xml',

        'views/webclient_templates.xml',
    ],

    'license': 'LGPL-3',
    
    'external_dependencies': {
        'python': ['docxtpl', 'htmldocx'],
    }
    
}
