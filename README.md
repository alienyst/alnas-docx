# Docx Report Generator

The Docx Report Generator is a module that helps you create reports using only a .docx template and Jinja syntax.

This module inspired from [Report Xlsx](https://apps.odoo.com/apps/modules/16.0/report_xlsx).

## Prerequisites

Before installing this module, make sure to install the following libraries:

- `pip install docxtpl htmldocx beautifulsoup4 num2words Babel`

## Usage

For usage instructions, you can refer to the following video: [Link](https://www.youtube.com/watch?v=dZvak8yiD5Q)  
![Video Preview](assets/preview.gif)

Example template use for sale order: [Link](https://github.com/alienyst/alnas-docx/raw/16.0/alnas_docx/static/description/example/example.docx)

Documentation on writing syntax in the document: [Link](https://docxtpl.readthedocs.io/en/stable/)

## Field Naming Convention

To call and write the field name, use the following format: `{{docs.field_name}}`, starting with the word "docs".

### Useful Functions

- `{{spelled_out(docs.numeric_field)}}`: Spell out numbers
- `{{formatdate(docs.date_field)}}`: Format dates
- `{{parsehtml(docs.html_field)}}`: Parse HTML fields
- `{{convert_currency(docs.monetary_field, docs.currency_id)}}`: Show monetary field
- `{{render_image(docs.image_field)}} or {{render_image(docs.image_field, width=10, height=10)}}`: Render Image in Mm.

Note: The functions will be updated as needed.

lang default is lang='id_ID' change if need, example = `{{spelled_out(docs.numeric_field, lang='en_US')}}`

## Feedback

We welcome any feedback and suggestions, especially for improving this module. Thank you!
