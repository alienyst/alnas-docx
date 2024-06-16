# Docx Report Generator

The Docx Report Generator is a module that helps you create reports using only a .docx template and Jinja syntax.

This module is inspired from [Report Xlsx](https://apps.odoo.com/apps/modules/16.0/report_xlsx).

## Prerequisites

Before installing this module, make sure to install the following libraries:

- `pip install docxtpl docxcompose html2docx num2words Babel`

## Usage

For usage instructions, you can refer to the following video:

[![YouTube video player](https://img.youtube.com/vi/dZvak8yiD5Q/0.jpg)](https://www.youtube.com/embed/dZvak8yiD5Q?si=BArrT3n33ZDkkhKm)

Documentation on writing syntax in the document: [https://docxtpl.readthedocs.io/en/stable/](https://docxtpl.readthedocs.io/en/stable/)

## Field Naming Convention

To call and write the field name, use the following format: `{{docs.field_name}}`, starting with the word "docs".

### Useful Functions (Indonesian Language)

- `{{spelled_out(docs.numeric_field)}}`: Spell out numbers
- `{{formatdate(docs.date_field)}}`: Format dates
- `{{parsehtml(docs.html_field)}}`: Parse HTML fields

Note: The functions will be updated as needed.

## Feedback

We welcome any feedback and suggestions, especially for improving this module. Thank you!
