# Docx Report Generator

The Docx Report Generator is a module that helps you create reports using only a .docx template and Jinja syntax.

This module inspired from [Report Xlsx](https://apps.odoo.com/apps/modules/16.0/report_xlsx).

## Prerequisites

Before installing this module, make sure to install the following libraries:

- `pip install docxcompose docxtpl htmldocx`



### Security Note
For security reasons, creating or editing DOCX Report Configurations is restricted to the **Report Editor** group (System Administrators by default). Regular users can only print published reports.

## Usage

For usage instructions, you can refer to the following video: [Link](https://www.youtube.com/watch?v=dZvak8yiD5Q)  
![Video Preview](assets/preview.gif)

Example template use for sale order: [Link](https://github.com/alienyst/alnas-docx/raw/16.0/alnas_docx/static/description/example/example.docx)

Documentation on writing syntax in the document: [Link](https://docxtpl.readthedocs.io/en/stable/)

### Template Playground

Open a saved DOCX report configuration and select an **Example Record (BETA)** below **Autoescape**. The non-stored field renders the DOCX automatically in the form with [docx-preview](https://github.com/VolodymyrBaydalka/docxjs). Browser rendering may differ slightly from native Microsoft Word.

## Field Naming Convention

To call and write the field name, use the following format: `{{docs.field_name}}`, starting with the word "docs".

### Available Functions & Variables

Because this module inherits Odoo's native mail rendering context, you have full access to Odoo's built-in formatting functions along with powerful DOCX manipulation tools.

#### 1. Odoo Native Formatting (Inherited)
- `{{ format_amount(docs.amount_total, docs.currency_id) }}`: Format currency automatically based on the user's language and currency symbol.
- `{{ format_datetime(docs.datetime_field) }}`: Format datetime fields with the correct timezone of the current user.
- `{{ format_date(docs.date_field) }}`: Format a date field according to the user's language.
- `{{ format_time(docs.datetime_field) }}`: Format only the time from a datetime field.
- `{{ format_duration(docs.duration_float) }}`: Format a float duration into HH:MM (e.g., `1.5` becomes `01:30`).

#### 2. Environment & Python Globals (Inherited)
- `{{ user.name }}` / `{{ user.email }}`: Access the current user's profile who is printing the report.
- `{{ ctx }}`: Access the current environment context (e.g., `ctx.get('lang')`).
- `{{ is_html_empty(docs.html_field) }}`: Return `True` if the HTML field is completely empty (safely ignoring empty tags like `<p><br></p>`).
- `{{ datetime.datetime.now() }}`: Python's native datetime module.
- `{{ formatdate(docs.date_field + relativedelta(months=1)) }}`: Python's relativedelta for easy date math.
- `{{ slug(object) }}`: Generate a URL-friendly slug from a record.
- Standard Python functions: `len()`, `abs()`, `min()`, `max()`, `sum()`, `round()`, `hasattr()`, `quote()`, `urlencode()`.

#### 3. Text & Data Conversion (Custom)
- `{{ spelled_out(docs.numeric_field) }}`: Spell out numbers into words.
- `{{ html2plaintext(docs.html_field) }}` : Render HTML content as plain text (Strips HTML tags safely).
- `{{ r rich_text(docs.text_field) }}`: Show Rich Text directly in DOCX.
- `{{ format_selection(docs, 'state') }}`: Return the translated label of a selection field (e.g., 'Draft' instead of 'draft').
- `{{ render_qrcode('https://odoo.com', width=20, height=20) }}`: Generate a QR Code.
- `{{ render_barcode('12345678', barcode_type='Code128', width=40, height=10) }}`: Generate a Barcode using Odoo's native generator.

#### 4. Document & Image Manipulation (Custom)
- `{{ render_image(docs.image_field) }}` or `{{ render_image(docs.image_field, width=10, height=10) }}`: Render an image (size in Mm).
- `{{p html2docx(docs.html_field) }}`: Render HTML as a formatted subdocument.
- `{{p add_subdoc(docs.docx_binary_field) }}`: Embed another DOCX file as a subdocument.
- `{{ replace_image('file_name_in_word', docs.image_field) }}`: Replace a dummy picture in the word document.
- `{{ replace_media('file_name_in_word', docs.image_field) }}`: Replace media (dummy file must exist in the template directory).
- `{{ replace_embedded('file_name_in_word', docs.binary_field) }}`: Replace embedded objects like an embedded DOCX.
- `{{ replace_zipname('file_path_in_word', docs.binary_field) }}`: Alternative for replacing embedded files using zipname replacement.
- `linked_attachments(docs)`: Returns binary attachments linked to the record. Use in a loop: `{% for att in linked_attachments(docs) %}{{ p add_subdoc(att.datas) }}{% endfor %}`.
- `{{ add_pdf(docs.pdf_attachment) }}`: **PDF mode only.** Queue an extra PDF to merge with the report output. Use `position='before'` or `position='after'`.

Note: The functions will be updated as needed.

lang default is lang='id_ID' change if need, example = `{{spelled_out(docs.numeric_field, lang='en_US')}}`

### Docx Mode

There are three modes for generating `.docx` reports:

1. **composer**: Generate a `.docx` file
2. **zip**: Generate a `.zip` containing the `.docx` file
3. **pdf**: Convert the `.docx` file to PDF using LibreOffice

#### PDF Mode

If you want to use the "pdf" option, ensure that LibreOffice is installed. Then set the LibreOffice path in **Settings** => **Technical** => **Parameters** => **System Parameters**, and search for the key `default_libreoffice_path`. Set the value according to your LibreOffice installation path:

- **Linux**: `/usr/bin/libreoffice`
- **Windows**: `C:\Program Files\LibreOffice\program\soffice.exe`

In PDF mode, the template also exposes `add_pdf` so you can merge additional PDF files with the PDF produced from your DOCX (for example cover pages or terms appended after the report). The main report PDF sits between any PDFs added with `position='before'` and those with `position='after'` (or the default). Only valid PDF data is accepted; add one PDF per call.

## Contributors

Thank you to the following contributors who have helped develop, fix bugs, and update features for this module:

- [@alienyst](https://github.com/alienyst)
- [@salvorapi](https://github.com/salvorapi)
- [@joachimnasution](https://github.com/joachimnasution)
- [@jankkm](https://github.com/jankkm)
- [@lpolitanski-tempoconsulting](https://github.com/lpolitanski-tempoconsulting)
- George

## Feedback

We welcome any feedback and suggestions, especially for improving this module. Thank you!
