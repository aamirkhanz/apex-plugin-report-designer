# Report Designer for Oracle APEX

Users design report templates in the page and generate them as PDF, Word and Excel files from your data.

[![Live demo](https://img.shields.io/badge/Live_demo-0f7b72?style=for-the-badge)](https://oracleapex.com/ords/r/aamirkhanz/plugins/report-designer) &nbsp; Sign in as `demo` with the password `Demo`.

![Report Designer](screenshots/preview.png)

| | |
|---|---|
| **Type** | Region plug-in |
| **Internal name** | `COM.TECHREVX.REPORT_DESIGNER` |
| **Version** | 1.0 |
| **Requires** | Oracle APEX 26.2 or later |
| **License** | MIT |

## What it does

Report Designer lets users lay out their own reports inside your application and download them as PDF, Word or Excel files: invoices, order confirmations, price lists, sales summaries. The data comes from the region's SQL query; the templates are saved in a table, so a report designed once can be downloaded again by anyone it is shared with.

- **Designer**: drag blocks onto a page that looks like the printed result: heading, text, fields, table, image, line, space and page break. Click a block to change it; undo, redo and keyboard shortcuts work as expected.
- **Fields and totals**: insert `{CUSTOMER}` into any text, or totals such as `{SUM:AMOUNT|currency}`, from a searchable list.
- **Tables**: pick columns, headings, formats and widths; totals (sum, average, count, minimum, maximum), groups with subtotals, striped rows, lines or a grid.
- **One section per record**: repeat the layout for each order, customer or employee, each on a new page, and in Excel on its own sheet.
- **Page setup**: paper size, orientation, margins, colors, logo, header and footer with page numbers, date format and currency symbol.
- **Three formats from one template**: PDF and Word keep the layout; Excel keeps real numbers and dates, with SUBTOTAL formulas and filters.
- **Report mode**: a download-only view for end users: pick a template, preview, download.

## At a glance

- **Source**: any region source (SQL query, table or REST source). Its rows are the report data; Page Items to Submit and Faceted Search filters apply.
- **Designer**: drag blocks onto the page (heading, text, fields, table, image, line, space, page break), click a block to edit it, insert fields such as {CUSTOMER} or totals such as {SUM:AMOUNT|currency}. Tables have column formats, totals, row groups with subtotals, striped rows and borders. The report can make one section per value of a column (for example one invoice per order), each on a new page and, in Excel, on its own sheet. Page size, orientation, margins, colors, logo, header and footer with page numbers, date format and currency symbol.
- Templates are saved in a table (default TX_REPORT_TEMPLATES, see the Reference page for the script), per Template Group and Owner; shared templates can be used by everyone and copied with Save As. Report mode shows the saved templates with a preview and download buttons only.
- Files are made in the browser with pdfmake 0.3.12, docx 9.12.0 and ExcelJS 4.4.0 (included, loaded on first download). Excel tables keep real numbers and dates, with SUBTOTAL formulas for totals.
- **JavaScript**: apex.region( "id" ).generate( "pdf" | "docx" | "xlsx" ), refresh(), openTemplate( id ), getTemplate(), setTemplate( json ), save().
- **Events**: Report Designer - Report Generated and Template Saved.

## Installation

1. Download `plugin/region_type_plugin_com_techrevx_report_designer.sql` from the [latest release](https://github.com/aamirkhanz/apex-plugin-report-designer/releases/latest) (or from this repository).
2. In App Builder, open your application and go to **Shared Components > Plug-ins**.
3. Choose **Import**, select the file and finish the wizard.
4. Create the template table (script on the Reference page), then a region of type **Report Designer [Plug-in]** with the query that holds your report data.

Everything the plug-in needs is inside the SQL file, including pdfmake and Roboto font and docx and ExcelJS; it loads nothing from other sites.

## Example

```sql
select o.order_no, o.order_date, c.name as customer, c.city,
l.product, l.quantity, l.unit_price, l.quantity * l.unit_price as amount
from orders o
join customers c on c.customer_id = o.customer_id
join order_lines l on l.order_id = o.order_id
order by o.order_no, l.line_no
```

Every column is a field in the designer. Page Items to Submit and Faceted Search filters narrow the rows, so the same template prints one order, one month or one region.

## Attributes

### Templates

| Attribute | Type | Default | Description |
|---|---|---|---|
| **Templates Table** (required) | Text | `TX_REPORT_TEMPLATES` | Table that stores the templates, with the columns TEMPLATE_ID (generated), TEMPLATE_GROUP, NAME, DESCRIPTION, TEMPLATE (CLOB, JSON), OWNER, IS_SHARED, CREATED_ON and UPDATED_ON. The script is on the plugin's Reference page. |
| **Template Group** (required) | Text | `REPORTS` | Templates are saved and listed per group, so each report region can have its own set, for example ORDERS or EMPLOYEES. Use the same group in a Designer region and a Report region to offer the designed templates for download elsewhere. |
| **Owner** | Text | `&APP_USER.` | Name saved with new templates. Users can change and delete only templates with their own Owner. Use an item substitution such as &P1_OWNER. to group users differently, for example per department. |
| **Allow Sharing** (required) | Select list | `N` | Yes: the Save dialog offers Share with everyone, and shared templates appear for every user of the Template Group (read only; Save As makes a copy). No: users only see their own templates and templates shared by an administrator in the table. Values: `Y` Yes, `N` No. |

### Report

| Attribute | Type | Default | Description |
|---|---|---|---|
| **Mode** (required) | Select list | `designer` | Designer: users create, change, save and download templates. Report: users pick a saved template, see a preview and download it; nothing can be changed. Values: `designer` Designer, `report` Report (Download Only). |
| **Height** (required) | Integer | `760` | Height of the region in pixels. The page preview scrolls inside it. |
| **Formats** (required) | Checkboxes | `pdf:docx:xlsx` | The download buttons. PDF and Word keep the page layout; Excel puts the report on a sheet with real numbers, dates and SUBTOTAL formulas. Values: `pdf` PDF, `docx` Word (DOCX), `xlsx` Excel (XLSX). |
| **Template Item** | Page item |  | Optional page item that holds the ID of the open template. It is set when the user picks or saves a template, and when the page loads with an ID in it, that template opens first. |
| **Maximum Rows** (required) | Integer | `5000` | Most rows read from the region source. The designer previews the first rows of each table; downloads contain every row up to this number. |

## Events

Available in Dynamic Actions, and as jQuery events on the region or item:

- **Report Designer - Report Generated** (`tx-report-generate`)
- **Report Designer - Template Saved** (`tx-report-save`)

## Template table

Run this once in your parsing schema. Several regions can share the table: each lists only its own Template Group. Users change and delete only templates whose OWNER matches the region's Owner attribute.

```sql
create table tx_report_templates (
template_id    number generated by default on null as identity primary key,
template_group varchar2(255 char) not null,
name           varchar2(255 char) not null,
description    varchar2(4000 char),
template       clob check ( template is json ),
owner          varchar2(255 char) not null,
is_shared      varchar2(1 char) default 'N' not null check ( is_shared in ( 'Y', 'N' ) ),
created_on     timestamp with local time zone default systimestamp not null,
updated_on     timestamp with local time zone default systimestamp not null );

create index tx_report_templates_i1 on tx_report_templates ( template_group, owner );
```

To give everyone a ready-made template, save it in the designer, then mark it shared: `update tx_report_templates set is_shared = 'Y' where template_id = 1`.

## Fields in texts

```text
{CUSTOMER}                 value of the column in the first row (of the section)
{ORDER_DATE|date}          with a format: text, number, integer, currency, percent, date, datetime
{SUM:AMOUNT|currency}      total of the rows: SUM, AVG, MIN, MAX, COUNT
{COUNT}                    number of rows
{TODAY} {NOW}              current date, date and time
{REPORT} {APP_USER}        template name, user name
{GROUP}                    value of the section column
{PAGE} {PAGES}             page number and page count (header and footer)
```

Unknown fields are marked red in the designer. Dates use the template's date format; currency uses its currency symbol.

## Template JSON

The TEMPLATE column holds JSON, so templates can be copied between groups, versioned or created with SQL:

```text
{
"version": 1,
"page":     { "size": "A4", "orientation": "portrait", "margin": 40 },
"style":    { "fontSize": 10, "accent": "#0f766e", "text": "#1f2937" },
"settings": { "groupBy": "ORDER_NO", "groupPageBreak": true, "sheetPerGroup": true,
"dateFormat": "DD MMM YYYY", "currency": "$", "fileName": "Invoice {ORDER_NO}" },
"header":   { "left": "Your company", "right": "{TODAY}", "logo": null, "line": true },
"footer":   { "left": "{REPORT}", "right": "Page {PAGE} of {PAGES}" },
"blocks": [
{ "type": "heading", "text": "Invoice {ORDER_NO}", "level": 1 },
{ "type": "fields", "columns": 2, "items": [ { "label": "Customer", "value": "{CUSTOMER}" } ] },
{ "type": "table", "columns": [
{ "field": "PRODUCT", "label": "Product", "width": 2.4 },
{ "field": "AMOUNT", "label": "Amount", "align": "right", "format": "currency", "total": "sum" } ],
"groupBy": "", "totals": true, "striped": true, "border": "lines" },
{ "type": "text", "text": "Thank you for your order.", "size": 9, "italic": true }
]
}
```

Other blocks: `image` (src as a data URL, width in points), `divider`, `spacer` (height) and `pagebreak`. Sizes are in points (72 per inch).

## Events and JavaScript

```text
Report Designer - Report Generated   after a download; this.data = { format, fileName, templateId, templateName }
Report Designer - Template Saved     after a save; this.data = { id, name }
```

```javascript
apex.region( "orders" ).generate( "pdf" );      // or "docx", "xlsx": downloads the open template
apex.region( "orders" ).openTemplate( 12 );     // open a template by ID
apex.region( "orders" ).refresh();              // read the data again (Page Items to Submit)
apex.region( "orders" ).getTemplate();          // the template JSON
apex.region( "orders" ).setTemplate( json );    // load a template from JSON (unsaved)
apex.region( "orders" ).save();
```

The files are made in the browser, so they hold the rows the region read (up to Maximum Rows). The libraries load on the first download and stay cached.

## Included libraries

| Library | Version | License |
|---|---|---|
| [pdfmake](https://pdfmake.github.io/docs/) | 0.3.12 | MIT |
| [Roboto font (in pdfmake's vfs_fonts.js)](https://fonts.google.com/specimen/Roboto) | pdfmake 0.3.12 | Apache-2.0 |
| [docx](https://docx.js.org) | 9.12.0 | MIT |
| [ExcelJS](https://github.com/exceljs/exceljs) | 4.4.0 | MIT |

See [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Repository contents

- `plugin/region_type_plugin_com_techrevx_report_designer.sql`: the plug-in, ready to import
- `src/`: the plug-in's own source (PL/SQL in `plugin.pls`, JavaScript and CSS), for reading and reviewing; the SQL file already contains it
- `apexplugin.json`: the plug-in's details for [apex.world](https://apex.world)

## License

MIT, see [LICENSE](LICENSE). The included libraries keep their own licenses.

## More plug-ins

Documentation, updates and more plug-ins for Oracle APEX by [TechRevX](https://www.techrevx.com).
