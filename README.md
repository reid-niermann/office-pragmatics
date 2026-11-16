# Office Pragmatics

Free, open-source productivity add-in for Microsoft Excel, PowerPoint and Word, built for finance professionals in the fields of private equity, investment banking, consulting, and FP&A.

MIT licensed.

## Download

Get `Office-Pragmatics-Setup.exe` from the [latest release](../../releases/latest) and run it. It
installs all three add-ins (administrator rights required).

**Requirements:** Windows with desktop Microsoft 365 or Office 2016 or later.

## What's included

### Excel

- **Auditing:** Trace precedents and dependents across ranges, sheets and workbooks, with arrows and a jump back to your starting cell.
- **Formula tools:** Flip sign, scale by 1,000, cycle `$` anchoring, wrap in `IFERROR` or custom text, smart fill, convert repeated formulas to `LET`, and replace lookups (`XLOOKUP`, `INDEX`/`MATCH`, `SUMIFS`, `OFFSET`) with direct links.
- **Paste tools:** Paste formulas exactly, smart paste, or paste only formats, fills, borders, column widths or row heights.
- **Model formatting:** Keyboard cycles for number, currency, percent, date, multiple and binary formats, plus borders, decimals and sizing. Blue/black and full modeling color conventions (auto-colored by cell type), accounting dots, sum lines, indents and footnote numbering.
- **Company formats:** Save a table and chart as your house style, then apply in one click.
- **Range tools:** Center across selection, group hidden rows/columns, reverse or delete blank rows/columns, hardcode formulas, calculate CAGR.
- **Charts:** Build waterfall, Marimekko, Gantt and bubble charts. Add CAGR arrows, totals, stat lines and metric bubbles. Sort stacked columns, maximize the plot area, center pies, set Y-axis minimum to 0, and duplicate charts across sheets.
- **Shapes:** Align, distribute, match size, extend, fit and snap shapes, with PowerPoint-style guides and slide-sized resizing.
- **Workbook tools:** Searchable sheet selector, deep search (text, formulas, values, formatting), visual formula-consistency check, and a cleaner for broken names, unused formats and bloated ranges.
- **File tools:** Version up (`_v2` to `_v3`), save as PDF, export a sheet, show in Explorer, copy file path, hide page breaks.
- **Shortcuts:** Keyboard shortcuts for most commands, customizable in Settings.

### PowerPoint

- **Alignment and sizing:** Align and distribute to the slide or last-selected shape, snap, match size or rotation, extend, fit, resize by ratio, swap positions, arrange into a grid, and measure with on-screen distances.
- **Logos, icons and maps:** Insert Logo (web search, auto-trim and size), Insert Icon (SVGs, recolorable), and Plot Map (US locations). Sort logos alphabetically, remove white backgrounds, and balance logo sizes.
- **Pitch tools:** Smart eye dropper, select similar shapes, redaction overlay, “To Be Updated” boxes, footnotes, flatten theme colors to RGB, zero text margins, toggle AutoFit, line spacing, find and replace (including alt text), and duplicate shapes across slides.
- **Charts:** Paste chart formatting, copy underlying data, snap labels or logos to axes, maximize the plot area, center pies.
- **Excel linking:** Live-linked charts (native, vector or picture), ranges (table or picture) and cell text, refreshed all at once or managed in the Link Manager.
- **Library:** Shared team folder of reusable shapes, slides and decks.
- **File tools:** Version up, save as PDF, export a selection, custom color palettes, go to slide, show in Explorer, copy file path.

### Word

- **Alignment and sizing:** Same shape tools as PowerPoint.
- **Document tools:** Redline compare (DOCX, DOC, PDF), form builder for fillable templates, smart eye dropper, number to words (“123” becomes “one hundred twenty-three (123)”), flatten theme colors, line spacing.
- **Excel linking:** Live-linked charts, tables and text, with the Link Manager and refresh.
- **Other:** Library, go to page, version up, save as PDF, export a selection, and file tools.

## API keys

Everything works offline except three PowerPoint features, which call third-party services directly
using API keys you add in **Settings > API Keys**:

| Feature | Service | What you need |
| --- | --- | --- |
| Insert Logo | [Brave Search API](https://brave.com/search/api/) | A Brave Search API key (image search is part of Brave's Search plan, which includes a free monthly credit) |
| Insert Icon | [Freepik API](https://www.freepik.com/api) | A Freepik API key |
| Plot Map | [Mapbox](https://account.mapbox.com/access-tokens/) | A Mapbox access token |

Keys are encrypted with Windows DPAPI for your Windows user account and stored under
`HKEY_CURRENT_USER\Software\Office Pragmatics\ApiKeys`. Requests go straight from your computer to
the provider, so usage counts against your own account. Insert Logo then downloads the chosen image
from the website that hosts it. Leave a key blank to turn that feature off.

## License

[MIT](LICENSE). Office Pragmatics is not affiliated with or endorsed by Microsoft. Logos inserted with
Insert Logo belong to their owners; check that your use of them is allowed.
