# KPivot Excel Add-in

A high-performance Excel add-in built with **Excel-DNA** that provides advanced pattern-based filtering and layout management for Data Model pivot tables.

**Current version: 2.3.0** — see [What's New](#whats-new-in-230).

## What is Excel-DNA?

**Excel-DNA** is a free, open-source library that makes it easy to create high-performance Excel add-ins using .NET (C#, VB.NET, or F#). It provides a lightweight alternative to traditional COM add-ins and VSTO.

### Key Advantages:
- **No Installation Required**: Excel-DNA add-ins are deployed as `.xll` files that users can simply open in Excel—no registry modifications or admin rights needed
- **High Performance**: Native integration with Excel's C API for fast execution
- **Easy Distribution**: Single `.xll` file contains all dependencies (DNA packing)
- **Full .NET Support**: Access to the entire .NET ecosystem while maintaining Excel compatibility
- **Ribbon Customization**: Full support for Office Fluent UI (Ribbon) via XML
- **Simple Deployment**: Just copy the `.xll` file and open it in Excel

## Technology Stack

### Core Framework
- **Excel-DNA 1.9.0** - Add-in framework and Excel integration layer
- **.NET Framework 4.8** - Runtime environment (targets `net48`)
- **C# (latest)** - Primary development language

### Excel Integration
- **Microsoft.Office.Interop.Excel 15.0.4795.1001** - Excel Object Model access for pivot table manipulation
- **ExcelDna.AddIn** - Handles all Office interop, ribbon integration, and add-in lifecycle

### Data Access
- **ADO (late-bound)** - ActiveX Data Objects for querying Excel's internal Data Model via DAX queries
  - Used to retrieve unique values from Power Pivot tables
  - Enables pattern matching against Data Model data
  - Called through `dynamic`, so there is no `adodb.dll` dependency; the Excel interop types are embedded too

### UI Components
- **Windows Forms** - Dialog interfaces (e.g., Pattern Filter Dialog)
- **Microsoft.VisualBasic** - Provides `InputBox` for quick user prompts
- **Custom Ribbon XML** - Office Fluent UI customization embedded as resource

### Additional Libraries
- **Microsoft.CSharp** - Dynamic language runtime support for advanced C# features

## Features

- **Pattern-based filtering**: Filter pivot items by text patterns (e.g., `;11;`)
- **Segment matching**: `;27;` selects every value containing the segment `27`, such as seller tree ids like `;1;27;10040;`
- **Multiple patterns**: Combine patterns with commas (e.g., `;11;,;2883;`)
- **Export filter values**: Export all, selected or unselected filter/slicer values to cells for review
- **Include/Exclude modes**: Either show matching items or hide them
- **Data Model support**: Full support for Excel's internal Data Model (Power Pivot)
- **Automatic slicer detection**: Finds and uses existing slicers for multi-value filtering
- **Layout management**: Advanced pivot table layout and formatting services
- **Totals on Top**: Shows a live copy of the Grand Total row directly above the pivot table, so totals stay visible on long pivots
- **Custom Ribbon**: Dedicated KPivot tab with all functions

## Architecture

```
KPivot/
├── KPivotRibbon.cs            # Ribbon callbacks and command handlers
├── KPivot.csproj              # Project file (Excel-DNA configuration)
├── Services/
│   ├── LayoutService.cs       # Pivot table layout management
│   ├── PatternFilterService.cs # Pattern filter logic (DAX queries, slicer control, export)
│   ├── PivotService.cs        # Pivot table navigation and management
│   └── TopTotalsService.cs    # "Totals on Top" live Grand Total row
├── UI/
│   ├── PatternFilterDialog.cs # WinForms dialog for pattern filter
│   └── ExportFilterDialog.cs  # WinForms dialog for exporting filter values
├── Ribbon/
│   └── KPivotRibbon.xml       # Custom ribbon UI definition (embedded resource)
├── Helpers/                    # Utility classes
└── build.bat                   # Build script
```

## Requirements

- Windows 10/11
- .NET Framework 4.8
- Microsoft Excel 2016 or later
- Visual Studio 2022 (for development)

## Building

### Option 1: Using Visual Studio
1. Open `KPivot.csproj` in Visual Studio
2. Build → Build Solution (Ctrl+Shift+B)
3. Output: `bin\Release\net48\publish\KPivot-AddIn64-packed.xll` (or `KPivot-AddIn-packed.xll` for 32-bit)

### Option 2: Using Command Line
```batch
dotnet build KPivot.csproj -c Release
```

Or run `build.bat`

> **Close Excel first** (or unload KPivot) if it has the add-in loaded. Otherwise the build fails with
> *"The process cannot access the file ... KPivot-AddIn64.xll because it is being used by another process"*.

The version number is set in `KPivot.csproj` (`Version`, `AssemblyVersion`, `FileVersion`) and is shown by the **Info** button.

## Installation

### Simple Deployment (Recommended)

1. **Build the add-in** (see above)
2. **Copy the packed `.xll` file** from `bin\Release\net48\publish\` to a location on the user's machine
3. **Open in Excel**: Double-click the `.xll` file or use File → Open in Excel
4. The **KPivot** tab will appear in the ribbon

### Permanent Installation

To load the add-in automatically when Excel starts:

1. Open Excel
2. Go to **File → Options → Add-ins**
3. At the bottom, select **Excel Add-ins** from the "Manage" dropdown and click **Go**
4. Click **Browse** and select the `KPivot-AddIn64-packed.xll` file
5. Check the box next to **KPivot** and click **OK**

**No registry modifications or administrator rights required!**

## Usage

### Ribbon Commands

| Button | Function |
|--------|----------|
| **Pattern Filter** | Opens dialog for advanced pattern filtering |
| **Quick Include** | Prompt to quickly include items matching a pattern |
| **Quick Exclude** | Prompt to quickly exclude items matching a pattern |
| **Clear Filters** | Clears all Page/Filter field filters |
| **Export Filter / Selected / Unselected** | Export all, selected or unselected values of a filter or slicer to cells |
| **Go to Pivot** | Navigate to the pivot table on the sheet |
| **Go to Data** | Navigate to the data source |
| **Refresh** | Refresh the active pivot table |
| **Tabular / Outline / Compact** | Change the pivot table's row layout |
| **Subtotals / Grand Totals** | Toggle subtotals for all row fields / grand totals |
| **Totals on Top** | Toggle a live Grand Total row above the active pivot table |
| **Diagnostics** | Show diagnostic information |
| **Info** | Show the KPivot version |

### Pattern Syntax

- Single pattern: `;11;`
- Multiple patterns: `;11;,;2883;,;456;` — commas mean **OR** (a value is kept if it matches any pattern)
- Patterns are case-insensitive
- **Patterns wrapped in `;` match a segment** of `;`-separated values (contains):
  - `;27;` matches `;27;`, `;1;27;350;`, `1;27;5` — but **not** `;127;` or `;270;`
  - Multi-segment patterns work too: `;27;10040;` matches `;1;27;10040;5;`
- **Any other pattern must match the whole value**: `423` matches only `423`, not `1423` or `54236`

Example — select all seller tree ids containing `;27;` or `;10040;`: open **Pattern Filter**, pick the seller tree id field, enter `;27;, ;10040;` and choose **Include matches**.

## Technical Details

### How it works for Data Model Pivots

1. **Connects to Data Model** via ADODB using the workbook's Model connection
2. **Executes DAX query** to get all unique values: `EVALUATE VALUES('Table'[Column])`
3. **Matches values** against the specified pattern(s)
4. **Applies filter** using `SlicerCache.VisibleSlicerItemsList` with MDX member names

### Totals on Top

- Writes a "Grand Total" label and one formula per value column into the row directly above the pivot table (for example `=IF($A$25="Grand Total",$B$25,"")`).
- KPivot finds the Grand Total row from the pivot itself, so a row item named "Grand Total" or a matching label below the pivot is never picked up. When the pivot grows, shrinks or is filtered, KPivot rewrites the references. On a PC without KPivot, a resized pivot makes the row show blanks rather than wrong numbers.
- If the row above the pivot is not empty, KPivot offers to insert a worksheet row.
- The setting is stored in the workbook as a hidden name (`_KPivotTopTotals_...`). After every pivot update KPivot re-aligns the row when the Grand Total row moves or value columns are added or removed.
- If the pivot's grand totals are switched off, the top row shows blanks rather than stale numbers.

### Why Slicers are Required

For Data Model (OLAP) pivot tables, the standard VBA/COM APIs have limitations:
- `PivotField.PivotItems` collection is empty
- `PivotField.VisibleItemsList` silently fails for Page fields

The workaround is to use `SlicerCache.VisibleSlicerItemsList`, which works reliably but requires a slicer to exist for the field.

## Excel-DNA vs Traditional Approaches

| Feature | Excel-DNA | COM Add-in | VSTO |
|---------|-----------|------------|------|
| Deployment | `.xll` file | Registry + DLL | ClickOnce/Installer |
| Admin Rights | Not required | Required | Often required |
| Performance | Native C API | COM interop | COM interop |
| Distribution | Copy file | Run installer | Run installer |
| .NET Support | Full | Full | Full |
| Debugging | Full VS debugger | Full VS debugger | Full VS debugger |
| Ribbon | CustomUI XML | CustomUI XML | Designer + XML |

## Troubleshooting

### Add-in doesn't appear in Excel
1. Check if the `.xll` file is opened: Look for it in File → Options → Add-ins → Active Application Add-ins
2. Try opening the `.xll` file directly from Excel: File → Open
3. Check Excel's disabled add-ins: File → Options → Add-ins → Manage: Disabled Items

### "Could not connect to Data Model"
- Ensure the workbook has a Data Model (Power Pivot data)
- The pivot table must be based on the Data Model

### Filter doesn't apply
- KPivot creates a hidden slicer (`KPIVOT_...`) automatically when the field has none; if that fails, add a Slicer for the field manually
- The field must be in the pivot's **Filters** area (or select a cell of the field you want to filter)
- Check that the pattern is right: plain patterns need an exact match, use `;value;` to match a segment

## What's New in 2.3.0

- Upgraded Excel-DNA from 1.7.0 to 1.9.0. KPivot targets .NET Framework 4.8 and ships packed XLLs for 32-bit and 64-bit Excel.

## What's New in 2.2.4

**Fixes**
- **Totals on Top could show a value from the wrong row.** The top row searched the sheet for the first "Grand Total" label, so a row item with that name, or a matching label below the pivot after grand totals were switched off, could supply the value. KPivot now takes the Grand Total row from the pivot itself, and the top row stays blank when grand totals are off.
- **Export Selected / Export Unselected no longer accept a partly matched selection.** If any selected Data Model item can't be matched to a Data Model value, nothing is exported and the message lists the unmatched items. Before, the matched items were exported and the rest were silently treated as unselected.
- **Item keys containing `.&[` or `].[`** are now read correctly instead of being cut short.

## What's New in 2.2.3

**Fixes**
- **A failed pattern search no longer clears the existing filter** on a regular (non–Data Model) pivot. KPivot now checks for matches before changing the field; if nothing matches (or Exclude would hide everything), the current selection is left as it was and the message says so.
- **Export Selected / Export Unselected no longer guess when the selection can't be read.** Previously, if Excel failed to report a Data Model slicer's or filter field's selection, KPivot treated it as unfiltered: Export Selected exported every value and Export Unselected exported none, with no warning. KPivot now treats a field as unfiltered only when Excel confirms it; otherwise it exports nothing and explains why. The same applies when a selection is read but can't be matched to any Data Model value.
- **Export Unselected with no filter applied** now says that every value is selected, instead of a generic "No values found".

## What's New in 2.2.2

**New**
- Segment matching for delimited patterns: `;27;` now selects every value containing that segment (e.g. seller tree ids such as `;1;27;10040;`). Plain patterns keep exact matching.

**Fixes**
- **Subtotals** toggle always showed "off": Excel's subtotal state was read with the wrong type.
- **Subtotals / Grand Totals** toggles stayed visually flipped when the action failed (e.g. no pivot selected).
- **Go to Data** selected the wrong cells for whole-column sources (e.g. columns A:C were read as cells C1:C3).
- **Quick Include / Exclude** tried to filter the value field when a value cell was selected; it now falls back to the page field.
- Filters on regular (non–Data Model) slicers could fail halfway; items are now selected before others are deselected, with a clear error if nothing matches.
- The automatic hidden slicer failed when a second pivot was filtered on the same field (slicer name already taken).
- **Export**: a multi-cell destination wrote the header into every cell; exported values are now written as text so codes like `00123` keep their leading zeros and `1-2` doesn't become a date.
- Removed double releases of Excel COM objects during export, a possible cause of intermittent "COM object separated from its RCW" errors.

## Development

### Debugging
1. Set KPivot as the startup project
2. Press F5 to start debugging
3. Excel will launch with the add-in loaded
4. Set breakpoints in your code as needed

### Project Structure
- Excel-DNA automatically packs the add-in into a single `.xll` file
- The ribbon XML is embedded as a resource
- All dependencies are included in the packed `.xll`

## License

Copyright © Kirill Samoylov. All rights reserved.

This software is proprietary. No part of it may be copied, modified, distributed or used without the prior written permission of the copyright holder.
