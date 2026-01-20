# KPivot
KPivot is an Excel add-in that facilitates work with pivot tables in Excel.

v1.0 released in January 2026 based on Excel VBA.
v2.0 released in January based on .NET framework and Excel-DNA tool.

# KPivot Excel Add-in

A high-performance Excel add-in built with **Excel-DNA** that provides advanced pattern-based filtering and layout management for Data Model pivot tables.

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
- **Excel-DNA 1.7.0** - Add-in framework and Excel integration layer
- **.NET Framework 4.8** - Runtime environment (targets `net48`)
- **C# (latest)** - Primary development language

### Excel Integration
- **Microsoft.Office.Interop.Excel 15.0.4795.1001** - Excel Object Model access for pivot table manipulation
- **ExcelDna.AddIn** - Handles all Office interop, ribbon integration, and add-in lifecycle

### Data Access
- **ADODB 7.10.3077** - ActiveX Data Objects for querying Excel's internal Data Model via DAX queries
  - Used to retrieve unique values from Power Pivot tables
  - Enables pattern matching against Data Model data

### UI Components
- **Windows Forms** - Dialog interfaces (e.g., Pattern Filter Dialog)
- **Microsoft.VisualBasic** - Provides `InputBox` for quick user prompts
- **Custom Ribbon XML** - Office Fluent UI customization embedded as resource

### Additional Libraries
- **Microsoft.CSharp** - Dynamic language runtime support for advanced C# features

## Features

- **Pattern-based filtering**: Filter pivot items by text patterns (e.g., `;11;`)
- **Multiple patterns**: Combine patterns with commas (e.g., `;11;,;2883;`)
- **Include/Exclude modes**: Either show matching items or hide them
- **Data Model support**: Full support for Excel's internal Data Model (Power Pivot)
- **Automatic slicer detection**: Finds and uses existing slicers for multi-value filtering
- **Layout management**: Advanced pivot table layout and formatting services
- **Custom Ribbon**: Dedicated KPivot tab with all functions

## Architecture

```
KPivot/
├── KPivotRibbon.cs            # Ribbon callbacks and command handlers
├── KPivot.csproj              # Project file (Excel-DNA configuration)
├── Services/
│   ├── LayoutService.cs       # Pivot table layout management
│   ├── PatternFilterService.cs # Pattern filter logic (DAX queries, slicer control)
│   └── PivotService.cs        # Pivot table navigation and management
├── UI/
│   └── PatternFilterDialog.cs # WinForms dialog for pattern filter
├── Ribbon/
│   └── KPivotRibbon.xml       # Custom ribbon UI definition (embedded resource)
├── Helpers/                    # Utility classes
└── build.bat                   # Build script
```

## Requirements

- Windows 10/11
- .NET Framework 4.8
- Microsoft Excel 2016 or later
- Visual Studio 2022 or 2026 (for development)


## Installation

### Simple Deployment (Recommended)

1. **Build the add-in** (see above)
2. **Copy the `.xll` file** from `bin\Release\net48\` to a location on the user's machine
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
| **Go to Pivot** | Navigate to the pivot table on the sheet |
| **Go to Data** | Navigate to the data source |
| **Refresh** | Refresh the active pivot table |
| **Diagnostics** | Show diagnostic information |

### Pattern Syntax

- Single pattern: `;11;`
- Multiple patterns: `;11;,;2883;,;456;`
- Patterns are case-insensitive
- Uses substring matching (Contains)

## Technical Details

### How it works for Data Model Pivots

1. **Connects to Data Model** via ADODB using the workbook's Model connection
2. **Executes DAX query** to get all unique values: `EVALUATE VALUES('Table'[Column])`
3. **Matches values** against the specified pattern(s)
4. **Applies filter** using `SlicerCache.VisibleSlicerItemsList` with MDX member names

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
- Add a Slicer for the field you want to filter
- Check that the pivot table is based on the Data Model (not a regular range)

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

MIT License

