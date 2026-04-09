# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Multitool** is a C# WinForms application (.NET Framework 4.7.2) that acts as a KOMPAS-3D plugin. It automates export of CAD documents (DXF, PDF) and generation of Excel specification sheets from sheet metal 3D models. In Release mode it compiles as a COM-visible DLL registered as a KOMPAS-3D plugin.

## Build

```bash
# Debug build (produces Multitool.exe)
msbuild Multitool.sln /p:Configuration=Debug /p:Platform="Any CPU"

# Release build (produces Multitool.dll registered for COM interop)
msbuild Multitool.sln /p:Configuration=Release /p:Platform="Any CPU"
```

No automated tests exist. All functionality is validated manually with KOMPAS-3D v22+ running.

## Runtime Requirements

- **KOMPAS-3D v22** must be installed at `C:\Program Files\ASCON\KOMPAS-3D v22\`
- A valid `bug` license file must be present in the executable directory (encodes hardware fingerprint)
- Excel templates in `Шаблоны Excel/` are copied to output directory on build

## Architecture

### Core flow

`MainForm.cs` is the single large form (~970 lines) containing all business logic via four main button handlers:

| Button | Method | What it does |
|--------|--------|--------------|
| Create DXF | `СreateDXF_Click` | Gets active sheet metal body from KOMPAS via COM, creates unfolded 2D view, exports `.dxf`, copies to configured directory, opens in eDrawingHost viewer |
| Create PDF | `СreatePDF_Click` | Finds the `.cdw` drawing paired with the active model, invokes KOMPAS's `Pdf2d.dll` converter, displays in PdfViewerControl |
| Create Excel | `СreateExcel_Click` | Reads document properties (designation, name, material, mass, thickness), fills `PartTemplate.xlsx` or `AssemblyTemplate.xlsx`, saves to output directory |
| Fix Model | `button4_Click` | Validates spec section assignments, checks filename matches document properties, cross-checks global thickness variable vs. sheet metal body |

### COM interop layer

`Multitool.cs` is the COM entry point (`ProgId = "Multitool.Multitool"`). It handles:
- License verification (decodes `bug` file using Base64/hex against OS username + machine fingerprint)
- Registry self-registration under `HKLM\Software\ASCON\KOMPAS-3D\...`
- Launching `MainForm` via `GetInstance()` singleton

KOMPAS is accessed through these COM interfaces: `IApplication`, `IKompasDocument3D`, `IPart7`, `ISheetMetalContainer`, `ISheetMetalBody`.

### Custom controls

- `Controls/PdfViewerControl.cs` — wraps PdfiumViewer for in-app PDF display
- `Controls/eDrawingHost.cs` — `AxHost` wrapper embedding the eDrawings ActiveX control for DXF viewing
- `Controls/Settings.cs` — settings dialog, UI built programmatically (not in designer); saves paths to `Settings.xml` via `XmlSerializer`

### Excel templates

Located in `Шаблоны Excel/` (Russian: "Excel Templates"). Three templates exist: `PartTemplate.xlsx`, `AssemblyTemplate.xlsx`, `AssemblyTemplateWeld.xlsx`. Manipulated via **ClosedXML**.

## Key Dependencies

| Package | Purpose |
|---------|---------|
| ClosedXML.Signed 0.95.4 | Read/write Excel files |
| PdfiumViewer 2.0.0 | Render PDFs in the viewer panel |
| Irony / XLParser | Excel formula parsing |
| eDrawings Interop | Embedded DXF/drawing viewer (COM ActiveX) |
| Kompas6API5, KompasAPI7 | KOMPAS-3D COM API references |

## Important Notes

- Code comments, variable names, and UI strings are primarily in **Russian**.
- The `DELETE/` folder contains abandoned/experimental code — ignore it.
- `Settings.xml` (output paths for DXF/PDF) is runtime-generated, not committed.
- COM registration requires administrator privileges.
- The hardcoded PDF converter path is `C:\Program Files\ASCON\KOMPAS-3D v22\Bin\Pdf2d.dll`.
