# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build

This is a .NET Framework 4.8 solution (Windows-only, x64). Build with MSBuild or Visual Studio:

```bash
msbuild SPS.TaxDeterminationImporter.sln /p:Configuration=Debug /p:Platform="Any CPU"
msbuild SPS.TaxDeterminationImporter.sln /p:Configuration=Release /p:Platform="Any CPU"
```

There are no tests in this project. Validation is done manually inside a running SAP Business One environment.

## External dependencies not in NuGet

The project references two sets of DLLs via relative HintPaths that must exist on the developer machine:

| Reference | Expected path from solution root |
|---|---|
| `Interop.SAPbobsCOM.dll` | `../../../../Alfa/DLL/` |
| `Interop.SAPbouiCOM.dll` | `../../../../Alfa/DLL/` |
| `SBO.Hub.dll` | `../../../../Util/DLL/` |

These are SAP Business One SDK COM interop assemblies plus a proprietary wrapper (`SBO.Hub`). The solution will not compile without them.

## Architecture

This is a **SAP Business One Add-on** — a Windows desktop app that attaches to a running SAP B1 client via COM and extends the Tax Code Determination form (SAP internal form `80401`).

**Solution projects:**

- `SPS.TaxDeterminationImporter` — WinExe entry point. Connects to SAP, calls `InitializeBLL.Initialize()`, starts the SAP event listener thread, then calls `Application.Run()`.
- `SPS.TaxDeterminationImporter.Core` — Class library containing all business logic, data access, forms, and models.

**Layer structure inside Core:**

```
BLL/   — Business logic. TaxDeterminationBLL is the core class.
DAO/   — Data access. Scripts.cs selects the right ResourceManager (SQL Server vs HANA).
Forms/ — SAP UI event handlers. Not WinForms Forms — they inherit SBO.Hub base classes.
Model/ — Plain C# model classes with [SBO.Hub.Attributes] for ORM mapping.
Enum/  — TaxKeyFieldTypeEnum maps SAP's integer key field types to named values.
Views/ — .srf XML files define custom SAP dialog layouts (FrmImportLog, FrmRemoveTax).
```

## Key patterns

### SAP UI event handling

Forms in `Forms/` are not standard WinForms. They extend `SystemForm` or `BaseForm` from `SBO.Hub` and override `ItemEvent()`. Event registration is configured in `EventFilterBLL.SetDefaultEvents()` using `EventFilterHelper`. Only registered form/event combinations fire. The `80402` form handlers are commented out — re-enable them in `EventFilterBLL` if needed.

`f80401` injects three custom buttons into SAP's native form `80401` on `et_FORM_LOAD`. All logic executes only when `!ItemEventInfo.BeforeAction` (post-action events).

### Database-agnostic queries

All SQL is stored as string resources in two `.resx` files:
- `DAO/SQL.resx` — SQL Server queries
- `DAO/Hana.resx` — SAP HANA queries

`Scripts.SetResourceManager()` selects the correct one at startup based on `SBOApp.Company.DbServerType`. All query access goes through `Scripts.Resource.GetString("QueryKey")`.

When adding a new query: add the string to **both** `SQL.resx` and `Hana.resx` with the same key.

### COM object lifecycle

All SAP SDK objects (`TaxCodeDeterminationsTCDService`, `TaxCodeDeterminationTCD`, `ProgressBar`, etc.) must be explicitly released with `Marshal.ReleaseComObject()` in `finally` blocks. Failing to do so causes memory leaks in the SAP client process. See `TaxDeterminationBLL.ImportData()` for the established pattern.

### Import Excel format

The expected Excel column layout (row 1 = header, row 2+ = data):

| A | B | C | D | E | F | G | H | I | G+3 … |
|---|---|---|---|---|---|---|---|---|---|
| Valor1 | Valor2 | Valor3 | Valor4 | Efetivo De | Efetivo Ate | *Usage name* | IVA - *Usage* | DA - *Usage* | *(repeats per usage)* |

Columns 7+ repeat in groups of 3 for each tax usage. The usage name in the header cell must match an existing `OUSG` record (case-insensitive). The export (`ExportData`) writes files in this same format.

### Validation caching

`TaxDeterminationBLL` caches validation lists (BPs, items, NCM codes, etc.) as static fields. These are populated lazily per import run. `BusinesPartnerList` is always re-fetched per import; the others are null-checked and reused across calls within the same process session.

### Key field types

`TaxKeyFieldTypeEnum` mirrors SAP's internal numbering for key field types. Types `NcmCode`, `ItemGroup`, `CustomerGroup`, and `SupplierGroup` require description-to-ID conversion (`ConvertDescriptionToId`) before writing to the SAP service. Other types (BP, Item, State, Branch, UDF) are passed as-is.
