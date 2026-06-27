# Draft-code

Development scratch repo — VBA modules and scripts in active iteration for Excel data management at True Footage. These are working drafts; stable versions are promoted to [VBA-working](https://github.com/kzelenakas/VBA-working).

## Table of Contents

- [Purpose](#purpose)
- [File Reference](#file-reference)
- [Workflow](#workflow)

---

## Purpose

This repo is a staging area for VBA modules and scripts being actively developed or tested. Files here may be incomplete, experimental, or mid-revision. Once a module is stable and tested, it moves to [kzelenakas/VBA-working](https://github.com/kzelenakas/VBA-working).

---

## File Reference

### Claude AI Integration

| File | Description |
|------|-------------|
| `ClaudeModule1-8-8a.bas` | Claude API integration module — sends text to Claude and processes the response within Excel |
| `ClaudeModuleKeywords8-8a.bas` | Claude-powered keyword extraction and classification module |

### Data Import & Cleaning

| File | Description |
|------|-------------|
| `ImportCSVandClean826.bas` / `_v3.bas` | Imports CSV files and applies cleaning rules (version 8.26) |
| `ImportCSVandClean826-1.txt` | Text version of the CSV import/clean script |
| `ImportCleanCSV-826.bas` | Import and clean CSV — variant build |
| `ImportandcleanCSV.bas` | Base import and clean module |
| `Importandclean2CSV.bas` | Handles two-CSV import and merge |
| `ImportandcleanCSV-DataStaff.bas` | Import and clean for Staff data specifically |
| `ImportandcleanCSV2data.bas` / `.txt` | Import clean CSV to data sheets |
| `DataMergeModule-Works.bas` | Merges data from multiple sources into a single sheet — confirmed working |

### Excel to Data Sheet Import

| File | Description |
|------|-------------|
| `ImportXLSToDataAMC-826.bas` | Imports XLS to AMC (Appraisal Management Company) data sheet |
| `ImportXLSToDataStaff-826.bas` | Imports XLS to Staff data sheet |
| `Sub ImportExcelToDataAMC.bas` | Sub-module for AMC Excel import |
| `Sub ImportExcelToStatsSheet.bas` / `Sub ImportExcelToStatsSheet2.bas` | Sub-modules for Stats sheet import |
| `Module4XslToDataAMC.bas` | Module 4 — XLS to AMC data sheet |
| `Module4XslToDataStaff.bas` | Module 4 — XLS to Staff data sheet |
| `Module4XslToStats.bas` | Module 4 — XLS to Stats sheet |
| `Module4csvToData.bas` | Module 4 — CSV to data sheet |

### Categorization

| File | Description |
|------|-------------|
| `ModuleCat-KWsheet.bas` | Categorizes data using keyword reference sheet |
| `ModuleCatArrayKW.bas` | Array-based keyword categorization — faster for large datasets |
| `ModuleCatCategorize.bas` | Core categorization logic module |
| `ModuleWeightedCat.bas` | Weighted categorization — assigns scores by keyword weight |
| `ModuleWeightedCat-need to mod for sheet name.bas` | Weighted categorization — needs sheet name update before use |

### Text Processing

| File | Description |
|------|-------------|
| `Split numberedANDlettered.v3.bas` | Splits entries that have both numbers and letters as prefixes — version 3 |
| `XLStoStaff-per.bas` | Staff import with per-record processing |

### Misc

| File | Description |
|------|-------------|
| `Sub ImportCSVToDataAMCSheet().txt` | Script snippet for CSV to AMC import |
| `Sub ImportCSVToDataAMCSheet.bas` | CSV to AMC data sheet import module |
| `Sub SaveActiveSheetAsPDF.txt` | Save active sheet as PDF snippet |
| `Sub SaveAllSheetsAsPDF().txt` | Save all sheets as PDF snippet |

---

## Workflow

```
Draft-code (this repo)          →    VBA-working
(development & testing)              (stable, production-ready)
```

Files in this repo may have open issues, incomplete logic, or version variants being compared. Check file names for version indicators (`v2`, `v3`, `_Version2`, `-826`) — higher numbers are generally more recent.
