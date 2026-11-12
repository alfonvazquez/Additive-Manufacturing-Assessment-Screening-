# Additive Manufacturing Assessment (Screening)

## Overview

Screening is a Python data-processing project that consolidates large datasets about parts and applies configurable business rules to help assess whether selected parts may be suitable for additive manufacturing (3D printing).

The tool brings together engineering, portfolio, procurement, inventory, forecast, and obsolescence information. It produces structured results for an initial review; it does not replace engineering validation, manufacturing analysis, or formal approval.

## What the tool does

- Combines records from multiple Excel workbooks and, for most sources, stores reusable copies in local SQLite databases.
- Cleans and standardizes fields so that records from different systems can be compared and joined.
- Builds separate screening datasets for PlaneA, PlaneB, and PlaneC.
- Adds design, material, dimension, ATA, pricing, SAP-history, and forecast information where available.
- Applies configurable rules to classify material, dimensions, functional class, and commercial data.
- Exports the resulting datasets to Excel for further review.

## Data sources

The script expects workbooks from the original operating environment. They are grouped by their role in the process:

| Group | Examples of information used |
| --- | --- |
| Common | Part portfolio, missing-parts records, descriptions to discard, material/document rules, ATA descriptions, procurement conditions, and forecast |
| PlaneA | Design data, SAP exports, PlaneA pricing, and location data used for ATA lookup |
| PlaneB and PlaneC | Design data, pricing, obsolescence records, and a PlaneB-to-ATA reference |

The source paths, workbook names, sheet names, and some header positions are configured near the beginning of `Screening.py`. The expected workbook structure must match the fields selected by the processing functions.

## Processing workflow

### 1. Load source data

At startup, the script asks whether to refresh the SQLite databases. With a refresh, it reads the configured workbooks, loads independent sources concurrently, and saves most of the resulting tables to local SQLite files. Without a refresh, it reuses those cached tables where configured; several reference or export files are still read directly from Excel.

This cache is intended to reduce repeated workbook-loading time. Refresh it after changing the source files or their structure.

### 2. Normalize and prepare records

Each dataset is reduced to the columns needed by the screening and mapped to consistent internal names. The preparation functions also clean blank values, normalize text and numeric fields, remove or consolidate duplicate records, and prepare values for matching.

Examples of derived data include:

- Converting design weights to a common unit and parsing dimension strings into length, width, diameter, and thickness.
- Aggregating SAP order records by part number, including order counts, quantities, average costs, and an AOG indicator.
- Aggregating purchasing data and calculating forecast quantities.
- Preparing the material/document rules and ATA references used later in the joins.

### 3. Build the PlaneA, PlaneB, and PlaneC populations

The script creates a separate part population for each plane, starting primarily from the portfolio and missing-parts data. PlaneB and PlaneC also incorporate eligible obsolescence records. A source-tracking field records which input dataset or datasets contributed each part.

Each plane applies its own part-number format checks and eligibility filters. Description-based discard rules are then applied, and ATA codes are extracted or looked up using the relevant reference data.

### 4. Enrich parts with engineering and commercial information

For each plane, the part population is joined with the relevant design and reference tables. The process identifies raw and normalized material information, classifies materials from description keywords, and attaches available dimensions and functional-class data.

The script then joins purchasing conditions, plane-specific pricing, SAP history, and forecast information by part number. Missing information is retained where possible so reviewers can distinguish an unavailable value from a negative assessment.

### 5. Evaluate candidate and study status

The screening rules evaluate several dimensions of the decision:

- **Material:** whether the material is recognized by the configured material roadmap.
- **Dimensions:** whether usable design dimensions exist and satisfy the configured, material-specific limits.
- **Functional class:** whether the part's functional classification satisfies the configured material requirement, where applicable.
- **Commercial information:** whether pricing and lead-time data are available and within the configured thresholds.

The `Candidate_X` field provides the overall screening outcome. Depending on the available data and rule results, it can indicate a positive preliminary signal, a rejection, a need for analysis, missing information, or a functional-class restriction. The `Study` field provides additional context, such as why further analysis or data is needed.

These are rule-based preliminary results, not a prediction that a part can be manufactured successfully. The rules, thresholds, source quality, and anonymized identifiers all affect the result.

### 6. Export results

The full workflow writes `Screening.xlsx` with three worksheets: `PlaneA`, `PlaneB`, and `PlaneC`. Each sheet contains the selected part information, source and ATA data, design/material fields, available commercial and historical data, and the candidate/study results.

## Optional P_Type_1 workflow

The first menu option creates `P_Type_1.xlsx` with separate PlaneB and PlaneC worksheets. It combines portfolio, missing-parts, and eligible obsolescence records, then filters them using the configured part-number patterns.

This feature is incomplete: it identifies a preliminary list based on naming conventions and does not resolve assembly-to-child-part relationships. Treat its output as an intermediate list, not as a completed assembly analysis. The script can offer to continue to the full Screening workflow afterward.

## Repository structure

```text
Screening/
├── Screening.py
├── aux_func/
│   ├── db_creator_reader_module.py
│   └── df_creator_module.py
├── inputs/       # External workbooks; not included
├── sqlite3/      # Local cache generated from source data
└── outputs/      # Workspace results folder; the script currently targets "output/"
```

### Main components

- **`Screening.py`** — Configures the source and cache paths, loads and prepares datasets, runs the plane-specific workflows, applies the screening rules, and exports results.
- **`aux_func/df_creator_module.py`** — Reads supported tabular files and standardizes column names.
- **`aux_func/db_creator_reader_module.py`** — Creates, reads, and inspects the SQLite tables used as local caches.

## Setup and execution

### Requirements

- Python 3
- pandas
- NumPy
- SciPy
- python-dateutil
- openpyxl
- `pyxlsb` and/or `xlrd` if the source workbooks use those formats

Install the listed packages with:

```bash
pip install pandas numpy scipy python-dateutil openpyxl pyxlsb xlrd
```

The repository does not currently pin dependency versions in a requirements file.

### Before running

1. Provide the required source workbooks in the expected locations.
2. Update the environment-specific paths and confirm the workbook sheet names and columns match the script.
3. Ensure the SQLite cache and output directories exist, and that the configured output path points to the intended folder.
4. Refresh the database after changing any source workbook.

The intended entry point is:

```bash
python Screening.py
```

At startup, choose whether to refresh the database, then select `1` for the optional P_Type_1 workflow or `2` for the full Screening workflow.

**Execution status:** the public, anonymized copy has not been validated as a standalone run. In addition to configuring the external files and paths, the current script may require code-level corrections before it runs successfully.

## Data privacy

For obvious confidentiality reasons, the data and identifiers in this public version have been anonymized. Production workbooks and databases are not included. Part-number patterns and rule values have also been generalized, so the public version may not reproduce the original screening results.

## Limitations

- The process depends on the quality, completeness, and consistency of the source workbooks.
- Input paths, workbook layouts, sheet names, and some rules are environment-specific.
- No production datasets or pinned dependency versions are provided.
- Screening outputs are preliminary decision-support results and must be reviewed by qualified engineering and manufacturing stakeholders before any production decision.
