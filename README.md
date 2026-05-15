# Semi-Structured Report Pipeline

A Python ETL pipeline for extracting and standardising semi-structured Excel reporting data.

---

## Overview

This project automates the extraction, cleaning, and consolidation of historical reporting data stored across multiple Excel workbooks and folder structures.

Many business reporting systems rely on manually maintained Excel files that evolve over time and contain inconsistent layouts, subtotal rows, merged sections, and changing schemas. This pipeline is designed to standardise those semi-structured reports into clean analytical datasets.

The system scans workbook directories, identifies valid reporting files, extracts required metrics, removes noise and summary rows, and exports consolidated historical datasets ready for analytics and visualisation tools.

---

## Features

- Automated Excel workbook ingestion
- Semi-structured data extraction
- Dynamic worksheet detection
- Historical dataset consolidation
- Subtotal and summary row removal
- Section/category detection
- Multi-folder batch processing
- Standardised analytical outputs
- Extraction logging and validation

---

## Example Workflow

```text
Excel Workbooks
       ↓
Folder Scanner
       ↓
Workbook Detection
       ↓
Worksheet Extraction
       ↓
Data Cleaning
       ↓
Schema Standardisation
       ↓
Master Historical Dataset
