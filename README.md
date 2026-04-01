# Product Characteristics Extraction (NLP, no LLM APIs)

## Overview

This project extracts product parameters from unstructured procurement files (`.docx`, `.pdf`, `.doc`) and saves a single consolidated Excel file.

The solution is built with traditional NLP methods (no OpenAI/Claude APIs).

## Imports Explained

- `argparse`: reads command-line flags (`--dataset`, `--lang`, `--output-dir`)
- `re`: regex engine for extracting voltage/capacity/dimensions/weight
- `dataclasses.dataclass`: clean container (`MappingResult`) for mapped columns
- `datetime`: builds output filename with required timestamp format
- `pathlib.Path`: safe cross-platform path handling
- `typing`: type hints for maintainable code
- `pandas`: builds final table and writes `.xlsx`
- `pdfplumber`: reads tabular content from PDF files
- `python-docx` (`Document`): reads tables from `.docx` files
- `rapidfuzz.fuzz`: fuzzy string matching for unstable header names
- `TfidfVectorizer` + `cosine_similarity`: header similarity (NLP Task 1)

## What Was Implemented

### 1) Automatic Folder Processing

- Iterates through all files in selected dataset folder
- Supports `.docx` and `.pdf` parsing directly
- `.doc` is recognized as supported by assignment but may be skipped in this environment if parser/converter is unavailable

### 2) Product Filtering (Task Requirement)

- Keeps only Lamps/Lights and Batteries/Accumulators
- Filtering is keyword/rule-based for both English and Russian texts

### 3) Intelligent Column Mapping (NLP Task 1)

- Column positions can change between files
- Header mapping uses hybrid similarity:
  - TF-IDF (character n-grams) + cosine similarity
  - fuzzy string matching (RapidFuzz)
  - small rule bonus for direct substring matches
- Maps arbitrary headers to standardized output fields

### 4) Multi-Column Text Parsing (NLP Task 2)

- Extracts structured attributes from product description/specification text
- Implemented with regex + rule-based parsing:
  - Voltage
  - Capacity
  - Dimensions (L/W/H)
  - Diameter
  - Weight
- If a value is absent in source text, output stays empty (as required)

### 5) Data Cleaning and Standardization

- Trims whitespace and normalizes decimal separators
- Removes fully empty records after extraction
- Produces one clean consolidated output table

### 6) Output Files

- Main output (strict assignment naming):  
  `params-YYYY-MM-DD-HH-MM-SS.xlsx`
- Validation report:
  - `report-eng-YYYY-MM-DD-HH-MM-SS.txt`
  - `report-rus-YYYY-MM-DD-HH-MM-SS.txt`

## Important Functions and Why They Matter

- `map_headers(headers)`: core of NLP Task 1; maps random input headers to standard fields
- `best_synonym_score(header, synonyms)`: hybrid score from fuzzy + TF-IDF cosine + rule bonus
- `parse_docx(path)` / `parse_pdf(path)`: parse each supported format into shared intermediate rows
- `is_target_product(text)`: keeps only lamps and accumulators
- `parse_voltage`, `parse_capacity`, `parse_dimensions`, `parse_weight`, `parse_diameter`: NLP Task 2 extractors
- `process_dataset(dataset_dir)`: loops all files and creates one consolidated dataset
- `to_language_columns(df, lang)`: outputs required headers in EN or RU
- `save_output(df, output_dir)`: creates final `params-...xlsx`
- `save_validation_report(...)`: writes `report-*.txt` with quality and TZ checklist

## Output Columns

### English (`--lang en`)

- ID
- Product Name
- Quantity, pcs, Unit
- Voltage, V
- Capacity
- Length (Height), mm
- Width, mm
- Height, mm
- Diameter
- Weight g/kg

### Russian (`--lang ru`)

- ID
- Наименование товара
- Количество, шт, Ед. Измерения
- Напряжение, В
- Емкость
- Длина (высота), мм
- Ширина, мм
- Высота, мм
- Диаметр
- Вес г/кг

## How To Run

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Run on English dataset:

   ```bash
   python extract_products.py --dataset "eng dataset" --output-dir . --lang en --label eng
   ```

3. Run on Russian dataset:

   ```bash
   python extract_products.py --dataset "rus dataset" --output-dir . --lang ru --label rus
   ```

## How To Verify Quality

1. Open generated report file (`report-*.txt`)
2. Check `TZ checklist`:
   - folder scan done
   - required 10 columns present
   - target product filtering done
   - NLP column mapping done
   - multi-column parsing done
   - single consolidated Excel generated
3. Check `Quality report`:
   - fill rate (%) per output column
   - technical columns can be partially empty if source files do not contain them

## Defense Notes (Short)

- Explain that this is a traditional NLP pipeline, not LLM extraction
- Show robust mapping on different column names across files
- Show one final Excel generated from all files in selected folder
- Show validation report to justify extraction quality and completeness
