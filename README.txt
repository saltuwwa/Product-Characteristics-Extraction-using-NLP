Product Characteristics Extraction (NLP, no LLM APIs)

Overview
This project extracts product parameters from unstructured procurement files
(.docx, .pdf, .doc) and saves a single consolidated Excel file.
The solution is built with traditional NLP methods (no OpenAI/Claude... APIs).

Imports explained
- `argparse`: reads command-line flags (`--dataset`, `--lang`, `--output-dir`).
- `re`: regex engine for extracting voltage/capacity/dimensions/weight.
- `dataclasses.dataclass`: clean container (`MappingResult`) for mapped columns.
- `datetime`: builds output file name with required timestamp format.
- `pathlib.Path`: safe cross-platform path handling.
- `typing`: type hints for maintainable code.
- `pandas`: builds the final table and writes `.xlsx`.
- `pdfplumber`: reads tabular content from PDF files.
- `python-docx` (`Document`): reads tables from `.docx` files.
- `rapidfuzz.fuzz`: fuzzy string matching for unstable header names.
- `TfidfVectorizer` + `cosine_similarity`: semantic-like header similarity (NLP Task 1).

What was implemented
1) Automatic folder processing
- Iterates through all files in selected dataset folder.
- Supports .docx and .pdf parsing directly.
- .doc is recognized as supported by assignment but may be skipped in this
  environment if parser/converter is unavailable.

2) Product filtering (Task requirement)
- Keeps only Lamps/Lights and Batteries/Accumulators.
- Filtering is keyword/rule based for both English and Russian texts.

3) Intelligent column mapping (NLP Task 1)
- Column positions can change between files.
- Header mapping uses hybrid similarity:
  - TF-IDF (character n-grams) + cosine similarity
  - fuzzy string matching (RapidFuzz)
  - small rule bonus for direct substring matches
- This maps arbitrary headers to standardized output fields.

4) Multi-column text parsing (NLP Task 2)
- Extracts structured attributes from product description/specification text.
- Implemented with regex + rule-based parsing:
  - Voltage
  - Capacity
  - Dimensions (L/W/H)
  - Diameter
  - Weight
- If a value is absent in source text, output stays empty (as required by assignment).

5) Data cleaning and standardization
- Trims whitespace and normalizes decimal separators.
- Removes fully empty records after extraction.
- Produces one clean consolidated output table.

6) Output files
- Main output (strict assignment naming):
  params-YYYY-MM-DD-HH-MM-SS.xlsx
- Validation report:
  report-eng-YYYY-MM-DD-HH-MM-SS.txt
  or
  report-rus-YYYY-MM-DD-HH-MM-SS.txt

Important functions and why they matter
- `map_headers(headers)`: core of NLP Task 1. It maps random input headers
  to standard fields (`id/name/spec/quantity/unit`) even when positions differ.
- `best_synonym_score(header, synonyms)`: calculates a hybrid score from
  fuzzy matching + TF-IDF cosine + rule bonus. This is the "intelligent mapping" engine.
- `parse_docx(path)` / `parse_pdf(path)`: parse each supported file format and
  convert rows into a shared intermediate structure.
- `is_target_product(text)`: filters only lamps and accumulators as required by TZ.
- `parse_voltage`, `parse_capacity`, `parse_dimensions`, `parse_weight`,
  `parse_diameter`: core of NLP Task 2 (multi-column extraction from one text field).
- `process_dataset(dataset_dir)`: loops all files in chosen folder and creates
  one consolidated dataset (required for defense).
- `to_language_columns(df, lang)`: outputs required headers in EN or RU form.
- `save_output(df, output_dir)`: creates final `params-YYYY-MM-DD-HH-MM-SS.xlsx`.
- `save_validation_report(...)`: writes `report-*.txt` with quality metrics and TZ checklist.

Output columns
For English run (--lang en):
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

For Russian run (--lang ru):
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

How to run
1) Install dependencies:
   pip install -r requirements.txt

2) Run on English dataset:
   python extract_products.py --dataset "eng dataset" --output-dir . --lang en --label eng

3) Run on Russian dataset:
   python extract_products.py --dataset "rus dataset" --output-dir . --lang ru --label rus

How to verify quality
1) Open generated report file (report-*.txt).
2) Check "TZ checklist" section:
   - folder scan done
   - required 10 columns present
   - target product filtering done
   - NLP column mapping done
   - multi-column parsing done
   - single consolidated Excel generated
3) Check "Quality report" section:
   - fill rate (%) per output column
   - technical columns can be partially empty if source files do not contain them

Defense notes (short)
- Explain that this is a traditional NLP pipeline, not LLM extraction.
- Show mapping robustness on different column names across files.
- Show final one-file Excel generated from all files in folder.
- Show validation report to justify extraction quality and completeness.
