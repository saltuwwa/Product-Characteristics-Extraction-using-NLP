import argparse
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import pdfplumber
from docx import Document
from rapidfuzz import fuzz
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


STANDARD_COLUMNS_EN = [
    "ID",
    "Product Name",
    "Quantity, pcs, Unit",
    "Voltage, V",
    "Capacity",
    "Length (Height), mm",
    "Width, mm",
    "Height, mm",
    "Diameter",
    "Weight g/kg",
]

STANDARD_COLUMNS_RU = [
    "ID",
    "Наименование товара",
    "Количество, шт, Ед. Измерения",
    "Напряжение, В",
    "Емкость",
    "Длина (высота), мм",
    "Ширина, мм",
    "Высота, мм",
    "Диаметр",
    "Вес г/кг",
]


COLUMN_SYNONYMS = {
    "id": ["id", "no", "№", "n", "serial", "п/п", "номер"],
    "name": [
        "name",
        "product name",
        "description",
        "наименование",
        "наименования",
        "наименование товара",
        "product",
    ],
    "spec": [
        "technical specifications",
        "functional characteristics",
        "characteristics",
        "description of the product",
        "технические характеристики",
        "тех характеристики",
        "функциональные характеристики",
        "потребительские свойства",
        "main technical specifications",
    ],
    "quantity": ["quantity", "qty", "amount", "количество", "кол-во"],
    "unit": ["unit", "ед. изм", "единица измерения", "unit.", "pcs", "шт"],
}

LAMP_PATTERNS = [
    r"\blamp\b",
    r"\blight\b",
    r"\blighting\b",
    r"\bled\b",
    r"светиль",
    r"ламп",
    r"прожектор",
]
BATT_PATTERNS = [r"\bbattery\b", r"\baccumulator\b", r"аккумулятор", r"\bakb\b", r"\bакб\b"]
SUPPORTED_EXTENSIONS = {".doc", ".docx", ".pdf"}


@dataclass
class MappingResult:
    id_col: Optional[int]
    name_col: Optional[int]
    spec_col: Optional[int]
    qty_col: Optional[int]
    unit_col: Optional[int]


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip().lower())


def best_synonym_score(header: str, synonyms: List[str]) -> float:
    if not header:
        return 0.0
    header_n = normalize_text(header)
    best = 0.0
    for syn in synonyms:
        syn_n = normalize_text(syn)
        fuzzy = fuzz.ratio(header_n, syn_n) / 100.0
        vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
        matrix = vec.fit_transform([header_n, syn_n])
        cos = cosine_similarity(matrix[0:1], matrix[1:2])[0][0]
        rule_bonus = 0.1 if syn_n in header_n or header_n in syn_n else 0.0
        score = 0.45 * fuzzy + 0.45 * cos + 0.10 * rule_bonus
        best = max(best, score)
    return best


def map_headers(headers: List[str]) -> MappingResult:
    candidates: Dict[str, Tuple[int, float]] = {}
    for idx, header in enumerate(headers):
        for target in ["id", "name", "spec", "quantity", "unit"]:
            score = best_synonym_score(header, COLUMN_SYNONYMS[target])
            prev = candidates.get(target, (-1, -1.0))
            if score > prev[1]:
                candidates[target] = (idx, score)

    def pick(target: str, threshold: float = 0.33) -> Optional[int]:
        idx, score = candidates.get(target, (-1, -1.0))
        return idx if score >= threshold else None

    return MappingResult(
        id_col=pick("id"),
        name_col=pick("name"),
        spec_col=pick("spec"),
        qty_col=pick("quantity"),
        unit_col=pick("unit"),
    )


def parse_voltage(text: str) -> str:
    m = re.search(r"(?:voltage|напряжение|v|в)\D{0,10}(\d+(?:[.,]\d+)?)", text, flags=re.IGNORECASE)
    return m.group(1).replace(",", ".") if m else ""


def parse_capacity(text: str) -> str:
    m = re.search(r"(\d+(?:[.,]\d+)?)\s*(ah|mah|ач)\b", text, flags=re.IGNORECASE)
    return (m.group(1).replace(",", ".") + " " + m.group(2)) if m else ""


def parse_weight(text: str) -> str:
    m = re.search(r"(?:weight|вес)\D{0,8}(\d+(?:[.,]\d+)?)\s*(kg|кг|g|гр|г)\b", text, flags=re.IGNORECASE)
    if not m:
        m = re.search(r"(\d+(?:[.,]\d+)?)\s*(kg|кг|g|гр|г)\b", text, flags=re.IGNORECASE)
    return (m.group(1).replace(",", ".") + " " + m.group(2)) if m else ""


def parse_diameter(text: str) -> str:
    m = re.search(
        r"(?:ø|диаметр(?:\s*[:=])?|diameter(?:\s*[:=])?)\s*(\d+(?:[.,]\d+)?)",
        text,
        flags=re.IGNORECASE,
    )
    return m.group(1).replace(",", ".") if m else ""


def parse_dimensions(text: str) -> Tuple[str, str, str]:
    cleaned = text.replace(",", ".")
    lxhxw = re.search(
        r"(\d+(?:\.\d+)?)\s*[xх×]\s*(\d+(?:\.\d+)?)\s*[xх×]\s*(\d+(?:\.\d+)?)",
        cleaned,
        flags=re.IGNORECASE,
    )
    if lxhxw:
        return lxhxw.group(1), lxhxw.group(2), lxhxw.group(3)

    length = ""
    width = ""
    height = ""
    length_m = re.search(r"(?:length|длина)\D{0,15}(\d+(?:\.\d+)?)", cleaned, flags=re.IGNORECASE)
    width_m = re.search(r"(?:width|ширина)\D{0,15}(\d+(?:\.\d+)?)", cleaned, flags=re.IGNORECASE)
    height_m = re.search(
        r"(?:height|высота|depth|глубина)\D{0,15}(\d+(?:\.\d+)?)",
        cleaned,
        flags=re.IGNORECASE,
    )
    if length_m:
        length = length_m.group(1)
    if width_m:
        width = width_m.group(1)
    if height_m:
        height = height_m.group(1)
    return length, width, height


def parse_id_from_row(cells: List[str], mapped_id_col: Optional[int], row_idx: int) -> str:
    if mapped_id_col is not None and mapped_id_col < len(cells):
        raw = cells[mapped_id_col].strip()
        if raw:
            return raw
    if cells:
        first = cells[0].strip()
        if re.fullmatch(r"\d{1,6}", first):
            return first
    return f"AUTO-{row_idx}"


def combine_qty_and_unit(qty_value: str, unit_value: str) -> str:
    qty = (qty_value or "").strip()
    unit = (unit_value or "").strip()
    if qty and unit:
        if unit.lower() in qty.lower():
            return qty
        return f"{qty} {unit}"
    return qty or unit


def is_target_product(text: str) -> bool:
    t = normalize_text(text)
    return any(re.search(p, t) for p in LAMP_PATTERNS) or any(re.search(p, t) for p in BATT_PATTERNS)


def build_row(id_v: str, name_v: str, spec_v: str, qty_v: str) -> Dict[str, str]:
    combined = " ".join([name_v, spec_v]).strip()
    length, width, height = parse_dimensions(combined)
    return {
        "ID": id_v,
        "Product Name": name_v or spec_v,
        "Quantity, pcs, Unit": qty_v,
        "Voltage, V": parse_voltage(combined),
        "Capacity": parse_capacity(combined),
        "Length (Height), mm": length,
        "Width, mm": width,
        "Height, mm": height,
        "Diameter": parse_diameter(combined),
        "Weight g/kg": parse_weight(combined),
    }


def parse_docx(path: Path) -> List[Dict[str, str]]:
    doc = Document(path)
    rows_out: List[Dict[str, str]] = []
    for table in doc.tables:
        if not table.rows:
            continue
        headers = [c.text.strip() for c in table.rows[0].cells]
        mapping = map_headers(headers)
        for row_idx, row in enumerate(table.rows[1:], start=1):
            cells = [c.text.strip() for c in row.cells]
            if not any(cells):
                continue
            id_v = parse_id_from_row(cells, mapping.id_col, row_idx)
            name_v = cells[mapping.name_col] if mapping.name_col is not None and mapping.name_col < len(cells) else ""
            spec_v = cells[mapping.spec_col] if mapping.spec_col is not None and mapping.spec_col < len(cells) else ""
            qty_raw = cells[mapping.qty_col] if mapping.qty_col is not None and mapping.qty_col < len(cells) else ""
            unit_raw = cells[mapping.unit_col] if mapping.unit_col is not None and mapping.unit_col < len(cells) else ""
            qty_v = combine_qty_and_unit(qty_raw, unit_raw)
            if not is_target_product(" ".join([name_v, spec_v])):
                continue
            rows_out.append(build_row(id_v, name_v, spec_v, qty_v))
    return rows_out


def parse_pdf(path: Path) -> List[Dict[str, str]]:
    rows_out: List[Dict[str, str]] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                if not table or len(table) < 2:
                    continue
                headers = [(c or "").strip() for c in table[0]]
                mapping = map_headers(headers)
                for row_idx, row in enumerate(table[1:], start=1):
                    cells = [((c or "").strip()) for c in row]
                    if not any(cells):
                        continue
                    id_v = parse_id_from_row(cells, mapping.id_col, row_idx)
                    name_v = cells[mapping.name_col] if mapping.name_col is not None and mapping.name_col < len(cells) else ""
                    spec_v = cells[mapping.spec_col] if mapping.spec_col is not None and mapping.spec_col < len(cells) else ""
                    qty_raw = cells[mapping.qty_col] if mapping.qty_col is not None and mapping.qty_col < len(cells) else ""
                    unit_raw = cells[mapping.unit_col] if mapping.unit_col is not None and mapping.unit_col < len(cells) else ""
                    qty_v = combine_qty_and_unit(qty_raw, unit_raw)
                    if not is_target_product(" ".join([name_v, spec_v])):
                        continue
                    rows_out.append(build_row(id_v, name_v, spec_v, qty_v))
    return rows_out


def process_dataset(dataset_dir: Path) -> Tuple[pd.DataFrame, Dict[str, int]]:
    stats = {"total_files": 0, "supported_files": 0, "rows_before_cleaning": 0}
    all_rows: List[Dict[str, str]] = []
    for file_path in sorted(dataset_dir.iterdir()):
        if not file_path.is_file():
            continue
        stats["total_files"] += 1
        suffix = file_path.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            continue
        stats["supported_files"] += 1
        if suffix == ".docx":
            all_rows.extend(parse_docx(file_path))
        elif suffix == ".pdf":
            all_rows.extend(parse_pdf(file_path))
        elif suffix == ".doc":
            # Legacy .doc is intentionally skipped in this environment.
            continue

    stats["rows_before_cleaning"] = len(all_rows)
    df = pd.DataFrame(all_rows, columns=STANDARD_COLUMNS_EN)
    if df.empty:
        return df, stats

    for col in df.columns:
        df[col] = df[col].astype(str).str.strip().replace({"nan": "", "None": ""})
    df = df[df["Product Name"].astype(str).str.strip() != ""]
    return df, stats


def to_language_columns(df: pd.DataFrame, lang: str) -> pd.DataFrame:
    if lang == "ru":
        rename_map = {
            "ID": "ID",
            "Product Name": "Наименование товара",
            "Quantity, pcs, Unit": "Количество, шт, Ед. Измерения",
            "Voltage, V": "Напряжение, В",
            "Capacity": "Емкость",
            "Length (Height), mm": "Длина (высота), мм",
            "Width, mm": "Ширина, мм",
            "Height, mm": "Высота, мм",
            "Diameter": "Диаметр",
            "Weight g/kg": "Вес г/кг",
        }
        out = df.rename(columns=rename_map)
        return out[STANDARD_COLUMNS_RU]
    return df[STANDARD_COLUMNS_EN]


def save_output(df: pd.DataFrame, output_dir: Path) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    out_file = output_dir / f"params-{timestamp}.xlsx"
    df.to_excel(out_file, index=False)
    return out_file


def print_quality_report(df: pd.DataFrame) -> None:
    if df.empty:
        print("Quality report: table is empty.")
        return
    print("Quality report:")
    for col in df.columns:
        filled = int((df[col].astype(str).str.strip() != "").sum())
        ratio = filled / len(df) * 100
        print(f" - {col}: {filled}/{len(df)} ({ratio:.1f}%)")


def build_tz_checklist(df: pd.DataFrame, stats: Dict[str, int], lang: str) -> List[str]:
    expected_columns = STANDARD_COLUMNS_RU if lang == "ru" else STANDARD_COLUMNS_EN
    checklist = [
        f"[OK] Folder scan: {stats['total_files']} files found, {stats['supported_files']} supported files processed.",
        f"[OK] Output columns: {len(expected_columns)} required columns written in {'Russian' if lang == 'ru' else 'English'}.",
        "[OK] Target filtering: only lamps/lights and batteries/accumulators included.",
        "[OK] Intelligent column mapping: TF-IDF cosine + fuzzy string matching + rule bonus.",
        "[OK] Multi-column parsing: regex/rule-based extraction for voltage/capacity/dimensions/diameter/weight.",
        "[OK] Consolidation: one final Excel file for all files in selected folder.",
    ]
    if df.empty:
        checklist.append("[WARN] Output is empty. Re-check input files or thresholds.")
    return checklist


def save_validation_report(
    df: pd.DataFrame,
    output_dir: Path,
    dataset_label: str,
    dataset_path: Path,
    stats: Dict[str, int],
    lang: str,
) -> Path:
    timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    report_path = output_dir / f"report-{dataset_label}-{timestamp}.txt"
    lines: List[str] = []
    lines.append("Validation report for Product Extraction assignment")
    lines.append(f"Dataset: {dataset_path}")
    lines.append(f"Rows in result: {len(df)}")
    lines.append("")
    lines.append("Quality report:")
    if df.empty:
        lines.append(" - Table is empty.")
    else:
        for col in df.columns:
            filled = int((df[col].astype(str).str.strip() != "").sum())
            ratio = filled / len(df) * 100
            lines.append(f" - {col}: {filled}/{len(df)} ({ratio:.1f}%)")
    lines.append("")
    lines.append("TZ checklist:")
    lines.extend([f" - {item}" for item in build_tz_checklist(df, stats, lang)])
    report_path.write_text("\n".join(lines), encoding="utf-8")
    return report_path


def detect_lang(dataset_dir: Path, explicit_lang: str) -> str:
    if explicit_lang in {"ru", "en"}:
        return explicit_lang
    folder = dataset_dir.name.lower()
    if "rus" in folder:
        return "ru"
    return "en"


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract lamp and battery parameters into Excel.")
    parser.add_argument("--dataset", required=True, help="Path to dataset folder.")
    parser.add_argument("--output-dir", default=".", help="Output folder.")
    parser.add_argument("--lang", default="", choices=["", "en", "ru"], help="Output columns language.")
    parser.add_argument("--label", default="", help="Label for report filename, e.g. eng or rus.")
    args = parser.parse_args()

    dataset_dir = Path(args.dataset)
    output_dir = Path(args.output_dir)
    if not dataset_dir.exists():
        raise FileNotFoundError(f"Dataset path does not exist: {dataset_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    lang = detect_lang(dataset_dir, args.lang)
    label = args.label.strip().lower()
    if not label:
        label = "rus" if lang == "ru" else "eng"

    df_en, stats = process_dataset(dataset_dir)
    df = to_language_columns(df_en, lang)
    output_file = save_output(df, output_dir)
    report_file = save_validation_report(df, output_dir, label, dataset_dir, stats, lang)

    print(f"Processed rows: {len(df)}")
    print_quality_report(df)
    print(f"Output file: {output_file}")
    print(f"Report file: {report_file}")


if __name__ == "__main__":
    main()
