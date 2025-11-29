from typing import List, Dict, Any, Tuple
import io
import re

from openpyxl import load_workbook

# --- Regex auxiliares ---
CURRENCY_RX = re.compile(r"R\$\s*")
SPACES_RX = re.compile(r"\s+")
NON_DIGIT_RX = re.compile(r"[^\d]")  # usado no novo parse_int_flex

# Cabeçalhos canônicos e aliases
CANON_KEYS = ["ITEM", "DESCRIÇÃO", "UNID.", "QUANT.", "VALOR UNIT.", "VALOR TOTAL"]
HEADER_ALIASES = {
    "ITEM": {"ITEM", "ITEN", "ITENS"},
    "DESCRIÇÃO": {
        "DESCRIÇÃO",
        "DESCRICAO",
        "DESCRIÇÃO/ESPECIFICAÇÃO",
        "DESCRIÇÃO DO ITEM",
    },
    "UNID.": {"UNID.", "UNIDADE", "UND", "UNID"},
    "QUANT.": {"QUANT.", "QTD", "QUANTIDADE"},
    "VALOR UNIT.": {
        "VALOR UNIT.",
        "VALOR UNITARIO",
        "VALOR UNITÁRIO",
        "VLR UNIT.",
        "PREÇO UNIT.",
        "PREÇO UNIT",
        "V. UNIT.",   # usado na planilha
        "V UNIT.",
    },
    "VALOR TOTAL": {
        "VALOR TOTAL",
        "VLR TOTAL",
        "TOTAL",
        "PREÇO TOTAL",
        "V. TOTAL",   # usado na planilha
        "V TOTAL",
    },
}

INCH_VARIANTS = ['"', '\\"', "”", "“", "″"]
INCH_GLYPH = "″"

VALOR_TOTAL_SUFFIX_RX = re.compile(
    r"(VALOR\s+TOTAL\s*R\$\s*[\d\.\,]+)$", flags=re.IGNORECASE
)
VALOR_PRECO_TRAILING_RX = re.compile(r"(R\$\s*[\d\.\,]+)$", flags=re.IGNORECASE)


# -------------------- Utilitários --------------------
def normalize_inches(text: str) -> str:
    if not text:
        return text
    for v in INCH_VARIANTS:
        text = text.replace(v, '"')
    text = re.sub(r"\s+\"", "\"", text)
    text = re.sub(
        r"(\d+(?:[\.,]\d+)?(?:/\d+(?:[\.,]\d+)?)?)\"",
        r"\1" + INCH_GLYPH,
        text,
    )
    return text.replace('"', INCH_GLYPH)


def strip_valor_total_from_desc(text: str) -> str:
    if not text:
        return text
    text = VALOR_TOTAL_SUFFIX_RX.sub("", text).strip()
    text = VALOR_PRECO_TRAILING_RX.sub("", text).strip()
    return text


def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = s.replace("\xa0", " ")
    s = SPACES_RX.sub(" ", s).strip()
    return normalize_inches(s)


def parse_money(s: str):
    if not s:
        return None
    s = normalize_text(s)
    if s.upper() == "R$":
        return None
    s = CURRENCY_RX.sub("", s)
    s = s.replace(".", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


# 🔥 NOVA FUNÇÃO FLEXÍVEL PARA NÚMEROS (Item e Quant)
def parse_int_flex(s: str):
    if not s:
        return None

    s = normalize_text(s)

    # remove tudo que não for dígito (exceto dígitos)
    only_digits = NON_DIGIT_RX.sub("", s)

    if not only_digits:
        return None

    try:
        return int(only_digits)
    except ValueError:
        return None


def looks_like_total_row(row: List[str]) -> bool:
    joined = " ".join(normalize_text(c).upper() for c in row)
    return ("VALOR TOTAL" in joined) and (parse_money(joined) is not None)


# -------------------- Mapeamento --------------------
def _header_match_score(cell: str, target_key: str) -> int:
    cell_u = normalize_text(cell).upper()
    return int(any(alias in cell_u for alias in HEADER_ALIASES[target_key]))


def _find_col_index(header_row: List[str], key: str):
    for j, cell in enumerate(header_row):
        cell_u = normalize_text(cell).upper()
        for alias in HEADER_ALIASES[key]:
            if alias in cell_u:
                return j
    return None


def detect_header_and_map(rows: List[List[str]]) -> Tuple[int, Dict[str, int]]:
    best_idx = -1
    best_score = -1

    for i, row in enumerate(rows):
        score = sum(
            any(_header_match_score(c, key) for c in row) for key in CANON_KEYS
        )
        if score > best_score:
            best_score = score
            best_idx = i

    if best_idx < 0 or best_score < 4:
        raise ValueError("Cabeçalho da tabela não foi encontrado com confiança suficiente.")

    header_row = rows[best_idx]

    col_map = {}
    for key in CANON_KEYS:
        idx = _find_col_index(header_row, key)
        if idx is not None:
            col_map[key] = idx

    missing = [k for k in CANON_KEYS if k not in col_map]
    if missing:
        raise ValueError(f"Não foi possível mapear todas as colunas: faltam {missing}")

    return best_idx, col_map


# -------------------- Parsing --------------------
def parse_row_with_map(row: List[str], col_map: Dict[str, Any]) -> Dict[str, Any]:

    def get(key):
        idx = col_map[key]
        if idx < len(row):
            return normalize_text(row[idx])
        return ""

    item = parse_int_flex(get("ITEM"))
    descricao = strip_valor_total_from_desc(get("DESCRIÇÃO"))
    unid = get("UNID.").upper().rstrip(".,")
    quant = parse_int_flex(get("QUANT."))
    valor_unit = parse_money(get("VALOR UNIT."))
    valor_total = parse_money(get("VALOR TOTAL"))

    return {
        "item": item,
        "descricao": descricao,
        "unid": unid,
        "quant": quant,
        "valor_unit": valor_unit,
        "valor_total": valor_total,
    }


def validate_and_fix(rec: Dict[str, Any]):
    vu = rec.get("valor_unit")
    qt = rec.get("quant")

    if vu is not None and qt is not None:
        calc = round(vu * qt, 2)
        rec["valor_total"] = calc

    return rec


def build_issues(rec: Dict[str, Any]):
    missing = [
        k
        for k in ["item", "descricao", "unid", "quant", "valor_unit", "valor_total"]
        if rec.get(k) in ("", None)
    ]
    if missing:
        return {
            "item": rec.get("item"),
            "missing": missing,
        }
    return None


# -------------------- Extração principal --------------------
def _extract_from_all_rows(all_rows: List[List[str]], empty_msg: str):
    if not all_rows:
        return {"rows": [], "issues": [{"error": empty_msg}]}

    header_idx, col_map = detect_header_and_map(all_rows)

    data_rows = []
    for row in all_rows[header_idx + 1:]:
        if not row or not any(normalize_text(c) for c in row):
            continue
        if looks_like_total_row(row):
            break
        data_rows.append(row)

    records = []
    issues = []

    for row in data_rows:
        safe_row = list(row)
        rec = parse_row_with_map(safe_row, col_map)
        rec = validate_and_fix(rec)

        issue = build_issues(rec)
        if issue:
            issues.append(issue)

        records.append(rec)

    records.sort(key=lambda r: (r["item"] if r["item"] is not None else 10**9))

    return {"rows": records, "issues": issues}


def extract_table_from_xlsx(file_bytes: bytes):
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        return {
            "rows": [],
            "issues": [{"error": "Não foi possível ler o arquivo XLSX."}],
        }

    sheet = wb.active

    all_rows = []
    for row in sheet.iter_rows(values_only=True):
        raw = ["" if c is None else str(c) for c in row]
        norm = [normalize_text(c) for c in raw]
        if any(norm):
            all_rows.append(norm)

    return _extract_from_all_rows(all_rows, "Nenhuma tabela encontrada na planilha XLSX.")
