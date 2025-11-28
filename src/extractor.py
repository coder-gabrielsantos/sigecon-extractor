from typing import List, Dict, Any, Tuple
import io
import re

from openpyxl import load_workbook

# --- Regex / sets auxiliares ---
CURRENCY_RX = re.compile(r"R\$\s*")
SPACES_RX = re.compile(r"\s+")
INT_RX = re.compile(r"^\d{1,4}$")

# UNIDADES conhecidas (apenas para warnings leves, não é obrigatório bater)
UNIT_SET = {
    "UN", "CJ", "UM", "M", "BAR", "TUB", "CX", "KIT", "PAR", "PÇ", "PC", "JG",
    "KG", "LT", "L", "G", "MM", "CM", "MT", "ROL", "SAC", "FD", "SC", "TB", "BL",
    "CONJ", "BLOCO"
}

# Cabeçalhos canônicos e aliases
CANON_KEYS = ["ITEM", "DESCRIÇÃO", "UNID.", "QUANT.", "VALOR UNIT.", "VALOR TOTAL"]
HEADER_ALIASES = {
    "ITEM": {"ITEM", "ITEN", "ITENS"},
    "DESCRIÇÃO": {
        "DESCRIÇÃO",
        "DESCRICAO",
        "DESCRIÇÃO/ESPECIFICAÇÃO",
        "DESCRIÇÃO DO ITEM",
        "DESCRIÇAO",
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
        "V. UNIT.",   # usado na planilha do contrato
        "V UNIT.",    # variação sem ponto
    },
    "VALOR TOTAL": {
        "VALOR TOTAL",
        "VLR TOTAL",
        "TOTAL",
        "PREÇO TOTAL",
        "V. TOTAL",   # usado na planilha do contrato
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
    # remove espaço antes de "
    text = re.sub(r"\s+\"", "\"", text)
    # transforma 100" em 100″
    text = re.sub(
        r"(\d+(?:[\.,]\d+)?(?:/\d+(?:[\.,]\d+)?)?)\"",
        r"\1" + INCH_GLYPH,
        text,
    )
    text = text.replace('"', INCH_GLYPH)
    return text


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
    s = normalize_inches(s)
    return s


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


def parse_int(s: str):
    s = normalize_text(s)
    try:
        return int(s) if INT_RX.match(s) else None
    except Exception:
        return None


def looks_like_total_row(row: List[str]) -> bool:
    joined = " ".join(normalize_text(c).upper() for c in row)
    return "VALOR TOTAL" in joined and parse_money(joined) is not None


# -------------------- Mapeamento de Cabeçalho --------------------
def _header_match_score(cell: str, target_key: str) -> int:
    cell_u = normalize_text(cell).upper()
    for alias in HEADER_ALIASES[target_key]:
        if alias in cell_u:
            return 1
    return 0


def _find_col_index(header_row: List[str], key: str):
    for j, cell in enumerate(header_row):
        cell_u = normalize_text(cell).upper()
        for alias in HEADER_ALIASES[key]:
            if alias in cell_u:
                return j
    return None


def detect_header_and_map(rows: List[List[str]]) -> Tuple[int, Dict[str, int]]:
    """
    Encontra a linha de cabeçalho e mapeia as colunas para:
    ITEM, DESCRIÇÃO, UNID., QUANT., VALOR UNIT., VALOR TOTAL.
    """
    best_idx = -1
    best_score = -1
    for i, row in enumerate(rows):
        score = 0
        for key in CANON_KEYS:
            if any(_header_match_score(c, key) for c in row):
                score += 1
        if score > best_score:
            best_score = score
            best_idx = i

    if best_idx < 0 or best_score < 4:
        raise ValueError(
            "Cabeçalho da tabela não foi encontrado com confiança suficiente."
        )

    header_row = rows[best_idx]
    col_map: Dict[str, int] = {}
    for key in CANON_KEYS:
        idx = _find_col_index(header_row, key)
        if idx is not None:
            col_map[key] = idx

    missing = [k for k in CANON_KEYS if k not in col_map]
    if missing:
        raise ValueError(f"Não foi possível mapear todas as colunas: faltam {missing}")

    return best_idx, col_map


# -------------------- Parsing de Linhas --------------------
def parse_row_with_map(row: List[str], col_map: Dict[str, Any]) -> Dict[str, Any]:
    get = (
        lambda key: normalize_text(row[col_map[key]])
        if col_map[key] < len(row)
        else ""
    )

    item = parse_int(get("ITEM"))
    descricao = strip_valor_total_from_desc(get("DESCRIÇÃO"))
    unid_raw = normalize_text(get("UNID.")).upper()
    unid = unid_raw.rstrip(".,")
    quant = parse_int(get("QUANT."))
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


def validate_and_fix(rec: Dict[str, Any]) -> Dict[str, Any]:
    vu = rec.get("valor_unit")
    vt = rec.get("valor_total")
    qt = rec.get("quant")

    # recalcula valor_total se possível
    if vu is not None and qt is not None:
        calc = round(vu * qt, 2)
        if vt is None or abs(calc - vt) <= 0.05:
            rec["valor_total"] = calc

    return rec


def build_issues(rec: Dict[str, Any]):
    missing = []
    for k in ["item", "descricao", "unid", "quant", "valor_unit", "valor_total"]:
        v = rec.get(k)
        if v in (None, ""):
            missing.append(k)

    if not missing:
        unid = (rec.get("unid") or "").upper()
        if unid and unid not in UNIT_SET:
            return {
                "item": rec.get("item"),
                "warning": f"Unidade '{unid}' fora da lista conhecida",
                "row": {
                    k: rec.get(k)
                    for k in ["descricao", "unid", "quant", "valor_unit", "valor_total"]
                },
            }
        return None

    return {
        "item": rec.get("item"),
        "missing": missing,
        "row": {k: rec.get(k) for k in ["descricao", "unid", "valor_unit", "valor_total"]},
    }


# -------------------- Função comum de extração --------------------
def _extract_from_all_rows(
    all_rows: List[List[str]], empty_msg: str
) -> Dict[str, Any]:
    """
    Recebe todas as linhas da planilha (já como texto), detecta cabeçalho e
    devolve:

    {
      "rows": [...],
      "issues": [...]
    }
    """
    if not all_rows:
        return {"rows": [], "issues": [{"error": empty_msg}]}

    # Detecta cabeçalho e mapeia
    header_idx, col_map = detect_header_and_map(all_rows)

    # Linhas de dados = abaixo do cabeçalho (não devolvemos o cabeçalho!)
    data_rows: List[List[str]] = []
    for row in all_rows[header_idx + 1 :]:
        if not row or not any(normalize_text(c) for c in row):
            continue
        if looks_like_total_row(row):
            # para quando encontrar totalização geral
            break
        data_rows.append(row)

    records: List[Dict[str, Any]] = []
    issues: List[Dict[str, Any]] = []

    for row in data_rows:
        safe_row = list(row)
        max_idx = max(col_map.values())
        if len(safe_row) <= max_idx:
            # completa com strings vazias se a linha for curta
            safe_row = safe_row + [""] * (max_idx + 1 - len(safe_row))

        rec = parse_row_with_map(safe_row, col_map)
        rec = validate_and_fix(rec)

        issue = build_issues(rec)
        if issue:
            issues.append(issue)

        records.append(
            {
                "item": rec.get("item"),
                "descricao": rec.get("descricao") or "",
                "unid": (rec.get("unid") or "").upper().rstrip(".,"),
                "quant": rec.get("quant"),
                "valor_unit": rec.get("valor_unit"),
                "valor_total": rec.get("valor_total"),
            }
        )

    # Ordena por item (quando existir)
    records.sort(key=lambda x: (x["item"] if x["item"] is not None else 10**9,))

    return {"rows": records, "issues": issues}


# -------------------- Função principal pública --------------------
def extract_table_from_xlsx(file_bytes: bytes) -> Dict[str, Any]:
    """
    Lê um arquivo XLSX, detecta o cabeçalho e extrai os campos
    ITEM, DESCRIÇÃO, UNID., QUANT., VALOR UNIT., VALOR TOTAL.

    Retorna NO MESMO FORMATO que o extrator antigo de PDF:
    {
      "rows": [...],
      "issues": [...]
    }
    """
    try:
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception:
        return {
            "rows": [],
            "issues": [{"error": "Não foi possível ler o arquivo XLSX."}],
        }

    # usa a planilha ativa (no Google Sheets normalmente é a única)
    sheet = wb.active

    all_rows: List[List[str]] = []
    for row in sheet.iter_rows(values_only=True):
        raw_row = ["" if cell is None else str(cell) for cell in row]
        norm_row = [normalize_text(c) for c in raw_row]
        if any(norm_row):
            all_rows.append(norm_row)

    return _extract_from_all_rows(
        all_rows, "Nenhuma tabela encontrada na planilha XLSX."
    )
