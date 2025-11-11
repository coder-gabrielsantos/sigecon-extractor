from typing import List, Dict, Any, Optional, Tuple
import io, re
import pdfplumber

# --- Regex / sets auxiliares ---
CURRENCY_RX = re.compile(r"R\$\s*")
SPACES_RX = re.compile(r"\s+")
INT_RX = re.compile(r"^\d{1,4}$")

# Aceita UNIDADES comuns (usada apenas para validação leve; não inferimos da descrição)
UNIT_SET = {
    "UN", "CJ", "UM", "M", "BAR", "TUB", "CX", "KIT", "PAR", "PÇ", "PC", "JG",
    "KG", "LT", "L", "G", "MM", "CM", "MT", "ROL", "SAC", "FD", "SC", "TB", "BL", "CONJ"
}

# Cabeçalhos canônicos e aliases
CANON_KEYS = ["ITEM", "DESCRIÇÃO", "UNID.", "QUANT.", "VALOR UNIT.", "VALOR TOTAL"]
HEADER_ALIASES = {
    "ITEM": {"ITEM", "ITEN", "ITENS"},
    "DESCRIÇÃO": {"DESCRIÇÃO", "DESCRICAO", "DESCRIÇÃO/ESPECIFICAÇÃO", "DESCRIÇÃO DO ITEM"},
    "UNID.": {"UNID.", "UNIDADE", "UND", "UNID"},
    "QUANT.": {"QUANT.", "QTD", "QUANTIDADE"},
    "VALOR UNIT.": {"VALOR UNIT.", "VALOR UNITARIO", "VALOR UNITÁRIO", "VLR UNIT.", "PREÇO UNIT.", "PREÇO UNIT"},
    "VALOR TOTAL": {"VALOR TOTAL", "VLR TOTAL", "TOTAL", "PREÇO TOTAL"},
}

INCH_VARIANTS = ['"', '\"', '”', '“', '″']
INCH_GLYPH = "″"

VALOR_TOTAL_SUFFIX_RX = re.compile(r"(VALOR\s+TOTAL\s*R\$\s*[\d\.\,]+)$", flags=re.IGNORECASE)
VALOR_PRECO_TRAILING_RX = re.compile(r"(R\$\s*[\d\.\,]+)$", flags=re.IGNORECASE)


# -------------------- Utilitários --------------------
def normalize_inches(text: str) -> str:
    if not text:
        return text
    for v in INCH_VARIANTS:
        text = text.replace(v, '"')
    text = re.sub(r"\s+\"", "\"", text)
    text = re.sub(r"(\d+(?:[\.,]\d+)?(?:/\d+(?:[\.,]\d+)?)?)\"", r"\1" + INCH_GLYPH, text)
    text = text.replace('"', INCH_GLYPH)
    return text


def strip_valor_total_from_desc(text: str) -> str:
    if not text:
        return text
    text = VALOR_TOTAL_SUFFIX_RX.sub("", text).strip()
    text = VALOR_PRECO_TRAILING_RX.sub("", text).strip()
    return text


def normalize_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = s.replace("\xa0", " ")
    s = SPACES_RX.sub(" ", s).strip()
    s = normalize_inches(s)
    return s


def parse_money(s: str) -> Optional[float]:
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


def parse_int(s: str) -> Optional[int]:
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
    """
    Retorna 1 se a célula bate em algum alias do target_key, caso contrário 0.
    """
    cell_u = normalize_text(cell).upper()
    for alias in HEADER_ALIASES[target_key]:
        if alias in cell_u:
            return 1
    return 0


def detect_header_and_map(rows: List[List[str]]) -> Tuple[int, Dict[str, int]]:
    """
    Encontra o índice da linha de cabeçalho e mapeia os índices de coluna para as chaves canônicas.
    Retorna (header_index, col_map). Lança ValueError se não encontrar algo minimamente válido.
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
        raise ValueError("Cabeçalho da tabela não foi encontrado com confiança suficiente.")

    header_row = rows[best_idx]
    col_map: Dict[str, int] = {}
    for key in CANON_KEYS:
        idx = _find_col_index(header_row, key)
        if idx is not None:
            col_map[key] = idx

    # Exigimos as 6 colunas canônicas
    missing = [k for k in CANON_KEYS if k not in col_map]
    if missing:
        raise ValueError(f"Não foi possível mapear todas as colunas: faltam {missing}")

    return best_idx, col_map


def _find_col_index(header_row: List[str], key: str) -> Optional[int]:
    """
    Procura o índice da coluna cujo texto casa com algum alias do 'key'.
    """
    for j, cell in enumerate(header_row):
        cell_u = normalize_text(cell).upper()
        for alias in HEADER_ALIASES[key]:
            if alias in cell_u:
                return j
    return None


# -------------------- Parsing de Linhas baseadas no mapeamento --------------------
def parse_row_with_map(row: List[str], col_map: Dict[str, int]) -> Dict[str, Any]:
    """
    Extrai campos exclusivamente pelas colunas mapeadas.
    NÃO infere unidade via descrição. Não força UN. Apenas normaliza e valida.
    """
    get = lambda key: normalize_text(row[col_map[key]]) if col_map[key] < len(row) else ""

    item = parse_int(get("ITEM"))
    descricao = strip_valor_total_from_desc(get("DESCRIÇÃO"))
    unid_raw = normalize_text(get("UNID.")).upper()
    # Limpa pontuação/resíduos triviais na unidade (ex.: "UN." -> "UN")
    unid = unid_raw.rstrip(".,")

    quant = parse_int(get("QUANT."))
    valor_unit = parse_money(get("VALOR UNIT."))
    valor_total = parse_money(get("VALOR TOTAL"))

    rec = {
        "item": item,
        "descricao": descricao,
        "unid": unid,
        "quant": quant,
        "valor_unit": valor_unit,
        "valor_total": valor_total,
    }
    return rec


def validate_and_fix(rec: Dict[str, Any]) -> Dict[str, Any]:
    """
    Valida consistência básica de valores. NÃO altera 'unid' por inferência textual.
    Completa valor_total a partir de valor_unit * quant quando aplicável.
    """
    vu = rec.get("valor_unit")
    vt = rec.get("valor_total")
    qt = rec.get("quant")

    # Recalcula total se unitário e quantidade estiverem presentes
    if vu is not None and qt is not None:
        calc = round(vu * qt, 2)
        # Se vt ausente ou muito próximo do calculado, substitui
        if vt is None or abs(calc - vt) <= 0.05:
            rec["valor_total"] = calc

    return rec


def build_issues(rec: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Reporta campos essenciais ausentes. Não marca 'unidade fora do set' como erro fatal,
    apenas informa se for útil.
    """
    missing = []
    for k in ["item", "descricao", "unid", "quant", "valor_unit", "valor_total"]:
        v = rec.get(k)
        if v in (None, ""):
            missing.append(k)

    if not missing:
        # Sinalização leve: unidade inesperada (não bloqueia)
        unid = (rec.get("unid") or "").upper()
        if unid and unid not in UNIT_SET:
            return {
                "item": rec.get("item"),
                "warning": f"Unidade '{unid}' fora da lista conhecida",
                "row": {k: rec.get(k) for k in ["descricao", "unid", "quant", "valor_unit", "valor_total"]},
            }
        return None

    return {
        "item": rec.get("item"),
        "missing": missing,
        "row": {k: rec.get(k) for k in ["descricao", "unid", "quant", "valor_unit", "valor_total"]},
    }


# -------------------- Função principal --------------------
def extract_table_from_pdf(file_bytes: bytes) -> Dict[str, Any]:
    """
    Lê o PDF, detecta o cabeçalho, mapeia as colunas e extrai os registros APENAS pelas colunas.
    Sem inferência por descrição. Sem fallback de unidade para 'UN'.
    """
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_rows: List[List[str]] = []
        for page in pdf.pages:
            tables = page.extract_tables(table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
                "snap_tolerance": 3,
                "join_tolerance": 3,
                "edge_min_length": 20,
            }) or []
            for tb in tables:
                for raw_row in tb:
                    row = [normalize_text(c) for c in (raw_row or [])]
                    if any(cell for cell in row):
                        all_rows.append(row)

    if not all_rows:
        return {"rows": [], "issues": [{"error": "Nenhuma tabela encontrada no PDF."}]}

    # Detecta cabeçalho e mapeamento
    header_idx, col_map = detect_header_and_map(all_rows)

    # Consome linhas de dados até o final ou até encontrar totalização
    data_rows = []
    for row in all_rows[header_idx + 1:]:
        if not row or not any(normalize_text(c) for c in row):
            continue
        if looks_like_total_row(row):
            # Paramos ao encontrar a linha de totalização geral
            break
        data_rows.append(row)

    records: List[Dict[str, Any]] = []
    issues: List[Dict[str, Any]] = []

    for row in data_rows:
        # Garante comprimento mínimo para indexação segura
        # (Se a tabela vier com células a menos em alguma linha)
        safe_row = list(row)
        max_idx = max(col_map.values())
        if len(safe_row) <= max_idx:
            # Completa com strings vazias
            safe_row = safe_row + [""] * (max_idx + 1 - len(safe_row))

        rec = parse_row_with_map(safe_row, col_map)
        rec = validate_and_fix(rec)

        issue = build_issues(rec)
        if issue:
            issues.append(issue)

        records.append({
            "item": rec.get("item"),
            "descricao": rec.get("descricao") or "",
            "unid": (rec.get("unid") or "").upper().rstrip(".,"),
            "quant": rec.get("quant"),
            "valor_unit": rec.get("valor_unit"),
            "valor_total": rec.get("valor_total"),
        })

    # Ordena por item (quando existir)
    records.sort(key=lambda x: (x["item"] if x["item"] is not None else 10 ** 9,))

    return {"rows": records, "issues": issues}
