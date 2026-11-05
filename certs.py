"""База сертификатов: чтение папки, номера сертификатов и сопоставление с материалами акта."""
import re
from pathlib import Path

from learn import split_materials, stems

EXTS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp", ".doc", ".docx", ".xls", ".xlsx"}

# Кириллица, похожая на латиницу, — в номерах сертификатов пишут то так, то так
LOOKALIKE = str.maketrans("АВЕКМНОРСТХУавекмнорстху", "ABEKMHOPCTXYabekmhopctxy")

# Номер сертификата/декларации: KZ.7500669.01.01.02034, ЕАЭС RU C-RU.АЖ03.В.00177/20, РОСС RU.32001...
NUM_RE = re.compile(
    r"(?:ЕАЭС|EAЭC|ЕАЭC|РОСС|POCC|KZ|RU|KG|BY|UZ)[\s.№:-]*[A-ZА-ЯЁ0-9][A-ZА-ЯЁ0-9.\-/ ]{5,40}[0-9]", re.I)


NAME_STOP = {"серти", "соотв", "серт", "декла", "паспо", "проту", "прото", "скан", "копия", "копии", "лист"}


def norm_key(number):
    """Ключ для сравнения номеров: только буквы и цифры, латиницей, заглавными."""
    return re.sub(r"[^A-Z0-9]", "", number.translate(LOOKALIKE).upper())


PREFIXES = {"ЕАЭС", "EAЭC", "ЕАЭC", "РОСС", "POCC", "KZ", "RU", "KG", "BY", "UZ", "№"}
# номер после «№ / N / No»: «№ 7500669.01.01.02034», «№123-45/2024», «N KZ 75…»
AFTER_NO_RE = re.compile(r"(?:№|\bN[oо]?\.?|#)\s*([A-ZА-ЯЁ0-9][A-ZА-ЯЁ0-9.\-/ ]*\d)", re.I)
# любой «номерной» токен: 6+ цифр, с точками/дефисами/дробями (KZ.7500669.01.01.02034, 398550150091224001)
TOKEN_RE = re.compile(r"[A-ZА-ЯЁ0-9][A-ZА-ЯЁ0-9.\-/]*\d[A-ZА-ЯЁ0-9.\-/]*", re.I)
DATE_RE = re.compile(r"^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}\s*г?\.?$")
SIZE_RE = re.compile(r"\d+\s*[xх×*]\s*\d+", re.I)


def _trim(chunk):
    """Токены номера: префиксы и токены с цифрами; слова после номера («MasterTile», «WP 666») отсекаются."""
    keep = []
    for tok in chunk.split():
        if tok.upper().translate(LOOKALIKE) in {p.translate(LOOKALIKE) for p in PREFIXES} or re.search(r"\d", tok) \
                or (not keep and re.fullmatch(r"[A-ZА-ЯЁ]{2,5}", tok)):   # буквенная приставка: «РПБ 20.30…», «ТС RU…»
            if keep and re.search(r"[.\-/]\d", keep[-1]) and re.fullmatch(r"\d+", tok):
                break                                  # «…02146 666» — второе число уже не номер
            keep.append(tok)
        else:
            break
    return " ".join(keep).strip(" .-/,;")


# Имя файла по образцу реестра: «Сертификат №KZ.7500669.01.01.02034 - Смеси сухие строительные Sika.pdf»
NAME_RE = re.compile(r"^(?P<type>.*?)\s*№\s*(?P<num>.+?)\s+[-–—]\s+(?P<mat>.+)$")


def parse_name(stem):
    """(вид документа, номер, материал) из имени файла «Вид №НОМЕР - Материал», иначе None."""
    m = NAME_RE.match(stem)
    if not m:
        return None
    num = re.sub(r"\s+от\s+\d{1,2}\.\d{1,2}\.\d{2,4}.*$", "", m.group("num"))     # «… от 22.02.2022»
    num = re.sub(r"\s*\(.*?\)\s*$", "", num).strip(" .")                          # «(приложение, лист 3)»
    return m.group("type").strip(" ._"), num, m.group("mat").strip()


def digits(number):
    return re.sub(r"\D", "", number)


def numbers_in(text, loose=True):
    """Номера сертификатов/писем в тексте: с приставкой (KZ, ЕАЭС, РОСС…), после «№», и длинные номерные токены.
    Даты и размеры (1200х600) не считаются."""
    text = SIZE_RE.sub(" ", text or "")
    found = []
    for m in NUM_RE.finditer(text):
        found.append(_trim(m.group(0)))
    for m in AFTER_NO_RE.finditer(text):
        found.append(_trim(m.group(1)))
    if loose:                                   # в именах файлов — любой длинный номерной токен
        for m in TOKEN_RE.finditer(text):
            tok = m.group(0).strip(".-/")
            if len(digits(tok)) >= 6 and not DATE_RE.match(tok):
                found.append(tok)
    out = []
    for n in found:
        if len(digits(n)) >= 4 and not DATE_RE.match(n) and n not in out:
            out.append(n)
    return out


def num_keys(numbers):
    """Ключи сравнения: полный (буквы+цифры) и только цифры."""
    keys = set()
    for n in numbers:
        k, d = norm_key(n), digits(n)
        if len(k) >= 4:
            keys.add(k)
        if len(d) >= 4:
            keys.add("#" + d)
    return keys


def pdf_text(path, pages=2):
    """Текст первых страниц PDF (если это не скан-картинка)."""
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        return " ".join((reader.pages[i].extract_text() or "") for i in range(min(pages, len(reader.pages))))
    except Exception:
        return ""


def make_entry(path, number="", material="", doc_type="", valid="", extra_numbers=()):
    nums = ([number] if number else []) + [n for n in extra_numbers if n != number]
    return {"file": str(path), "name": Path(path).stem, "number": number or (nums[0] if nums else ""),
            "numbers": nums[:12], "keys": sorted(num_keys(nums[:12])), "folder": Path(path).parent.name,
            "material": material, "type": doc_type, "valid": valid}


def scan_folder(folder):
    """Все файлы сертификатов в папке (с подпапками). Номер — из имени «Вид №НОМЕР - Материал»,
    иначе из любого номера в имени; плюс номера из текста PDF."""
    out = []
    for p in sorted(Path(folder).rglob("*")):
        if not p.is_file() or p.suffix.lower() not in EXTS or p.name.startswith("~$"):
            continue
        parsed = parse_name(p.stem)
        number, material, doc_type = (parsed[1], parsed[2], parsed[0]) if parsed else ("", p.stem, "")
        extra = numbers_in(p.stem) + (numbers_in(pdf_text(p)) if p.suffix.lower() == ".pdf" else [])
        out.append(make_entry(p, number, material, doc_type, extra_numbers=extra))
    return out


def load_registry(xlsx):
    """Реестр сертификатов из Excel (лист с колонками «Номер документа», «Наименование материала», «Файл»…).
    Файлы ищутся по ссылке в колонке «Файл» относительно папки реестра."""
    from urllib.parse import unquote
    from openpyxl import load_workbook
    wb = load_workbook(str(xlsx))
    base = Path(xlsx).parent
    out = []
    for ws in wb.worksheets:
        head_row, cols = None, {}
        for r in range(1, min(ws.max_row, 15) + 1):
            vals = [str(c.value or "").strip().lower() for c in ws[r]]
            if any("номер" in v for v in vals) and any(v.startswith("файл") for v in vals):
                head_row = r
                for i, v in enumerate(vals):
                    for key, word in (("num", "номер"), ("mat", "наименование"), ("type", "вид"),
                                      ("valid", "действителен"), ("file", "файл")):
                        if v.startswith(word) and key not in cols:
                            cols[key] = i
                break
        if not head_row:
            continue
        for row in ws.iter_rows(min_row=head_row + 1):
            get = lambda k: (row[cols[k]].value if k in cols and cols[k] < len(row) else None)
            fcell = row[cols["file"]] if "file" in cols else None
            if not fcell or not (fcell.value or fcell.hyperlink):
                continue
            target = unquote(fcell.hyperlink.target) if fcell.hyperlink and fcell.hyperlink.target else str(fcell.value)
            path = (base / target) if not Path(target).is_absolute() else Path(target)
            valid = get("valid")
            valid = valid.strftime("%d.%m.%Y") if hasattr(valid, "strftime") else str(valid or "")
            number = str(get("num") or "")
            out.append(make_entry(path, number, str(get("mat") or ""), str(get("type") or ""), valid,
                                  extra_numbers=numbers_in(Path(target).stem)))
    return out


def brands(text):
    """Марки и обозначения: латинские слова (SIMH, PROFI, Knauf) и слова заглавными (ВВГ, НАШИ)."""
    return {w.lower() for w in re.findall(r"\b[A-Za-z][A-Za-z0-9\-]{2,}\b|\b[А-ЯЁ]{3,}\b", text or "")}


def _partial(k, x):
    """Цифровые ключи «#…» одного номера, записанного не полностью или с опечаткой."""
    if not (k[0] == x[0] == "#"):
        return False
    n = min(len(k), len(x))
    return ((n >= 9 and (k in x or x in k))                              # часть номера внутри
            or (n >= 7 and (k.endswith(x[1:]) or x.endswith(k[1:])))      # хвост номера
            # опечатка в середине («KZ.7500985.05.59523» вместо «…05.01.59523»): совпали начало и конец
            or (n >= 13 and k[1:8] == x[1:8] and k[-5:] == x[-5:]))


def match_material(item, certs):
    """Сертификат для одной позиции материала: сначала по номеру (полному или по цифрам), затем по названию."""
    text = item["name"] + " " + item.get("cert", "")
    # в тексте материала номер документа — только с приставкой (KZ, ЕАЭС…) или после «№»;
    # «1700/17», «3*1.5» и прочие числа марки номером не считаются
    keys = num_keys(numbers_in(text, loose=False) + ([item["cert"]] if item.get("cert") else []))
    for c in certs:
        if keys & set(c.get("keys", [])):
            return c, "номер"
    # номер записан частично («ЦСП 389497» для «KZ.7100841.13.12.389497») — только если подходит ровно один документ
    part = [c for c in certs if any(_partial(k, x) for k in keys for x in c.get("keys", []))]
    if len(part) == 1:
        return part[0], "номер"
    # У материала свой номер, а такого сертификата в базе нет — по названию не подбираем, иначе подставится чужой
    if keys:
        return None, ""
    # по названию: слова материала из реестра/имени файла должны быть в названии материала акта,
    # а марка (SIMH, PROFI, X-PERT, ВВГ…) — совпасть обязательно
    words = stems(item["name"])
    item_brands = brands(item["name"])
    best, score = None, 0.0
    for c in certs:
        title = c.get("material") or c["name"]
        fw = stems(title) - NAME_STOP
        if not fw:
            continue
        cb = brands(title)
        if cb and not (cb & item_brands):
            continue
        s = len(fw & words) / len(fw)
        if s > score:
            best, score = c, s
    return (best, "название") if score >= 0.67 else (None, "")


def match_materials(materials_text, certs):
    """[(позиция материала, сертификат или None, как найден)] для текста «применены»."""
    return [(it, *match_material(it, certs)) for it in split_materials(materials_text or "")]
