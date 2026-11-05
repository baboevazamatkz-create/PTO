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
DATE_RE = re.compile(r"^\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}$")
SIZE_RE = re.compile(r"\d+\s*[xх×*]\s*\d+", re.I)


def _trim(chunk):
    """Токены номера: префиксы и токены с цифрами; слова после номера («MasterTile», «WP 666») отсекаются."""
    keep = []
    for tok in chunk.split():
        if tok.upper().translate(LOOKALIKE) in {p.translate(LOOKALIKE) for p in PREFIXES} or re.search(r"\d", tok):
            if keep and re.search(r"[.\-/]\d", keep[-1]) and re.fullmatch(r"\d+", tok):
                break                                  # «…02146 666» — второе число уже не номер
            keep.append(tok)
        else:
            break
    return " ".join(keep).strip(" .-/,;")


def digits(number):
    return re.sub(r"\D", "", number)


def numbers_in(text):
    """Номера сертификатов/писем в тексте: с приставкой (KZ, ЕАЭС, РОСС…), после «№», и длинные номерные токены.
    Даты и размеры (1200х600) не считаются."""
    text = SIZE_RE.sub(" ", text or "")
    found = []
    for m in NUM_RE.finditer(text):
        found.append(_trim(m.group(0)))
    for m in AFTER_NO_RE.finditer(text):
        found.append(_trim(m.group(1)))
    for m in TOKEN_RE.finditer(text):
        tok = m.group(0).strip(".-/")
        if len(digits(tok)) >= 6 and not DATE_RE.match(tok):
            found.append(tok)
    out = []
    for n in found:
        if len(digits(n)) >= 5 and not DATE_RE.match(n) and n not in out:
            out.append(n)
    return out


def num_keys(numbers):
    """Ключи сравнения: полный (буквы+цифры) и только цифры."""
    keys = set()
    for n in numbers:
        k, d = norm_key(n), digits(n)
        if len(k) >= 6:
            keys.add(k)
        if len(d) >= 6:
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


def scan_folder(folder):
    """Все файлы сертификатов в папке (с подпапками): номера из имени файла И из содержимого PDF."""
    out = []
    for p in sorted(Path(folder).rglob("*")):
        if not p.is_file() or p.suffix.lower() not in EXTS or p.name.startswith("~$"):
            continue
        from_name = numbers_in(p.stem)
        from_text = numbers_in(pdf_text(p)) if p.suffix.lower() == ".pdf" else []
        nums = from_name + [n for n in from_text if n not in from_name]
        out.append({"file": str(p), "name": p.stem, "number": nums[0] if nums else "",
                    "numbers": nums[:12], "keys": sorted(num_keys(nums[:12])), "folder": p.parent.name})
    return out


def match_material(item, certs):
    """Сертификат для одной позиции материала: сначала по номеру (полному или по цифрам), затем по названию."""
    text = item["name"] + " " + item.get("cert", "")
    keys = num_keys(numbers_in(text) + ([item["cert"]] if item.get("cert") else []))
    for c in certs:
        ck = set(c.get("keys", []))
        if keys & ck:
            return c, "номер"
        for k in keys:                                 # номер записан частично: «…01.01.02034» в «KZ.7500669.01.01.02034»
            for x in ck:
                if k[0] == x[0] == "#" and (
                        (min(len(k), len(x)) >= 9 and (k in x or x in k))             # часть номера внутри
                        or (min(len(k), len(x)) >= 7 and (k.endswith(x[1:]) or x.endswith(k[1:])))):  # хвост номера
                    return c, "номер"
    # по названию: слова из имени файла (без «сертификат», «соответствия»…) должны быть в названии материала
    words = stems(item["name"])
    best, score = None, 0.0
    for c in certs:
        fw = stems(c["name"]) - NAME_STOP
        if not fw:
            continue
        s = len(fw & words) / len(fw)
        if s > score:
            best, score = c, s
    return (best, "название") if score >= 0.5 else (None, "")


def match_materials(materials_text, certs):
    """[(позиция материала, сертификат или None, как найден)] для текста «применены»."""
    return [(it, *match_material(it, certs)) for it in split_materials(materials_text or "")]
