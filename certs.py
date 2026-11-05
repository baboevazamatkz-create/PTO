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


def numbers_in(text):
    """Номера сертификатов в тексте; слова после номера («MasterTile 665», «WP 666») отсекаются."""
    out = []
    for m in NUM_RE.finditer(text or ""):
        keep = []
        for tok in m.group(0).split():
            if tok.upper().translate(LOOKALIKE) in {p.translate(LOOKALIKE) for p in PREFIXES} \
                    or re.search(r"[.\-/\d]", tok):
                if keep and re.search(r"[.\-/]\d", keep[-1]) and re.fullmatch(r"\d+", tok):
                    break                              # «…02146 666» — второе число уже не номер
                keep.append(tok)
            else:
                break
        num = " ".join(keep).strip(" .-/")
        if len(norm_key(num)) >= 8:
            out.append(num)
    return out


def pdf_text(path, pages=2):
    """Текст первых страниц PDF (если это не скан-картинка)."""
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        return " ".join((reader.pages[i].extract_text() or "") for i in range(min(pages, len(reader.pages))))
    except Exception:
        return ""


def scan_folder(folder):
    """Все файлы сертификатов в папке (с подпапками): номер, название, файл."""
    out = []
    for p in sorted(Path(folder).rglob("*")):
        if not p.is_file() or p.suffix.lower() not in EXTS or p.name.startswith("~$"):
            continue
        nums = numbers_in(p.stem)
        if not nums and p.suffix.lower() == ".pdf":
            nums = numbers_in(pdf_text(p))
        out.append({"file": str(p), "name": p.stem, "number": nums[0] if nums else "",
                    "keys": sorted({norm_key(n) for n in nums}), "folder": p.parent.name})
    return out


def match_material(item, certs):
    """Сертификат для одной позиции материала: сначала по номеру, затем по названию."""
    text = item["name"] + " " + item.get("cert", "")
    keys = {norm_key(n) for n in numbers_in(text)} | ({norm_key(item["cert"])} if item.get("cert") else set())
    keys = {k for k in keys if len(k) >= 8}
    for c in certs:
        for ck in c["keys"]:
            for k in keys:
                if ck == k or (len(k) >= 10 and (k in ck or ck in k)):
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
