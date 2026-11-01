"""База знаний из прошлых актов: извлечение полей и связанные подсказки."""
import re
from datetime import date, timedelta
from pathlib import Path

from docx import Document

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]
DATE_RE = r"«\s*(\d{1,2})\s*»\s*([а-яё]+)\s*(\d{4})"


def _date(m):
    """Совпадение «18» мая 2026 -> '18.05.2026'."""
    try:
        return f"{int(m.group(1)):02d}.{MONTHS.index(m.group(2).lower()) + 1:02d}.{m.group(3)}"
    except (ValueError, AttributeError):
        return ""


def to_date(s):
    try:
        d, m, y = (int(x) for x in s.split("."))
        return date(y, m, d)
    except (ValueError, AttributeError):
        return None


def act_num(act_no):
    m = re.match(r"\s*(\d+)", act_no or "")
    return int(m.group(1)) if m else 0


FIELD_STOPS = ("(наименование", "наименование", "при наличии", "должность, фамилия")


def _clean(t):
    t = re.sub(r"_{2,}", " ", t.replace("\xa0", " ").replace("\t", " "))
    return re.sub(r"\s+", " ", t).strip(" _")


def _after(text, marker):
    i = text.lower().find(marker)
    return text[i + len(marker):] if i >= 0 else None


def _collect(paras, i, first, stops):
    """Значение поля: остаток абзаца i + следующие абзацы до подстрочника/следующего пункта."""
    parts = [first]
    for t in paras[i + 1:]:
        lt = t.lower()
        if not t or lt.startswith(stops):
            break
        parts.append(t)
    return _clean(" ".join(parts))


def parse_act(path):
    """Извлекает поля одного акта .docx. Возвращает dict или None."""
    try:
        paras = [p.text.strip() for p in Document(str(path)).paragraphs]
    except Exception:
        return None
    rec = {"act_no": "", "works": "", "docs": "", "materials": "", "next_works": "",
           "start": "", "end": "", "signers": [], "file": str(path), "source": "импорт"}
    text = " ".join(paras)
    rec["start"] = _date(re.search(r"начал\w*\s+работ\s*" + DATE_RE, text, re.I))
    rec["end"] = _date(re.search(r"окончани\w*\s+работ\s*" + DATE_RE, text, re.I))
    in_signs = False
    for i, t in enumerate(paras):
        lt = t.lower()
        if not rec["act_no"] and lt.replace(" ", "").startswith("акт№"):
            rec["act_no"] = _clean(t.split("№", 1)[1])
        elif "предъявлены работы:" in lt and not rec["works"]:
            rec["works"] = _collect(paras, i, _after(t, "предъявлены работы:"), FIELD_STOPS)
        elif "проектной документации:" in lt and not rec["docs"]:
            rec["docs"] = _collect(paras, i, _after(t, "проектной документации:"), FIELD_STOPS)
        elif "работ применены:" in lt and not rec["materials"]:
            rec["materials"] = _collect(paras, i, _after(t, "работ применены:"),
                                        FIELD_STOPS + ("4", "при выполнении работ отсутств", "даты"))
        elif lt.startswith("решение комиссии"):
            in_signs = True
        elif in_signs and "/" in t:
            name = _clean(t.rsplit("/", 1)[1]).strip(" ._")
            if name and name not in rec["signers"]:
                rec["signers"].append(name)
        elif lt.startswith("на основании изложенного") and not rec["next_works"]:
            rest = _after(t, "(монтажу)") or ""
            rest = re.sub(r"наименование последующих работ и конструкций:?", "", rest).strip(" :")
            rec["next_works"] = _collect(paras, i, rest, FIELD_STOPS)
    if not rec["works"]:
        return None
    rec["materials"] = rec["materials"].rstrip(".") + "." if rec["materials"] else ""
    return rec


# ---------- материалы: разбор строки «применены» на отдельные позиции
CERT_RE = re.compile(r"\s*[–—-]\s*сертификат соответствия\s*№\s*(.+?)\.?$", re.I)


def split_materials(text):
    """Делит строку «применены» на позиции: по «;», по «. Заглавная» и по «, » после номера сертификата."""
    chunks = re.split(r";\s*|(?<=[\d\)»\w])\.\s+(?=[А-ЯЁA-Z])", text.strip())
    items = []
    for ch in chunks:
        buf = ""
        for part in re.split(r",\s+", ch):
            buf = f"{buf}, {part}" if buf else part
            if re.search(r"(№\s*\S*\d\S*|сертификац\w*\)?)$", buf.strip(" .)")):
                items.append(buf)
                buf = ""
        if buf:
            if items and not re.search(r"№|сертифик", buf, re.I) and len(buf) < 40:
                items[-1] += ", " + buf
            else:
                items.append(buf)
    out = []
    for it in items:
        it = it.strip().rstrip(".").strip()
        if len(it) < 4:
            continue
        m = CERT_RE.search(it)
        if m:
            out.append({"name": it[:m.start()].strip(), "cert": m.group(1).strip()})
        else:
            out.append({"name": it[:1].upper() + it[1:], "cert": ""})
    return out


# ---------- признаки для сравнения
ROOM_RE = re.compile(r"помещени[еяй]\s*(?:№\s*)*(\d+)", re.I)
ROOM_NAME_RE = re.compile(r"помещения\s+(.+?)\s*№\s*(\d+)", re.I)
WORD_RE = re.compile(r"[а-яёa-z]{3,}", re.I)
STOP = {"помещение", "этаж", "оси", "отм", "мм", "толщиной", "для", "под", "при", "все", "что"}


def room_of(text):
    m = ROOM_RE.search(text or "")
    return m.group(1) if m else ""


def rooms_of(text):
    """Все помещения акта: «Помещения №№ 16, 19, 22 (…)» -> ['16', '19', '22']."""
    m = re.search(r"помещени\w*\s*((?:№\s*)*\d+(?:\s*(?:,|и)\s*(?:№\s*)*\d+)*)", text or "", re.I)
    return re.findall(r"\d+", m.group(1)) if m else []


# Этапы работ в технологическом порядке (колонки сводной таблицы).
# При разборе берётся ПЕРВОЕ совпадение из STAGE_RULES — порядок правил важен.
STAGES = ["Дверные проёмы", "Каркас", "Электрика", "Вентиляция", "Теплоизоляция", "Обшивка ГКЛ/СМЛ",
          "Сетка Фарадея", "Греющий кабель", "Выравнивание, штукатурка", "Стяжка", "Гидроизоляция",
          "Трапы, сантехника", "Чистовая отделка", "Витраж", "Бассейн", "Прочее"]
STAGE_RULES = [
    ("Бассейн", r"бассейн"),
    ("Дверные проёмы", r"дверн"),
    ("Трапы, сантехника", r"трап|сантехн|смесител"),
    ("Вентиляция", r"вентиляц"),
    ("Витраж", r"витраж"),
    ("Сетка Фарадея", r"фараде"),
    ("Чистовая отделка", r"затирк|плитк|керамогранит|кафел|мозаик|камн|вагонк|слэб|чистов"),
    ("Выравнивание, штукатурка", r"оштукатур|штукатур"),
    ("Теплоизоляция", r"теплов|теплоизол"),
    ("Обшивка ГКЛ/СМЛ", r"облицовк|обшивк"),
    ("Гидроизоляция", r"гидроизол|герметиз"),
    ("Стяжка", r"стяжк"),
    ("Выравнивание, штукатурка", r"выравнива"),
    ("Греющий кабель", r"греющ|нагреват"),
    ("Электрика", r"электр|кабел"),
    ("Каркас", r"каркас|обрешетк|контррейк"),
]


# Типы помещений: у хаммамов (и пилингов) — металлокаркас, панели, греющий кабель, мозаика;
# у бань и саун — деревянный каркас, утеплитель с фольгой, обшивка, вагонка.
KIND_RULES = [("Хаммам", r"хам+ам|пилинг|тепидари"),
              ("Баня, сауна", r"парн|саун|бан[яи]|финск|парени|запарив")]
KIND_PLAN = {
    "Хаммам": ["Дверные проёмы", "Каркас", "Электрика", "Вентиляция", "Теплоизоляция", "Сетка Фарадея",
               "Греющий кабель", "Выравнивание, штукатурка", "Стяжка", "Гидроизоляция", "Трапы, сантехника",
               "Чистовая отделка"],
    "Баня, сауна": ["Дверные проёмы", "Каркас", "Электрика", "Вентиляция", "Теплоизоляция", "Обшивка ГКЛ/СМЛ",
                    "Гидроизоляция", "Трапы, сантехника", "Чистовая отделка"],
}

# Наименование работ по умолчанию, если в базе ещё нет такого этапа у помещений того же типа
KIND_WORK = {
    ("Хаммам", "Каркас"): "Устройство металлического каркаса стен, потолка и лежака из профильной трубы и профиля",
    ("Хаммам", "Теплоизоляция"): "Монтаж теплоизоляционных панелей на стены, потолок, сиденья и лежак",
    ("Хаммам", "Электрика"): "Прокладка кабелей в запотолочном и межстеновом пространствах",
    ("Хаммам", "Чистовая отделка"): "Облицовка стен, потолка и пола мозаикой с затиркой швов",
    ("Баня, сауна", "Каркас"): "Устройство деревянного каркаса стен и потолка из бруса с огнебиозащитной пропиткой",
    ("Баня, сауна", "Чистовая отделка"): "Чистовая отделка стен и потолка вагонкой",
}

# Сокращения для имени файла — как в архиве актов («обмаз. Г.И. пола», «вырав. ЦПС пола»)
ABBR = [(r"цементно-песчан\w*\s+стяжк\w*", "ЦПС"), (r"гидроизоляци\w*", "Г.И."), (r"гидроизол\w*", "Г.И."),
        (r"обмазочн\w*", "обмаз."), (r"выравнивающ\w*", "вырав."), (r"греющ\w*\s+кабел\w*", "гр. кабеля"),
        (r"нагревательн\w*\s+кабел\w*", "нагр. кабеля"), (r"восстановлени\w*", "восст."),
        (r"теплоизоляци\w*|теплов\w* изоляци\w*", "теплоизол."), (r"электрокабел\w*|электрическ\w* кабел\w*", "эл. кабеля"),
        (r"запотолочн\w*", "запотол."), (r"межстен\w*", "межстен."), (r"имеющ\w* пропитку средством огнебиозащиты", ""),
        (r"с выводом заземления и проверкой металлосвязи \(целостности цепи\)", "")]


def stage_of(works):
    wp = work_part(works).lower()
    return next((name for name, rx in STAGE_RULES if re.search(rx, wp)), "Прочее")


def floor_of(text):
    """Этаж из текста: «(2 этаж)», «(5 эт.)», «(1 блок, 2 этаж)», «5-й этаж»."""
    m = re.search(r"(\d+)\s*(?:-?й\s*)?эт", text or "")
    return m.group(1) if m else ""


def key_no(key):
    """Ключ помещения «38.2» -> номер «38»."""
    return key.split(".")[0]


def key_floor(key):
    return key.split(".")[1] if "." in key else ""


def typed_room(text):
    """Номер помещения из того, что набрал пользователь: «помещение 32», «п. 32», «№32» или просто «32»."""
    m = ROOM_RE.search(text) or re.search(r"(?:\bп\.|№)\s*(\d{1,3})\b", text) \
        or re.search(r"(?<![\d.,+])\b(\d{1,3})\b(?![.,]\d)", text)
    return (m.group(1), text[:m.start()] + text[m.end():]) if m else ("", text)


def work_part(works):
    """Наименование работ без «Помещение № … оси …»."""
    m = ROOM_RE.search(works)
    return works[:m.start()].strip(" .,") if m else works.strip()


def location_part(works):
    m = ROOM_RE.search(works)
    return works[m.start():].strip() if m else ""


def stems(text):
    return {w.lower()[:5] for w in WORD_RE.findall(text or "") if w.lower() not in STOP}


def sim(a, b):
    a, b = stems(a), stems(b)
    return len(a & b) / len(a | b) if a and b else 0.0


# ---------- проект: какие листы альбома ссылать для вида работ
def project_sheets(sheets, work, stage):
    """Номера листов альбома, подходящих к работам (по названиям листов)."""
    w = work.lower()
    want = []
    floor = re.search(r"\bпол(а|у|ов)?\b|наполь|стяжк|трап|фараде", w)
    if floor:
        want.append(r"план пол(а|ов)\b")
    if re.search(r"нагрев|греющ", w):
        want.append(r"нагревательн")
    if re.search(r"лежак|сиден|седуш|спин", w):
        want.append(r"лежак|сиден")
    if re.search(r"полк", w):
        want.append(r"полков")
    if re.search(r"купол", w):
        want.append(r"купол")
    if re.search(r"контр-?рейк", w):
        want.append(r"контр-?рейк")
    if re.search(r"стен|потол|купол", w) and stage in ("Выравнивание, штукатурка", "Гидроизоляция",
                                                       "Чистовая отделка"):
        want.append(r"разрез")                    # стены/потолок/купол в актах ссылаются на разрезы
    elif re.search(r"потол", w):
        want.append(r"каркас.*потол|потол.*каркас" if stage == "Каркас" else r"план потолка")
    if re.search(r"стен", w) and stage in ("Каркас", "Теплоизоляция", "Обшивка ГКЛ/СМЛ"):
        want.append(r"стенов")
    if stage == "Электрика":
        want += [r"освещени|подсветк", r"инженерн"]
    if stage in ("Вентиляция", "Трапы, сантехника"):
        want.append(r"инженерн")
    if not want and stage in ("Каркас", "Теплоизоляция", "Обшивка ГКЛ/СМЛ"):
        want.append(r"стенов|каркас")
    found = [no for no, name in sheets.items()
             if any(re.search(rx, name.lower()) for rx in want) and not re.search(r"специфик|общие", name.lower())]
    if stage != "Чистовая отделка":               # для гидроизоляции и выравнивания хватает первого разреза
        cuts = [no for no in found if "разрез" in sheets[no].lower()]
        found = [no for no in found if no not in cuts[1:]]
    if re.search(r"контр-?рейк", w):              # контррейка — свой лист, без общего каркаса
        found = [no for no in found if re.search(r"контр-?рейк", sheets[no].lower())] or found
    return found[:5]


def sheets_text(nums):
    return ("лист " if len(nums) == 1 else "листы ") + ", ".join(nums)


def short_work(works, limit=70):
    """Короткое наименование работ для имени файла."""
    t = work_part(works)
    t = t[:1].lower() + t[1:]
    for rx, sub in ABBR:
        t = re.sub(rx, sub, t, flags=re.I)
    t = re.sub(r"\s+", " ", t).strip(" .,")
    if len(t) > limit:
        t = t[:limit].rsplit(" ", 1)[0].rstrip(" ,.(")
    return t


def safe_name(t):
    return re.sub(r'[\\/:*?"<>|]+', " ", t).strip()


class Knowledge:
    def __init__(self, records, project=None):
        self.project = project if project and project.get("rooms") else {"rooms": {}}
        # помещение = номер + этаж («38.2» и «38.5» — разные помещения); старые ключи «38» дополняем этажом
        rooms = {}
        for k, p in self.project["rooms"].items():
            fl = re.sub(r"\D", "", p.get("floor", ""))
            rooms[k if "." in k or not fl else f"{k}.{fl}"] = p
        self.project["rooms"] = rooms
        self.recs = []
        seen = set()
        for r in records:
            key = (r.get("act_no"), r.get("works"))
            if key not in seen and r.get("works"):
                seen.add(key)
                self.recs.append(r)

    def upsert(self, rec):
        """Добавить акт или обновить уже известный (тот же номер и работы). True — если новый."""
        key = (rec.get("act_no"), rec.get("works"))
        for i, r in enumerate(self.recs):
            if (r.get("act_no"), r.get("works")) == key:
                merged = {**r, **{k: v for k, v in rec.items() if v}}
                if r.get("source") == "сформирован":     # акт из программы остаётся «сформирован»
                    merged["source"] = "сформирован"
                self.recs[i] = merged
                return False
        if rec.get("works"):
            self.recs.append(rec)
            return True
        return False

    def add(self, rec):
        key = (rec.get("act_no"), rec.get("works"))
        if rec.get("works") and all((r.get("act_no"), r.get("works")) != key for r in self.recs):
            self.recs.append(rec)
            return True
        return False

    def rec_floor(self, r):
        return floor_of(location_part(r["works"])) or floor_of(r["works"]) or \
            (re.search(r"[\\/](\d+) этаж[\\/]", r.get("file", "")) or [None, ""])[1]

    def keys(self, r):
        """Помещения акта как ключи «номер.этаж»."""
        fl = self.rec_floor(r)
        out = []
        for n in rooms_of(r["works"]):
            if fl:
                out.append(f"{n}.{fl}")
            else:                                  # этаж не указан: если в проекте такой номер один — он
                same = [k for k in self.project["rooms"] if key_no(k) == n]
                out.append(same[0] if len(same) == 1 else n)
        return out

    @staticmethod
    def _uniq(items, limit):
        out, seen = [], set()
        for x in items:
            if x and x.lower() not in seen:
                seen.add(x.lower())
                out.append(x)
            if len(out) >= limit:
                break
        return out

    def related(self, works):
        """Акты, похожие по виду работ (и помещению), по убыванию сходства."""
        room, wp = typed_room(works)
        wp = work_part(re.sub(r"помещени\w*", "", wp, flags=re.I))
        scored = []
        for r in self.recs:
            s = sim(wp, work_part(r["works"])) if wp else 0.0
            if room and room_of(r["works"]) == room:
                s += 0.3
            scored.append((s, r))
        scored.sort(key=lambda x: -x[0])
        return [r for s, r in scored if s > 0.05]

    def suggest_works(self, typed, limit=12):
        typed = typed.strip()
        room, wp = typed_room(typed)
        wp = re.sub(r"помещени\w*", "", wp, flags=re.I).strip(" .,")
        loc = next((location_part(r["works"]) for r in self.recs
                    if room and room_of(r["works"]) == room), "")

        def score(works, room_bonus):
            part = work_part(works)
            s = sim(wp, part) if wp else 0.0
            if wp and wp.lower() in part.lower():
                s += 0.5
            return s + room_bonus

        cands = {}
        for r in self.recs:
            w = r["works"]
            same = bool(room) and room_of(w) == room
            cands[w] = max(cands.get(w, 0), score(w, 0.4 if same else 0.0))
            if room and loc and not same:
                # тот же вид работ, но в нужном помещении (оси и отметка — из его актов)
                sub = work_part(w) + ". " + loc
                cands[sub] = max(cands.get(sub, 0), score(sub, 0.35))
        ranked = sorted(cands.items(), key=lambda x: -x[1])
        return self._uniq([w for w, s in ranked if s > 0.05 or not typed], limit)

    def suggest_docs(self, works, limit=10):
        room = typed_room(works)[0]
        rel = self.related(works)
        same_room = [r["docs"] for r in rel if room and room_of(r["works"]) == room]
        rest = [r["docs"] for r in rel]
        if room and not same_room:
            # проект по номеру помещения из любого акта этого помещения
            same_room = [r["docs"] for r in self.recs if room_of(r["works"]) == room]
        return self._uniq(same_room + rest, limit)

    def suggest_next(self, works, limit=10):
        return self._uniq([r["next_works"] for r in self.related(works)], limit)

    def suggest_materials(self, works, limit=3):
        """Названия материалов из самых похожих актов."""
        names = []
        for r in self.related(works)[:limit]:
            names += [m["name"] for m in split_materials(r["materials"])]
        return self._uniq(names, 30)

    # ---------- помещения и цепочки работ
    def room_acts(self, room):
        """Акты помещения по возрастанию номера."""
        return sorted((r for r in self.recs if room in self.keys(r)), key=lambda r: act_num(r["act_no"]))

    def rooms(self):
        """{ключ: подпись}, например {'38.2': '38 · 2 этаж — Хамам мужское отделение'}."""
        names = {}
        for r in self.recs:
            for key in self.keys(r):
                n = key_no(key)
                name = next((m.group(1).strip(" «»") for m in ROOM_NAME_RE.finditer(r.get("docs", ""))
                             if m.group(2) == n), "")
                if len(name) > len(names.get(key, "")) or key not in names:
                    names[key] = name or names.get(key, "")
        for key, p in self.project["rooms"].items():   # проект важнее: точное название, все помещения
            names[key] = p["name"]

        def label(key):
            fl = key_floor(key)
            return f"{key_no(key)}" + (f" · {fl} этаж" if fl else "") + (f" — {names[key]}" if names[key] else "")

        order = sorted(names, key=lambda k: (int(key_no(k)), key_floor(k)))
        return {k: label(k) for k in order}

    def room_name(self, room):
        p = self.project["rooms"].get(room)
        if p:
            return p["name"]
        lbl = self.rooms().get(room, "")
        return lbl.split(" — ", 1)[1] if " — " in lbl else ""

    def room_kind(self, room):
        name = self.room_name(room).lower()
        return next((k for k, rx in KIND_RULES if re.search(rx, name)), "")

    def room_floor(self, room):
        if key_floor(room):
            return key_floor(room)
        p = self.project["rooms"].get(room)
        if p and p.get("floor"):
            return re.sub(r"\D", "", p["floor"])
        for r in self.room_acts(room):
            m = re.search(r"\((\d+)\s*эт", r["works"])
            if m:
                return m.group(1)
        return ""

    def planned(self, room):
        """Этапы помещения: из проекта, иначе — типовые для бани или хаммама."""
        p = self.project["rooms"].get(room)
        if p and "stages" in p:
            return list(p["stages"])
        return list(KIND_PLAN.get(self.room_kind(room), []))

    def room_status(self, room):
        """[(этап, [акты])] в технологическом порядке: и сделанные, и оставшиеся по плану."""
        done = {}
        for r in self.recs:
            if room in self.keys(r):
                done.setdefault(stage_of(r["works"]), []).append(r)
        plan = set(self.planned(room))
        return [(st, sorted(done.get(st, []), key=lambda r: act_num(r["act_no"])))
                for st in STAGES if st in plan or st in done]

    def file_path(self, rec, base, hint=""):
        """База актов: <этаж> этаж / <пом> этаж <этаж> / АСР №<n> п.<пом> <работы>.docx — как в архиве."""
        keys = self.keys(rec) or ([hint] if hint else [])
        room = keys[0] if keys else ""
        floor = self.room_floor(room) if room else ""
        no = key_no(room)
        folder = Path(base)
        if room:
            folder = folder / (f"{floor} этаж" if floor else "Без этажа") / \
                (f"{no} этаж {floor}" if floor else no)
        rooms_txt = ", ".join(key_no(k) for k in keys)
        name = f'АСР №{rec["act_no"]}' + (f" п.{rooms_txt}" if rooms_txt else "") + " " + short_work(rec["works"])
        return folder / (safe_name(name) + ".docx")

    def project_docs(self, room, work):
        p = self.project["rooms"].get(room)
        if not p:
            return ""
        nums = project_sheets(p["sheets"], work, stage_of(work))
        if not nums:
            return ""
        return (f'ТОО «АКВАСЕРВИС»: «Альбом помещения {p["name"]} № {p["album_no"]}, {p["block"]}, '
                f'{p["floor"]}. ТХ, {sheets_text(nums)}».')

    def typical_work(self, stage, room=""):
        """Самое частое наименование работ этапа — в первую очередь у помещений того же типа."""
        kind = self.room_kind(room) if room else ""
        c = {}
        for r in self.recs:
            if stage_of(r["works"]) == stage:
                wp = work_part(r["works"])
                same = bool(kind) and any(self.room_kind(x) == kind for x in self.keys(r))
                c[wp] = c.get(wp, 0) + (3 if same else 1)
        if (kind, stage) in KIND_WORK:                # при равенстве — ближе к типовому для этого типа помещения
            for wp in c:
                c[wp] += 2 * sim(wp, KIND_WORK[(kind, stage)])
        if kind and not any(v >= 3 for v in c.values()) and (kind, stage) in KIND_WORK:  # примеров этого типа нет
            return KIND_WORK[(kind, stage)]
        return max(c, key=c.get) if c else stage

    def best_work(self, text, exclude=""):
        """Наименование работ из базы, ближе всего к тексту (по смыслу слов)."""
        best, score = "", 0.0
        for r in self.recs:
            wp = work_part(r["works"])
            if exclude and sim(wp, exclude) > 0.8:
                continue
            s = sim(text, wp) + (0.3 if text and text.lower()[:12] in wp.lower() else 0)
            if s > score:
                best, score = wp, s
        return best if score > 0.15 else ""

    def next_stage(self, room):
        """Следующий этап для помещения: «последующие работы» его последнего акта -> вид работ."""
        acts = self.room_acts(room)
        done = {stage_of(r["works"]) for r in acts}
        todo = [s for s in self.planned(room) if s not in done]
        if acts:
            last = acts[-1]
            nxt = self.best_work(last.get("next_works", ""), exclude=work_part(last["works"]))
            if nxt and (not todo or stage_of(nxt) in todo or stage_of(nxt) not in self.planned(room)):
                return nxt
        if todo:                                   # по плану: первый этап, по которому акта ещё нет
            return self.typical_work(todo[0], room)
        return ""                                  # все этапы по плану закрыты

    def work_types(self, typed="", room="", limit=15):
        """Виды работ (без помещения) для подсказок: сначала из актов этого помещения."""
        scored = {}
        for r in self.recs:
            wp = work_part(r["works"])
            s = sim(typed, wp) if typed else 0.1
            if typed and typed.lower() in wp.lower():
                s += 0.5
            if room and room_of(r["works"]) == room:
                s += 0.05
            scored[wp] = max(scored.get(wp, 0), s)
        ranked = sorted(scored.items(), key=lambda x: -x[1])
        return self._uniq([w for w, s in ranked if s > 0.05], limit)

    def summary(self):
        """Сводная: {помещение: {этап: [акты]}} и список задействованных этапов по порядку."""
        table = {}
        for r in self.recs:
            st = stage_of(r["works"])
            for room in self.keys(r):
                table.setdefault(room, {}).setdefault(st, []).append(r)
        for cells in table.values():
            for acts in cells.values():
                acts.sort(key=lambda r: act_num(r["act_no"]))
        for room, p in self.project["rooms"].items():
            table.setdefault(room, {})
        plan = {room: set(self.planned(room)) for room in table}
        used = [s for s in STAGES if any(s in c for c in table.values()) or any(s in pl for pl in plan.values())]
        rooms = sorted(table, key=lambda k: (int(key_no(k)), key_floor(k)))
        return rooms, used, table, plan

    def next_act_no(self):
        return str(max((act_num(r["act_no"]) for r in self.recs), default=0) + 1)

    def draft(self, room, work):
        """Черновик акта на помещение по виду работ — из самых похожих прошлых актов."""
        room_acts = self.room_acts(room)

        def score(r):
            wp = work_part(r["works"])
            s = sim(work, wp) + (0.5 if work and work.lower() in wp.lower() else 0)
            return s + (0.05 if room in self.keys(r) else 0)   # помещение — только при равенстве

        tpl = max(self.recs, key=score, default=None)
        if not tpl:
            return {}
        p = self.project["rooms"].get(room)
        single = [r for r in room_acts if self.keys(r) == [room]]   # без актов на несколько помещений
        if single:
            loc = location_part(single[-1]["works"])
        elif p or key_floor(room):                 # помещение без актов: этаж из ключа/проекта, отметку и оси — вручную
            loc = f'Помещение № {key_no(room)} ({self.room_floor(room)} этаж) отм. ___ оси ___.'
        else:
            loc = location_part(tpl["works"])
        wp = work_part(tpl["works"]) if sim(work, work_part(tpl["works"])) > 0.3 else work
        works = (wp.rstrip(" .") + ". " + loc) if loc else wp
        # проект — из актов этого помещения, лист — от самых похожих работ
        docs_src = sorted(room_acts, key=lambda r: -sim(work, work_part(r["works"])))
        # лист — по ведомости альбома ЭТОГО помещения: на каком листе нарисован этот вид работ.
        # Прошлые акты — только если помещения нет в проекте или в альбоме не нашлось подходящего листа.
        same = next((r for r in docs_src if r.get("docs") and sim(work, work_part(r["works"])) >= 0.5), None)
        docs = self.project_docs(room, works) or (same or {}).get("docs") \
            or next((r["docs"] for r in docs_src if r.get("docs")), tpl.get("docs", ""))
        ends = [to_date(r.get("end", "")) for r in room_acts]
        ends = [d for d in ends if d]
        start = (max(ends) + timedelta(days=1)).strftime("%d.%m.%Y") if ends else ""
        return {"act_no": self.next_act_no(), "works": works, "docs": docs,
                "materials_text": tpl.get("materials", ""), "next_works": tpl.get("next_works", ""),
                "start": start, "end": "", "template": tpl.get("act_no", "")}

    def all_materials(self):
        out, seen = [], set()
        for r in self.recs:
            for m in split_materials(r["materials"]):
                if m["name"].lower() not in seen:
                    seen.add(m["name"].lower())
                    out.append(m)
        return out


def import_folder(folder):
    recs = []
    for p in sorted(Path(folder).rglob("*.docx")):
        if p.name.startswith("~$"):
            continue
        r = parse_act(p)
        if r:
            recs.append(r)
    return recs
