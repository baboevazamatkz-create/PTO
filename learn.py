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


class Knowledge:
    def __init__(self, records):
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
                self.recs[i] = {**r, **{k: v for k, v in rec.items() if v}}
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
        return sorted((r for r in self.recs if room_of(r["works"]) == room), key=lambda r: act_num(r["act_no"]))

    def rooms(self):
        """{номер: подпись для списка}, например {'16': '16 — Русская парная, 2 этаж'}."""
        out = {}
        for r in self.recs:
            room = room_of(r["works"])
            if not room:
                continue
            name = next((m.group(1).strip(" «»") for m in ROOM_NAME_RE.finditer(r.get("docs", ""))
                         if m.group(2) == room), "")
            floor = re.search(r"\((\d+)\s*этаж", r["works"])
            label = room + (f" — {name}" if name else "") + (f", {floor.group(1)} этаж" if floor else "")
            if len(label) > len(out.get(room, "")):
                out[room] = label
        return dict(sorted(out.items(), key=lambda x: int(x[0])))

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
        if not acts:
            return ""
        last = acts[-1]
        return self.best_work(last.get("next_works", ""), exclude=work_part(last["works"])) \
            or last.get("next_works", "")

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

    def next_act_no(self):
        return str(max((act_num(r["act_no"]) for r in self.recs), default=0) + 1)

    def draft(self, room, work):
        """Черновик акта на помещение по виду работ — из самых похожих прошлых актов."""
        room_acts = self.room_acts(room)

        def score(r):
            wp = work_part(r["works"])
            s = sim(work, wp) + (0.5 if work and work.lower() in wp.lower() else 0)
            return s + (0.05 if room_of(r["works"]) == room else 0)   # помещение — только при равенстве

        tpl = max(self.recs, key=score, default=None)
        if not tpl:
            return {}
        loc = location_part(room_acts[-1]["works"]) if room_acts else location_part(tpl["works"])
        wp = work_part(tpl["works"]) if sim(work, work_part(tpl["works"])) > 0.3 else work
        works = (wp.rstrip(" .") + ". " + loc) if loc else wp
        # проект — из актов этого помещения, лист — от самых похожих работ
        docs_src = sorted(room_acts, key=lambda r: -sim(work, work_part(r["works"])))
        docs = next((r["docs"] for r in docs_src if r.get("docs")), tpl.get("docs", ""))
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
