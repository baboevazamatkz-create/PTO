"""База знаний из прошлых актов: извлечение полей и связанные подсказки."""
import re
from pathlib import Path

from docx import Document

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
    rec = {"act_no": "", "works": "", "docs": "", "materials": "", "next_works": ""}
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
ROOM_RE = re.compile(r"помещени[еяй]\s*№?\s*(\d+)", re.I)
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
