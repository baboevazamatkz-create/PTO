"""Формирование акта: правит шаблон template.docx, сохраняя его оформление."""
import copy
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]


def base_dir():
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent))


def split_date(s):
    """'19.10.2026' -> ('19', 'октября', '2026'); пусто -> подчёркивания."""
    try:
        d, m, y = [x.strip() for x in s.strip().split(".")]
        return str(int(d)), MONTHS[int(m) - 1], y
    except Exception:
        return "____", "____________", "2026"


def find(doc, prefix, start=0):
    ps = doc.paragraphs
    for i in range(start, len(ps)):
        if ps[i].text.strip().startswith(prefix):
            return i
    raise ValueError(f"В шаблоне не найден абзац «{prefix}»")


def font_size(p, default=None):
    for r in p.runs:
        if r.font.size:
            return r.font.size
    return default


def set_runs(p, parts, size=None):
    """Заменяет содержимое абзаца. parts: [(текст, {'b','i','u'})]"""
    size = size or font_size(p)
    for ch in list(p._p):
        if ch.tag != qn("w:pPr"):
            p._p.remove(ch)
    for text, f in parts:
        r = p.add_run(text)
        if size:
            r.font.size = size
        if f.get("b"):
            r.bold = True
        if f.get("i"):
            r.italic = True
        if f.get("u"):
            r.underline = True


def clone_after(ref_p, proto_p):
    new = copy.deepcopy(proto_p._p)
    ref_p._p.addnext(new)
    return Paragraph(new, ref_p._parent)


def text_width(doc):
    s = doc.sections[0]
    return s.page_width - s.left_margin - s.right_margin


def labeled(p, label, value, vfmt=None):
    set_runs(p, [(label + " ", {}), (value, vfmt or {})])


def generate(data, out_path, template=None):
    doc = Document(str(template or base_dir() / "template.docx"))
    P = lambda prefix, start=0: doc.paragraphs[find(doc, prefix, start)]

    # --- номер акта
    p = P("АКТ")
    set_runs(p, [("АКТ № " + data["act_no"], {"b": True})])

    # --- город и дата
    p = doc.paragraphs[find(doc, "г. ")]
    d, m, y = split_date(data["act_date"])
    p.alignment = 0
    p.paragraph_format.tab_stops.clear_all()
    p.paragraph_format.tab_stops.add_tab_stop(text_width(doc), WD_TAB_ALIGNMENT.RIGHT)
    set_runs(p, [(f'{data["city"]}\t«{d}» {m} {y}г.', {})])

    # --- объект
    p = P("Объект:")
    set_runs(p, [("Объект: «" + data["object"] + "»", {"i": True, "u": True})])

    # --- представители (верх)
    reps = [r for r in data["reps"] if r.get("on", True)]
    a = find(doc, "Мы, нижеподписавшиеся") + 1
    b = find(doc, "произвели осмотр")
    ps = doc.paragraphs
    proto_line = copy.deepcopy(ps[a]._p)         # абзац «представители Заказчика»
    proto_cap = copy.deepcopy(ps[a + 1]._p)      # подпись мелким шрифтом
    for el in [ps[k]._p for k in range(a, b)]:
        el.getparent().remove(el)
    anchor = P("Мы, нижеподписавшиеся")
    for r in reps:
        parent = anchor._parent
        for proto, kind in ((proto_line, "line"), (proto_cap, "cap")):
            el = copy.deepcopy(proto)
            anchor._p.addnext(el)
            anchor = Paragraph(el, parent)
            if kind == "line":
                who = " ".join(x for x in (r["org"], r["position"], r["name"]) if x)
                set_runs(anchor, [(f'Представитель {r["role"]}: ', {"b": True}), (who, {})], Pt(10))
            else:
                set_runs(anchor, [("(организация, должность, фамилия, инициалы)", {})], Pt(8))

    # --- СМО
    p = P("произвели осмотр")
    set_runs(p, [("произвели осмотр работ, выполненных: ", {"b": True}),
                 (data["contractor_org"], {})])

    # --- работы, документация
    p = P("К освидетельствованию")
    set_runs(p, [("К освидетельствованию предъявлены работы: ", {}), (data["works"], {})])
    p = P("Работы выполнены по")
    set_runs(p, [("Работы выполнены по проектной документации: ", {}),
                 (data["docs"], {"u": True})])

    # --- материалы
    p = P("3. При выполнении")
    items = []
    for name in data["materials"]:
        items.append(name)
    set_runs(p, [("3. При выполнении работ применены: ", {}), (" ".join(items), {})])

    # --- отклонения
    i = find(doc, "4. При выполнении")
    ps = doc.paragraphs
    set_runs(ps[i + 1], [(data["deviations"] if data["deviations_on"] else "Отсутствуют", {})])

    # --- даты
    i = find(doc, "Даты:")
    d1, m1, y1 = split_date(data["start"])
    d2, m2, y2 = split_date(data["end"])
    ps = doc.paragraphs
    set_runs(ps[i], [(f"Даты: начала работ  «{d1}» {m1} {y1}г.", {})])
    set_runs(ps[i + 1], [(f"\t                окончание работ «{d2}» {m2} {y2}г.", {})])

    # --- последующие работы
    i = find(doc, "На основании изложенного")
    ps = doc.paragraphs
    set_runs(ps[i + 1], [(data["next_works"], {"u": True})])

    # --- подписи (низ)
    s = find(doc, "Представитель", i + 1)
    ps = doc.paragraphs
    last = ps[-1]
    proto_sig = copy.deepcopy(ps[s]._p)
    cap_i = s + 1
    proto_sigcap = copy.deepcopy(ps[cap_i]._p)
    for el in [ps[k]._p for k in range(s, len(ps) - 1)]:
        el.getparent().remove(el)
    width = text_width(doc)
    anchor = None
    for r in reps:
        for proto, kind in ((proto_sig, "sig"), (proto_sigcap, "cap"), (proto_sigcap, "gap")):
            el = copy.deepcopy(proto)
            if anchor is None:
                last._p.addprevious(el)
                anchor = Paragraph(el, last._parent)
            else:
                anchor._p.addnext(el)
                anchor = Paragraph(el, last._parent)
            if kind == "sig":
                anchor.paragraph_format.tab_stops.clear_all()
                anchor.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
                org = (" " + r["org"]) if r["org"] else ""
                set_runs(anchor, [(f'Представитель {r["role"]}{org}', {}),
                                  (f'\t_______________ / {r["name"]}', {})], Pt(11))
            elif kind == "cap":
                anchor.paragraph_format.tab_stops.clear_all()
                anchor.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
                set_runs(anchor, [("\t(подпись)            (расшифровка подписи)", {})], Pt(8))
                anchor.alignment = 0
            else:
                set_runs(anchor, [("", {})], Pt(8))
    doc.save(str(out_path))
