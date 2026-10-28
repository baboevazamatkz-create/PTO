"""Формирование акта: правит шаблон template.docx, сохраняя его оформление."""
import copy
import re
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

FONT = "Times New Roman"
SIZE, SMALL = Pt(11), Pt(9)

MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]

CAPTION_STARTS = ("(", "наименование ", "должность, фамилия", "при наличии отклонений")


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


def set_runs(p, parts):
    """Заменяет содержимое абзаца. parts: [(текст, {'b','i','u'})]"""
    for ch in list(p._p):
        if ch.tag != qn("w:pPr"):
            p._p.remove(ch)
    for text, f in parts:
        r = p.add_run(text)
        r.bold = True if f.get("b") else None
        r.italic = True if f.get("i") else None
        r.underline = True if f.get("u") else None


def text_width(doc):
    s = doc.sections[0]
    return s.page_width - s.left_margin - s.right_margin


def right_tab(p, width):
    p.paragraph_format.tab_stops.clear_all()
    p.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)


def is_caption(p):
    t = p.text.strip().lower()
    return bool(t) and t.startswith(CAPTION_STARTS)


def apply_fonts(doc):
    """Times New Roman; 11 пт везде, подстрочники 9 пт."""
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, SIZE
    for p in doc.paragraphs:
        size = SMALL if is_caption(p) else SIZE
        for r in p.runs:
            r.font.name = FONT
            r.font.size = size
            rf = r._r.get_or_add_rPr().rFonts
            for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rf.set(qn(a), FONT)
            for a in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
                if rf.get(qn(a)) is not None:
                    del rf.attrib[qn(a)]


def generate(data, out_path, template=None):
    doc = Document(str(template or base_dir() / "template.docx"))
    P = lambda prefix, start=0: doc.paragraphs[find(doc, prefix, start)]
    I = {"i": True}

    # --- строка «город … дата» (поля остаются как в шаблоне, выравнивание — по табуляции)
    p = doc.paragraphs[find(doc, "г. ")]
    set_runs(p, [(re.sub(r"\s{3,}", "\t", p.text.strip()), {})])
    p.alignment = 0
    right_tab(p, text_width(doc))

    # --- номер акта
    set_runs(P("АКТ"), [("АКТ № ", {"b": True}), (data["act_no"], {"b": True, "i": True})])

    # --- представители (верх)
    reps = [r for r in data["reps"] if r.get("on", True)]
    a = find(doc, "Мы, нижеподписавшиеся") + 1
    b = find(doc, "произвели осмотр")
    ps = doc.paragraphs
    proto_line = copy.deepcopy(ps[a]._p)
    proto_cap = copy.deepcopy(ps[a + 1]._p)
    for el in [ps[k]._p for k in range(a, b)]:
        el.getparent().remove(el)
    anchor = P("Мы, нижеподписавшиеся")
    for r in reps:
        for proto, kind in ((proto_line, "line"), (proto_cap, "cap")):
            el = copy.deepcopy(proto)
            anchor._p.addnext(el)
            anchor = Paragraph(el, anchor._parent)
            if kind == "line":
                who = " ".join(x for x in (r["org"], r["position"], r["name"]) if x)
                set_runs(anchor, [(f'Представитель {r["role"]}: ', {"b": True}), (who, I)])
            else:
                set_runs(anchor, [("(организация, должность, фамилия, инициалы)", {})])

    # --- работы, документация
    set_runs(P("К освидетельствованию"),
             [("К освидетельствованию предъявлены работы: ", {}), (data["works"], I)])
    set_runs(P("Работы выполнены по"),
             [("Работы выполнены по проектной документации: ", {}), (data["docs"], {"i": True, "u": True})])

    # --- материалы
    set_runs(P("3. При выполнении"),
             [("3. При выполнении работ применены: ", {}), (" ".join(data["materials"]), I)])

    # --- даты
    i = find(doc, "Даты:")
    d1, m1, y1 = split_date(data["start"])
    d2, m2, y2 = split_date(data["end"])
    ps = doc.paragraphs
    set_runs(ps[i], [("Даты: начала работ  ", {}), (f"«{d1}» {m1} {y1}г.", I)])
    set_runs(ps[i + 1], [("\t                окончание работ ", {}), (f"«{d2}» {m2} {y2}г.", I)])

    # --- последующие работы
    i = find(doc, "На основании изложенного")
    ps = doc.paragraphs
    set_runs(ps[i], [("На основании изложенного разрешается производство последующих работ "
                      "по устройству (монтажу)", {})])
    set_runs(ps[i + 1], [(data["next_works"], {"i": True, "u": True})])

    # --- подписи (низ)
    s = find(doc, "Представитель", i + 1)
    ps = doc.paragraphs
    last = ps[-1]
    proto_sig = copy.deepcopy(ps[s]._p)
    proto_sigcap = copy.deepcopy(ps[s + 1]._p)
    for el in [ps[k]._p for k in range(s, len(ps) - 1)]:
        el.getparent().remove(el)
    width = text_width(doc)
    anchor = None
    for r in reps:
        for proto, kind in ((proto_sig, "sig"), (proto_sigcap, "cap"), (proto_sigcap, "gap")):
            el = copy.deepcopy(proto)
            if anchor is None:
                last._p.addprevious(el)
            else:
                anchor._p.addnext(el)
            anchor = Paragraph(el, last._parent)
            anchor.alignment = 0
            right_tab(anchor, width)
            if kind == "sig":
                org = (" " + r["org"]) if r["org"] else ""
                set_runs(anchor, [(f'Представитель {r["role"]}', {}), (org, I),
                                  ("\t_______________ / ", {}), (r["name"], I)])
            elif kind == "cap":
                set_runs(anchor, [("\t(подпись)            (расшифровка подписи)", {})])
            else:
                set_runs(anchor, [("", {})])

    apply_fonts(doc)
    doc.save(str(out_path))
