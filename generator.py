"""Формирование акта: правит шаблон template.docx, сохраняя его оформление."""
import copy
import re
import sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.text.paragraph import Paragraph

FONT = "Times New Roman"
SIZE, SMALL = Pt(10), Pt(8)

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
        r.underline = f["u"] if "u" in f else None


def text_width(doc):
    s = doc.sections[0]
    return s.page_width - s.left_margin - s.right_margin


def add_tab(p, pos, align, leader=WD_TAB_LEADER.SPACES):
    p.paragraph_format.tab_stops.add_tab_stop(int(pos), align, leader)


def right_tab(p, width):
    p.paragraph_format.tab_stops.clear_all()
    p.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)


def to_edge(p, width):
    """Подчёркивание значения тянется табуляцией до правого поля."""
    p.paragraph_format.tab_stops.clear_all()
    add_tab(p, width, WD_TAB_ALIGNMENT.RIGHT)
    r = p.add_run("\t")
    r.underline = True


def bottom_line(p):
    """Линия на всю ширину — нижняя граница абзаца."""
    bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    for k, v in (("val", "single"), ("sz", "4"), ("space", "1"), ("color", "auto")):
        bottom.set(qn("w:" + k), v)
    bdr.append(bottom)
    pPr = p._p.get_or_add_pPr()
    before = {qn("w:" + t) for t in ("pStyle", "keepNext", "keepLines", "pageBreakBefore",
                                     "framePr", "widowControl", "numPr", "suppressLineNumbers")}
    for i, ch in enumerate(pPr):                 # pBdr должен стоять на своём месте по схеме
        if ch.tag not in before:
            pPr.insert(i, bdr)
            return
    pPr.append(bdr)


def is_caption(p):
    t = p.text.strip().lower()
    return bool(t) and t.startswith(CAPTION_STARTS)


def apply_fonts(doc):
    """Times New Roman; 11 пт везде, подстрочники 9 пт."""
    st = doc.styles["Normal"]
    st.font.name, st.font.size = FONT, SIZE
    for p in doc.paragraphs:
        cap = is_caption(p)
        size = SMALL if cap else SIZE
        pPr = p._p.pPr
        if pPr is not None and pPr.find(qn("w:pBdr")) is not None:
            pPr.remove(pPr.find(qn("w:pBdr")))      # линии-подчёркивания абзаца заменены подчёркиванием значения
        for r in p.runs:
            if cap:
                r.italic = True
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
    IU = {"i": True, "u": True}      # заполняемое значение: курсив + подчёркивание
    W = text_width(doc)
    edge = []                        # абзацы, где подчёркивание тянется до правого поля

    # --- строка «город … дата» (поля остаются как в шаблоне, выравнивание — по табуляции)
    p = doc.paragraphs[find(doc, "г. ")]
    set_runs(p, [(re.sub(r"\s{3,}", "\t", p.text.strip()), {})])
    p.alignment = 0
    right_tab(p, text_width(doc))

    # --- номер акта
    set_runs(P("АКТ"), [("АКТ № ", {"b": True}), (data["act_no"], {"b": True})])

    # --- объект и организация: подчёркивается только значение после двоеточия
    p = P("Объект:")
    label, _, value = p.text.partition(": ")
    set_runs(p, [(label + ": ", {}), (value, IU)])
    p.alignment = 3                  # по ширине
    edge.append(p)
    p = P("произвели осмотр")
    label, _, value = p.text.partition(": ")
    set_runs(p, [(label + ": ", {"b": True}), (value.strip(), IU)])
    edge.append(p)

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
                set_runs(anchor, [(f'Представитель {r["role"]}: ', {"b": True}), (who, IU)])
                edge.append(anchor)
            else:
                set_runs(anchor, [("(организация, должность, фамилия, инициалы)", {})])

    # --- работы, документация
    p = P("К освидетельствованию")
    set_runs(p, [("К освидетельствованию предъявлены работы: ", {}), (data["works"], IU)])
    edge.append(p)
    p = P("Работы выполнены по")
    set_runs(p, [("Работы выполнены по проектной документации: ", {}), (data["docs"], IU)])
    edge.append(p)

    # --- материалы
    p = P("3. При выполнении")
    set_runs(p, [("3. При выполнении работ применены: ", {}), (" ".join(data["materials"]), IU)])
    edge.append(p)

    # --- отклонения (значение подчёркивается)
    i = find(doc, "4. При выполнении")
    p = doc.paragraphs[i + 1]
    set_runs(p, [(p.text.strip(), {"i": True, "u": False})])
    dev_p = p

    # --- даты
    i = find(doc, "Даты:")
    d1, m1, y1 = split_date(data["start"])
    d2, m2, y2 = split_date(data["end"])
    ps = doc.paragraphs
    num_ref = ps[find(doc, "3. При выполнении")].paragraph_format
    pf = ps[i].paragraph_format
    numPr = ps[i]._p.pPr.find(qn("w:numPr"))
    if numPr is not None:                       # автонумерация давала «1.» — пункт должен быть 5
        ps[i]._p.pPr.remove(numPr)
    pf.left_indent, pf.first_line_indent = num_ref.left_indent, num_ref.first_line_indent
    set_runs(ps[i], [("5. Даты: начала работ  ", {}), (f"«{d1}» {m1} {y1}г.", {"i": True})])
    set_runs(ps[i + 1], [("\t                окончание работ ", {}), (f"«{d2}» {m2} {y2}г.", {"i": True})])

    # --- последующие работы
    i = find(doc, "На основании изложенного")
    ps = doc.paragraphs
    set_runs(ps[i], [("На основании изложенного разрешается производство последующих работ "
                      "по устройству (монтажу)", {})])
    set_runs(ps[i + 1], [(data["next_works"], {"i": True})])
    next_p = ps[i + 1]

    # --- подписи (низ)
    s = find(doc, "Представитель", i + 1)
    ps = doc.paragraphs
    last = ps[-1]
    proto_sig = copy.deepcopy(ps[s]._p)
    proto_sigcap = copy.deepcopy(ps[s + 1]._p)
    for el in [ps[k]._p for k in range(s, len(ps) - 1)]:
        el.getparent().remove(el)
    S0, S1 = W * 0.47, W * 0.72                  # начало и конец линии подписи
    C1, C2 = (S0 + S1) / 2, (S1 + W) / 2         # центры подстрочников
    anchor = None

    def new_par(kind_proto):
        nonlocal anchor
        el = copy.deepcopy(kind_proto)
        if anchor is None:
            last._p.addprevious(el)
        else:
            anchor._p.addnext(el)
        anchor = Paragraph(el, last._parent)
        anchor.alignment = 0
        anchor.paragraph_format.tab_stops.clear_all()
        return anchor

    for r in reps:
        label = f'Представитель {r["role"]}' + (f' {r["org"]}' if r["org"] else "")
        if len(label) > 46:                       # длинная подпись — отдельной строкой
            set_runs(new_par(proto_sig), [(label, {})])
            p = new_par(proto_sig)
            parts = []
        else:
            p = new_par(proto_sig)
            parts = [(label, {})]
        add_tab(p, S0, WD_TAB_ALIGNMENT.LEFT)
        add_tab(p, S1, WD_TAB_ALIGNMENT.LEFT, WD_TAB_LEADER.LINES)
        add_tab(p, C2, WD_TAB_ALIGNMENT.CENTER)
        set_runs(p, parts + [("\t\t/\t", {}), (r["name"], {"u": True})])
        p = new_par(proto_sigcap)
        add_tab(p, C1, WD_TAB_ALIGNMENT.CENTER)
        add_tab(p, C2, WD_TAB_ALIGNMENT.CENTER)
        set_runs(p, [("\t(подпись)\t(расшифровка подписи)", {})])
        set_runs(new_par(proto_sigcap), [("", {})])

    apply_fonts(doc)
    for p in edge:
        to_edge(p, W)
    dev_p.paragraph_format.space_after = Pt(0)
    bottom_line(dev_p)
    bottom_line(next_p)
    doc.save(str(out_path))
