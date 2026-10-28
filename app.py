"""Генератор актов освидетельствования скрытых работ (Windows, tkinter)."""
import json
import os
import re
import shutil
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, simpledialog, font as tkfont

from generator import base_dir, generate
from learn import Knowledge, import_folder, split_materials

if "__compiled__" in globals():                     # собрано Nuitka: папка, где лежит akt.exe
    APP_DIR = Path(sys.argv[0]).resolve().parent
elif getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent
# Данные — в скрытой папке .data рядом с программой: не видно, переносится вместе с .exe
DATA_DIR = APP_DIR / ".data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
if os.name == "nt":
    import ctypes
    ctypes.windll.kernel32.SetFileAttributesW(str(DATA_DIR), 0x2)   # FILE_ATTRIBUTE_HIDDEN
MAT_FILE = DATA_DIR / "materials.json"
SIGN_FILE = DATA_DIR / "signers.json"
LAST_FILE = DATA_DIR / "last_project.json"
KNOW_FILE = DATA_DIR / "knowledge.json"          # акты, которые программа изучила у пользователя

# данные прежних версий (рядом с программой или в %APPDATA%\AktsPTO) переносятся в .data
_OLD_DIRS = [APP_DIR, Path(os.environ.get("APPDATA") or Path.home()) / "AktsPTO"]
for _f in (MAT_FILE, SIGN_FILE, LAST_FILE, KNOW_FILE):
    for _d in _OLD_DIRS:
        _old = _d / _f.name
        if _old.exists() and not _f.exists():
            try:
                shutil.move(str(_old), str(_f))
            except OSError:
                pass

ROLES = ["Заказчика", "Подрядчика", "Технического надзора", "Авторского надзора",
         "Генподрядной организации", "Субподрядчика"]

# Палитра «Ivory»: тёплая бумага, чернила и золото-шампань
BG, CARD, FIELD, LINE, SOFT = "#F4EFE6", "#FBF9F5", "#FFFFFF", "#E4DDD0", "#F1EADD"
TEXT, MUTED, ACCENT, ACCENT_D = "#1B1813", "#8A8277", "#9A7B44", "#7E6334"
SAVED = "#44705B"                 # изумрудный из палитры Altyn
INK_ON, CHAMPAGNE, HOVER = "#F6F2EA", "#C8A86B", "#EFE3CB"
PAPER_TOP, PAPER_BOTTOM = "#F6F2EA", "#EDE6DA"
SERIF = "Georgia"


def spaced(text):
    """Подпись в стиле «микролейбла»: заглавные с разрядкой."""
    return "\u200a".join(text.upper())

DEFAULT = {
    "act_no": "", "works": "", "docs": "", "materials_text": "",
    "start": "", "end": "", "next_works": "",
    "reps": [],
}

DEFAULT_SIGNERS = [
    {"role": "Заказчика", "org": "", "position": "", "name": ""},
    {"role": "подрядчика", "org": "", "position": "", "name": ""},
    {"role": "технического надзора", "org": "", "position": "", "name": ""},
    {"role": "авторского надзора", "org": "", "position": "", "name": ""},
]


def rep_label(r):
    label = f'{r["role"].capitalize()}: {r["name"]}'
    return label + (f' — {r["org"]}' if r["org"] else "")


def load_json(path, default):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def save_json(path, obj):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def fmt_date(v):
    """Только цифры -> ДД.ММ.ГГГГ."""
    d = re.sub(r"\D", "", v)[:8]
    return ".".join(x for x in (d[:2], d[2:4], d[4:]) if x)


def date_entry(parent, var, **kw):
    """Поле даты: вводятся только цифры, точки расставляются сами, курсор не прыгает."""
    def fix():
        raw = var.get()
        new = fmt_date(raw)
        if new == raw:
            return
        n = len(re.sub(r"\D", "", raw[:e.index("insert")]))     # цифр слева от курсора
        var.set(new)
        pos, seen = 0, 0
        while pos < len(new) and seen < n:
            seen += new[pos].isdigit()
            pos += 1
        e.icursor(pos)

    e = ttk.Entry(parent, textvariable=var, **kw)
    var.trace_add("write", lambda *_: e.after_idle(fix))
    var.set(fmt_date(var.get()))
    return e


def cap(s):
    """Первая буква — заглавная."""
    s = s.strip()
    return s[:1].upper() + s[1:]


def auto_cap(t):
    """Text: первая буква автоматически заглавная."""
    def on_key(_):
        ch = t.get("1.0")
        if ch.islower():
            t.delete("1.0")
            t.insert("1.0", ch.upper())
    t.bind("<KeyRelease>", on_key, add="+")


def install_clipboard(root):
    """Ctrl+C/V/X/A работают и при русской раскладке, плюс меню по правой кнопке."""
    codes = {67: "<<Copy>>", 86: "<<Paste>>", 88: "<<Cut>>"}      # коды клавиш Windows

    def select_all(w):
        if isinstance(w, tk.Text):
            w.tag_add("sel", "1.0", "end-1c")
        elif hasattr(w, "select_range"):
            w.select_range(0, "end")
            w.icursor("end")

    def on_ctrl(e):
        if e.keysym.lower() in ("c", "v", "x", "a"):       # латиница — работает штатно
            return None
        w = e.widget
        if e.keycode == 65:
            select_all(w)
            return "break"
        if e.keycode in codes and w.winfo_class() in ("Entry", "TEntry", "TCombobox", "Text"):
            w.event_generate(codes[e.keycode])
            return "break"
        return None

    root.bind_all("<Control-KeyPress>", on_ctrl)

    menu = tk.Menu(root, tearoff=0)
    target = []
    for label, ev in (("Вырезать", "<<Cut>>"), ("Копировать", "<<Copy>>"), ("Вставить", "<<Paste>>")):
        menu.add_command(label=label, command=lambda ev=ev: target and target[0].event_generate(ev))
    menu.add_command(label="Выделить всё", command=lambda: target and select_all(target[0]))

    def popup(e):
        target[:] = [e.widget]
        e.widget.focus_set()
        menu.tk_popup(e.x_root, e.y_root)

    for cls in ("Entry", "TEntry", "TCombobox", "Text"):
        root.bind_class(cls, "<Button-3>", popup)


def setup_style(root):
    fam = "Segoe UI" if "Segoe UI" in tkfont.families() else "DejaVu Sans"
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        tkfont.nametofont(name).configure(family=fam, size=9)
    root.option_add("*Font", (fam, 9))
    root.configure(bg=BG)
    s = ttk.Style(root)
    s.theme_use("clam")
    s.configure(".", background=BG, foreground=TEXT, font=(fam, 9), borderwidth=0)
    s.configure("TFrame", background=BG)
    s.configure("Card.TFrame", background=CARD)
    s.configure("TLabel", background=BG, foreground=TEXT)
    s.configure("Muted.TLabel", foreground=MUTED, font=(fam, 9))
    s.configure("CardMuted.TLabel", background=CARD, foreground=MUTED, font=(fam, 8))
    for name in ("TEntry", "TCombobox"):
        s.configure(name, fieldbackground=FIELD, background=FIELD, bordercolor=LINE,
                    lightcolor=LINE, darkcolor=LINE, padding=5, relief="flat", arrowcolor=ACCENT,
                    selectbackground=HOVER, selectforeground=TEXT)
        s.map(name, bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)],
              darkcolor=[("focus", ACCENT)])
    s.map("TCombobox", fieldbackground=[("readonly", FIELD)],
          selectbackground=[("readonly", FIELD)], selectforeground=[("readonly", TEXT)],
          bordercolor=[("focus", ACCENT)])
    s.configure("TCheckbutton", background=CARD, focuscolor=CARD, indicatorbackground=FIELD,
                indicatorforeground=ACCENT, bordercolor=LINE, indicatorrelief="flat")
    s.map("TCheckbutton", background=[("active", CARD)],
          indicatorbackground=[("selected", FIELD)])
    # кнопки: вторичные — тонкая линия по контуру, главная — «чернильная»
    s.configure("TButton", background=CARD, foreground=TEXT, padding=(12, 5), relief="flat",
                borderwidth=1, bordercolor=LINE, lightcolor=CARD, darkcolor=CARD, focuscolor=CARD)
    s.map("TButton", background=[("active", HOVER)], bordercolor=[("active", ACCENT)])
    s.configure("Head.TButton", background=CARD, foreground=TEXT, padding=(12, 5), borderwidth=1,
                bordercolor=LINE, lightcolor=CARD, darkcolor=CARD, focuscolor=CARD)
    s.map("Head.TButton", background=[("active", HOVER)], bordercolor=[("active", ACCENT)],
          foreground=[("active", ACCENT_D)])
    for name in ("Accent.TButton", "Ok.TButton"):
        s.configure(name, background=TEXT, foreground=INK_ON, padding=(18, 8), borderwidth=0,
                    font=(fam, 9, "bold"), focuscolor=TEXT, lightcolor=TEXT, darkcolor=TEXT)
        s.map(name, background=[("active", ACCENT)], foreground=[("active", "#FFFFFF")])
    s.configure("Link.TButton", background=CARD, foreground=ACCENT, padding=(0, 2), borderwidth=0,
                focuscolor=CARD, lightcolor=CARD, darkcolor=CARD)
    s.map("Link.TButton", background=[("active", CARD)], foreground=[("active", ACCENT_D)])
    s.configure("Icon.TButton", background=CARD, foreground=MUTED, padding=(5, 3), borderwidth=0,
                focuscolor=CARD, lightcolor=CARD, darkcolor=CARD)
    s.map("Icon.TButton", background=[("active", HOVER)], foreground=[("active", ACCENT_D)])
    s.configure("Treeview", background=FIELD, fieldbackground=FIELD, rowheight=26,
                bordercolor=LINE, borderwidth=0)
    s.configure("Treeview.Heading", background=CARD, foreground=MUTED, font=(fam, 8, "bold"),
                relief="flat")
    s.map("Treeview", background=[("selected", HOVER)], foreground=[("selected", TEXT)])
    s.configure("Vertical.TScrollbar", background="#D9CFBE", troughcolor=BG, bordercolor=BG,
                lightcolor="#D9CFBE", darkcolor="#D9CFBE", arrowsize=0, relief="flat", width=9)
    s.map("Vertical.TScrollbar", background=[("active", CHAMPAGNE)])
    return fam


SCROLLERS = {}        # путь виджета -> Canvas, который он прокручивает


def scrollable(parent, bg):
    """Canvas с вертикальной прокруткой; возвращает (внешняя рамка, внутренняя рамка, canvas)."""
    outer = tk.Frame(parent, bg=bg)
    cv = tk.Canvas(outer, bg=bg, highlightthickness=0)
    sb = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
    cv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    cv.pack(side="left", fill="both", expand=True)
    inner = tk.Frame(cv, bg=bg)
    win = cv.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
    cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
    SCROLLERS[str(outer)] = cv
    return outer, inner, cv


def install_wheel(root):
    """Колесо мыши прокручивает ту область, над которой находится курсор."""
    def on_wheel(e, step=None):
        w = root.winfo_containing(e.x_root, e.y_root)
        if w is None:
            return
        path = str(w)
        best = max((p for p in SCROLLERS if path.startswith(p)), key=len, default=None)
        if not best:
            return
        cv = SCROLLERS[best]
        try:
            if cv.yview() == (0.0, 1.0):
                return
            cv.yview_scroll(step if step else -1 * int(e.delta / 120 or (1 if e.delta > 0 else -1)),
                            "units")
        except tk.TclError:
            SCROLLERS.pop(best, None)
        return "break"

    root.bind_all("<MouseWheel>", on_wheel)
    root.bind_all("<Button-4>", lambda e: on_wheel(e, -1))     # Linux
    root.bind_all("<Button-5>", lambda e: on_wheel(e, 1))
    # поля не листаются колесом (выпадающий список не меняет значение, текст не прокручивается) —
    # вместо этого прокручивается окно
    for cls in ("TCombobox", "Text", "TEntry", "TSpinbox"):
        root.bind_class(cls, "<MouseWheel>", on_wheel)
        root.bind_class(cls, "<Button-4>", lambda e: on_wheel(e, -1))
        root.bind_class(cls, "<Button-5>", lambda e: on_wheel(e, 1))


def text_field(parent, height, value=""):
    t = tk.Text(parent, height=height, wrap="word", relief="flat", bg=FIELD, fg=TEXT,
                spacing1=0, spacing3=0,
                selectbackground=HOVER, selectforeground=TEXT,
                insertbackground=TEXT, padx=6, pady=3, highlightthickness=1,
                highlightbackground=LINE, highlightcolor=ACCENT, undo=True)
    t.insert("1.0", value)
    auto_cap(t)
    return t


def mat_text(m):
    return (f'{m["name"]} – сертификат соответствия № {m["cert"]}.' if m["cert"]
            else m["name"].rstrip(".") + ".")


class Suggest:
    """Выпадающие подсказки под текстовым полем.

    provider(typed) -> список строк. Выбор заменяет текст поля целиком.
    """

    NAV = {"Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R",
           "Control_L", "Control_R", "Alt_L", "Alt_R", "Left", "Right"}

    all = []

    @classmethod
    def install(cls, root):
        """Клик мышью в любом другом месте закрывает подсказки."""
        def on_click(e):
            for sg in cls.all:
                if sg.top and not str(e.widget).startswith((str(sg.top), str(sg.w))):
                    sg.hide()
        root.bind_all("<Button-1>", on_click, add="+")
        root.bind("<Configure>", lambda e: e.widget is root and [sg.hide() for sg in cls.all],
                  add="+")

    def __init__(self, widget, provider, on_pick=None):
        Suggest.all.append(self)
        self.w, self.provider, self.on_pick = widget, provider, on_pick
        self.top = None
        self.items, self.labels, self.idx = [], [], -1
        self.picked = None                       # текст, выбранный из подсказки
        widget.bind("<KeyRelease>", self.on_key, add="+")
        widget.bind("<FocusIn>", lambda e: widget.after(80, self.show), add="+")
        widget.bind("<FocusOut>", lambda e: widget.after(200, self.hide_if_unfocused), add="+")
        widget.bind("<Down>", lambda e: self.move(1), add="+")
        widget.bind("<Up>", lambda e: self.move(-1) if self.top else None, add="+")
        widget.bind("<Return>", self.on_enter, add="+")
        widget.bind("<Escape>", lambda e: self.hide(), add="+")

    def text(self):
        return self.w.get("1.0", "end").strip()

    def on_key(self, e):
        if e.keysym not in self.NAV:
            self.picked = None
            self.show()

    def on_enter(self, e):
        if self.top and self.idx >= 0:
            self.pick(self.idx)
            return "break"
        return None

    def move(self, d):
        if not self.top:
            self.show()
            return "break"
        self.idx = max(0, min(len(self.items) - 1, self.idx + d))
        self.paint()
        return "break"

    def show(self):
        if self.w.focus_get() is not self.w:
            return
        typed = self.text()
        items = [x for x in self.provider(typed) if x.strip() and x.strip() != typed]
        if not items or typed == self.picked:
            self.hide()
            return
        self.items, self.idx = items[:15], -1
        if not self.top:
            self.top = tk.Toplevel(self.w)
            self.top.wm_overrideredirect(True)
            self.top.configure(bg=CHAMPAGNE)
            try:
                self.top.attributes("-alpha", 0.97)
            except tk.TclError:
                pass
        for c in self.top.winfo_children():
            c.destroy()
        width = self.w.winfo_width()
        tk.Label(self.top, text=spaced("Подсказки из прошлых актов") + "      ↑↓  Enter  ·  Esc",
                 bg=CARD, fg=ACCENT, anchor="w", padx=10, pady=4,
                 font=("Segoe UI", 7, "bold")).pack(fill="x", padx=1, pady=(1, 0))
        self.labels = []
        outer, box, self.cv = scrollable(self.top, LINE)
        outer.pack(fill="both", expand=True, padx=1, pady=(0, 1))
        for i, it in enumerate(self.items):
            lb = tk.Label(box, text=it, bg=FIELD, fg=TEXT, anchor="w", justify="left",
                          wraplength=width - 44, padx=10, pady=5, cursor="hand2")
            lb.pack(fill="x", pady=(0, 1))
            lb.bind("<Button-1>", lambda e, i=i: self.pick(i))
            lb.bind("<Enter>", lambda e, i=i: (setattr(self, "idx", i), self.paint()))
            self.labels.append(lb)
        box.update_idletasks()
        self.cv.configure(height=min(box.winfo_reqheight(), 260))
        self.top.update_idletasks()
        x = self.w.winfo_rootx()
        y = self.w.winfo_rooty() + self.w.winfo_height()
        self.top.geometry(f"{width}x{self.top.winfo_reqheight()}+{x}+{y}")
        self.top.lift()
        if self.w.focus_get() is not self.w:
            self.w.focus_set()

    def paint(self):
        if self.idx >= 0 and self.labels:              # держать выделенную строку в видимой области
            lb = self.labels[self.idx]
            total = max(lb.master.winfo_height(), 1)
            top, bot = self.cv.yview()
            y0, y1 = lb.winfo_y() / total, (lb.winfo_y() + lb.winfo_height()) / total
            if y0 < top:
                self.cv.yview_moveto(y0)
            elif y1 > bot:
                self.cv.yview_moveto(y1 - (bot - top))
        for i, lb in enumerate(self.labels):
            lb.configure(bg=HOVER if i == self.idx else FIELD)

    def pick(self, i):
        val = self.items[i]
        self.w.delete("1.0", "end")
        self.w.insert("1.0", val)
        self.picked = val
        self.hide()
        self.w.focus_set()
        if self.on_pick:
            self.on_pick(val)

    def hide_if_unfocused(self):
        if self.w.focus_get() is not self.w:
            self.hide()

    def hide(self):
        if self.top:
            self.top.destroy()
            self.top = None


def filter_typed(items, typed):
    """Если пользователь уже что-то ввёл — оставить подсказки, содержащие его слова."""
    words = [w.lower() for w in re.findall(r"\w{2,}", typed)]
    if not words:
        return items
    return [x for x in items if all(w in x.lower() for w in words)] or \
        [x for x in items if any(w in x.lower() for w in words)]


class MaterialsDialog(tk.Toplevel):
    """Окно выбора материалов из базы."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Материалы")
        self.configure(bg=BG)
        self.geometry("820x520")
        self.transient(app)
        self.grab_set()
        app.chosen = set()
        top = ttk.Frame(self, padding=(16, 14, 16, 4))
        top.pack(fill="x")
        tk.Label(top, text="База материалов", bg=BG, fg=TEXT, font=(SERIF, 15, "bold")).pack(side="left")
        self.q = tk.StringVar()
        self.q.trace_add("write", lambda *_: self.refresh())
        ttk.Entry(top, textvariable=self.q, width=28).pack(side="right")
        ttk.Label(top, text="Поиск", style="Muted.TLabel").pack(side="right", padx=8)
        ttk.Label(self, text="Нажмите на строку, чтобы отметить материал для акта.",
                  style="Muted.TLabel").pack(anchor="w", padx=16)
        mid = ttk.Frame(self, padding=(16, 8))
        mid.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(mid, columns=("on", "name", "cert"), show="headings",
                                 selectmode="browse")
        for c, t, w, st in (("on", "", 40, False), ("name", "Материал", 420, True),
                            ("cert", "Сертификат №", 200, False)):
            self.tree.heading(c, text=t, anchor="w")
            self.tree.column(c, width=w, stretch=st, anchor="w")
        sb = ttk.Scrollbar(mid, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<ButtonRelease-1>", self.toggle)
        bar = ttk.Frame(self, padding=(16, 4, 16, 14))
        bar.pack(fill="x")
        for t, c in (("Добавить", self.add), ("Изменить", self.edit), ("Удалить", self.delete),
                     ("Импорт из файла", self.imp), ("Из прошлых актов", self.from_acts)):
            ttk.Button(bar, text=t, command=c).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="Вставить отмеченные", style="Accent.TButton",
                   command=self.insert).pack(side="right")
        self.refresh()

    def refresh(self):
        q = self.q.get().strip().lower()
        self.tree.delete(*self.tree.get_children())
        for i, m in enumerate(self.app.mats):
            if q and q not in (m["name"] + m["cert"]).lower():
                continue
            self.tree.insert("", "end", iid=str(i), values=(
                "✔" if m["name"] in self.app.chosen else "", m["name"], m["cert"]))
        save_json(MAT_FILE, self.app.mats)

    def cur(self):
        s = self.tree.selection()
        return int(s[0]) if s else None

    def insert(self):
        """Отмеченные материалы дописываются в поле «применены»."""
        add = " ".join(mat_text(m) for m in self.app.mats if m["name"] in self.app.chosen)
        if add:
            box = self.app.materials
            cur = box.get("1.0", "end").strip()
            box.delete("1.0", "end")
            box.insert("1.0", (cur + " " + add).strip())
        self.destroy()

    def from_acts(self):
        have = {m["name"].lower() for m in self.app.mats}
        new = [m for m in self.app.know.all_materials() if m["name"].lower() not in have]
        self.app.mats += new
        self.refresh()
        messagebox.showinfo("Материалы", f"Добавлено из прошлых актов: {len(new)}", parent=self)

    def toggle(self, ev):
        if self.tree.identify_region(ev.x, ev.y) != "cell":
            return
        i = self.cur()
        if i is not None:
            self.app.chosen ^= {self.app.mats[i]["name"]}
            self.refresh()
            self.tree.selection_set(str(i))

    def ask(self, name="", cert=""):
        n = simpledialog.askstring("Материал", "Наименование:", initialvalue=name, parent=self)
        if not n:
            return None
        c = simpledialog.askstring("Материал", "Номер сертификата соответствия:",
                                   initialvalue=cert, parent=self)
        return {"name": n.strip(), "cert": (c or "").strip()}

    def add(self):
        m = self.ask()
        if m:
            self.app.mats.append(m)
            self.refresh()

    def edit(self):
        i = self.cur()
        if i is not None:
            old = self.app.mats[i]
            m = self.ask(**old)
            if m:
                if old["name"] in self.app.chosen:
                    self.app.chosen.discard(old["name"])
                    self.app.chosen.add(m["name"])
                self.app.mats[i] = m
                self.refresh()

    def delete(self):
        i = self.cur()
        if i is not None:
            self.app.chosen.discard(self.app.mats[i]["name"])
            del self.app.mats[i]
            self.refresh()

    def clear(self):
        self.app.chosen.clear()
        self.refresh()

    def imp(self):
        p = filedialog.askopenfilename(parent=self, title="Файл: Наименование;Сертификат",
                                       filetypes=[("Текст/CSV", "*.txt *.csv"), ("Все", "*.*")])
        if not p:
            return
        raw = Path(p).read_bytes()
        try:
            text = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw.decode("cp1251")
        n = 0
        for line in text.splitlines():
            if line.strip():
                name, _, cert = line.partition(";")
                self.app.mats.append({"name": name.strip(), "cert": cert.strip()})
                n += 1
        self.refresh()
        messagebox.showinfo("Импорт", f"Добавлено: {n}", parent=self)


class SignersDialog(tk.Toplevel):
    """База подписантов: правка прямо в строках таблицы."""

    COLS = (("role", "Роль", 22), ("org", "Организация", 28), ("position", "Должность", 30),
            ("name", "Ф.И.О.", 20))

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("Подписанты")
        self.configure(bg=BG)
        self.geometry("980x480")
        self.transient(app)
        self.grab_set()
        tk.Label(self, text="Подписанты", bg=BG, fg=TEXT, font=(SERIF, 15, "bold")).pack(
            anchor="w", padx=16, pady=(12, 0))
        ttk.Label(self, text="Пишите прямо в полях. На главном экране подписант выбирается из списка.",
                  style="Muted.TLabel").pack(anchor="w", padx=16, pady=(0, 6))
        bar = ttk.Frame(self, padding=(16, 6, 16, 12))
        bar.pack(side="bottom", fill="x")
        ttk.Button(bar, text="＋ Добавить подписанта", command=self.add).pack(side="left")
        ttk.Button(bar, text="Сохранить", style="Accent.TButton", command=self.close).pack(side="right")
        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=(16, 32))
        for c, (k, t, w) in enumerate(self.COLS):
            head.columnconfigure(c, weight=w)
            tk.Label(head, text=t, bg=BG, fg=MUTED, font=("Segoe UI", 8, "bold"), anchor="w").grid(
                row=0, column=c, sticky="ew", padx=3)
        head.columnconfigure(len(self.COLS), minsize=30)
        outer, self.grid_box, _ = scrollable(self, BG)
        outer.pack(fill="both", expand=True, padx=(16, 16))
        self.rows = []
        for r in app.signers:
            self.add_row(r)
        self.protocol("WM_DELETE_WINDOW", self.close)

    def add_row(self, rep=None):
        rep = rep or {"role": "", "org": "", "position": "", "name": ""}
        f = tk.Frame(self.grid_box, bg=BG)
        f.pack(fill="x", pady=2)
        row = {"frame": f, "orig": dict(rep)}
        for c, (k, t, w) in enumerate(self.COLS):
            f.columnconfigure(c, weight=w)
            v = tk.StringVar(value=rep.get(k, ""))
            v.trace_add("write", lambda *_, v=v: v.get()[:1].islower()
                        and v.set(v.get()[:1].upper() + v.get()[1:]))
            w_ = (ttk.Combobox(f, textvariable=v, values=ROLES) if k == "role"
                  else ttk.Entry(f, textvariable=v))
            w_.grid(row=0, column=c, sticky="ew", padx=3)
            row[k] = v
        ttk.Button(f, text="✕", width=2, style="Icon.TButton",
                   command=lambda: self.remove(row)).grid(row=0, column=len(self.COLS))
        self.rows.append(row)
        return row

    def add(self):
        row = self.add_row()
        row["frame"].winfo_children()[0].focus_set()

    def remove(self, row):
        row["frame"].destroy()
        self.rows.remove(row)

    def close(self):
        new = []
        for row in self.rows:
            r = {k: cap(row[k].get()) for k, *_ in self.COLS}
            if not (r["name"] or r["org"]):
                continue
            if row["orig"].get("name") or row["orig"].get("org"):
                if row["orig"] != r:              # выбранные на главном экране строки обновятся
                    self.app.replace_rep(row["orig"], r)
            new.append(r)
        self.app.signers[:] = new
        save_json(SIGN_FILE, new)
        self.app.refresh_signer_lists()
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Акты скрытых работ")
        try:
            self.iconbitmap(default=str(base_dir() / "akt.ico"))
        except tk.TclError:
            pass
        self.geometry("860x680")
        self.minsize(760, 560)
        setup_style(self)
        install_clipboard(self)
        self.mats = load_json(MAT_FILE, [])
        self.signers = [{k: cap(v) for k, v in r.items()} for r in load_json(SIGN_FILE, DEFAULT_SIGNERS)]
        data = {**DEFAULT, **load_json(LAST_FILE, {})}
        if not data["reps"]:
            data["reps"] = [{"on": True, "role": "", "org": "", "position": "", "name": ""}]
        self.chosen = set()
        seed = load_json(base_dir() / "knowledge_seed.json", [])
        self.know = Knowledge(seed + load_json(KNOW_FILE, []))
        if data.get("materials_sel") and not data.get("materials_text"):
            data["materials_text"] = self.old_materials(data["materials_sel"])
        self.rows = []          # [{"on": BooleanVar, "var": StringVar, "rep": dict, "cb": Combobox}]
        self.reps_data = [dict(r) for r in data["reps"]]
        self.build(data)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- построение окна
    def build(self, data):
        head = tk.Canvas(self, height=50, highlightthickness=0, bg=BG)
        head.pack(fill="x")

        def paint_head(e):
            head.delete("grad")
            c1 = [int(PAPER_TOP[i:i + 2], 16) for i in (1, 3, 5)]
            c2 = [int(BG[i:i + 2], 16) for i in (1, 3, 5)]
            for y in range(0, 50, 2):            # мягкая вертикальная «бумага»
                t = y / 50
                col = "#%02x%02x%02x" % tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))
                head.create_rectangle(0, y, e.width, y + 2, fill=col, outline=col, tags="grad")
            head.create_line(16, 49, e.width - 16, 49, fill=LINE, tags="grad")
            head.tag_lower("grad")
            x = e.width - 16
            for b, win in reversed(btn_wins):
                head.coords(win, x, 25)
                x -= b.winfo_reqwidth() + 6
        head.bind("<Configure>", paint_head)
        head.create_text(20, 19, text="А К Т Ы", anchor="w", fill=ACCENT,
                         font=(SERIF, 14, "bold"))
        head.create_text(21, 37, text=spaced("освидетельствование скрытых работ"), anchor="w",
                         fill=MUTED, font=("Segoe UI", 7, "bold"))
        btn_wins = []
        for t, c in (("Импорт актов", self.import_acts), ("Подписанты", lambda: SignersDialog(self)),
                     ("Открыть", self.open_project), ("Сохранить", self.save_project)):
            b = ttk.Button(head, text=t, style="Head.TButton", command=c)
            btn_wins.append((b, head.create_window(0, 25, window=b, anchor="e")))

        foot = ttk.Frame(self, padding=(18, 6, 18, 8))
        foot.pack(side="bottom", fill="x")
        tk.Frame(self, height=1, bg=LINE).pack(side="bottom", fill="x")
        foot.configure(style="TFrame")
        ttk.Button(foot, text="Сформировать акт (Word)", style="Ok.TButton",
                   command=self.make).pack(side="right")
        self.saved_lbl = tk.Label(foot, text="", bg=BG, fg=SAVED, font=("Segoe UI", 10, "bold"))
        self.saved_lbl.pack(side="right", padx=12)

        wrap, body, cv = scrollable(self, BG)
        wrap.pack(fill="both", expand=True)
        body.configure(padx=16, pady=8)
        install_wheel(self)
        body.columnconfigure(0, weight=1)

        r = 0

        num = [0]

        def section(title, tint=None):
            """Карточка: номер раздела золотом и подпись-микролейбл."""
            nonlocal r
            num[0] += 1
            card = tk.Frame(body, bg=CARD, highlightthickness=1, highlightbackground=LINE,
                            highlightcolor=CHAMPAGNE)
            card.grid(row=r, column=0, sticky="ew", pady=(0, 6))
            r += 1
            inner = tk.Frame(card, bg=CARD, padx=14, pady=6)
            inner.pack(fill="both", expand=True)
            cap_ = tk.Frame(inner, bg=CARD)
            cap_.pack(anchor="w", pady=(0, 3))
            tk.Label(cap_, text=f"{num[0]:02d}", bg=CARD, fg=ACCENT,
                     font=(SERIF, 11, "bold")).pack(side="left")
            tk.Label(cap_, text="   " + spaced(title), bg=CARD, fg=MUTED,
                     font=("Segoe UI", 7, "bold")).pack(side="left")
            return inner

        # 1. номер акта
        c = section("Акт №")
        self.act_no = tk.StringVar(value=data["act_no"])
        ttk.Entry(c, textvariable=self.act_no, width=14).pack(anchor="w")

        # 2. представители
        c = section("Мы, нижеподписавшиеся — представители")
        self.reps_frame = ttk.Frame(c, style="Card.TFrame")
        self.reps_frame.pack(fill="x")
        ttk.Button(c, text="＋ Добавить представителя", style="Link.TButton",
                   command=self.rep_add).pack(anchor="w")
        self.rep_draw()

        # 3. работы
        c = section("К освидетельствованию предъявлены работы")
        self.works = text_field(c, 2, data["works"])
        self.works.pack(fill="x")

        c = section("Работы выполнены по проектной документации")
        self.docs = text_field(c, 2, data["docs"])
        self.docs.pack(fill="x")

        # 4. материалы
        c = section("При выполнении работ применены")
        self.materials = text_field(c, 2, data["materials_text"])
        self.materials.pack(fill="x")
        ttk.Button(c, text="＋ Выбрать из базы материалов", style="Link.TButton",
                   command=self.open_materials).pack(anchor="w")

        # 5. даты
        c = section("Даты")
        dates = ttk.Frame(c, style="Card.TFrame")
        dates.pack(anchor="w")
        self.start = tk.StringVar(value=data["start"])
        self.end = tk.StringVar(value=data["end"])
        ttk.Label(dates, text="начало работ", style="CardMuted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(dates, text="окончание работ", style="CardMuted.TLabel").grid(
            row=0, column=1, sticky="w", padx=(12, 0))
        date_entry(dates, self.start, width=14).grid(row=1, column=0)
        date_entry(dates, self.end, width=14).grid(row=1, column=1, padx=(12, 0))

        # 6. последующие работы
        c = section("Разрешается производство последующих работ")
        self.next_works = text_field(c, 1, data["next_works"])
        self.next_works.pack(fill="x")

        # подсказки: всё, кроме самих работ, подбирается по введённым работам
        k, works = self.know, lambda: self.works.get("1.0", "end")
        Suggest(self.works, lambda t: k.suggest_works(t))
        Suggest(self.docs, lambda t: filter_typed(k.suggest_docs(works()), t))
        Suggest(self.materials, lambda t: filter_typed(
            k._uniq([r["materials"] for r in k.related(works())], 10), t))
        Suggest(self.next_works, lambda t: filter_typed(k.suggest_next(works()), t))
        Suggest.install(self)

    # ---------- представители
    def signer_labels(self):
        return [rep_label(r) for r in self.signers]

    def rep_draw(self):
        for w in self.reps_frame.winfo_children():
            w.destroy()
        self.rows = []
        f = self.reps_frame
        f.columnconfigure(1, weight=1)
        for i, rep in enumerate(self.reps_data):
            row = {"on": tk.BooleanVar(value=rep.get("on", True)), "rep": dict(rep),
                   "var": tk.StringVar(value=rep_label(rep) if rep.get("role") else "")}
            self.rows.append(row)
            ttk.Checkbutton(f, variable=row["on"]).grid(row=i, column=0, padx=(0, 6), pady=2)
            cb = ttk.Combobox(f, textvariable=row["var"], values=self.signer_labels(),
                              state="readonly")
            cb.grid(row=i, column=1, sticky="ew", pady=2)
            cb.bind("<<ComboboxSelected>>", lambda e, row=row: self.rep_pick(row))
            row["cb"] = cb
            ttk.Button(f, text="↑", width=2, style="Icon.TButton",
                       command=lambda i=i: self.rep_move(i, -1)).grid(row=i, column=2)
            ttk.Button(f, text="↓", width=2, style="Icon.TButton",
                       command=lambda i=i: self.rep_move(i, 1)).grid(row=i, column=3)
            ttk.Button(f, text="✕", width=2, style="Icon.TButton",
                       command=lambda i=i: self.rep_del(i)).grid(row=i, column=4)

    def rep_pick(self, row):
        i = row["cb"].current()
        if i >= 0:
            row["rep"] = dict(self.signers[i])

    def rep_read(self):
        self.reps_data = [{"on": r["on"].get(), **{k: r["rep"].get(k, "") for k in
                           ("role", "org", "position", "name")}} for r in self.rows]

    def rep_add(self):
        self.rep_read()
        self.reps_data.append({"on": True, "role": "", "org": "", "position": "", "name": ""})
        self.rep_draw()

    def rep_del(self, i):
        self.rep_read()
        del self.reps_data[i]
        self.rep_draw()

    def rep_move(self, i, d):
        self.rep_read()
        j = i + d
        if 0 <= j < len(self.reps_data):
            self.reps_data[i], self.reps_data[j] = self.reps_data[j], self.reps_data[i]
            self.rep_draw()

    def refresh_signer_lists(self):
        for r in self.rows:
            r["cb"].configure(values=self.signer_labels())

    def replace_rep(self, old, new):
        """Подписант изменён в базе — обновить уже выбранные строки."""
        for r in self.rows:
            if all(r["rep"].get(k) == old[k] for k in old):
                r["rep"] = dict(new)
                r["var"].set(rep_label(new))

    # ---------- материалы
    def old_materials(self, names):
        by = {m["name"]: m for m in self.mats}
        return " ".join(mat_text(by[n]) for n in names if n in by)

    def import_acts(self):
        folder = filedialog.askdirectory(title="Папка с актами (.docx, включая подпапки)")
        if not folder:
            return
        self.config(cursor="watch")
        self.update()
        try:
            recs = import_folder(folder)
        finally:
            self.config(cursor="")
        added = [r for r in recs if self.know.add(r)]
        save_json(KNOW_FILE, load_json(KNOW_FILE, []) + added)
        messagebox.showinfo("Импорт актов", f"Прочитано актов: {len(recs)}\nНовых в базе: {len(added)}")

    def learn(self, d):
        rec = {k: d[k] for k in ("act_no", "works", "docs", "next_works")}
        rec["materials"] = d["materials_text"]
        if self.know.add(rec):
            save_json(KNOW_FILE, load_json(KNOW_FILE, []) + [rec])

    def open_materials(self):
        MaterialsDialog(self)

    # ---------- данные
    def collect(self):
        self.rep_read()
        g = lambda t: cap(t.get("1.0", "end"))
        return {"act_no": self.act_no.get().strip(), "works": g(self.works), "docs": g(self.docs),
                "materials_text": g(self.materials),
                "start": self.start.get().strip(), "end": self.end.get().strip(),
                "next_works": g(self.next_works),
                "reps": [r for r in self.reps_data if r["role"] or r["name"] or r["org"]]}

    def apply(self, d):
        d = {**DEFAULT, **d}
        self.act_no.set(d["act_no"])
        if d.get("materials_sel") and not d.get("materials_text"):
            d["materials_text"] = self.old_materials(d["materials_sel"])
        for t, k in ((self.works, "works"), (self.docs, "docs"), (self.next_works, "next_works"),
                     (self.materials, "materials_text")):
            t.delete("1.0", "end")
            t.insert("1.0", d[k])
        self.start.set(d["start"])
        self.end.set(d["end"])
        self.reps_data = [dict(r) for r in d["reps"]]
        self.rep_draw()

    def open_project(self):
        p = filedialog.askopenfilename(filetypes=[("Проект акта", "*.json")])
        if p:
            self.apply(load_json(p, {}))

    def save_project(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Проект акта", "*.json")],
                                         initialfile=f"akt_{self.act_no.get()}.json")
        if p:
            save_json(p, self.collect())
            self.flash_saved(f"✓  Проект сохранён: {Path(p).name}")

    def make(self):
        d = self.collect()
        d["materials"] = [d["materials_text"]]
        p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")],
                                         initialfile=f"Акт № {d['act_no']}.docx")
        if not p:
            return
        try:
            generate(d, p)
        except PermissionError:
            messagebox.showerror("Ошибка", "Файл открыт в Word. Закройте его и повторите.")
            return
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            return
        save_json(LAST_FILE, self.collect())
        self.learn(d)
        self.flash_saved(f"✓  Сохранено: {Path(p).name}")

    def flash_saved(self, text):
        """Галочка рядом с кнопкой вместо всплывающего окна; гаснет сама."""
        self.saved_lbl.configure(text=text)
        if getattr(self, "_saved_job", None):
            self.after_cancel(self._saved_job)
        self._saved_job = self.after(6000, lambda: self.saved_lbl.configure(text=""))

    def on_close(self):
        save_json(LAST_FILE, self.collect())
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
