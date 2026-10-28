"""Генератор актов освидетельствования скрытых работ (Windows, tkinter)."""
import json
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk, filedialog, messagebox, simpledialog

from generator import generate

APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent
MAT_FILE = APP_DIR / "materials.json"
LAST_FILE = APP_DIR / "last_project.json"

ROLES = ["Заказчика", "подрядчика", "технического надзора", "авторского надзора",
         "Генподрядной организации", "субподрядчика"]

DEFAULT = {
    "act_no": "370", "city": "г. Алматы", "act_date": "",
    "object": "Строительство спортивно-оздоровительного комплекса ...",
    "contractor_org": "ТОО «АКВАСЕРВИС»",
    "works": "", "docs": "", "materials": [],
    "deviations_on": False, "deviations": "Отсутствуют",
    "start": "", "end": "", "next_works": "",
    "reps": [
        {"on": True, "role": "Заказчика", "org": "ТОО «FORTUNATA»", "position": "", "name": "Тлеуханов Н.Б."},
        {"on": True, "role": "подрядчика", "org": "ТОО «АКВАСЕРВИС»", "position": "ВрИО руководителя строительства", "name": "Соловьев Е.В."},
        {"on": True, "role": "технического надзора", "org": "«ИП Орман»", "position": "Инженер по техническому надзору", "name": "Асанбаев К.Б."},
        {"on": True, "role": "авторского надзора", "org": "TOO «Vitality Group»", "position": "", "name": "Оспанов Н.А."},
        {"on": True, "role": "Генподрядной организации", "org": "ТОО «Modulic construction»", "position": "", "name": "Соломко М.Б."},
    ],
}


def load_json(path, default):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return default


def save_json(path, obj):
    Path(path).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Акты скрытых работ")
        self.geometry("900x680")
        self.data = {**DEFAULT, **load_json(LAST_FILE, {})}
        self.mats = load_json(MAT_FILE, [])          # база материалов
        self.reps = [dict(r) for r in self.data["reps"]]
        self.vars = {}

        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=8, pady=6)
        ttk.Button(bar, text="Открыть проект…", command=self.open_project).pack(side="left")
        ttk.Button(bar, text="Сохранить проект…", command=self.save_project).pack(side="left", padx=4)
        ttk.Button(bar, text="Сформировать акт (Word)", command=self.make).pack(side="right")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=4)
        self.tab_main(nb)
        self.tab_reps(nb)
        self.tab_mats(nb)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------- вкладка «Основное»
    def tab_main(self, nb):
        f = ttk.Frame(nb, padding=8)
        nb.add(f, text="Основное")
        f.columnconfigure(1, weight=1)
        rows = [("act_no", "Номер акта", 1), ("city", "Город", 1),
                ("act_date", "Дата акта (ДД.ММ.ГГГГ)", 1),
                ("object", "Объект", 3), ("contractor_org", "СМО (выполнившая работы)", 1),
                ("works", "Предъявленные работы", 4), ("docs", "Работы выполнены по проектной документации", 3),
                ("start", "Начало работ (ДД.ММ.ГГГГ)", 1), ("end", "Окончание работ (ДД.ММ.ГГГГ)", 1),
                ("next_works", "Последующие работы", 3)]
        r = 0
        for key, label, h in rows:
            ttk.Label(f, text=label).grid(row=r, column=0, sticky="nw", pady=3, padx=(0, 8))
            if h == 1:
                v = tk.StringVar(value=self.data[key])
                ttk.Entry(f, textvariable=v).grid(row=r, column=1, sticky="ew", pady=3)
                self.vars[key] = v
            else:
                t = tk.Text(f, height=h, wrap="word")
                t.insert("1.0", self.data[key])
                t.grid(row=r, column=1, sticky="ew", pady=3)
                self.vars[key] = t
            r += 1
        self.dev_on = tk.BooleanVar(value=self.data["deviations_on"])
        ttk.Checkbutton(f, text="Есть отклонения от проекта (иначе «Отсутствуют»)",
                        variable=self.dev_on).grid(row=r, column=0, columnspan=2, sticky="w", pady=3)
        r += 1
        self.dev_text = tk.Text(f, height=2, wrap="word")
        self.dev_text.insert("1.0", self.data["deviations"])
        self.dev_text.grid(row=r, column=1, sticky="ew")
        ttk.Label(f, text="Текст отклонений").grid(row=r, column=0, sticky="nw")

    # ---------- вкладка «Представители»
    def tab_reps(self, nb):
        f = ttk.Frame(nb, padding=8)
        nb.add(f, text="Представители")
        cols = ("on", "role", "org", "position", "name")
        self.tree = ttk.Treeview(f, columns=cols, show="headings", height=8, selectmode="browse")
        for c, t, w in (("on", "Вкл", 40), ("role", "Роль", 150), ("org", "Организация", 200),
                        ("position", "Должность", 200), ("name", "Ф.И.О.", 150)):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w)
        self.tree.pack(fill="x")
        self.tree.bind("<<TreeviewSelect>>", self.rep_select)
        self.tree.bind("<Double-1>", self.rep_toggle)

        form = ttk.LabelFrame(f, text="Выбранный представитель", padding=8)
        form.pack(fill="x", pady=8)
        form.columnconfigure(1, weight=1)
        self.rv = {k: tk.StringVar() for k in ("role", "org", "position", "name")}
        for i, (k, t) in enumerate((("role", "Роль"), ("org", "Организация"),
                                    ("position", "Должность"), ("name", "Ф.И.О. (Иванов И.И.)"))):
            ttk.Label(form, text=t).grid(row=i, column=0, sticky="w", pady=2, padx=(0, 8))
            if k == "role":
                ttk.Combobox(form, textvariable=self.rv[k], values=ROLES).grid(row=i, column=1, sticky="ew")
            else:
                ttk.Entry(form, textvariable=self.rv[k]).grid(row=i, column=1, sticky="ew")
        b = ttk.Frame(f)
        b.pack(fill="x")
        for t, c in (("Добавить", self.rep_add), ("Применить к выбранному", self.rep_apply),
                     ("Вкл/Выкл", self.rep_toggle), ("Удалить", self.rep_del),
                     ("↑", lambda: self.rep_move(-1)), ("↓", lambda: self.rep_move(1))):
            ttk.Button(b, text=t, command=c).pack(side="left", padx=2)
        ttk.Label(f, text="Порядок и состав верхнего списка и блока подписей берётся отсюда. "
                          "Двойной клик по строке — включить/отключить.",
                  foreground="gray").pack(anchor="w", pady=6)
        self.rep_refresh()

    def rep_refresh(self, sel=None):
        self.tree.delete(*self.tree.get_children())
        for i, r in enumerate(self.reps):
            self.tree.insert("", "end", iid=str(i), values=("✔" if r["on"] else "", r["role"],
                                                            r["org"], r["position"], r["name"]))
        if sel is not None and 0 <= sel < len(self.reps):
            self.tree.selection_set(str(sel))

    def rep_idx(self):
        s = self.tree.selection()
        return int(s[0]) if s else None

    def rep_select(self, _=None):
        i = self.rep_idx()
        if i is not None:
            for k in self.rv:
                self.rv[k].set(self.reps[i][k])

    def rep_add(self):
        self.reps.append({"on": True, **{k: v.get() for k, v in self.rv.items()}})
        self.rep_refresh(len(self.reps) - 1)

    def rep_apply(self):
        i = self.rep_idx()
        if i is not None:
            self.reps[i].update({k: v.get() for k, v in self.rv.items()})
            self.rep_refresh(i)

    def rep_toggle(self, _=None):
        i = self.rep_idx()
        if i is not None:
            self.reps[i]["on"] = not self.reps[i]["on"]
            self.rep_refresh(i)

    def rep_del(self):
        i = self.rep_idx()
        if i is not None:
            del self.reps[i]
            self.rep_refresh(min(i, len(self.reps) - 1))

    def rep_move(self, d):
        i = self.rep_idx()
        if i is not None and 0 <= i + d < len(self.reps):
            self.reps[i], self.reps[i + d] = self.reps[i + d], self.reps[i]
            self.rep_refresh(i + d)

    # ---------- вкладка «Материалы»
    def tab_mats(self, nb):
        f = ttk.Frame(nb, padding=8)
        nb.add(f, text="Материалы")
        ttk.Label(f, text="База материалов. Двойной клик — отметить для акта (✔). "
                          "Порядок отметки не важен, в акт идут сверху вниз.").pack(anchor="w")
        self.mtree = ttk.Treeview(f, columns=("on", "name", "cert"), show="headings",
                                  height=14, selectmode="browse")
        for c, t, w in (("on", "В акт", 50), ("name", "Материал", 450), ("cert", "Сертификат №", 200)):
            self.mtree.heading(c, text=t)
            self.mtree.column(c, width=w)
        self.mtree.pack(fill="both", expand=True, pady=6)
        self.mtree.bind("<Double-1>", self.mat_toggle)
        self.chosen = set(self.data["materials_sel"]) if "materials_sel" in self.data else set()
        b = ttk.Frame(f)
        b.pack(fill="x")
        for t, c in (("Добавить", self.mat_add), ("Изменить", self.mat_edit), ("Удалить", self.mat_del),
                     ("Импорт из файла…", self.mat_import), ("Снять все отметки", self.mat_clear)):
            ttk.Button(b, text=t, command=c).pack(side="left", padx=2)
        self.mat_refresh()

    def mat_refresh(self):
        self.mtree.delete(*self.mtree.get_children())
        for i, m in enumerate(self.mats):
            self.mtree.insert("", "end", iid=str(i),
                              values=("✔" if m["name"] in self.chosen else "", m["name"], m["cert"]))
        save_json(MAT_FILE, self.mats)

    def mat_sel(self):
        s = self.mtree.selection()
        return int(s[0]) if s else None

    def mat_toggle(self, _=None):
        i = self.mat_sel()
        if i is not None:
            n = self.mats[i]["name"]
            self.chosen ^= {n}
            self.mat_refresh()

    def mat_ask(self, name="", cert=""):
        n = simpledialog.askstring("Материал", "Наименование:", initialvalue=name, parent=self)
        if not n:
            return None
        c = simpledialog.askstring("Материал", "Номер сертификата соответствия:",
                                   initialvalue=cert, parent=self)
        return {"name": n, "cert": c or ""}

    def mat_add(self):
        m = self.mat_ask()
        if m:
            self.mats.append(m)
            self.mat_refresh()

    def mat_edit(self):
        i = self.mat_sel()
        if i is not None:
            m = self.mat_ask(**self.mats[i])
            if m:
                self.chosen.discard(self.mats[i]["name"])
                self.mats[i] = m
                self.mat_refresh()

    def mat_del(self):
        i = self.mat_sel()
        if i is not None:
            del self.mats[i]
            self.mat_refresh()

    def mat_import(self):
        p = filedialog.askopenfilename(title="Текстовый файл: Наименование;Сертификат",
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
            if not line.strip():
                continue
            name, _, cert = line.partition(";")
            self.mats.append({"name": name.strip(), "cert": cert.strip()})
            n += 1
        self.mat_refresh()
        messagebox.showinfo("Импорт", f"Добавлено: {n}")

    def mat_clear(self):
        self.chosen.clear()
        self.mat_refresh()

    # ---------- данные
    def collect(self):
        d = {}
        for k, v in self.vars.items():
            d[k] = v.get("1.0", "end").strip() if isinstance(v, tk.Text) else v.get().strip()
        d["deviations_on"] = self.dev_on.get()
        d["deviations"] = self.dev_text.get("1.0", "end").strip()
        d["reps"] = [dict(r) for r in self.reps]
        d["materials_sel"] = [m["name"] for m in self.mats if m["name"] in self.chosen]
        return d

    def apply(self, d):
        d = {**DEFAULT, **d}
        for k, v in self.vars.items():
            if isinstance(v, tk.Text):
                v.delete("1.0", "end")
                v.insert("1.0", d[k])
            else:
                v.set(d[k])
        self.dev_on.set(d["deviations_on"])
        self.dev_text.delete("1.0", "end")
        self.dev_text.insert("1.0", d["deviations"])
        self.reps = [dict(r) for r in d["reps"]]
        self.chosen = set(d.get("materials_sel", []))
        self.rep_refresh()
        self.mat_refresh()

    def open_project(self):
        p = filedialog.askopenfilename(filetypes=[("Проект акта", "*.json")])
        if p:
            self.apply(load_json(p, {}))

    def save_project(self):
        p = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("Проект акта", "*.json")],
                                         initialfile=f"akt_{self.vars['act_no'].get()}.json")
        if p:
            save_json(p, self.collect())

    def make(self):
        d = self.collect()
        by_name = {m["name"]: m for m in self.mats}
        d["materials"] = [
            f'{by_name[n]["name"]} – сертификат соответствия № {by_name[n]["cert"]}.'
            if by_name[n]["cert"] else by_name[n]["name"] + "."
            for n in d["materials_sel"]]
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
        messagebox.showinfo("Готово", f"Акт сохранён:\n{p}")

    def on_close(self):
        save_json(LAST_FILE, self.collect())
        self.destroy()


if __name__ == "__main__":
    App().mainloop()
