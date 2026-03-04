#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Восстановление повреждённых .docx файлов
=========================================
Запуск:  python3 docx_recovery.py
Система: Alt Linux, Ubuntu, Debian и другие Linux-дистрибутивы

Скрипт сам проверяет и устанавливает нужные зависимости при первом запуске.
"""

import sys
import os
import subprocess



def _install_tkinter():
    print(">>> tkinter не найден. Пробуем установить автоматически...")
    candidates = [
        ["apt-get", "install", "-y", "python3-module-tkinter"],  # Alt Linux
        ["apt-get", "install", "-y", "python3-tk"],              # Debian/Ubuntu
        ["dnf",     "install", "-y", "python3-tkinter"],         # Fedora
        ["zypper",  "install", "-y", "python3-tk"],              # openSUSE
    ]
    for cmd in candidates:
        try:
            subprocess.run(["sudo"] + cmd, check=True, timeout=120)
            print(f">>> Установлено через: {' '.join(cmd)}")
            return True
        except Exception:
            continue
    return False


def _check_zip():
    try:
        subprocess.run(["zip", "--version"], capture_output=True, timeout=5)
        return True
    except Exception:
        return False


def _install_zip():
    for cmd in [
        ["apt-get", "install", "-y", "zip"],
        ["dnf",     "install", "-y", "zip"],
        ["zypper",  "install", "-y", "zip"],
    ]:
        try:
            subprocess.run(["sudo"] + cmd, check=True, timeout=120)
            return True
        except Exception:
            continue
    return False


# Проверяем DISPLAY перед попыткой открыть окно
if not os.environ.get("DISPLAY") and sys.platform != "win32":
    print("[ОШИБКА] Переменная DISPLAY не установлена.")
    print("Запускайте скрипт из графической сессии (не через SSH без -X).")
    sys.exit(1)

# Проверяем tkinter
try:
    import tkinter as _tk_test
    del _tk_test
except ModuleNotFoundError:
    ok = _install_tkinter()
    if not ok:
        print(
            "\n[ОШИБКА] Не удалось установить tkinter автоматически.\n"
            "Установите вручную:\n"
            "  Alt Linux:     sudo apt-get install python3-module-tkinter\n"
            "  Ubuntu/Debian: sudo apt-get install python3-tk\n"
            "  Fedora:        sudo dnf install python3-tkinter\n"
        )
        sys.exit(1)
    os.execv(sys.executable, [sys.executable] + sys.argv)

if not _check_zip():
    print(">>> Утилита zip не найдена. Устанавливаем...")
    _install_zip()

# ─────────────────────────────────────────────────────────────────────────────
# Основные импорты
# ─────────────────────────────────────────────────────────────────────────────
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import shutil
import zipfile
import struct
import zlib
import tempfile
import time


# ─────────────────────────────────────────────────────────────────────────────
# Шрифты с fallback (нет гарантии что DejaVu есть на каждом Alt Linux)
# ─────────────────────────────────────────────────────────────────────────────

def _pick_font(candidates, size, weight="normal"):
    """Возвращает первый доступный шрифт из списка"""
    try:
        import tkinter.font as tkfont
        root = tk.Tk()
        root.withdraw()
        available = set(tkfont.families())
        root.destroy()
        for name in candidates:
            if name in available:
                return (name, size, weight)
    except Exception:
        pass
    return ("TkDefaultFont", size, weight)


def make_fonts():
    sans = ["DejaVu Sans", "Liberation Sans", "FreeSans",
            "Helvetica", "Arial", "TkDefaultFont"]
    mono = ["DejaVu Sans Mono", "Liberation Mono", "FreeMono",
            "Courier New", "Courier", "TkFixedFont"]
    return {
        "title": _pick_font(sans, 15, "bold"),
        "h2":    _pick_font(sans, 11, "bold"),
        "body":  _pick_font(sans, 10),
        "small": _pick_font(sans, 9),
        "btn":   _pick_font(sans, 11, "bold"),
        "mono":  _pick_font(mono, 9),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Цвета
# ─────────────────────────────────────────────────────────────────────────────
BG      = "#1C2333"
CARD    = "#242E42"
ACCENT  = "#4A86E8"
SUCCESS = "#50C878"
ERROR   = "#E05C5C"
WARN    = "#E8B84A"
TEAL    = "#50BFA0"
TEXT    = "#E2E8F4"
SUBTEXT = "#8B96A8"
BORDER  = "#2D3A52"
WHITE   = "#FFFFFF"

# Этапы (ASCII-иконки вместо emoji — X11/Alt Linux надёжнее их рендерит)
STEPS = [
    ("[1]", "Анализ файла",
     "Определяем размер, проверяем сигнатуру ZIP и BOM UTF-16"),
    ("[2]", "Метод А: zip -FF",
     "Восстанавливаем ZIP-структуру штатной утилитой"),
    ("[3]", "Метод Б: частичное чтение",
     "Читаем ZIP-записи по одной, обходя повреждённые блоки"),
    ("[4]", "Метод В: сканирование байт",
     "Ищем PK-сигнатуры и XML-теги напрямую, поддержка UTF-16"),
    ("[5]", "Сборка содержимого",
     "Объединяем найденные папки word/, _rels/, docProps/"),
    ("[6]", "Чистый шаблон",
     "Создаём корректную OOXML-основу для пересборки"),
    ("[7]", "Сохранение",
     "Упаковываем восстановленный .docx и записываем на диск"),
]


# ─────────────────────────────────────────────────────────────────────────────
# Движок восстановления
# ─────────────────────────────────────────────────────────────────────────────

def run_cmd(cmd, input_text=None):
    try:
        proc = subprocess.Popen(cmd, shell=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
            stdin=subprocess.PIPE, text=True)
        out, err = proc.communicate(input=input_text, timeout=60)
        return proc.returncode, (out + err).strip()
    except Exception as e:
        return -1, str(e)


def try_open_zip(path):
    try:
        with zipfile.ZipFile(path, "r") as z:
            return True, z.namelist()
    except Exception as e:
        return False, str(e)


def repair_zip_ff(src, dst):
    return run_cmd(f'zip -FF "{src}" --out "{dst}"', input_text="y\n")


def partial_read_zip(path):
    result = {}
    try:
        with zipfile.ZipFile(path, "r") as z:
            for name in z.namelist():
                try:
                    result[name] = z.read(name)
                except Exception:
                    pass
    except zipfile.BadZipFile:
        pass
    return result


def scan_pk_signatures(data):
    """Ищет ZIP Local File Headers (PK\\x03\\x04) в сырых байтах"""
    result = {}
    SIG = b"PK\x03\x04"
    i = 0
    while i < len(data) - 30:
        pos = data.find(SIG, i)
        if pos == -1:
            break
        try:
            (_, flags, method, _, _, crc32v, comp_size, uncomp_size,
             fname_len, extra_len) = struct.unpack_from("<HHHHHIIIIHH", data, pos + 4)

            fname_end  = pos + 30 + fname_len
            data_start = fname_end + extra_len
            data_end   = data_start + comp_size

            if (fname_len == 0 or fname_end > len(data) or
                    data_end > len(data) or comp_size > 50_000_000):
                i = pos + 1
                continue

            fname = data[pos + 30 : fname_end].decode("utf-8", errors="replace")
            if fname.endswith("/") or (comp_size == 0 and uncomp_size == 0):
                i = max(data_end, pos + 1)
                continue

            raw = data[data_start:data_end]
            if method == 0:
                content = raw
            elif method == 8:
                try:
                    content = zlib.decompress(raw, -15)
                except zlib.error:
                    try:
                        content = zlib.decompress(raw)
                    except zlib.error:
                        i = pos + 1
                        continue
            else:
                i = pos + 1
                continue

            result[fname] = content
            i = max(data_end, pos + 1)
        except (struct.error, ValueError):
            i = pos + 1
    return result


def scan_xml_fragments(data):
    """
    Ищет XML-теги Word прямо в тексте.
    Поддерживает UTF-8, UTF-16 LE/BE (Windows Notepad сохраняет UTF-16).
    """
    result = {}
    encodings = ["utf-8", "utf-16", "utf-16-le", "utf-16-be", "cp1251", "latin-1"]
    patterns = [
        ("<w:document",    "</w:document>",    "word/document.xml"),
        ("<w:styles",      "</w:styles>",      "word/styles.xml"),
        ("<Relationships", "</Relationships>", "word/_rels/document.xml.rels"),
        ("<Types",         "</Types>",         "[Content_Types].xml"),
    ]
    for enc in encodings:
        try:
            raw = data[2:] if (enc in ("utf-16-le","utf-16-be") and
                               data[:2] in (b"\xff\xfe", b"\xfe\xff")) else data
            text = raw.decode(enc, errors="ignore")
        except Exception:
            continue

        found_any = False
        for start_tag, end_tag, fname in patterns:
            s = text.find(start_tag)
            e = text.find(end_tag)
            if s != -1 and e != -1 and e > s:
                fragment = text[s : e + len(end_tag)]
                if not fragment.startswith("<?xml"):
                    fragment = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + fragment
                if fname not in result:
                    result[fname] = fragment.encode("utf-8")
                    found_any = True
        if found_any:
            break
    return result


def build_dir_from_dict(file_dict, dest_dir):
    for fname, content in file_dict.items():
        full = os.path.join(dest_dir, fname.replace("/", os.sep))
        os.makedirs(os.path.dirname(full), exist_ok=True)
        if isinstance(content, str):
            content = content.encode("utf-8")
        with open(full, "wb") as f:
            f.write(content)


def extract_zip(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path, "r") as z:
        z.extractall(dest_dir)


def create_minimal_docx(path):
    """Минимальный корректный .docx. document.xml идёт последним."""
    CT = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
          '</Types>')
    RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            '</Relationships>')
    WRELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
             '</Relationships>')
    STYLES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="styles"/>')
    CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"'
            ' xmlns:dc="http://purl.org/dc/elements/1.1/">'
            '<dc:creator>Recovery Tool</dc:creator></cp:coreProperties>')
    APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>Recovery Tool</Application></Properties>')
    DOC = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body><w:p><w:r><w:t>.</w:t></w:r></w:p></w:body></w:document>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",          CT)
        z.writestr("_rels/.rels",                  RELS)
        z.writestr("docProps/core.xml",            CORE)
        z.writestr("docProps/app.xml",             APP)
        z.writestr("word/styles.xml",              STYLES)
        z.writestr("word/_rels/document.xml.rels", WRELS)
        z.writestr("word/document.xml",            DOC)  # последним


def pack_dir_to_docx(src_dir, out_path):
    all_files = []
    for root, _, files in os.walk(src_dir):
        for f in files:
            full    = os.path.join(root, f)
            arcname = os.path.relpath(full, src_dir).replace(os.sep, "/")
            all_files.append((full, arcname))
    last   = [(f, a) for f, a in all_files if a == "word/document.xml"]
    others = [(f, a) for f, a in all_files if a != "word/document.xml"]
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for full, arcname in others + last:
            zf.write(full, arcname)


# ─────────────────────────────────────────────────────────────────────────────
# GUI
# ─────────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.F = make_fonts()
        self.title("Восстановление .docx")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.src_var = tk.StringVar()
        self.dst_var = tk.StringVar()
        self._steps  = []
        self._build()
        self._center()

    def _center(self):
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w = max(self.winfo_reqwidth(), 720)
        h = min(self.winfo_reqheight() + 10, sh - 80)
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.minsize(560, 400)

    # ── построение UI ────────────────────────────────────────────────────────

    def _build(self):
        outer = tk.Frame(self, bg=BG)
        outer.pack(fill="both", expand=True)

        vsb = tk.Scrollbar(outer, orient="vertical")
        vsb.pack(side="right", fill="y")

        canvas = tk.Canvas(outer, bg=BG, highlightthickness=0,
                           yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.config(command=canvas.yview)

        inner = tk.Frame(canvas, bg=BG)
        cwin  = canvas.create_window((0, 0), window=inner, anchor="nw")

        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(cwin, width=e.width))
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        def scroll(e):
            d = -1 if (getattr(e, "num", 0) == 4 or
                       getattr(e, "delta", 0) > 0) else 1
            canvas.yview_scroll(d, "units")

        canvas.bind_all("<MouseWheel>", scroll)
        canvas.bind_all("<Button-4>",   scroll)
        canvas.bind_all("<Button-5>",   scroll)

        self._fill(inner)

    def _fill(self, p):
        F   = self.F
        PAD = 22

        # Шапка
        hdr = tk.Frame(p, bg=ACCENT, pady=16)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Восстановление документа .docx",
                 font=F["title"], bg=ACCENT, fg=WHITE,
                 anchor="w").pack(padx=PAD, anchor="w")
        tk.Label(hdr,
                 text="Автоматическое восстановление повреждённых Word-файлов",
                 font=F["small"], bg=ACCENT, fg="#c8d8f8",
                 anchor="w").pack(padx=PAD, anchor="w")

        # Карточка выбора файлов
        fc = tk.Frame(p, bg=CARD, padx=PAD, pady=14)
        fc.pack(fill="x", padx=PAD, pady=(PAD, 0))

        self._file_row(fc, "Повреждённый файл (.docx):",
                       self.src_var, self._pick_src)
        self._file_row(fc, "Сохранить восстановленный файл как:",
                       self.dst_var, self._pick_dst, top=10)

        # Заголовок шагов
        tk.Label(p, text="Этапы восстановления",
                 font=F["h2"], bg=BG, fg=SUBTEXT,
                 anchor="w").pack(padx=PAD, pady=(PAD, 5), anchor="w")

        sf = tk.Frame(p, bg=BG)
        sf.pack(fill="x", padx=PAD)

        for icon, title, desc in STEPS:
            row = tk.Frame(sf, bg=CARD, pady=7, padx=10)
            row.pack(fill="x", pady=2)
            row.columnconfigure(1, weight=1)

            num = tk.Label(row, text=icon,
                           font=F["small"], bg=BORDER, fg=SUBTEXT,
                           width=5, anchor="center", relief="flat")
            num.grid(row=0, column=0, rowspan=2, padx=(0, 10),
                     pady=2, sticky="ns")

            tk.Label(row, text=title,
                     font=F["h2"], bg=CARD, fg=TEXT,
                     anchor="w").grid(row=0, column=1, sticky="w")

            tk.Label(row, text=desc,
                     font=F["small"], bg=CARD, fg=SUBTEXT,
                     wraplength=440, justify="left",
                     anchor="w").grid(row=1, column=1, sticky="w")

            st = tk.Label(row, text="", font=F["small"],
                          bg=CARD, fg=SUBTEXT, width=14, anchor="e")
            st.grid(row=0, column=2, rowspan=2, padx=6, sticky="e")

            self._steps.append((num, st))

        # Журнал
        tk.Label(p, text="Журнал",
                 font=F["h2"], bg=BG, fg=SUBTEXT,
                 anchor="w").pack(padx=PAD, pady=(PAD, 4), anchor="w")

        log_frame = tk.Frame(p, bg="#0F1520", padx=8, pady=6)
        log_frame.pack(fill="x", padx=PAD)

        lsb = tk.Scrollbar(log_frame, orient="vertical")
        self.log = tk.Text(
            log_frame,
            height=7, font=F["mono"],
            bg="#0F1520", fg=SUBTEXT,
            insertbackground=TEXT, relief="flat", bd=0,
            wrap="word", state="disabled",
            yscrollcommand=lsb.set
        )
        lsb.config(command=self.log.yview)
        lsb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)

        # Кнопка (внутри scroll — всегда доступна)
        bf = tk.Frame(p, bg=BG)
        bf.pack(fill="x", padx=PAD, pady=PAD)

        self.btn = tk.Button(
            bf,
            text="  НАЧАТЬ ВОССТАНОВЛЕНИЕ  ",
            font=F["btn"],
            bg=ACCENT, fg=WHITE,
            activebackground="#3a6dcc", activeforeground=WHITE,
            relief="flat", bd=0, cursor="hand2",
            pady=12, command=self._start
        )
        self.btn.pack(fill="x")
        self.btn.bind("<Enter>", lambda e: self.btn.config(bg="#3a6dcc"))
        self.btn.bind("<Leave>", lambda e: self.btn.config(bg=ACCENT))

    def _file_row(self, parent, label, var, cmd, top=0):
        tk.Label(parent, text=label,
                 font=self.F["small"], bg=CARD, fg=SUBTEXT,
                 anchor="w").pack(anchor="w", pady=(top, 2))
        row = tk.Frame(parent, bg=CARD)
        row.pack(fill="x", pady=(0, 2))
        tk.Entry(row, textvariable=var,
                 font=self.F["body"], bg=BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", bd=6,
                 highlightthickness=1,
                 highlightcolor=ACCENT,
                 highlightbackground=BORDER
                 ).pack(side="left", fill="x", expand=True, ipady=3)
        tk.Button(row, text="Выбрать...",
                  font=self.F["small"], bg=BORDER, fg=TEXT,
                  activebackground=ACCENT, activeforeground=WHITE,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=cmd).pack(side="left", padx=(6, 0))

    # ── диалоги ──────────────────────────────────────────────────────────────

    def _pick_src(self):
        p = filedialog.askopenfilename(
            title="Выберите повреждённый .docx файл",
            filetypes=[("Word документы", "*.docx"), ("Все файлы", "*.*")]
        )
        if p:
            self.src_var.set(p)
            if not self.dst_var.get():
                self.dst_var.set(os.path.splitext(p)[0] + "_восстановлен.docx")

    def _pick_dst(self):
        p = filedialog.asksaveasfilename(
            title="Сохранить восстановленный файл",
            defaultextension=".docx",
            filetypes=[("Word документы", "*.docx")]
        )
        if p:
            self.dst_var.set(p)

    # ── логирование ──────────────────────────────────────────────────────────

    def _log(self, text, color=None):
        self.log.configure(state="normal")
        tag = f"t{time.monotonic_ns()}"
        self.log.insert("end", text + "\n", tag)
        if color:
            self.log.tag_config(tag, foreground=color)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.update_idletasks()

    # ── шаги ─────────────────────────────────────────────────────────────────

    def _set_step(self, idx, state):
        cfg = {
            "wait": (BORDER,  SUBTEXT, ""),
            "run":  (ACCENT,  WHITE,   ">> в работе"),
            "ok":   (SUCCESS, WHITE,   "OK  готово"),
            "skip": (BORDER,  SUBTEXT, "--  пропущен"),
            "err":  (ERROR,   WHITE,   "!!  ошибка"),
        }
        num, st = self._steps[idx]
        bg, fg, label = cfg[state]
        num.configure(bg=bg, fg=fg)
        st.configure(text=label, fg=fg if state != "wait" else SUBTEXT)
        self.update_idletasks()

    def _reset(self):
        for i in range(len(STEPS)):
            self._set_step(i, "wait")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ── запуск ────────────────────────────────────────────────────────────────

    def _start(self):
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()
        if not src:
            messagebox.showerror("Нет файла",
                                 "Выберите повреждённый .docx файл.")
            return
        if not os.path.isfile(src):
            messagebox.showerror("Файл не найден",
                                 f"Файл не существует:\n{src}")
            return
        if not dst:
            messagebox.showerror("Нет пути",
                                 "Укажите, куда сохранить результат.")
            return
        self._reset()
        self.btn.configure(state="disabled", text="  Восстановление...  ")
        threading.Thread(target=self._run,
                         args=(src, dst), daemon=True).start()

    # ── основная логика (фоновый поток) ──────────────────────────────────────

    def _run(self, src, dst):
        tmp = tempfile.mkdtemp(prefix="docx_fix_")
        try:
            self._recover(src, dst, tmp)
        except Exception as e:
            self._log(f"\n[ОШИБКА] {e}", ERROR)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            self.after(0, lambda: self.btn.configure(
                state="normal",
                text="  НАЧАТЬ ВОССТАНОВЛЕНИЕ  "))

    def _recover(self, src, dst, tmp):

        # Шаг 1: Анализ
        self.after(0, self._set_step, 0, "run")
        self._log("Шаг 1. Анализ файла...")
        time.sleep(0.2)

        src_copy = os.path.join(tmp, "input.zip")
        shutil.copy2(src, src_copy)

        with open(src_copy, "rb") as f:
            raw = f.read()

        self._log(f"  Размер: {len(raw):,} байт ({len(raw)//1024} КБ)", TEAL)

        if raw[:2] == b"PK":
            self._log("  ZIP-сигнатура PK обнаружена.", TEAL)
        elif raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
            self._log("  BOM UTF-16 — файл сохранён Блокнотом (Windows)!", WARN)
        else:
            self._log(f"  Начало файла: {raw[:4].hex()} (ZIP-сигнатуры нет).", WARN)

        ok, info = try_open_zip(src_copy)
        if ok:
            self._log(f"  Архив читается нормально. Файлов: {len(info)}", TEAL)
            working_zip = src_copy
            recovered   = None
        else:
            self._log(f"  Архив повреждён: {str(info)[:80]}", WARN)
            working_zip = None
            recovered   = None

        self.after(0, self._set_step, 0, "ok")

        # Шаг 2: zip -FF
        self.after(0, self._set_step, 1, "run")
        time.sleep(0.2)

        if working_zip:
            self._log("\nШаг 2. Архив цел — zip -FF не нужен.")
            self.after(0, self._set_step, 1, "skip")
        else:
            self._log("\nШаг 2. Запускаем zip -FF...")
            fixed = os.path.join(tmp, "fixed.zip")
            _, out = repair_zip_ff(src_copy, fixed)
            if out:
                self._log(f"  {out[:120]}", SUBTEXT)
            ok2, info2 = try_open_zip(fixed)
            if ok2 and len(info2) > 0:
                self._log(f"  zip -FF восстановил архив! Файлов: {len(info2)}", SUCCESS)
                working_zip = fixed
                self.after(0, self._set_step, 1, "ok")
            else:
                self._log("  zip -FF не помог, следующий метод.", WARN)
                self.after(0, self._set_step, 1, "skip")

        # Шаг 3: Частичное чтение
        self.after(0, self._set_step, 2, "run")
        time.sleep(0.2)

        if working_zip:
            self._log("\nШаг 3. Частичное чтение не нужно.")
            self.after(0, self._set_step, 2, "skip")
        else:
            self._log("\nШаг 3. Частичное чтение ZIP...")
            partial = partial_read_zip(src_copy)
            if partial:
                self._log(f"  Прочитано файлов: {len(partial)}", SUCCESS)
                for n in list(partial.keys())[:6]:
                    self._log(f"    + {n}", TEAL)
                if not any("document.xml" in k for k in partial):
                    self._log("  word/document.xml не прочитан — применим метод В.", WARN)
                recovered = partial
                self.after(0, self._set_step, 2, "ok")
            else:
                self._log("  Частичное чтение не дало результатов.", WARN)
                self.after(0, self._set_step, 2, "skip")

        # Шаг 4: Сканирование байт
        self.after(0, self._set_step, 3, "run")
        time.sleep(0.3)

        need_scan = (working_zip is None) and (
            recovered is None or
            not any("document.xml" in k for k in (recovered or {}))
        )

        if not need_scan:
            self._log("\nШаг 4. Сканирование байт не требуется.")
            self.after(0, self._set_step, 3, "skip")
        else:
            self._log(f"\nШаг 4. Сканирование {len(raw):,} байт...")

            pk = scan_pk_signatures(raw)
            if pk:
                has_doc = any("document.xml" in k for k in pk)
                self._log(f"  PK-сигнатуры: найдено {len(pk)} файл(ов)"
                          + (" [document.xml есть]" if has_doc else " [document.xml НЕТ]"),
                          SUCCESS if has_doc else WARN)
                for n in list(pk.keys())[:5]:
                    self._log(f"    + {n}", TEAL)
                recovered = dict(recovered or {})
                recovered.update(pk)
                self.after(0, self._set_step, 3, "ok")
            else:
                self._log("  PK-сигнатур нет. Ищем XML-фрагменты (UTF-8/UTF-16)...", WARN)
                xf = scan_xml_fragments(raw)
                if xf:
                    self._log(f"  XML-фрагменты найдены: {list(xf.keys())}", SUCCESS)
                    recovered = dict(recovered or {})
                    recovered.update(xf)
                    self.after(0, self._set_step, 3, "ok")
                elif recovered:
                    self._log("  Дополнительных данных нет, используем шаг 3.", WARN)
                    self.after(0, self._set_step, 3, "skip")
                else:
                    self._log("  Данные не найдены ни одним методом.", ERROR)
                    self._log("  Файл полностью уничтожен — восстановление невозможно.", ERROR)
                    self.after(0, self._set_step, 3, "err")
                    return

        # Шаг 5: Сборка содержимого
        self.after(0, self._set_step, 4, "run")
        self._log("\nШаг 5. Сборка содержимого...")
        time.sleep(0.2)

        src_dir = os.path.join(tmp, "src_content")
        os.makedirs(src_dir)

        if working_zip:
            try:
                extract_zip(working_zip, src_dir)
                self._log("  Распаковано из ZIP.", TEAL)
            except Exception as e:
                self._log(f"  Ошибка распаковки: {e}", ERROR)
                self.after(0, self._set_step, 4, "err")
                return
        else:
            build_dir_from_dict(recovered, src_dir)
            self._log(f"  Записано {len(recovered)} файл(ов).", TEAL)

        word_dir = os.path.join(src_dir, "word")
        if not os.path.isdir(word_dir):
            self._log("  Папка word/ не найдена — восстановление невозможно.", ERROR)
            self.after(0, self._set_step, 4, "err")
            return

        self._log(f"  word/: {', '.join(os.listdir(word_dir)[:7])}", SUCCESS)
        self.after(0, self._set_step, 4, "ok")

        # Шаг 6: Чистый шаблон
        self.after(0, self._set_step, 5, "run")
        self._log("\nШаг 6. Создание чистого шаблона OOXML...")
        time.sleep(0.2)

        base_docx = os.path.join(tmp, "base.docx")
        base_dir  = os.path.join(tmp, "base_content")
        os.makedirs(base_dir)

        try:
            create_minimal_docx(base_docx)
            extract_zip(base_docx, base_dir)
        except Exception as e:
            self._log(f"  Ошибка создания шаблона: {e}", ERROR)
            self.after(0, self._set_step, 5, "err")
            return

        dst_word = os.path.join(base_dir, "word")
        shutil.rmtree(dst_word, ignore_errors=True)
        try:
            shutil.copytree(word_dir, dst_word)
        except Exception as e:
            self._log(f"  Ошибка копирования word/: {e}", ERROR)
            self.after(0, self._set_step, 5, "err")
            return

        for fname in ("[Content_Types].xml",):
            src_f = os.path.join(src_dir, fname)
            if os.path.isfile(src_f):
                shutil.copy2(src_f, os.path.join(base_dir, fname))
                self._log(f"  {fname} — из оригинала.", TEAL)

        for dname in ("_rels", "docProps"):
            src_d = os.path.join(src_dir, dname)
            if os.path.isdir(src_d):
                dst_d = os.path.join(base_dir, dname)
                shutil.rmtree(dst_d, ignore_errors=True)
                shutil.copytree(src_d, dst_d)
                self._log(f"  {dname}/ — из оригинала.", TEAL)

        self.after(0, self._set_step, 5, "ok")

        # Шаг 7: Сохранение
        self.after(0, self._set_step, 6, "run")
        self._log("\nШаг 7. Сохранение файла...")
        time.sleep(0.2)

        try:
            pack_dir_to_docx(base_dir, dst)
            kb = os.path.getsize(dst) // 1024
            self._log(f"  Сохранено: {dst}", SUCCESS)
            self._log(f"  Размер: {kb} КБ", TEAL)
            self.after(0, self._set_step, 6, "ok")
        except Exception as e:
            self._log(f"  Ошибка сохранения: {e}", ERROR)
            self.after(0, self._set_step, 6, "err")
            return

        self._log("\n" + "-" * 48, TEAL)
        self._log("  ВОССТАНОВЛЕНИЕ ЗАВЕРШЕНО УСПЕШНО!", SUCCESS)
        self._log(f"  {dst}", TEAL)
        self._log("-" * 48, TEAL)

        self.after(0, self._done, dst)

    def _done(self, path):
    	folder = os.path.dirname(path)
    	if messagebox.askyesno(
        "Готово!",
        f"Файл восстановлен:\n{path}\n\nОткрыть папку с файлом?"
    ):
        # Пробуем файловые менеджеры по очереди (Alt Linux, GNOME, KDE, Xfce...)
      	  for fm in ["xdg-open", "nautilus", "dolphin", "thunar", "nemo", "pcmanfm"]:
            try:
                subprocess.Popen(
                    [fm, folder],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
                break
            except FileNotFoundError:
                continue


# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    App().mainloop()
