#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
textcheck.py - локальная проверка курсовой/диплома. Работает офлайн, текст никуда не отправляется.
Форматы: .docx, .doc, .pdf, .txt, .md (формат определяется по содержимому файла, а не по расширению).

Что нужно для чтения форматов
-----------------------------
  .docx, .txt, .md  - ничего, только Python 3.8+.
  .doc              - ничего: Word 97-2003 разбирается встроенным кодом. Если файл нестандартный,
                      программа попробует LibreOffice (если он установлен).
  .pdf              - нужна одна из библиотек:  pip install pypdf   (или  pip install pymupdf),
                      либо установленный Poppler (команда pdftotext). PDF-сканы (картинки без текста)
                      прочитать нельзя: сначала распознайте их (OCR) в FineReader, OCRmyPDF и т.п.

Окно с кнопками
---------------
  Запустите файл без параметров (двойной щелчок или  python textcheck.py ): откроется окно с кнопками
  «Добавить файл» и «Проверить». Нужен tkinter: он входит в обычный Python с python.org для Windows и macOS,
  в Linux:  sudo apt install python3-tk . Чтобы окно открывалось без чёрной консоли в Windows,
  переименуйте файл в textcheck.pyw.

Команды (для командной строки)
------------------------------
  python textcheck.py audit   работа.docx [--fix] [--limit 40] [--allow "іїє"] [--report отчет.txt]
      Ищет: невидимые символы (нулевой ширины, мягкий перенос, управляющие), необычные пробелы,
      подмену букв (латиница/греческие вместо кириллицы), цифры вместо букв (0 вместо о),
      скрытый текст в docx (атрибут «скрытый», белый шрифт, шрифт <= 2 пт).
      --fix сохраняет исправленную копию: «работа.clean.docx» для docx и «работа.clean.txt» для
      doc/pdf/txt (оригинал не трогается). Скрытый текст автоматически НЕ удаляется.
      Для .doc и .pdf скрытый текст (белый шрифт и т.п.) не проверяется - только символы.

  python textcheck.py compare работа.pdf папка_с_источниками [файл ...] [--shingle 5]
                         [--no-legal] [--skip-quotes] [--limit 30] [--report отчет.txt]
      Сравнивает работу с источниками, которые вы сами положили в папку (.docx .doc .pdf .txt .md).
      Автоматически исключаются: титульный лист и содержание (всё до «Введения»), список
      литературы и приложения (всё с заголовка «Список литературы» и далее), предложения со ссылками
      на законы, ГОСТы, статьи кодексов (отключается --no-legal), цитаты в кавычках (--skip-quotes).
      Перед сравнением подмены символов и невидимые символы убираются, чтобы они не маскировали совпадения.

Проверка по открытым источникам в интернете (команда compare с ключом --web)
------------------------------------------------------------------------------
  python textcheck.py compare работа.docx --web --engine yandex --queries 60
  python textcheck.py compare работа.docx папка_с_источниками --web          (вместе с вашими файлами)

  Как это работает: из работы выбираются фразы по 9 слов (равномерно по всему тексту), каждая
  отправляется в поисковик в кавычках, найденные страницы скачиваются и сравниваются с текстом так же,
  как ваши файлы. Фразы, уже найденные на ранее скачанных страницах, повторно не ищутся.
  ВАЖНО: фразы из вашей работы уходят в поисковую систему. Включается только явным --web.

  Поисковики (--engine):
    ddg      DuckDuckGo, без ключа. Страница результатов разбирается как HTML: может просить капчу.
    yandex   Yandex Search API: переменные YANDEX_SEARCH_API_KEY и YANDEX_FOLDER_ID (или --yandex-key,
             --yandex-folder). Лучший вариант для русскоязычных текстов; платный, нужен аккаунт Yandex Cloud.
    brave    Brave Search API: переменная BRAVE_API_KEY (или --brave-key).
    searxng  свой сервер SearXNG с включённым JSON: --searxng-url http://localhost:8080 (или SEARXNG_URL).
  Google Custom Search и Bing Search API закрыты (Bing - август 2025, Google - для новых клиентов
  с января 2026, полностью с 01.01.2027), поэтому их здесь нет.

  Результаты поиска и страницы кэшируются в папке .textcheck_cache (повторный запуск не ходит в сеть).
  Сайты с законами (consultant.ru, garant.ru, pravo.gov.ru и т.п.) исключены по умолчанию.

Ограничения
-----------
  * Без --web программа знает только те источники, которые вы ей дали. Базы вуза она не видит никогда.
  * С --web проверяется не весь текст, а выборка фраз (--queries), поэтому процент - нижняя оценка.
    Найти получится только то, что есть в открытом интернете и находится по точной фразе.
  * Определения ИИ-текста здесь нет: надёжного локального способа не существует.
  * В PDF абзацы восстанавливаются по эвристике (переносы строк, заголовки), поэтому границы
    абзацев и поиск «Введения»/«Списка литературы» там менее точны, чем в Word.
"""
import argparse
import base64
import collections
import gzip
import hashlib
import html
import io
import json
import os
import random
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from html.parser import HTMLParser
from xml.sax.saxutils import escape as xml_escape

REPORT = []
STATS = {}                  # итоговые числа последней проверки (для окна)
CANCEL = threading.Event()  # кнопка «Остановить» в окне


def emit(line=""):
    REPORT.append(line)
    print(line)


# ----------------------------------------------------------------------------------------------
# Символы
# ----------------------------------------------------------------------------------------------
NBSPS = {0x00A0, 0x202F}  # неразрывные пробелы: в нормальных документах встречаются часто
INVISIBLE_EXTRA = {0x115F, 0x1160, 0x3164, 0xFFA0, 0x034F, 0x17B4, 0x17B5}
RUSSIAN = set("абвгдеёжзийклмнопрстуфхцчшщъыьэюя")

LAT2CYR = dict(zip("AaBCcEeHKkMOoPpTXxyY", "АаВСсЕеНКкМОоРрТХхуУ"))
CYR2LAT = {v: k for k, v in LAT2CYR.items()}
GRE2CYR = dict(zip("ΑΒΕΗΚΜΟΡΤΧΥορ", "АВЕНКМОРТХУор"))
GRE2LAT = dict(zip("ΑΒΕΖΗΙΚΜΝΟΡΤΥΧονρ", "ABEZHIKMNOPTYXovp"))
MAPS = {("LAT", "CYR"): LAT2CYR, ("GRE", "CYR"): GRE2CYR,
        ("CYR", "LAT"): CYR2LAT, ("GRE", "LAT"): GRE2LAT}

WORD_RE = re.compile(r"[^\W_]+", re.U)


def invisible(ch):
    if ch in "\t\n\r":
        return False
    cp = ord(ch)
    return (unicodedata.category(ch) in ("Cf", "Cc") or cp in INVISIBLE_EXTRA
            or 0xFE00 <= cp <= 0xFE0F or 0xE0100 <= cp <= 0xE01EF)


def odd_space(ch):
    return (unicodedata.category(ch) in ("Zs", "Zl", "Zp") and ch != " ") or ord(ch) == 0x2800


def script_of(ch):
    if not ch.isalpha():
        return None
    name = unicodedata.name(ch, "")
    if name.startswith("CYRILLIC"):
        return "CYR"
    if name.startswith("LATIN"):
        return "LAT"
    if name.startswith("GREEK"):
        return "GRE"
    return "OTH"


def analyze(text, allowed=frozenset()):
    """Возвращает (находки, план_исправлений). Находка: (вид, уровень, [индексы], (с, по), пояснение)."""
    finds, plan = [], {}
    for i, ch in enumerate(text):
        cp = ord(ch)
        if invisible(ch):
            name = unicodedata.name(ch, "управляющий символ")
            finds.append(("invis", "warn", [i], (i, i + 1), "U+%04X %s" % (cp, name)))
            plan[i] = ""
        elif cp in NBSPS:
            finds.append(("nbsp", "info", [i], (i, i + 1), ""))
        elif odd_space(ch):
            finds.append(("space", "warn", [i], (i, i + 1),
                          "U+%04X %s" % (cp, unicodedata.name(ch, "?"))))
            plan[i] = " "

    kept = [(i, ch) for i, ch in enumerate(text) if not invisible(ch)]
    s = "".join(c for _, c in kept)
    letters = [x for x in map(script_of, s) if x]
    cyr_share = (sum(1 for x in letters if x == "CYR") / len(letters)) if letters else 0.0

    for m in WORD_RE.finditer(s):
        word = m.group()
        idx = [kept[t][0] for t in range(m.start(), m.end())]
        span = (idx[0], idx[-1] + 1)
        scr = [script_of(c) for c in word]
        present = {x for x in scr if x in ("CYR", "LAT", "GRE")}

        if len(present) > 1:
            cnt = collections.Counter(x for x in scr if x in present).most_common()
            if cnt[0][1] == cnt[1][1]:
                major = "CYR" if cyr_share >= 0.5 else "LAT"
                if major not in present:
                    major = cnt[0][0]
            else:
                major = cnt[0][0]
            marks, notes = [], []
            for i, c, sc in zip(idx, word, scr):
                if sc in present and sc != major:
                    marks.append(i)
                    tgt = MAPS.get((sc, major), {}).get(c)
                    notes.append("%s(%s)->%s" % (c, sc.lower(), tgt or "?"))
                    if tgt:
                        plan[i] = tgt
            finds.append(("mixed", "warn", marks, span, "слово «%s»: %s" % (word, ", ".join(notes))))

        elif present == {"LAT"}:
            if (cyr_share >= 0.6 and len(word) >= 3 and word.isalpha() and not word.isupper()
                    and all(c in LAT2CYR for c in word)):
                finds.append(("lookalike", "maybe", idx, span,
                              "«%s» целиком из латинских букв, похожих на русские" % word))

        elif present == {"CYR"}:
            digs = [c for c in word if c.isdigit()]
            ncyr = sum(1 for x in scr if x == "CYR")
            if (len(digs) == 1 and digs[0] in "036" and ncyr >= 2
                    and not (word[0].isdigit() and ncyr <= 2)):
                pos = [i for i, c in zip(idx, word) if c.isdigit()]
                finds.append(("digit", "maybe", pos, span, "цифра %s внутри слова «%s»" % (digs[0], word)))
            for i, c in zip(idx, word):
                if script_of(c) == "CYR" and c.lower() not in RUSSIAN and c not in allowed:
                    finds.append(("foreign", "warn", [i], span,
                                  "%s (%s) - не русская буква" % (c, unicodedata.name(c, "?"))))
    return finds, plan


def apply_plan(text, plan):
    return "".join(plan.get(i, ch) for i, ch in enumerate(text))


def show_char(ch):
    cp = ord(ch)
    if invisible(ch) or odd_space(ch) or cp in NBSPS:
        return "\u27e6U+%04X\u27e7" % cp
    return " " if ch in "\t\n\r" else ch


def snippet(text, span, marks, pad=25):
    lo, hi = max(0, span[0] - pad), min(len(text), span[1] + pad)
    ms = set(marks)
    out = []
    for i in range(lo, hi):
        piece = show_char(text[i])
        out.append("\u27e8%s\u27e9" % piece if (i in ms and not piece.startswith("\u27e6")) else piece)
    return ("\u2026" if lo > 0 else "") + "".join(out) + ("\u2026" if hi < len(text) else "")


# ----------------------------------------------------------------------------------------------
# Чтение файлов
# ----------------------------------------------------------------------------------------------
DOCX_PARTS = re.compile(r"^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")
P_RE = re.compile(r"<w:p(?:\s[^>]*)?(?<!/)>.*?</w:p>", re.S)
R_RE = re.compile(r"<w:r(?:\s[^>]*)?(?<!/)>.*?</w:r>", re.S)
T_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)", re.S)
TOKEN_RE = re.compile(r"<w:t(?:\s[^>]*)?>(.*?)</w:t>|<w:tab/>|<w:br/>|<w:noBreakHyphen/>", re.S)
RPR_RE = re.compile(r"<w:rPr>.*?</w:rPr>", re.S)
VANISH_RE = re.compile(r'<w:vanish(?:\s+w:val="(?!0"|false"|off")[^"]*")?\s*/>')
WHITE_RE = re.compile(r'<w:color\s[^>]*w:val="(?i:ffffff)"')
SZ_RE = re.compile(r'<w:sz\s+w:val="(\d+)"')
PART_LABEL = {"document": "", "footnotes": "сноска", "endnotes": "концевая сноска",
              "header": "колонтитул", "footer": "колонтитул"}


class Para:
    __slots__ = ("label", "text", "hidden")

    def __init__(self, label, text, hidden):
        self.label, self.text, self.hidden = label, text, hidden


def run_text(rxml):
    parts = []
    for m in TOKEN_RE.finditer(rxml):
        if m.group(1) is not None:
            parts.append(html.unescape(m.group(1)))
        elif m.group(0) == "<w:tab/>":
            parts.append("\t")
        elif m.group(0) == "<w:br/>":
            parts.append("\n")
        else:
            parts.append("-")
    return "".join(parts)


def load_docx(path):
    paras = []
    with zipfile.ZipFile(path) as z:
        names = sorted((n for n in z.namelist() if DOCX_PARTS.match(n)),
                       key=lambda n: (n != "word/document.xml", n))
        for name in names:
            kind = re.sub(r"\d+$", "", os.path.basename(name)[:-4])
            label = PART_LABEL.get(kind, "")
            xml = z.read(name).decode("utf-8", "replace")
            for pm in P_RE.finditer(xml):
                texts, hidden = [], []
                for rm in R_RE.finditer(pm.group(0)):
                    rxml = rm.group(0)
                    t = run_text(rxml)
                    texts.append(t)
                    if not t.strip():
                        continue
                    rpr = RPR_RE.search(rxml)
                    if not rpr:
                        continue
                    props = rpr.group(0)
                    if VANISH_RE.search(props):
                        hidden.append((t, "атрибут «скрытый текст»"))
                    if WHITE_RE.search(props):
                        hidden.append((t, "белый цвет шрифта (проверьте, не на белом ли фоне)"))
                    sz = SZ_RE.search(props)
                    if sz and int(sz.group(1)) <= 4:
                        hidden.append((t, "размер шрифта %s пт" % (int(sz.group(1)) / 2)))
                paras.append(Para(label, "".join(texts), hidden))
    return paras


def read_plain(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "replace")


OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)  # Windows: без мигающей консоли


def sniff(path):
    """Определяет формат по содержимому: docx / doc / pdf / text / иное."""
    with open(path, "rb") as f:
        head = f.read(4096)
    if head.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(path) as z:
                return "docx" if "word/document.xml" in z.namelist() else "zip"
        except zipfile.BadZipFile:
            return "zip"
    if head.startswith(OLE_MAGIC):
        return "doc"
    if b"%PDF" in head[:1024]:
        return "pdf"
    if head.startswith(b"{\\rtf"):
        return "rtf"
    if b"\x00" in head:
        return "binary"
    return "text"


# ---- .doc (Word 97-2003): разбор контейнера OLE2 и таблицы кусков текста, только stdlib ----------
class Ole:
    END = 0xFFFFFFFA  # значения >= END - служебные (конец цепочки, свободный сектор)

    def __init__(self, data):
        if data[:8] != OLE_MAGIC:
            raise ValueError("не OLE-файл")
        self.data = data
        u16 = lambda o: struct.unpack_from("<H", data, o)[0]
        u32 = lambda o: struct.unpack_from("<I", data, o)[0]
        self.sec = 1 << u16(30)
        self.msec = 1 << u16(32)
        n_fat, dir_start, self.cutoff = u32(44), u32(48), u32(56)
        minifat_start, difat_start, n_difat = u32(60), u32(68), u32(72)
        difat = list(struct.unpack_from("<109I", data, 76))
        s = difat_start
        for _ in range(n_difat):
            if s >= self.END:
                break
            vals = struct.unpack_from("<%dI" % (self.sec // 4), data, (s + 1) * self.sec)
            difat.extend(vals[:-1])
            s = vals[-1]
        self.fat = []
        for s in [x for x in difat if x < self.END][:n_fat]:
            self.fat.extend(struct.unpack_from("<%dI" % (self.sec // 4), data, (s + 1) * self.sec))
        d = self._chain_bytes(dir_start)
        self.entries = {}
        self.root = None
        for off in range(0, len(d) - 127, 128):
            nlen = struct.unpack_from("<H", d, off + 64)[0]
            typ = d[off + 66]
            name = d[off:off + max(0, nlen - 2)].decode("utf-16le", "replace")
            start = struct.unpack_from("<I", d, off + 116)[0]
            size = struct.unpack_from("<I", d, off + 120)[0]
            if typ == 5 and self.root is None:
                self.root = (start, size)
            elif typ == 2 and name not in self.entries:
                self.entries[name] = (start, size)
        self.minifat = []
        if minifat_start < self.END:
            mf = self._chain_bytes(minifat_start)
            self.minifat = list(struct.unpack_from("<%dI" % (len(mf) // 4), mf, 0))
        self._ministream = None

    def _chain(self, start, table):
        out, s, guard = [], start, 0
        while s < self.END and guard <= len(table):
            out.append(s)
            s = table[s] if s < len(table) else self.END
            guard += 1
        return out

    def _chain_bytes(self, start):
        return b"".join(self.data[(s + 1) * self.sec:(s + 2) * self.sec] for s in self._chain(start, self.fat))

    def stream(self, name):
        if name not in self.entries:
            raise KeyError(name)
        start, size = self.entries[name]
        if size < self.cutoff and self.root:
            if self._ministream is None:
                self._ministream = self._chain_bytes(self.root[0])[:self.root[1]]
            chunks = [self._ministream[s * self.msec:(s + 1) * self.msec]
                      for s in self._chain(start, self.minifat)]
            return b"".join(chunks)[:size]
        return self._chain_bytes(start)[:size]


def doc_raw_text(data):
    """Основной текст .doc (без сносок и колонтитулов) как строка с управляющими символами Word."""
    ole = Ole(data)
    wd = ole.stream("WordDocument")
    u16 = lambda b, o: struct.unpack_from("<H", b, o)[0]
    u32 = lambda b, o: struct.unpack_from("<I", b, o)[0]
    if u16(wd, 0) != 0xA5EC:
        raise ValueError("не документ Word")
    if u16(wd, 2) < 0x00C1:
        raise ValueError("слишком старый формат Word (до 97)")
    flags = u16(wd, 0x0A)
    if flags & 0x0100:
        raise ValueError("документ зашифрован")
    table = ole.stream("1Table" if flags & 0x0200 else "0Table")
    pos = 32
    csw = u16(wd, pos)
    pos += 2 + csw * 2
    cslw = u16(wd, pos)
    lw_start = pos + 2
    pos = lw_start + cslw * 4
    blob = pos + 2
    ccp_text = u32(wd, lw_start + 3 * 4)
    fc_clx, lcb_clx = u32(wd, blob + 33 * 8), u32(wd, blob + 33 * 8 + 4)
    clx = table[fc_clx:fc_clx + lcb_clx]
    p = 0
    while p < len(clx) and clx[p] == 1:
        p += 3 + u16(clx, p + 1)
    if p >= len(clx) or clx[p] != 2:
        raise ValueError("не найдена таблица кусков текста")
    lcb = u32(clx, p + 1)
    plc = clx[p + 5:p + 5 + lcb]
    n = (lcb - 4) // 12
    cps = struct.unpack_from("<%dI" % (n + 1), plc, 0)
    parts = []
    for i in range(n):
        if cps[i] >= ccp_text:
            break
        cnt = min(cps[i + 1], ccp_text) - cps[i]
        fc = u32(plc, (n + 1) * 4 + i * 8 + 2)
        if fc & 0x40000000:
            fc = (fc & 0x3FFFFFFF) // 2
            parts.append(wd[fc:fc + cnt].decode("cp1252", "replace"))
        else:
            parts.append(wd[fc:fc + cnt * 2].decode("utf-16le", "replace"))
    return "".join(parts)


def doc_paragraphs(raw):
    out, stack = [], []  # stack: True - мы внутри кода поля (HYPERLINK, TOC, PAGEREF...), его пропускаем
    for ch in raw:
        if ch == "\x13":
            stack.append(True)
        elif ch == "\x14":
            if stack:
                stack[-1] = False
        elif ch == "\x15":
            if stack:
                stack.pop()
        elif not any(stack):
            out.append(ch)
    s = "".join(out)
    s = s.replace("\x07", "\r").replace("\x0c", "\r").replace("\x0b", " ").replace("\x1e", "-")
    s = "".join(c for c in s if c in "\r\t" or ord(c) >= 0x20)
    return s.split("\r")


def find_soffice():
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def load_doc(path):
    problem = None
    try:
        with open(path, "rb") as f:
            paras = [Para("", t, []) for t in doc_paragraphs(doc_raw_text(f.read()))]
        if any(p.text.strip() for p in paras):
            return paras
        problem = "текст не найден"
    except Exception as e:  # нестандартный файл: пробуем LibreOffice
        problem = str(e)
    exe = find_soffice()
    if exe:
        tmp = tempfile.mkdtemp(prefix="textcheck_")
        try:
            subprocess.run([exe, "--headless", "--convert-to", "docx", "--outdir", tmp, path],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=240,
                           creationflags=NO_WINDOW)
            conv = os.path.join(tmp, os.path.splitext(os.path.basename(path))[0] + ".docx")
            if os.path.exists(conv):
                return load_docx(conv)
        except Exception as e:
            problem = "%s; LibreOffice: %s" % (problem, e)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    raise SystemExit("Не удалось прочитать .doc (%s). Откройте файл в Word и сохраните как .docx." % problem)


# ---- .pdf ----------------------------------------------------------------------------------------
LIGATURES = {"\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl"}
LEADER_RE = re.compile(r"(?:\.{4,}|…{2,}|(?:\.\s){4,})\s*\d+\s*$")
NUMBERED_RE = re.compile(r"^\d+(?:\.\d+)*\.?\s+\S")
BULLET_RE = re.compile(r"^(?:[–—•\-]\s|\d+[.)]\s)")
PAGE_NO_RE = re.compile(r"^[-–—\s]*\d{1,4}[-–—\s]*$")
STRUCT_HEADS = {"введение", "заключение", "содержание", "оглавление", "реферат", "аннотация", "выводы"}
TERMINAL = (".", "!", "?", "…", ":", ";", "»", "”", '"', ")")


def pdf_page_lines(path):
    """Возвращает (страницы, имя_библиотеки); страница - список строк. Порядок: PyMuPDF, pdftotext, pypdf.
    pypdf поставлен последним: на PDF с выключкой по ширине он нередко склеивает слова без пробелов."""
    tried = []
    try:
        import fitz  # PyMuPDF
        with fitz.open(path) as doc:
            if getattr(doc, "needs_pass", False):
                raise SystemExit("PDF защищён паролем: снимите защиту и повторите.")
            return [pg.get_text("text").splitlines() for pg in doc], "pymupdf"
    except ImportError:
        tried.append("pymupdf не установлен")
    except SystemExit:
        raise
    except Exception as e:
        tried.append("pymupdf: %s" % e)
    exe = shutil.which("pdftotext")
    if exe:
        try:
            r = subprocess.run([exe, "-enc", "UTF-8", path, "-"], capture_output=True, timeout=300,
                               creationflags=NO_WINDOW)
            if r.returncode == 0:
                pages = r.stdout.decode("utf-8", "replace").split("\f")
                if pages and not pages[-1].strip():
                    pages.pop()
                return [pg.splitlines() for pg in pages], "pdftotext"
            tried.append("pdftotext: код %d" % r.returncode)
        except Exception as e:
            tried.append("pdftotext: %s" % e)
    else:
        tried.append("pdftotext не найден")
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise SystemExit("PDF защищён паролем: снимите защиту и повторите.")
        return [(pg.extract_text() or "").splitlines() for pg in reader.pages], "pypdf"
    except ImportError:
        tried.append("pypdf не установлен")
    except SystemExit:
        raise
    except Exception as e:
        tried.append("pypdf: %s" % e)
    raise SystemExit("Не удалось прочитать PDF (%s).\nУстановите библиотеку: pip install pymupdf  "
                     "(или pip install pypdf), либо Poppler с командой pdftotext." % "; ".join(tried))


def strip_running_lines(pages):
    """Убирает номера страниц и повторяющиеся колонтитулы (в первых/последних двух строках страниц)."""
    cnt = collections.Counter()
    edge = []
    for lines in pages:
        idx = [i for i, l in enumerate(lines) if l.strip()]
        keep = set(idx[:2] + idx[-2:])
        edge.append(keep)
        cnt.update({re.sub(r"\d+", "#", lines[i].strip()) for i in keep})
    thr = max(3, int(len(pages) * 0.4))
    repeated = {k for k, v in cnt.items() if v >= thr and len(k) < 120}
    out = []
    for lines, keep in zip(pages, edge):
        out.append([l for i, l in enumerate(lines)
                    if not (i in keep and (PAGE_NO_RE.match(l) or re.sub(r"\d+", "#", l.strip()) in repeated))])
    return out


def is_heading(line):
    t = line.strip()
    if not t:
        return False
    if LEADER_RE.search(t):
        return True
    h = norm_head(t)
    if h in STRUCT_HEADS or (len(h.split()) <= 12 and END_RE.match(h)):
        return True
    low = t.lower()
    if re.match(r"^(?:глава|раздел|часть)\s+\d", low):
        return True
    return len(t) <= 90 and bool(NUMBERED_RE.match(t)) and not t.endswith((".", ",", ";", ":"))


def join_lines(lines):
    out = lines[0]
    for nxt in lines[1:]:
        if out.endswith("\u00ad"):
            out = out[:-1] + nxt
        elif re.search(r"[^\W\d_][-\u2010\u2011]$", out) and nxt[:1].islower():
            out = out[:-1] + nxt  # перенос слова в конце строки
        else:
            out += " " + nxt
    return out


def pdf_paragraphs(pages):
    pages = strip_running_lines(pages)
    flat = []
    for pn, lines in enumerate(pages, 1):
        for l in lines:
            for lig, rep in LIGATURES.items():
                l = l.replace(lig, rep)
            flat.append((pn, l.strip()))
    lens = sorted(len(l) for _, l in flat if len(l) > 30)
    full = lens[int(len(lens) * 0.75)] if lens else 80
    paras, cur, cur_page = [], [], 1

    def flush():
        nonlocal cur
        if cur:
            paras.append((cur_page, join_lines(cur)))
            cur = []

    for pn, s in flat:
        if not s:
            flush()
            continue
        if is_heading(s):
            flush()
            paras.append((pn, s))
            continue
        if cur and ((cur[-1].endswith(TERMINAL) and len(cur[-1]) < 0.7 * full) or BULLET_RE.match(s)):
            flush()
        if not cur:
            cur_page = pn
        cur.append(s)
    flush()
    return paras


def load_pdf(path):
    pages, backend = pdf_page_lines(path)
    chars = sum(len(l) for pg in pages for l in pg)
    if not pages or chars < 100 * max(1, len(pages)) * 0.2:
        raise SystemExit("В PDF нет текстового слоя (похоже на скан). Распознайте его (OCR) в FineReader, "
                         "OCRmyPDF или Adobe Acrobat и повторите.")
    bad = sum(1 for pg in pages for l in pg for c in l if c == "\ufffd" or unicodedata.category(c) == "Co")
    if bad > 0.02 * chars:
        print("ВНИМАНИЕ: текстовый слой PDF повреждён (%d%% нечитаемых символов), результаты ненадёжны. "
              "Лучше проверять исходный .docx." % round(bad * 100 / chars), file=sys.stderr)
    sample = [w for pg in pages[:20] for l in pg for w in l.split()]
    if sample and sum(len(w) for w in sample) / len(sample) > 11:
        print("ВНИМАНИЕ: слова в извлечённом тексте PDF слиплись (читалка: %s). Установите  pip install pymupdf  "
              "или Poppler (pdftotext) - иначе проверка будет неточной." % backend, file=sys.stderr)
    return [Para("стр. %d" % pn, t, []) for pn, t in pdf_paragraphs(pages)]


def load_any(path):
    kind = sniff(path)
    if kind == "docx":
        return load_docx(path)
    if kind == "doc":
        return load_doc(path)
    if kind == "pdf":
        return load_pdf(path)
    if kind == "text":
        return [Para("", line, []) for line in read_plain(path).splitlines()]
    raise SystemExit("Формат файла «%s» не поддерживается (%s). Сохраните как .docx, .pdf или .txt." %
                     (os.path.basename(path), kind))


# ----------------------------------------------------------------------------------------------
# audit
# ----------------------------------------------------------------------------------------------
KINDS = [
    ("invis", "Невидимые символы (нулевой ширины, мягкий перенос, управляющие)"),
    ("space", "Необычные пробелы"),
    ("mixed", "Смешение алфавитов в слове: вероятная подмена букв"),
    ("foreign", "Нерусские буквы кириллицы (і, ї, ѕ и т.п.)"),
    ("lookalike", "Слова целиком латиницей из букв, похожих на русские (возможная подмена)"),
    ("digit", "Цифры вместо букв внутри слов (0 вместо о, 3 вместо з)"),
    ("hidden", "Скрытый текст в docx"),
]
SERIOUS = {"invis", "space", "mixed", "foreign", "hidden"}


def fix_xml(xml, allowed):
    total = 0

    def fix_par(pm):
        nonlocal total
        pxml = pm.group(0)
        tms = list(T_RE.finditer(pxml))
        if not tms:
            return pxml
        texts = [html.unescape(m.group(2)) for m in tms]
        _, plan = analyze("".join(texts), allowed)
        if not plan:
            return pxml
        total += len(plan)
        out, last, pos = [], 0, 0
        for m, t in zip(tms, texts):
            new = "".join(plan.get(pos + k, ch) for k, ch in enumerate(t))
            pos += len(t)
            out.append(pxml[last:m.start(2)])
            out.append(xml_escape(new) if new != t else m.group(2))
            last = m.end(2)
        out.append(pxml[last:])
        return "".join(out)

    return P_RE.sub(fix_par, xml), total


def fix_file(src, allowed):
    base, ext = os.path.splitext(src)
    kind = sniff(src)
    if kind == "docx":
        dst = base + ".clean.docx"
        changed = 0
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if DOCX_PARTS.match(item.filename):
                    xml, n = fix_xml(data.decode("utf-8"), allowed)
                    changed += n
                    data = xml.encode("utf-8")
                zout.writestr(item, data, compress_type=item.compress_type)
    else:
        # .doc и .pdf нельзя аккуратно переписать на месте: сохраняем очищенный текст
        dst = base + ".clean" + (ext if kind == "text" else ".txt")
        lines, changed = [], 0
        for p in load_any(src):
            _, plan = analyze(p.text, allowed)
            changed += len(plan)
            lines.append(apply_plan(p.text, plan))
        with open(dst, "w", encoding="utf-8", newline="\n") as f:
            f.write("\n".join(lines) + "\n")
    return dst, changed


def run_audit(path, limit, fix, allowed):
    paras = load_any(path)
    found = collections.defaultdict(list)
    nbsp = 0
    for no, p in enumerate(paras, 1):
        finds, _ = analyze(p.text, allowed)
        for kind, level, marks, span, detail in finds:
            if kind == "nbsp":
                nbsp += 1
            else:
                found[kind].append((no, p, span, marks, detail))
        for t, why in p.hidden:
            found["hidden"].append((no, p, None, [], "«%s» - %s" % (t.strip()[:60], why)))

    emit("Файл: %s (формат: %s)" % (path, sniff(path)))
    emit("Абзацев: %d" % len(paras))
    if sniff(path) in ("doc", "pdf"):
        emit("Примечание: в .doc и .pdf проверяются только символы; скрытый текст не определяется.")
    emit("")
    for kind, title in KINDS:
        emit("%-82s %d" % (title + ":", len(found[kind])))
    emit("%-82s %d" % ("Неразрывные пробелы (обычно норма, не считаются нарушением):", nbsp))
    emit("")

    for kind, title in KINDS:
        items = found[kind]
        if not items:
            continue
        emit("=== %s ===" % title)
        for no, p, span, marks, detail in items[:limit]:
            where = "абз. %d%s" % (no, " (%s)" % p.label if p.label else "")
            ctx = snippet(p.text, span, marks) if span else ""
            emit("  %s | %s%s" % (where, ctx + "  |  " if ctx else "", detail))
        if len(items) > limit:
            emit("  ... и ещё %d (увеличьте --limit)" % (len(items) - limit))
        emit("")

    serious = sum(len(found[k]) for k in SERIOUS)
    if serious:
        emit("ИТОГ: найдено подозрительных мест: %d." % serious)
    else:
        emit("ИТОГ: невидимых символов, подмен и скрытого текста не найдено.")
    emit("Значки: \u27e6U+XXXX\u27e7 - невидимый символ, \u27e8x\u27e9 - подозрительная буква.")

    if fix:
        dst, changed = fix_file(path, allowed)
        emit("")
        emit("Исправленная копия: %s (правок символов: %d). Скрытый текст не удалялся." % (dst, changed))
    elif serious:
        emit("Чтобы получить исправленную копию, добавьте --fix.")
    STATS["serious"] = serious
    return 1 if serious else 0


# ----------------------------------------------------------------------------------------------
# compare
# ----------------------------------------------------------------------------------------------
def tokens_orig(text):
    return re.findall(r"[^\W_]+", text)


def norm_word(w):
    return w.lower().replace("ё", "е")


def words_of(text):
    return [norm_word(w) for w in tokens_orig(text)]


def norm_head(t):
    return re.sub(r"[\d.\s_…·\-–—:]+", " ", t.lower().replace("ё", "е")).strip()


END_RE = re.compile(
    r"^(?:список (?:использованн\w+ )?(?:источник\w*|литератур\w*)(?: и литератур\w*)?"
    r"|библиографическ\w+ список|библиографи\w+|литература|приложени\w+(?: [а-яa-z])?)\b")

SENT_SPLIT = re.compile(r"(?<=[.!?…])\s+(?=[А-ЯЁA-Z«\"“(\[0-9])")
LEGAL_CS = re.compile(
    r"\b(?:ГОСТ|СНиП|СанПиН|ФЗ|ФКЗ|ТК РФ|ГК РФ|УК РФ|НК РФ|КоАП|ГПК|АПК|УПК|СК РФ|ЖК РФ|ЗК РФ|"
    r"ФСБУ|МСФО|ПБУ|ОКВЭД|ОКПД)\b|№\s*\d")
LEGAL_CI = re.compile(
    r"федеральн\w+\s+(?:конституционн\w+\s+)?закон|конституци\w+|кодекс\w*|"
    r"постановлени\w+\s+(?:правительства|пленума)|\bст\.\s*\d|\bстат(?:ья|ьи|ье|ью|ьёй|ей)\s+\d",
    re.I)
QUOTE_RE = re.compile(r"«[^»]*»|“[^”]*”|\"[^\"]*\"")


def count_words(paras):
    return sum(len(words_of(p)) for p in paras)


def select_body(texts):
    """Режет титул+содержание и список литературы+приложения. Возвращает (тело, до, после, начало найдено)."""
    start, found_start = 0, False
    for i, p in enumerate(texts):
        if norm_head(p) == "введение":
            nxt = next((q for q in texts[i + 1:i + 6] if q.strip()), "")
            if len(words_of(nxt)) >= 15:  # дальше идёт связный текст, а не строка содержания
                start, found_start = i + 1, True
                break
    min_pos = start if found_start else int(len(texts) * 0.4)
    end = len(texts)
    for i in range(min_pos, len(texts)):
        h = norm_head(texts[i])
        if h and len(h.split()) <= 12 and len(h) <= 100 and END_RE.match(h):
            end = i
            break
    return texts[start:end], count_words(texts[:start]), count_words(texts[end:]), found_start


def strip_legal(paras):
    kept, nsent, nwords = [], 0, 0
    for p in paras:
        keep = []
        for s in SENT_SPLIT.split(p):
            if LEGAL_CS.search(s) or LEGAL_CI.search(s):
                nsent += 1
                nwords += len(words_of(s))
            else:
                keep.append(s)
        kept.append(" ".join(keep))
    return kept, nsent, nwords


def strip_quotes(paras):
    kept, nwords = [], 0
    for p in paras:
        for q in QUOTE_RE.findall(p):
            nwords += len(words_of(q))
        kept.append(QUOTE_RE.sub(" ", p))
    return kept, nwords


def collect_sources(args_sources, target):
    files = []
    tgt = os.path.abspath(target)
    for s in args_sources:
        if os.path.isdir(s):
            for root, _, names in os.walk(s):
                for n in sorted(names):
                    if os.path.splitext(n)[1].lower() in (".docx", ".doc", ".pdf", ".txt", ".md"):
                        files.append(os.path.join(root, n))
        elif os.path.isfile(s):
            files.append(s)
    return [f for f in files if os.path.abspath(f) != tgt]


def clean_para_texts(paras, allowed):
    out = []
    for p in paras:
        text = p.text if isinstance(p, Para) else p
        _, plan = analyze(text, allowed)
        out.append(apply_plan(text, plan))
    return out


# ----------------------------------------------------------------------------------------------
# Сопоставление: индексируем фрагменты работы, а источники пропускаем потоком
# ----------------------------------------------------------------------------------------------
class Matcher:
    def __init__(self, words, pid, k):
        self.words, self.pid, self.k, self.n = words, pid, k, len(words)
        self.pos_by_hash = collections.defaultdict(list)
        for i in range(self.n - k + 1):
            if pid[i] == pid[i + k - 1]:
                self.pos_by_hash[hash(" ".join(words[i:i + k]))].append(i)
        self.covered = [False] * self.n
        self.pos_src = collections.defaultdict(set)
        self.src_cov = collections.defaultdict(set)
        self.names = []

    def add_source(self, name, texts):
        sid = len(self.names)
        self.names.append(name)
        k, hits = self.k, 0
        for text in texts:
            w = words_of(text)
            for i in range(len(w) - k + 1):
                for pos in self.pos_by_hash.get(hash(" ".join(w[i:i + k])), ()):
                    for j in range(pos, pos + k):
                        self.covered[j] = True
                        self.pos_src[j].add(sid)
                    self.src_cov[sid].update(range(pos, pos + k))
                    hits += 1
        return hits

    def covered_share(self, i, length):
        return sum(self.covered[i:i + length]) / float(length)


# ----------------------------------------------------------------------------------------------
# Интернет
# ----------------------------------------------------------------------------------------------
DEFAULT_EXCLUDE_DOMAINS = ("consultant.ru", "garant.ru", "pravo.gov.ru", "cntd.ru", "kodeks.ru")
UA = "Mozilla/5.0 (compatible; textcheck/1.0; personal plagiarism self-check)"


class SearchError(Exception):
    pass


def short_err(e):
    if isinstance(e, urllib.error.HTTPError):
        return "HTTP %d" % e.code
    if isinstance(e, urllib.error.URLError):
        return "сеть: %s" % getattr(e, "reason", e)
    return "%s: %s" % (type(e).__name__, str(e)[:80])


def http_get(url, headers=None, data=None, timeout=20, max_bytes=15000000):
    h = {"User-Agent": UA, "Accept-Encoding": "gzip", "Accept-Language": "ru,en;q=0.7"}
    h.update(headers or {})
    req = urllib.request.Request(url, data=data, headers=h)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        raw = r.read(max_bytes + 1)
        if len(raw) > max_bytes:
            raise ValueError("файл больше %d МБ" % (max_bytes // 1000000))
        if "gzip" in (r.headers.get("Content-Encoding") or "").lower():
            raw = gzip.GzipFile(fileobj=io.BytesIO(raw)).read(max_bytes * 2)
        return raw, r.headers, r.geturl()


def decode_bytes(raw, content_type):
    cands = []
    m = re.search(r"charset=([\w\-]+)", content_type or "", re.I)
    if m:
        cands.append(m.group(1))
    m = re.search(r"<meta[^>]+charset=[\"']?([\w\-]+)", raw[:4096].decode("ascii", "ignore"), re.I)
    if m:
        cands.append(m.group(1))
    cands += ["utf-8", "cp1251"]
    for c in cands:
        try:
            return raw.decode(c)
        except (LookupError, UnicodeDecodeError):
            continue
    return raw.decode("utf-8", "replace")


class _TextExtractor(HTMLParser):
    SKIP = {"script", "style", "noscript", "svg", "template", "head", "iframe"}
    BLOCK = {"p", "div", "br", "li", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "table", "section",
             "article", "blockquote", "pre", "main", "dd", "dt", "td", "th", "nav", "footer", "header", "aside"}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.skip, self.buf, self.paras = 0, [], []

    def flush(self):
        t = " ".join("".join(self.buf).split())
        if t:
            self.paras.append(t)
        self.buf = []

    def handle_starttag(self, tag, attrs):
        if tag == "body":
            self.skip = 0  # защита от незакрытого <head>
        if tag in self.SKIP:
            self.skip += 1
        elif tag in self.BLOCK:
            self.flush()

    def handle_endtag(self, tag):
        if tag in self.SKIP:
            self.skip = max(0, self.skip - 1)
        elif tag in self.BLOCK:
            self.flush()

    def handle_data(self, data):
        if not self.skip:
            self.buf.append(data)


def html_to_paragraphs(page):
    ex = _TextExtractor()
    ex.feed(page)
    ex.close()
    ex.flush()
    return [p for p in ex.paras if len(words_of(p)) >= 3]


def bytes_to_paragraphs(raw, content_type, url):
    """Текст страницы по её байтам: HTML, PDF, DOC/DOCX или простой текст."""
    ct = (content_type or "").lower()
    head = raw[:1024]
    kind = None
    if raw[:5] == b"%PDF-" or b"%PDF" in head[:1024] and "pdf" in ct:
        kind = ".pdf"
    elif raw.startswith(OLE_MAGIC):
        kind = ".doc"
    elif raw.startswith(b"PK\x03\x04"):
        kind = ".docx"
    if kind:
        tmp = tempfile.mkdtemp(prefix="textcheck_")
        try:
            path = os.path.join(tmp, "page" + kind)
            with open(path, "wb") as f:
                f.write(raw)
            return [p.text for p in load_any(path)]
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    text = decode_bytes(raw[:3000000], ct)
    if "html" in ct or re.search(r"<(?:html|body|p|div)\b", text[:5000], re.I):
        return html_to_paragraphs(text)
    return [ln for ln in text.splitlines() if ln.strip()]


class Cache:
    def __init__(self, directory, enabled=True, ttl_days=14):
        self.dir, self.enabled, self.ttl = directory, enabled, ttl_days * 86400
        if enabled:
            os.makedirs(directory, exist_ok=True)

    def _path(self, kind, key):
        return os.path.join(self.dir, "%s_%s.json" % (kind, hashlib.sha1(key.encode("utf-8")).hexdigest()))

    def get(self, kind, key):
        if not self.enabled:
            return None
        p = self._path(kind, key)
        try:
            if time.time() - os.path.getmtime(p) > self.ttl:
                return None
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except (OSError, ValueError):
            return None

    def put(self, kind, key, value):
        if self.enabled:
            try:
                with open(self._path(kind, key), "w", encoding="utf-8") as f:
                    json.dump(value, f, ensure_ascii=False)
            except OSError:
                pass


class Provider:
    name = "?"
    delay = 1.0

    def search(self, query, n):
        raise NotImplementedError


class DuckDuckGo(Provider):
    name = "ddg"
    delay = 3.0

    def __init__(self, url="https://html.duckduckgo.com/html/"):
        self.url = url

    def search(self, query, n):
        try:
            raw, headers, _ = http_get(self.url, data=urllib.parse.urlencode({"q": query, "kl": "ru-ru"}).encode(),
                                       headers={"Content-Type": "application/x-www-form-urlencoded"})
        except Exception as e:
            raise SearchError(short_err(e))
        page = decode_bytes(raw, headers.get("Content-Type"))
        urls = []
        for m in re.finditer(r"<a\b[^>]*>", page):
            tag = m.group(0)
            if "result__a" not in tag:
                continue
            hm = re.search(r'href="([^"]+)"', tag)
            if not hm:
                continue
            href = html.unescape(hm.group(1))
            if href.startswith("//"):
                href = "https:" + href
            q = urllib.parse.parse_qs(urllib.parse.urlparse(href).query)
            urls.append(q["uddg"][0] if "uddg" in q else href)
        if not urls and re.search(r"anomaly|captcha|challenge|unusual traffic|robot", page, re.I):
            raise SearchError("DuckDuckGo требует проверку (подозрение на бота). Подождите или возьмите другой --engine")
        return urls[:n]


class YandexAPI(Provider):
    name = "yandex"
    delay = 0.5

    def __init__(self, api_key, folder_id, url="https://searchapi.api.cloud.yandex.net/v2/web/search"):
        self.key, self.folder, self.url = api_key, folder_id, url

    def search(self, query, n):
        body = json.dumps({"query": {"searchType": "SEARCH_TYPE_RU", "queryText": query},
                           "folderId": self.folder, "responseFormat": "FORMAT_XML"}).encode("utf-8")
        try:
            raw, _, _ = http_get(self.url, data=body, headers={"Authorization": "Api-Key " + self.key,
                                                                "Content-Type": "application/json"})
            xml = base64.b64decode(json.loads(raw)["rawData"]).decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            raise SearchError("Yandex Search API: HTTP %d %s" % (e.code, e.read()[:200].decode("utf-8", "replace")))
        except Exception as e:
            raise SearchError(short_err(e))
        urls = [html.unescape(u.strip()) for u in re.findall(r"<url>(.*?)</url>", xml, re.S)]
        return urls[:n]


class BraveAPI(Provider):
    name = "brave"
    delay = 1.1  # бесплатный тариф: не чаще 1 запроса в секунду

    def __init__(self, api_key, url="https://api.search.brave.com/res/v1/web/search"):
        self.key, self.url = api_key, url

    def search(self, query, n):
        q = urllib.parse.urlencode({"q": query, "count": min(20, max(n, 1)), "country": "ru", "search_lang": "ru"})
        try:
            raw, _, _ = http_get(self.url + "?" + q, headers={"X-Subscription-Token": self.key,
                                                              "Accept": "application/json"})
            data = json.loads(raw)
        except urllib.error.HTTPError as e:
            raise SearchError("Brave API: HTTP %d" % e.code)
        except Exception as e:
            raise SearchError(short_err(e))
        return [r["url"] for r in (data.get("web") or {}).get("results", []) if r.get("url")][:n]


class SearXNG(Provider):
    name = "searxng"
    delay = 1.0

    def __init__(self, base_url):
        self.base = base_url.rstrip("/")

    def search(self, query, n):
        q = urllib.parse.urlencode({"q": query, "format": "json", "language": "ru"})
        try:
            raw, _, _ = http_get(self.base + "/search?" + q)
            data = json.loads(raw)
        except urllib.error.HTTPError as e:
            raise SearchError("SearXNG: HTTP %d (в настройках сервера должен быть включён формат json)" % e.code)
        except Exception as e:
            raise SearchError(short_err(e))
        return [r["url"] for r in data.get("results", []) if r.get("url")][:n]


def make_provider(args):
    eng = args.engine
    if eng == "yandex":
        key = args.yandex_key or os.environ.get("YANDEX_SEARCH_API_KEY")
        folder = args.yandex_folder or os.environ.get("YANDEX_FOLDER_ID")
        if not key or not folder:
            raise SystemExit("Для --engine yandex нужны YANDEX_SEARCH_API_KEY и YANDEX_FOLDER_ID "
                             "(или --yandex-key и --yandex-folder).")
        return YandexAPI(key, folder)
    if eng == "brave":
        key = args.brave_key or os.environ.get("BRAVE_API_KEY")
        if not key:
            raise SystemExit("Для --engine brave нужен BRAVE_API_KEY (или --brave-key).")
        return BraveAPI(key)
    if eng == "searxng":
        url = args.searxng_url or os.environ.get("SEARXNG_URL")
        if not url:
            raise SystemExit("Для --engine searxng укажите --searxng-url http://адрес:порт (или SEARXNG_URL).")
        return SearXNG(url)
    return DuckDuckGo()


def window_ok(orig, pid, i, w):
    if pid[i] != pid[i + w - 1]:
        return False
    toks = orig[i:i + w]
    if sum(1 for t in toks if t.isdigit()) > 2:
        return False
    return sum(len(t) for t in toks) / float(w) >= 4.0


def plan_windows(orig, pid, w, budget):
    """Равномерно выбирает начала окон из w слов по всему тексту, не больше budget штук."""
    n = len(orig)
    if n < w or budget < 1:
        return []
    step = max(w, n // budget)
    out, s = [], 0
    while s <= n - w:
        for i in range(s, min(s + step, n - w + 1)):
            if window_ok(orig, pid, i, w):
                out.append(i)
                break
        s += step
    return out[:budget]


def host_excluded(url, domains):
    host = (urllib.parse.urlparse(url).hostname or "").lower()
    return any(host == d or host.endswith("." + d) for d in domains)


def run_web(matcher, orig, opts, provider, cache):
    """Ищет фразы работы в интернете, скачивает найденные страницы и добавляет их в Matcher."""
    w = opts.window
    plan = plan_windows(orig, matcher.pid, w, opts.queries)
    delay = provider.delay if opts.delay is None else opts.delay
    st = {"engine": provider.name, "planned": len(plan), "queries": 0, "cached": 0, "skipped": 0,
          "loaded": 0, "failed": [], "excluded": 0, "aborted": None, "window": w}
    seen, excluded, errors_in_row, last_net = set(), set(), 0, 0.0
    for idx, i in enumerate(plan, 1):
        if CANCEL.is_set():
            st["aborted"] = "остановлено пользователем"
            break
        if len(seen) >= opts.max_pages:
            break
        if matcher.covered_share(i, w) >= 0.6:
            st["skipped"] += 1
            continue
        query = '"%s"' % " ".join(orig[i:i + w])
        key = provider.name + "|" + query
        urls = cache.get("search", key)
        if urls is None:
            wait = delay * random.uniform(0.8, 1.4) - (time.time() - last_net)
            if wait > 0:
                time.sleep(wait)
            try:
                urls = provider.search(query, opts.per_query)
                errors_in_row = 0
            except SearchError as e:
                errors_in_row += 1
                print("  поиск не удался (%d/3): %s" % (errors_in_row, e), file=sys.stderr)
                last_net = time.time()
                if errors_in_row >= 3:
                    st["aborted"] = "3 ошибки подряд: %s" % e
                    break
                continue
            last_net = time.time()
            cache.put("search", key, urls)
            st["queries"] += 1
        else:
            st["cached"] += 1
        got = 0
        for u in urls[:opts.per_query]:
            if len(seen) >= opts.max_pages:
                break
            if u in seen or not u.lower().startswith(("http://", "https://")):
                continue
            if host_excluded(u, opts.exclude_domains):
                excluded.add(u)
                st["excluded"] = len(excluded)
                continue
            seen.add(u)
            paras = cache.get("page", u)
            if paras is None:
                try:
                    raw, headers, _ = http_get(u, timeout=15)
                    paras = bytes_to_paragraphs(raw, headers.get("Content-Type"), u)
                    cache.put("page", u, paras)
                except (Exception, SystemExit) as e:
                    st["failed"].append((u, short_err(e) if isinstance(e, Exception) else str(e)[:80]))
                    continue
            paras = clean_para_texts(paras, frozenset())
            hits = matcher.add_source(u, paras)
            st["loaded"] += 1
            got += 1 if hits else 0
        print("  [%d/%d] запросов: %d, страниц загружено: %d, покрыто текста: %d%%" %
              (idx, len(plan), st["queries"] + st["cached"], st["loaded"],
               100 * sum(matcher.covered) // max(1, matcher.n)), file=sys.stderr)
    return st


# ----------------------------------------------------------------------------------------------
def run_compare(path, sources, k, legal, skip_quotes, limit, allowed, web=None):
    files = collect_sources(sources, path) if sources else []
    if not files and not web:
        raise SystemExit("Источники не найдены: положите .docx/.doc/.pdf/.txt файлы в указанную папку "
                         "или добавьте --web для поиска в интернете.")

    texts = clean_para_texts(load_any(path), allowed)
    total_all = count_words(texts)
    body, w_before, w_after, found_start = select_body(texts)
    ex_legal = ex_quotes = n_sent = 0
    if legal:
        body, n_sent, ex_legal = strip_legal(body)
    if skip_quotes:
        body, ex_quotes = strip_quotes(body)

    words, orig, pid = [], [], []
    for n, p in enumerate(body):
        for t in tokens_orig(p):
            orig.append(t)
            words.append(norm_word(t))
            pid.append(n)
    if len(words) < k:
        raise SystemExit("После исключений в работе почти не осталось текста для проверки.")

    matcher = Matcher(words, pid, k)
    skipped, n_files = [], 0
    for f in files:
        if CANCEL.is_set():
            break
        print("Чтение источника: %s" % os.path.relpath(f), file=sys.stderr)
        try:
            src_texts = clean_para_texts(load_any(f), allowed)
        except (Exception, SystemExit) as e:
            skipped.append("%s - %s" % (os.path.relpath(f), str(e).splitlines()[0] if str(e) else "ошибка"))
            continue
        matcher.add_source(os.path.relpath(f), src_texts)
        n_files += 1
    if files and not n_files and not web:
        raise SystemExit("Ни один источник не удалось прочитать:\n  " + "\n  ".join(skipped))

    wst = None
    if web:
        provider = make_provider(web)
        print("Поиск в интернете (%s): фразы из вашей работы отправляются в поисковую систему." % provider.name,
              file=sys.stderr)
        cache = Cache(web.cache_dir, enabled=not web.no_cache)
        wst = run_web(matcher, orig, web, provider, cache)

    covered, src_cov, pos_src, names = matcher.covered, matcher.src_cov, matcher.pos_src, matcher.names
    matched = sum(covered)
    pct = matched * 100.0 / len(words)
    STATS.update(pct=pct, matched=matched, checked=len(words), web=bool(wst),
                web_aborted=(wst or {}).get("aborted"), web_loaded=(wst or {}).get("loaded", 0),
                web_failed=len((wst or {}).get("failed", [])))

    emit("Файл: %s" % path)
    if files:
        emit("Файлов-источников: %d (длина фрагмента: %d слов)" % (n_files, k))
    for sk in skipped:
        emit("  НЕ ПРОЧИТАН: %s" % sk)
    emit("")
    emit("Слов в работе всего:                         %d" % total_all)
    if found_start:
        emit("  исключено титул + содержание:             %d" % w_before)
    else:
        emit("  заголовок «Введение» не найден: титул и содержание НЕ исключены")
    emit("  исключено список литературы/приложения:   %d" % w_after)
    if legal:
        emit("  исключено предложений со ссылками на законы/ГОСТы: %d (%d слов)" % (n_sent, ex_legal))
    if skip_quotes:
        emit("  исключено слов в кавычках:                %d" % ex_quotes)
    emit("Проверено слов:                              %d" % len(words))
    emit("")
    if wst:
        emit("=== Поиск в интернете (%s) ===" % wst["engine"])
        emit("  фраз выбрано для поиска: %d (по %d слов); отправлено запросов: %d, взято из кэша: %d"
             % (wst["planned"], wst["window"], wst["queries"], wst["cached"]))
        emit("  пропущено, т.к. уже найдено на скачанных страницах: %d" % wst["skipped"])
        emit("  страниц скачано и сравнено: %d; исключено по домену: %d; не скачано: %d"
             % (wst["loaded"], wst["excluded"], len(wst["failed"])))
        for u, why in wst["failed"][:8]:
            emit("    не скачано: %s (%s)" % (u[:90], why))
        if wst["aborted"]:
            emit("  ПОИСК ПРЕРВАН (%s)" % wst["aborted"])
        emit("  Процент ниже - нижняя оценка: искалась выборка фраз, а не весь текст.")
        emit("")
    emit("Совпадения: %d слов = %.1f%%" % (matched, pct))
    emit("Оригинальность (относительно ваших источников%s): %.1f%%" %
         (" и найденных страниц" if wst else "", 100 - pct))
    emit("")

    if src_cov:
        emit("=== Источники по объёму совпадений ===")
        for s_, cov in sorted(src_cov.items(), key=lambda x: -len(x[1]))[:limit]:
            emit("  %5d слов (%4.1f%%)  %s" % (len(cov), len(cov) * 100.0 / len(words), names[s_]))
        emit("")

        emit("=== Совпавшие фрагменты ===")
        shown, i = 0, 0
        while i < len(words):
            if not covered[i]:
                i += 1
                continue
            j = i
            while j < len(words) and covered[j] and pid[j] == pid[i]:
                j += 1
            srcs = collections.Counter()
            for t in range(i, j):
                srcs.update(pos_src[t])
            if shown < limit:
                frag = " ".join(words[i:min(j, i + 40)]) + (" \u2026" if j - i > 40 else "")
                top = ", ".join(names[s_] for s_, _ in srcs.most_common(3))
                emit("  [%d слов] %s" % (j - i, frag))
                emit("      источник: %s" % top)
            shown += 1
            i = j
        if shown > limit:
            emit("  ... и ещё %d фрагментов (увеличьте --limit)" % (shown - limit))
    else:
        emit("Совпадений не найдено.")
    emit("")
    emit("Важно: проверка идёт только по тем источникам, которые вы дали%s. Это не замена системе вуза."
         % (" или которые нашлись в открытом интернете" if wst else ""))
    return 0


# ----------------------------------------------------------------------------------------------
# Окно с кнопками
# ----------------------------------------------------------------------------------------------
class _Writer:
    """Подменяет sys.stderr/sys.stdout на время проверки: строки уходят в окно."""

    def __init__(self, callback):
        self.cb = callback

    def write(self, text):
        for line in str(text).replace("\r", "\n").split("\n"):
            if line.strip():
                self.cb(line.strip())
        return len(text)

    def flush(self):
        pass


def gui_check(path, opts):
    """Полная проверка для окна. opts: web, engine, yandex_key, yandex_folder, brave_key, searxng_url,
    queries, sources (список папок), skip_quotes. Возвращает словарь с отчётом и числами."""
    REPORT.clear()
    STATS.clear()
    bar = "=" * 70
    emit(bar)
    emit("ПРОВЕРКА 1 из 2: скрытые символы и подмена букв")
    emit(bar)
    run_audit(path, 40, False, frozenset())
    emit("")
    emit(bar)
    emit("ПРОВЕРКА 2 из 2: совпадения с источниками")
    emit(bar)
    if opts.get("web") or opts.get("sources"):
        web = None
        if opts.get("web"):
            web = argparse.Namespace(
                engine=opts.get("engine", "ddg"), yandex_key=opts.get("yandex_key"),
                yandex_folder=opts.get("yandex_folder"), brave_key=opts.get("brave_key"),
                searxng_url=opts.get("searxng_url"), queries=opts.get("queries", 60), window=9,
                per_query=5, max_pages=40, delay=None, exclude_domains=list(DEFAULT_EXCLUDE_DOMAINS),
                cache_dir=os.path.join(os.path.expanduser("~"), ".textcheck_cache"), no_cache=False)
        try:
            run_compare(path, opts.get("sources") or [], 5, True, bool(opts.get("skip_quotes")), 30,
                        frozenset(), web)
        except SystemExit as e:
            emit("Сравнение не выполнено: %s" % e)
    else:
        emit("Сравнение не выполнялось: включите поиск в интернете или выберите папку с источниками")
        emit("(кнопка «Дополнительно»).")
    incomplete = None
    if STATS.get("web_aborted"):
        incomplete = "поиск в интернете прерван (%s)" % STATS["web_aborted"]
    elif STATS.get("web") and not STATS.get("web_loaded") and STATS.get("web_failed"):
        incomplete = "ни одну найденную страницу не удалось скачать"
    return {"report": "\n".join(REPORT), "serious": STATS.get("serious", 0), "pct": STATS.get("pct"),
            "checked": STATS.get("checked", 0), "web": bool(STATS.get("web")), "incomplete": incomplete}


ENGINES = [("DuckDuckGo (без ключа)", "ddg"), ("Яндекс (нужен ключ)", "yandex"),
           ("Brave (нужен ключ)", "brave"), ("SearXNG (свой сервер)", "searxng")]
GUI_CONFIG = os.path.join(os.path.expanduser("~"), ".textcheck_gui.json")


class App:
    def __init__(self, root, tk, ttk, filedialog, messagebox, tkfont):
        import queue
        self.root, self.tk, self.ttk = root, tk, ttk
        self.fd, self.mb = filedialog, messagebox
        self.q = queue.Queue()
        self.path, self.sources, self.worker, self.last = None, None, None, None
        self.cfg = self.load_cfg()

        root.title("Антиплагиат: локальная проверка")
        root.geometry("880x720")
        root.minsize(740, 560)
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=11)
        big = tkfont.Font(family=base.actual("family"), size=13, weight="bold")
        ttk.Style().configure("Big.TButton", font=big, padding=(20, 10))

        main = ttk.Frame(root, padding=16)
        main.pack(fill="both", expand=True)

        ttk.Label(main, text="1. Выберите файл вашей работы (Word или PDF)").pack(anchor="w")
        row1 = ttk.Frame(main)
        row1.pack(fill="x", pady=(4, 14))
        self.btn_add = ttk.Button(row1, text="Добавить файл", style="Big.TButton", command=self.add_file)
        self.btn_add.pack(side="left")
        self.lbl_file = ttk.Label(row1, text="файл не выбран", foreground="#666666")
        self.lbl_file.pack(side="left", padx=14)

        ttk.Label(main, text="2. Нажмите «Проверить»").pack(anchor="w")
        row2 = ttk.Frame(main)
        row2.pack(fill="x", pady=(4, 8))
        self.btn_check = ttk.Button(row2, text="Проверить", style="Big.TButton", command=self.start)
        self.btn_check.pack(side="left")
        self.btn_stop = ttk.Button(row2, text="Остановить", command=self.stop, state="disabled")
        self.btn_stop.pack(side="left", padx=14)

        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.pack(fill="x", pady=(4, 2))
        self.lbl_status = ttk.Label(main, text="", foreground="#666666")
        self.lbl_status.pack(anchor="w")
        self.lbl_head = ttk.Label(main, text="", font=big, justify="left")
        self.lbl_head.pack(anchor="w", pady=(8, 4))

        self.btn_adv = ttk.Button(main, text="Дополнительно \u25be", command=self.toggle_adv)
        self.btn_adv.pack(anchor="w")
        self.adv = ttk.LabelFrame(main, text="Дополнительно", padding=10)
        self.adv_open = False
        self.build_adv()

        self.txt_frame = ttk.Frame(main)
        self.txt_frame.pack(fill="both", expand=True, pady=(8, 8))
        self.txt = tk.Text(self.txt_frame, wrap="word", height=12, state="disabled")
        sb = ttk.Scrollbar(self.txt_frame, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)

        row3 = ttk.Frame(main)
        row3.pack(fill="x")
        self.btn_save = ttk.Button(row3, text="Сохранить отчёт", command=self.save_report, state="disabled")
        self.btn_save.pack(side="left")
        self.btn_fix = ttk.Button(row3, text="Исправить символы в копии файла", command=self.fix_symbols,
                                  state="disabled")
        self.btn_fix.pack(side="left", padx=10)

    # ---- настройки -------------------------------------------------------------------------
    def load_cfg(self):
        cfg = {"web": True, "engine": "ddg", "queries": 60, "searxng_url": "", "skip_quotes": False,
               "remember": False, "consent": False, "yandex_key": "", "yandex_folder": "", "brave_key": ""}
        try:
            with open(GUI_CONFIG, encoding="utf-8") as f:
                cfg.update(json.load(f))
        except (OSError, ValueError):
            pass
        return cfg

    def save_cfg(self, opts):
        cfg = {k: self.cfg.get(k) for k in ("consent",)}
        cfg.update(web=opts["web"], engine=opts["engine"], queries=opts["queries"],
                   searxng_url=opts["searxng_url"], skip_quotes=opts["skip_quotes"], remember=self.v_remember.get())
        if self.v_remember.get():
            cfg.update(yandex_key=opts["yandex_key"], yandex_folder=opts["yandex_folder"],
                       brave_key=opts["brave_key"])
        self.cfg.update(cfg)
        try:
            with open(GUI_CONFIG, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False)
        except OSError:
            pass

    # ---- блок «Дополнительно» ----------------------------------------------------------------
    def build_adv(self):
        tk, ttk, adv = self.tk, self.ttk, self.adv
        self.v_web = tk.BooleanVar(value=bool(self.cfg["web"]))
        self.v_quotes = tk.BooleanVar(value=bool(self.cfg["skip_quotes"]))
        self.v_remember = tk.BooleanVar(value=bool(self.cfg["remember"]))
        self.v_queries = tk.StringVar(value=str(self.cfg["queries"]))
        self.v_vars = {k: tk.StringVar(value=self.cfg.get(k) or "")
                       for k in ("yandex_key", "yandex_folder", "brave_key", "searxng_url")}

        ttk.Checkbutton(adv, text="Искать в интернете", variable=self.v_web).grid(
            row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(adv, text="Короткие фразы (по 9 слов) из вашей работы отправляются в выбранный поисковик.",
                  foreground="#666666").grid(row=1, column=0, columnspan=3, sticky="w")
        ttk.Label(adv, text="Поисковик:").grid(row=2, column=0, sticky="w", pady=4)
        self.cb_engine = ttk.Combobox(adv, state="readonly", width=30, values=[n for n, _ in ENGINES])
        self.cb_engine.grid(row=2, column=1, sticky="w")
        cur = [n for n, c in ENGINES if c == self.cfg["engine"]] or [ENGINES[0][0]]
        self.cb_engine.set(cur[0])
        self.cb_engine.bind("<<ComboboxSelected>>", lambda e: self.update_engine_fields())

        self.field_widgets = {}
        labels = [("yandex_key", "Ключ Яндекса:", True), ("yandex_folder", "Каталог Яндекса:", False),
                  ("brave_key", "Ключ Brave:", True), ("searxng_url", "Адрес SearXNG:", False)]
        for n, (key, text, secret) in enumerate(labels):
            lab = ttk.Label(adv, text=text)
            ent = ttk.Entry(adv, textvariable=self.v_vars[key], width=46, show="*" if secret else "")
            lab.grid(row=3 + n, column=0, sticky="w")
            ent.grid(row=3 + n, column=1, sticky="w", pady=2)
            self.field_widgets[key] = (lab, ent)
        self.update_engine_fields()

        ttk.Label(adv, text="Запросов не больше:").grid(row=7, column=0, sticky="w", pady=4)
        ttk.Spinbox(adv, from_=10, to=500, increment=10, textvariable=self.v_queries, width=6).grid(
            row=7, column=1, sticky="w")
        ttk.Label(adv, text="Папка с источниками:").grid(row=8, column=0, sticky="w", pady=4)
        fr = ttk.Frame(adv)
        fr.grid(row=8, column=1, sticky="w")
        ttk.Button(fr, text="Выбрать\u2026", command=self.pick_sources).pack(side="left")
        ttk.Button(fr, text="Убрать", command=self.clear_sources).pack(side="left", padx=6)
        self.lbl_src = ttk.Label(fr, text="не выбрана", foreground="#666666")
        self.lbl_src.pack(side="left", padx=6)
        ttk.Checkbutton(adv, text="Не учитывать текст в кавычках (цитаты)", variable=self.v_quotes).grid(
            row=9, column=0, columnspan=3, sticky="w")
        ttk.Checkbutton(adv, text="Запомнить ключи на этом компьютере (хранятся в открытом виде)",
                        variable=self.v_remember).grid(row=10, column=0, columnspan=3, sticky="w")

    def engine_code(self):
        label = self.cb_engine.get()
        return ([c for n, c in ENGINES if n == label] or ["ddg"])[0]

    def update_engine_fields(self):
        need = {"yandex": ("yandex_key", "yandex_folder"), "brave": ("brave_key",),
                "searxng": ("searxng_url",)}.get(self.engine_code(), ())
        for key, (lab, ent) in self.field_widgets.items():
            if key in need:
                lab.grid()
                ent.grid()
            else:
                lab.grid_remove()
                ent.grid_remove()

    def toggle_adv(self):
        if self.adv_open:
            self.adv.pack_forget()
            self.btn_adv.config(text="Дополнительно \u25be")
        else:
            self.adv.pack(fill="x", pady=6, before=self.txt_frame)
            self.btn_adv.config(text="Дополнительно \u25b4")
        self.adv_open = not self.adv_open

    def pick_sources(self):
        d = self.fd.askdirectory(title="Папка с источниками (.docx, .doc, .pdf, .txt)")
        if d:
            self.sources = d
            self.lbl_src.config(text=os.path.basename(d.rstrip("/\\")) or d)

    def clear_sources(self):
        self.sources = None
        self.lbl_src.config(text="не выбрана")

    # ---- кнопки --------------------------------------------------------------------------------
    def add_file(self):
        p = self.fd.askopenfilename(title="Выберите файл работы", filetypes=[
            ("Документы", "*.docx *.doc *.pdf *.txt"), ("Все файлы", "*.*")])
        if p:
            self.path = p
            self.lbl_file.config(text=os.path.basename(p), foreground="black")
            self.lbl_head.config(text="")
            self.set_text("")
            self.btn_save.config(state="disabled")
            self.btn_fix.config(state="disabled")

    def collect_opts(self):
        try:
            queries = max(5, min(1000, int(str(self.v_queries.get()).strip())))
        except ValueError:
            queries = 60
        return {"web": bool(self.v_web.get()), "engine": self.engine_code(),
                "yandex_key": self.v_vars["yandex_key"].get().strip(),
                "yandex_folder": self.v_vars["yandex_folder"].get().strip(),
                "brave_key": self.v_vars["brave_key"].get().strip(),
                "searxng_url": self.v_vars["searxng_url"].get().strip(),
                "queries": queries, "sources": [self.sources] if self.sources else [],
                "skip_quotes": bool(self.v_quotes.get())}

    def settings_problem(self, o):
        if not o["web"]:
            return None
        env = os.environ.get
        if o["engine"] == "yandex" and not ((o["yandex_key"] or env("YANDEX_SEARCH_API_KEY")) and
                                            (o["yandex_folder"] or env("YANDEX_FOLDER_ID"))):
            return "Для Яндекса укажите ключ и каталог (кнопка «Дополнительно»)."
        if o["engine"] == "brave" and not (o["brave_key"] or env("BRAVE_API_KEY")):
            return "Для Brave укажите ключ (кнопка «Дополнительно»)."
        if o["engine"] == "searxng" and not (o["searxng_url"] or env("SEARXNG_URL")):
            return "Для SearXNG укажите адрес сервера (кнопка «Дополнительно»)."
        return None

    def start(self):
        if not self.path:
            self.mb.showinfo("Нет файла", "Сначала нажмите кнопку «Добавить файл».")
            return
        o = self.collect_opts()
        problem = self.settings_problem(o)
        if problem:
            self.mb.showwarning("Не хватает настроек", problem)
            return
        if o["web"] and not self.cfg.get("consent"):
            if not self.mb.askokcancel(
                    "Поиск в интернете",
                    "Чтобы найти заимствования, программа отправит в поисковую систему короткие фразы "
                    "(по 9 слов) из вашей работы. Файл целиком никуда не отправляется.\n\n"
                    "Продолжить? (Отключить поиск можно в «Дополнительно».)"):
                return
            self.cfg["consent"] = True
        if o["engine"] == "yandex":  # ключи из полей или из переменных окружения
            o["yandex_key"] = o["yandex_key"] or None
            o["yandex_folder"] = o["yandex_folder"] or None
        self.save_cfg(o)
        CANCEL.clear()
        self.set_running(True)
        self.set_text("")
        self.lbl_head.config(text="")
        self.worker = threading.Thread(target=self.work, args=(self.path, o), daemon=True)
        self.worker.start()
        self.root.after(150, self.poll)

    def stop(self):
        CANCEL.set()
        self.lbl_status.config(text="Останавливаю\u2026")
        self.btn_stop.config(state="disabled")

    def set_running(self, running):
        self.btn_check.config(state="disabled" if running else "normal")
        self.btn_add.config(state="disabled" if running else "normal")
        self.btn_stop.config(state="normal" if running else "disabled")
        if running:
            self.btn_save.config(state="disabled")
            self.btn_fix.config(state="disabled")
            self.lbl_status.config(text="Идёт проверка\u2026")
            self.progress.start(12)
        else:
            self.progress.stop()

    def set_text(self, text):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", text)
        self.txt.config(state="disabled")

    # ---- фоновая работа ---------------------------------------------------------------------
    def work(self, path, opts):
        old_err, old_out = sys.stderr, sys.stdout
        sys.stderr = _Writer(lambda s: self.q.put(("status", s)))
        sys.stdout = _Writer(lambda s: None)
        try:
            self.q.put(("done", gui_check(path, opts)))
        except SystemExit as e:
            self.q.put(("error", str(e) or "Ошибка"))
        except OSError as e:
            self.q.put(("error", "Не удалось открыть файл: %s" % e))
        except Exception as e:
            self.q.put(("error", "Непредвиденная ошибка: %s: %s" % (type(e).__name__, e)))
        finally:
            sys.stderr, sys.stdout = old_err, old_out

    def poll(self):
        import queue
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "status":
                    self.lbl_status.config(text=val[:150])
                elif kind == "done":
                    self.finish(val)
                    return
                elif kind == "error":
                    self.fail(val)
                    return
        except queue.Empty:
            pass
        self.root.after(150, self.poll)

    def finish(self, res):
        self.last = res
        self.set_running(False)
        self.lbl_status.config(text="Готово.")
        if CANCEL.is_set():
            self.lbl_status.config(text="Проверка остановлена, показан неполный результат.")
        lines = []
        incomplete = res.get("incomplete") or ("проверка остановлена" if CANCEL.is_set() else None)
        if res["pct"] is None:
            lines.append("Совпадения не проверялись")
        elif incomplete:
            lines.append("Результат неполный: %s" % incomplete)
            lines.append("Найдено совпадений: %.1f%% (на самом деле может быть больше)" % res["pct"])
        else:
            lines.append("Найдено совпадений: %.1f%%" % res["pct"])
            if res.get("web"):
                lines.append("Проверена выборка фраз, реальное значение может быть выше")
        if res["serious"]:
            lines.append("Символы: найдено подозрительных мест \u2014 %d" % res["serious"])
        else:
            lines.append("Символы: подозрительных мест нет")
        self.lbl_head.config(text="\n".join(lines))
        self.set_text(res["report"])
        self.btn_save.config(state="normal")
        self.btn_fix.config(state="normal" if res["serious"] else "disabled")

    def fail(self, message):
        self.set_running(False)
        self.lbl_status.config(text="")
        self.mb.showerror("Не получилось", message)

    def save_report(self):
        if not self.last:
            return
        p = self.fd.asksaveasfilename(defaultextension=".txt", initialfile="отчёт.txt",
                                      filetypes=[("Текстовый файл", "*.txt")])
        if p:
            with open(p, "w", encoding="utf-8") as f:
                f.write(self.last["report"] + "\n")
            self.mb.showinfo("Готово", "Отчёт сохранён:\n%s" % p)

    def fix_symbols(self):
        try:
            dst, changed = fix_file(self.path, frozenset())
        except (Exception, SystemExit) as e:
            self.mb.showerror("Не получилось", str(e))
            return
        self.mb.showinfo("Готово", "Исправленная копия сохранена:\n%s\n\nИсправлено символов: %d.\n"
                         "Оригинал не изменён. Скрытый текст не удалялся." % (dst, changed))


def gui():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, font as tkfont
    except ImportError:
        print("Не найден tkinter, без него окно не открыть.\n"
              "Windows/macOS: установите Python с python.org (там tkinter есть).\n"
              "Linux: sudo apt install python3-tk\n"
              "Командная строка работает и без него: python textcheck.py --help")
        sys.exit(1)
    try:  # чёткое отображение на экранах с масштабированием (Windows)
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()

    def show_error(exc, val, tb):
        messagebox.showerror("Ошибка программы", "%s: %s" % (exc.__name__, val))

    root.report_callback_exception = show_error
    App(root, tk, ttk, filedialog, messagebox, tkfont)
    root.mainloop()


# ----------------------------------------------------------------------------------------------
def main():
    for st in (sys.stdout, sys.stderr):
        try:
            st.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if len(sys.argv) == 1 or sys.argv[1] == "gui":
        return gui()

    ap = argparse.ArgumentParser(description="Локальная проверка курсовой/диплома (см. docstring в начале файла)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("audit", help="скрытые символы, подмена букв, скрытый текст")
    a.add_argument("file")
    a.add_argument("--fix", action="store_true", help="сохранить исправленную копию *.clean.*")
    a.add_argument("--limit", type=int, default=40)
    a.add_argument("--allow", default="", help="разрешённые нерусские буквы кириллицы, например: іїє")
    a.add_argument("--report", help="сохранить отчёт в файл (UTF-8)")

    c = sub.add_parser("compare", help="сравнение с вашими источниками и/или с открытым интернетом")
    c.add_argument("file")
    c.add_argument("sources", nargs="*", help="папки и/или файлы .docx/.doc/.pdf/.txt/.md (можно не указывать с --web)")
    c.add_argument("--shingle", type=int, default=5, help="длина совпадающего фрагмента в словах")
    c.add_argument("--no-legal", action="store_true", help="не исключать предложения про законы/ГОСТы")
    c.add_argument("--skip-quotes", action="store_true", help="не учитывать текст в кавычках")
    c.add_argument("--limit", type=int, default=30)
    c.add_argument("--allow", default="")
    c.add_argument("--report", help="сохранить отчёт в файл (UTF-8)")
    c.add_argument("--web", action="store_true", help="искать фразы работы в интернете (они уйдут в поисковик!)")
    c.add_argument("--engine", choices=["ddg", "yandex", "brave", "searxng"], default="ddg")
    c.add_argument("--yandex-key")
    c.add_argument("--yandex-folder")
    c.add_argument("--brave-key")
    c.add_argument("--searxng-url")
    c.add_argument("--queries", type=int, default=60, help="максимум поисковых запросов (по умолчанию 60)")
    c.add_argument("--window", type=int, default=9, help="слов во фразе-запросе (по умолчанию 9)")
    c.add_argument("--per-query", type=int, default=5, help="сколько результатов брать из поиска на запрос")
    c.add_argument("--max-pages", type=int, default=40, help="максимум скачиваемых страниц")
    c.add_argument("--delay", type=float, default=None, help="пауза между запросами, секунд")
    c.add_argument("--exclude-domain", action="append", default=[], help="не скачивать страницы этого домена")
    c.add_argument("--no-default-excludes", action="store_true", help="не исключать сайты с законами по умолчанию")
    c.add_argument("--cache-dir", default=".textcheck_cache")
    c.add_argument("--no-cache", action="store_true")

    args = ap.parse_args()
    allowed = frozenset(args.allow)
    if args.cmd == "audit":
        code = run_audit(args.file, args.limit, args.fix, allowed)
    else:
        if not args.sources and not args.web:
            ap.error("укажите папку с источниками и/или ключ --web")
        web = None
        if args.web:
            args.exclude_domains = list(args.exclude_domain) + ([] if args.no_default_excludes
                                                                 else list(DEFAULT_EXCLUDE_DOMAINS))
            web = args
        code = run_compare(args.file, args.sources, args.shingle, not args.no_legal,
                           args.skip_quotes, args.limit, allowed, web)
    if args.report:
        with open(args.report, "w", encoding="utf-8") as f:
            f.write("\n".join(REPORT) + "\n")
        print("\nОтчёт сохранён: %s" % args.report)
    sys.exit(code)


if __name__ == "__main__":
    main()
