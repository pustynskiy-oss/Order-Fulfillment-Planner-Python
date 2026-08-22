# HANDOFF: Order Fulfillment Planner (передача задачи)  

## 0. Среда и ограничения  
- Windows 10, Python 3.12, GUI PySide6, Excel openpyxl. Пользователь НЕ программист: правки вносить ТОЛЬКО полным файлом (ручные «патчи» неоднократно ломали файл: слипшиеся строки, лишние точки, код до определения функций).  
- В заголовке окна и в шапке ОБЯЗАТЕЛен тег сборки BUILD (сейчас v21) — по нему проверять, что запущен актуальный файл.  
- Запись профилей/БЗ — только в
\`%LOCALAPPDATA%\\OrderFulfillmentPlanner\\\` (запись в сетевой диск N:\\ давала PermissionError).  
- Файлы \`.xls\` не читать — выводить просьбу сохранить как \`.xlsx\`.  

## 1. Назначение приложения  
Ежедневный расчёт перемещений ГП между двумя РЦ‑складами: «Склад основной» (Рязань) и «Склад Скопин РЦ». Результат — карточки рейсов по каждому направлению с позициями, паллетоместами, контролем высоты и вместимости; экспорт в Excel.  

## 2. Хронология требований пользователя (ключевые цитаты/смысл)  
1. Импорт «Обеспечение заказов» (1С), справочников; расчёт потребности по дням и уровням; рейсы, drag‑and‑drop, ручные строки, показатели, предупреждения.  
2. «вначале осуществляй проверку потребности по позиции, и только если эту позицию необходимо переместить выдавай ошибку» — предупреждения только по перемещаемым позициям.  
3. «Не делай перемещение номенклатуры со склада основного на склад скопин рц, если нет отрицательного конечного остатка на складе получателе»; «в данном направлении не нужно формировать страховой запас, нужно только закрыть потребность» (целые коробки, без докладки).  
4. Скопин→Основной: страховой запас только при свободном месте в рейсе 1 (или «Рейсов»>1) и только если источник без дефицита; страховой не создаёт рейсов и не перегружает.  
5. «всегда стремись к 1 рейсу… расчёт 1 раз в день»; при фиксированном «Рейсов: 1» новые рейсы НЕ создаются (уровень 3 идёт в рейс 2 только если рейсов задано ≥2, иначе в рейс 1 с предупреждением о перегрузе); авторасширение только при «Рейсов: 0».  
6. UI: кнопка «Выгрузка 1с»→ позже замена на 3 плоских файла; рамки кнопок импорта зелёные при загрузке; «Рассчитать» крупная; блок «Импорт» одной рамкой с подписью внизу; карточка рейса на всю высоту при 1 рейсе; нижний блок
(Предупреждения/Недостаточные/Недогрузы)
растягивается; столбцы таблицы: №, Номенклатура (2‑я), Уп/слой, Упак к перемещению, Паллет мест, «Ур. потребности» (мин. закрываемый уровень; зелёный/красный фон), крестик удаления с
подтверждением; выравнивание №/Номенклатура влево, остальные по центру; Ctrl+C копирует выделенные ячейки; строка «Итог» между шапкой и данными.  
7. «Переделай код, чтобы я выгружал… три файла вместо 1 „Выгрузка 1с“» — плоские таблицы вместо сводной 1С.  
8. Производство: новая структура отчёта (вертикальные группы цехов, колонки «Размещено в
заказе|Произведено|Осталось выполнить»); карта поступления выпуска: Цех №2 производство→Скопин РЦ; Цех №1 производство→Основной; Цех №1
упаковка→Скопин РЦ; Цех №2 упаковка→Основной.  
9. «Не придумывай недостающие бизнес‑правила молча. Если данных нет, перечисли их в вопросах».  

## 3. Финальные бизнес‑правила (вшиты в кнопку «Правила»)  
- Дата из ОС; первая дата файлов = сегодня, иначе ошибка.  
- Расчёт ТОЛЬКО по складским строкам; итоговая строка номенклатуры = сумма складов (не используется).  
- Производственные склады‑источники (Цех Рязань основной склад; Склад Мастер Цех №2 смена №1‑4) добавляют источнику положительный «Конечный остаток» (в 3‑файловой схеме их заменяет
«Производство»).  
- Игнор: Склад ГП цех 2, Пандус Ряз. (упаковка), Империя Холода, Колибри; служебные наименования («не использовать», «снята…», «эксперимент», «(уд)»…).  
- Уровни потребности: 1 и 2 — всегда рейс 1; 3 — рейс 2 (если рейсов ≥2, иначе рейс 1); 4 — рейс 1 при месте; ≥5 игнор.  
- Объём = min(потребность, доступный источник); оба минуса → «Недостаточные остатки».  
- Направление по дефициту: спросовое перемещение только туда, где у получателя отрицательный конечный остаток (хранящийся ИЛИ пересчитанный: начальный+движение+план поступления−потребность).  
- Основной→Скопин: только под потребность, целые коробки, без докладки и без страховки.  
- Страховой запас (≤ 3,5×контролируемого) — только Скопин→Основной, по правилам п.4.  
- Высота: слои = floor((2250−145)/H); базовый паллет =
(слои−1)×упак_в_ряду; докладка до целого слоя только Скопин→Основной; контроль ≤2250 мм.  
- Паллетоместа: ceil с квантом 0,5 при остатке источника >5 паллет, иначе 0,1.  
- кг = Вес×Коэффициент, если «Фасовка/Вес»=«Вес», иначе Вес.  
- Сопоставление имён: точное (нормализация: нижний регистр, ё→е, сжатие пробелов), при неудаче difflib 0.86. Базы знаний заменяются целиком при каждом импорте.  

## 4. Форматы входных файлов  
### 4.1 «Товары на складах» (плоский): Номенклатура | Склад | Контролируемый остаток | Уровень потребности | Дата | Начальный остаток | Конечный остаток.  
### 4.2 «Заказы»: Номенклатура | Склад | Дата | Потребность к отгрузке.  
### 4.3 «Производство» (сводный): строки‑группы «Подразделение» (имена цехов из карты) и под ними номенклатура; колонки «Размещено в
заказе|Произведено|Осталось выполнить». В
поступление берётся «Осталось выполнить»; даты в отчёте НЕТ → объём относится на первый день горизонта (ДОПУЩЕНИЕ; при появлении колонки даты — привязать по дате). Поддерживается и старая структура (шапка цехов по горизонтали + «Дата исполнения») как запасной парсер.  
### 4.4 Справочники: «Номенклатура‑справочник» (Ссылка|Код|Вид|Вес|Фасовка/Вес|Коэффициент|Вес брутто|упак на поддоне Низком|Высоком|упак в ряду|Производственный вид); «Гофра высота»
(Ссылка|высота, мм); «Соответствие
номенклатура‑гофра» (Номенклатура|Код|Гофра).  
### 4.5 Старый сводный «Обеспечение заказов» (если понадобится): колонки A имя, B контролируемый, C уровень, D показатель («Начальный
остаток/Движение/Потребность к
отгрузке/Выпущено/План поступления/Конечный остаток»), E… даты (11 дат); блоки: тотальные строки + складские строки «Склад основной»/«Склад Скопин РЦ»; порядок строк ненадёжен (складские могут идти до тотальной) → прикрепление по контролю суммы (сумма B складских строк блока ≈ B тотальной). Пример: \`| Склад Скопин РЦ|25.16|1|Конечный остаток|-204|…\`.  

## 5. Архитектура (модули)  
- \`load_wb()\` — ремонт xlsx 1С: нормализация имён zip‑частей, принудительный \`xl/sharedStrings.xml\`; отказ для \`.xls\`.  
- \`num()\` — русские числа: «1,234.5»→1234.5; «2,005»→2005; «0,25»→0.25; юникод‑минус/тире/скобки.  
- \`read_table()\` — поиск шапки плоской таблицы по ключевым словам.  
- \`SupplyData\` — import_stock/import_orders/import_prod + \`build(today)\` → даты и блоки {номенклатура: {склад: WhBlock(indicators, level)}}; производство добавляется в «План поступления» целевого склада по карте PROD_TO_WH.  
- \`Planner\` — \`_balance()\` (пересчёт конечного остатка), \`_recv_neg()\` (дефицит получателя по пересчёту ИЛИ хранящемуся), \`_pack_info()\` (гофра/высоты/упак), \`_manual_item()\`, \`plan()\` (этапы: спрос по дефициту → размещение (уровни/рейсы) → страховая дозагрузка только Скопин→Основной), \`_place()\`/\`_place_soft()\`.  
- GUI — 6 кнопок импорта в одной рамке «Импорт» (зелёные рамки при загрузке), «Рассчитать», «Правила», «Экспорт Excel»; две панели направлений со сплиттерами карточек рейсов; нижний растягиваемый блок с вкладками; тег BUILD.  

## 6. История багов и фиксов (не повторять!)  
1. \`importlib.util\` без импорта → AttributeError → \`_has()\` через try/except \`__import__\`.  
2. openpyxl KeyError \`xl/sharedStrings.xml\` у файлов 1С → zip‑ремонт; затем IndexError (t="s" при пустой таблице) → нормализация имён частей архива.  
3. PermissionError записи JSON на N:\\ → DATA_DIR в LOCALAPPDATA.  
4. Ручные правки ломали синтаксис → переход на поставку полным файлом + BUILD‑тег.  
5. NameError \`PROD_TO_WH_N\` (вычисление до определения \`norm\`) → перенести строку \`PROD_TO_WH_N = {norm(k): k for k in PROD_TO_WH}\` ПОСЛЕ \`def norm\`.  
6. Неприкрепление складских строк в сводном файле → сначала pending‑буфер/сумма‑контроль, затем ПОЛНЫЙ переход на 3 плоских файла (класс бага устранён конструктивно).  
7. Ложные страховые перевозки Основной→Скопин → отключены; спрос только при дефиците получателя.  
8. Вторые рейсы при «Рейсов: 1» → запрещены (см. правила).  

## 7. Открытые вопросы/допущения (ждать подтверждения пользователя)  
- Дата поступления выпуска из «Производство» (сейчас первый день горизонта).  
- Возможное задвоение выпуска, если он уже сидит в «План поступления» других файлов (проверить на живых данных; при задвоении — режим «заменять, а не добавлять»).  
- Имена цехов в реальном отчёте должны совпадать с картой PROD_TO_WH (иначе строки игнорируются).  

## 8. Эталонный код (v21 = v20 + фикс порядка PROD_TO_WH_N)  
\`\`\`python  
# -\*- coding: utf-8 -\*-  
"""  
ORDER FULFILLMENT PLANNER — перемещения Склад Основной <-> Склад Скопин РЦ.  
Windows 10 | Python 3.12. Библиотеки ставятся автоматически: PySide6, openpyxl.  

ВХОДНЫЕ ФАЙЛЫ:  
1) «Товары на складах»: Номенклатура | Склад | Контролируемый остаток |  
Уровень потребности | Дата | Начальный остаток | Конечный остаток.  
2) «Заказы»: Номенклатура | Склад | Дата | Потребность к отгрузке.  
3) «Производство»: вертикальные группы
«Подразделение» (цех) и номенклатура;  
колонки «Размещено в заказе | Произведено | Осталось выполнить».  
Поступление = «Осталось выполнить» на первый день горизонта; карта цех->склад:  
Цех №2 производство -> Скопин РЦ; Цех №1
производство -> Основной;  
Цех №1 упаковка -> Скопин РЦ; Цех №2 упаковка -> Основной.  
Старая структура (шапка цехов + «Дата исполнения») поддерживается как запасная.  
Справочники: «Номенклатура-справочник», «Гофра» (высоты), «Номенклатура»  
(соответствие номенклатура-гофра). Файлы .xls не читаются — сохраняйте .xlsx.  

ВШИТАЯ ИНСТРУКЦИЯ (кнопка «Правила»): см. раздел 3 файла HANDOFF.  
"""  
import sys, subprocess  

BUILD = "v21 (12.08.2026)"  

def _has(mod):  
try:  
__import__(mod); return True  
except ImportError:  
return False  

_missing = [m for m in ("PySide6", "openpyxl") if not _has(m)]  
if _missing:  
subprocess.check_call([sys.executable, "-m", "pip", "install", \*_missing])  

import re, json, math, datetime, traceback, os, zipfile, tempfile, difflib  
from dataclasses import dataclass, field  
from typing import Dict, List, Optional, Tuple  

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,  
QLabel, QPushButton, QFileDialog, QMessageBox, QSplitter, QScrollArea, QFrame,  
QTreeWidget, QTreeWidgetItem, QSpinBox, QPlainTextEdit, QTableWidget, QTableWidgetItem,  
QTabWidget, QHeaderView, QAbstractItemView, QInputDialog, QSizePolicy)  
from PySide6.QtCore import Qt, QMimeData, QByteArray  
from PySide6.QtGui import QKeySequence, QColor, QFont  
import openpyxl  

# ============================== КОНСТАНТЫ ==============================  
WH_MAIN = "Склад основной"  
WH_SKOP = "Склад Скопин РЦ"  
D_SKOP_MAIN = f"{WH_SKOP}->{WH_MAIN}"  
D_MAIN_SKOP = f"{WH_MAIN}->{WH_SKOP}"  
PROD_TO_WH = {  
"Цех №2 производство": WH_SKOP,  
"Цех №1 производство": WH_MAIN,  
"Цех №1 упаковка": WH_SKOP,  
"Цех №2 упаковка": WH_MAIN,  
}  
ALL_WH = {WH_MAIN, WH_SKOP} | set(PROD_TO_WH.values())  
EXTRA_WH = {"Склад ГП цех 2", "Пандус Ряз. (упаковка)", "Империя Холода", "Колибри"}  
SKIP_PREFIX = ("склад", "пандус", "цех", "империя", "колибри", "дополнительные поля",  
"отбор", "номенклатура", "группировки", "планирование", "ссылка", "параметры")  
PALLET_MM, MAX_H_MM = 145, 2250  
DEF_CAP, MAX_CAP = 33, 36  
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),  
"OrderFulfillmentPlanner")  
try:  
os.makedirs(DATA_DIR, exist_ok=True)  
except Exception:  
DATA_DIR = tempfile.gettempdir()  
KB_PATH = os.path.join(DATA_DIR, "knowledge_base.json")  
PROFILE_PATH = os.path.join(DATA_DIR, "planner_profile.json")  
MIME = "application/x-planer-item"  
RULES_TEXT = __doc__  

def norm(s) -> str:  
if s is None: return ""  
return re.sub(r"\\s+", " ", str(s).lower().replace("ё", "е")).strip()  

def num(v) -> float:  
if v is None or v == "": return 0.0  
if isinstance(v, (int, float)): return float(v)  
s = str(v).replace("\\u00a0", "").replace(" ", "").strip()  
if not s: return 0.0  
neg = False  
if s.startswith("(") and s.endswith(")"):  
neg, s = True, s[1:-1]  
for ch in ("−", "–", "—"):  
s = s.replace(ch, "-")  
if s.startswith("-"):  
neg, s = True, s[1:]  
if "," in s and "." in s:  
s = s.replace(",", "")  
elif "," in s:  
a, b = s.split(",")  
s = a + b if len(b) == 3 else a + "." + b  
try: r = float(s)  
except: return 0.0  
return -r if neg else r  

PROD_TO_WH_N = {norm(k): k for k in PROD_TO_WH}          # ПОСЛЕ norm!  
CANON_ALL = {norm(x): x for x in ALL_WH}  
EXTRA_WH_N = {norm(x) for x in EXTRA_WH}  
EMPTY_SST = (''  
'             'count="0" uniqueCount="0"/>')  

def load_wb(path):  
if path.lower().endswith(".xls") and not path.lower().endswith(".xlsx"):  
raise RuntimeError("Файл .xls не читается. Сохраните как .xlsx и загрузите снова.")  
try:  
return openpyxl.load_workbook(path, data_only=True)  
except (KeyError, IndexError):  
tmp = os.path.join(tempfile.gettempdir(), "planer_repair.xlsx")  
with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:  
seen, wrote_ss = set(), False  
for item in zin.infolist():  
canon = item.filename.replace("\\\\", "/").strip().lstrip("/")  
if not canon: continue  
low = canon.lower()  
if low.endswith("sharedstrings.xml"):  
if not wrote_ss:  
zout.writestr("xl/sharedStrings.xml", zin.read(item)); wrote_ss = True  
continue  
if low in seen: continue  
seen.add(low)  
zout.writestr(canon, zin.read(item))  
if not wrote_ss:  
zout.writestr("xl/sharedStrings.xml", EMPTY_SST)  
try:  
return openpyxl.load_workbook(tmp, data_only=True)  
except Exception:  
raise RuntimeError("Файл не читается даже после ремонта: сохраните его в Excel как .xlsx.")  

def read_table(path):  
wb = load_wb(path); ws = wb.active  
hdr, hi = None, 0  
rows_all = list(ws.iter_rows(values_only=True))  
for i, r in enumerate(rows_all):  
cells = [str(c).strip() for c in r if c is not None and str(c).strip()]  
if len(cells) >= 3:  
hdr = [norm(str(c)) if c is not None else "" for c in r]; hi = i; break  
if hdr is None: return {}, []  
m = {}  
for i, h in enumerate(hdr):  
if not h: continue  
if "номенклатура" in h or "наименован" in h: m.setdefault("name", i)  
elif "склад" in h: m.setdefault("wh", i)  
elif "дата" in h: m.setdefault("date", i)  
elif "контролиру" in h: m.setdefault("ctrl", i)  
elif "уровень" in h: m.setdefault("level", i)  
elif "начальн" in h: m.setdefault("open", i)  
elif "конечн" in h: m.setdefault("close", i)  
elif "потребность" in h: m.setdefault("demand", i)  
elif "поступлен" in h or "выпущено" in h or "производ" in h: m.setdefault("receipt", i)  
elif "колич" in h or "заказ" in h: m.setdefault("qty", i)  
return m, rows_all[hi + 1:]  

# ============================== БАЗА ЗНАНИЙ ==============================  
class KnowledgeBase:  
def __init__(self):  
self.nomen, self.gofra, self.corr = {}, {}, {}  
self._fz_n, self._fz_c = {}, {}  
self.load()  
def save(self):  
try:  
json.dump({"nomen": self.nomen, "gofra": self.gofra, "corr": self.corr},  
open(KB_PATH, "w", encoding="utf-8"), ensure_ascii=False)  
except Exception as e:  
print("Не удалось сохранить базу знаний:", e)  
def load(self):  
try:  
d = json.load(open(KB_PATH, encoding="utf-8"))  
self.nomen, self.gofra, self.corr = (d.get("nomen", {}), d.get("gofra", {}), d.get("corr", {}))  
except Exception: pass  
def _fuzz(self, key, base, cache):  
if key in cache: return cache[key]  
m = difflib.get_close_matches(key, list(base), 1, 0.86)  
r = m[0] if m else None  
cache[key] = r; return r  
def import_nomen(self, path) -> str:  
ws = load_wb(path).active  
hdr = None  
for row in ws.iter_rows(values_only=True):  
cells = [str(c).strip() for c in row if c is not None and str(c).strip()]  
if len(cells) >= 6 and any("Код" in c for c in cells) and any("Высоком" in c or "ряду" in c for c in cells):  
hdr = [str(c) if c is not None else "" for c in row]; break  
if hdr:  
ci = lambda k: next((i for i, c in enumerate(hdr) if k in c), None)  
ix = {"code": ci("Код"), "w": ci("Вес"), "p": ci("Фасовка"), "k": ci("Коэффициент"),  
"low": ci("Низком"), "high": ci("Высоком"), "row": ci("в ряду"), "prod": ci("Производственный")}  
else:  
ix = {"code": 1, "w": 3, "p": 4, "k": 5, "low": 7, "high": 8, "row": 9, "prod": 10}  
new = {}  
for row in ws.iter_rows(values_only=True):  
name = row[0]  
if name is None or norm(name) == "" or norm(name).startswith(SKIP_PREFIX): continue  
code = str(row[ix["code"]] or "").strip() if ix["code"] is not None else ""  
if not code.isdigit(): continue  
rec = {"code": code, "weight": num(row[ix["w"]]), "pack": str(row[ix["p"]] or ""),  
"coef": num(row[ix["k"]]), "row": num(row[ix["row"]]),  
"high": num(row[ix["high"]]), "low": num(row[ix["low"]]),  
"prod": str(row[ix["prod"]] or "")}  
key = norm(name); old = new.get(key)  
if old is None: new[key] = rec  
elif (rec["row"] or rec["high"]) and not (old.get("row") or old.get("high")): new[key] = rec  
self.nomen = new; self._fz_n = {}; self.save()  
return f"Номенклатура: загружено {len(new)} (база заменена целиком)"  
def import_gofra(self, path) -> str:  
wb = load_wb(path); n = 0  
for ws in wb.worksheets:  
hdr = None  
for row in ws.iter_rows(values_only=True):  
cells = [str(c).lower() if c is not None else "" for c in row]  
if any("высот" in c for c in cells) and any("ссылка" in c for c in cells):  
hdr = cells; break  
i_h = next((i for i, c in enumerate(hdr) if "высот" in c), None) if hdr else None  
for row in ws.iter_rows(values_only=True):  
cells = list(row)  
nz = [(i, v) for i, v in enumerate(cells) if v not in (None, "")]  
if len(nz) < 2: continue  
name = nz[0][1]  
if not isinstance(name, str) or norm(name).startswith(SKIP_PREFIX): continue  
h = num(cells[i_h]) if (i_h is not None and i_h < len(cells)) else 0.0  
if not h:  
for i, v in nz[1:]:  
if num(v) > 0: h = num(v); break  
if not h: continue  
self.gofra[norm(name)] = h; n += 1  
self.save(); return f"Гофра высота: загружено {n}"  
def import_corr(self, path) -> str:  
wb = load_wb(path); new = {}  
for ws in wb.worksheets:  
hdr = i_n = i_g = None  
for row in ws.iter_rows(values_only=True):  
cells = [str(c).lower() if c is not None else "" for c in row]  
if any("гофра" in c for c in cells) and any("номенклатура" in c for c in cells):  
hdr = cells  
i_n = next((i for i, c in enumerate(cells) if "номенклатура" in c and "код" not in c), 0)  
i_g = next((i for i, c in enumerate(cells) if "гофра" in c and "номенклатура" not in c), None)  
break  
if hdr is None: continue  
for row in ws.iter_rows(values_only=True):  
low = [str(c).lower() if c is not None else "" for c in row]  
if low == hdr: continue  
name = row[i_n] if i_n is not None else row[0]  
g = str(row[i_g] or "").strip() if i_g is not None else ""  
if name is None or norm(name) == "" or not g: continue  
if norm(name).startswith(SKIP_PREFIX): continue  
new[norm(name)] = [g, 0.0]  
if not new:  
ws = wb.active  
for row in ws.iter_rows(values_only=True):  
if row[0] is None or norm(row[0]) == "": continue  
if "номенклатура" in norm(row[0]) and any("гофра" in norm(c or "") for c in row[1:]): continue  
new[norm(row[0])] = [str(row[1] or ""), num(row[2])]  
if not new: return "В файле не найдены колонки
«Номенклатура» и «Гофра»"  
self.corr = new; self._fz_c = {}; self.save()  
return f"Соответствия: загружено {len(new)} (база заменена целиком)"  
SERVICE = ("не использовать", "не исп", "не отгружать", "сняты", "снята с производства",  
"эксперимент", "(уд)", "удаленные", "архив")  
def is_service(self, name): return any(s in norm(name) for s in self.SERVICE)  
def pack_attrs(self, name):  
k = norm(name); r = self.nomen.get(k)  
if r is None and self.nomen:  
f = self._fuzz(k, self.nomen, self._fz_n)  
if f: r = self.nomen[f]  
return r  
def corr_of(self, name):  
k = norm(name); r = self.corr.get(k)  
if r is None and self.corr:  
f = self._fuzz(k, self.corr, self._fz_c)  
if f: r = self.corr[f]  
return r  
def kg_pack(self, name):  
r = self.pack_attrs(name)  
if not r: return None  
return r["weight"] \* r["coef"] if r["pack"] == "Вес" else r["weight"]  

# ============================== МОДЕЛИ ==============================  
@dataclass  
class WhBlock:  
indicators: Dict[str, List[float]] = field(default_factory=dict)  
level: int = 0  

@dataclass  
class NomenBlock:  
name: str  
controlled: float = 0.0  
wh: Dict[str, WhBlock] = field(default_factory=dict)  

@dataclass  
class LineItem:  
uid: int; name: str  
per_layer: float = 0; per_pallet: float = 0; qty: float = 0; pallets: float = 0  
gofra: str = ""; height_mm: float = 0; manual: bool = False  
level_min: int = 0  

@dataclass  
class Trip:  
index: int; direction: str; capacity: int = DEF_CAP  
items: List[LineItem] = field(default_factory=list)  
def occupied(self): return round(sum(i.pallets for i in self.items), 2)  
def free(self): return round(self.capacity - self.occupied(), 2)  
def max_height(self): return max([i.height_mm for i in self.items if not i.manual] + [0])  

@dataclass  
class Plan:  
dates: List[datetime.date]  
trips: Dict[str, List[Trip]] = field(default_factory=dict)  
warnings: List[str] = field(default_factory=list)  
insufficient: List[Tuple[str, str, float]] = field(default_factory=list)  
underload: List[Tuple[str, str, float]] = field(default_factory=list)  
metrics: Dict[str, float] = field(default_factory=dict)  
moved: int = 0  
uncovered: set = field(default_factory=set)  

UID = 0  
def new_uid():  
global UID; UID += 1; return UID  

# ============================== СБОР ДАННЫХ ==============================  
class SupplyData:  
def __init__(self): self.stock, self.orders, self.prod = [], [], []  
def import_stock(self, path) -> str:  
m, rows = read_table(path)  
if "name" not in m or "wh" not in m:  
return "Ошибка: в «Товары на складах» не найдены колонки Номенклатура/Склад"  
data = []  
for r in rows:  
name = r[m["name"]] if m["name"] < len(r) else None  
wh = r[m["wh"]] if m["wh"] < len(r) else None  
if name is None or wh is None or norm(name) == "": continue  
if norm(name).startswith(SKIP_PREFIX): continue  
data.append({"name": str(name), "wh": str(wh),  
"ctrl": num(r[m["ctrl"]]) if "ctrl" in m and m["ctrl"] < len(r) else 0.0,  
"level": int(num(r[m["level"]])) if "level" in m and m["level"] < len(r) else 0,  
"date": _date(r[m["date"]]) if "date" in m and m["date"] < len(r) else None,  
"open": num(r[m["open"]]) if "open" in m and m["open"] < len(r) else None,  
"close": num(r[m["close"]]) if "close" in m and m["close"] < len(r) else None})  
if not data: return "Ошибка: «Товары на складах» — пустая таблица"  
self.stock = data; return f"Товары на складах: строк {len(data)}"  
def import_orders(self, path) -> str:  
m, rows = read_table(path)  
if "name" not in m or "wh" not in m:  
return "Ошибка: в «Заказы» не найдены колонки Номенклатура/Склад"  
data = []  
for r in rows:  
name = r[m["name"]] if m["name"] < len(r) else None  
wh = r[m["wh"]] if m["wh"] < len(r) else None  
if name is None or wh is None or norm(name) == "": continue  
qcol = m.get("demand", m.get("qty"))  
data.append({"name": str(name), "wh": str(wh),  
"date": _date(r[m["date"]]) if "date" in m and m["date"] < len(r) else None,  
"qty": num(r[qcol]) if qcol is not None and qcol < len(r) else 0.0})  
if not data: return "Ошибка: «Заказы» — пустая таблица"  
self.orders = data; return f"Заказы: строк {len(data)}"  
def import_prod(self, path) -> str:  
wb = load_wb(path); ws = wb.active  
rows = list(ws.iter_rows(values_only=True))  
hi = None  
for i, r in enumerate(rows):  
cells = [str(c).strip().lower() if c is not None else "" for c in r]  
if any("дата исполнения" in c for c in cells): hi = i; break  
if hi is not None:  # старая структура  
shop_hdr = rows[hi]; col_shop = {}; cur = None  
for j, v in enumerate(shop_hdr):  
s = str(v).strip() if v is not None else ""  
if s: cur = s  
if cur: col_shop[j] = cur  
sub, si = None, None  
for i in range(hi + 1, min(hi + 5, len(rows))):  
cells = [str(c).strip().lower() if c is not None else "" for c in rows[i]]  
if any(("размещено" in c) or ("произведено" in c) or ("осталось" in c) for c in cells):  
sub, si = cells, i; break  
if si is not None:  
col_field = {}; cur = None  
for j, v in enumerate(sub):  
if v: cur = v  
if cur: col_field[j] = cur  
data = []  
for r in rows[si + 1:]:  
cells = list(r); date = None  
for v in cells:  
d = _date(v)  
if d: date = d; break  
name = None  
for j in (0, 1, 2):  
if j >= len(cells): continue  
v = cells[j]  
if v is None or isinstance(v, (int, float)): continue  
s = str(v).strip()  
if not s or s.lower().startswith(("итог", "заказ", "дата", "номенклатура")): continue  
if s.replace(".", "").replace("-", "").replace("#", "").isdigit(): continue  
name = s; break  
if name is None or date is None: continue  
if norm(name).startswith(SKIP_PREFIX): continue  
for j, v in enumerate(cells):  
shop = col_shop.get(j); fld = col_field.get(j)  
if not shop or not fld or "осталось" not in fld: continue  
wh = PROD_TO_WH.get(shop) or PROD_TO_WH.get(PROD_TO_WH_N.get(norm(shop), ""), None)  
if not wh: continue  
q = num(v)  
if q > 0: data.append({"name": name, "wh": wh, "date": date, "qty": q})  
if data:  
self.prod = data  
return f"Производство: строк {len(data)}
(поступление на склады по цехам)"  
si = None  # новая структура  
for i, r in enumerate(rows):  
cells = [str(c).strip().lower() if c is not None else "" for c in r]  
if any("размещено в заказе" in c for c in cells): si = i; break  
if si is None:  
return "Ошибка: в «Производство» не найдена шапка Размещено/Произведено/Осталось"  
hdr = [str(c).strip().lower() if c is not None else "" for c in rows[si]]  
i_left = next((j for j, h in enumerate(hdr) if "осталось" in h), 3)  
data = []; cur_shop = None  
for r in rows[si + 1:]:  
if not r: continue  
name = str(r[0]).strip() if r[0] is not None else ""  
if not name: continue  
low = norm(name)  
if low.startswith("итог"): cur_shop = None; continue  
if low in PROD_TO_WH_N: cur_shop = PROD_TO_WH_N[low]; continue  
if low.startswith(("подразделение", "номенклатура", "заказ", "дата", "отборы",  
"дополнительные", "период",
"показатели", "группировки", "ведомость")): continue  
if cur_shop is None: continue  
q = num(r[i_left]) if i_left < len(r) else 0.0  
if q > 0: data.append({"name": name, "wh": PROD_TO_WH[cur_shop], "date": None, "qty": q})  
if not data: return "Ошибка: «Производство» — нет строк с «Осталось выполнить» > 0"  
self.prod = data  
return f"Производство: строк {len(data)} (поступление на склады по цехам)"  
def ready(self): return bool(self.stock and self.orders)  
def build(self, today):  
dates = sorted({r["date"] for r in self.stock + self.orders  
if r["date"] is not None and r["date"] >= today})  
if not dates: return [], {}  
n = len(dates); idx = {d: i for i, d in enumerate(dates)}  
blocks: Dict[str, NomenBlock] = {}  
def wh_key(w): return CANON_ALL.get(norm(w), w)  
for r in self.stock:  
blk = blocks.setdefault(norm(r["name"]), NomenBlock(name=r["name"],
controlled=r["ctrl"]))  
wb_ = blk.wh.setdefault(wh_key(r["wh"]), WhBlock())  
if r["level"]: wb_.level = r["level"]  
di = idx.get(r["date"])  
if di is None: continue  
if r["open"] is not None: wb_.indicators.setdefault("Начальный остаток", [0.0]\*n)[di] = r["open"]  
if r["close"] is not None: wb_.indicators.setdefault("Конечный остаток", [0.0]\*n)[di] = r["close"]  
for r in self.orders:  
blk = blocks.setdefault(norm(r["name"]), NomenBlock(name=r["name"]))  
wb_ = blk.wh.setdefault(wh_key(r["wh"]), WhBlock())  
di = idx.get(r["date"])  
if di is None: continue  
wb_.indicators.setdefault("Потребность к отгрузке", [0.0]\*n)[di] += r["qty"]  
for r in self.prod:  
blk = blocks.setdefault(norm(r["name"]), NomenBlock(name=r["name"]))  
wb_ = blk.wh.setdefault(wh_key(r["wh"]), WhBlock())  
di = idx.get(r["date"]) if r.get("date") else 0  
if di is None: continue  
wb_.indicators.setdefault("План поступления", [0.0]\*n)[di] += r["qty"]  
return dates, blocks  

def _date(v):  
if isinstance(v, datetime.datetime): return v.date()  
if isinstance(v, datetime.date): return v  
m = re.search(r"(\\d{2})\\.(\\d{2})\\.(\\d{4})", str(v or ""))  
if m: return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))  
return None  

# ============================== ПЛАНЕР ==============================  
class Planner:  
def __init__(self, kb): self.kb = kb  
def _layers(self, h): return max(1, math.floor((MAX_H_MM - PALLET_MM) / h)) if h > 0 else 0  
def _balance(self, wb_, dates):  
if wb_ is None: return None  
n = len(dates)  
init = wb_.indicators.get("Начальный остаток", [0]\*n)  
move = wb_.indicators.get("Движение", [0]\*n)  
rec = wb_.indicators.get("План поступления", [0]\*n)  
dem = wb_.indicators.get("Потребность к отгрузке", [0]\*n)  
stored = wb_.indicators.get("Конечный остаток", [])  
if not (any(init) or any(move) or any(rec) or any(dem)):  
return list(stored) if stored else [0.0]\*n  
out, b = [], 0.0  
for d in range(n):  
b += (init[d] if d == 0 and d < len(init) else 0.0)  
b += (move[d] if d < len(move) else 0.0)  
b += (rec[d] if d < len(rec) else 0.0)  
b -= (dem[d] if d < len(dem) else 0.0)  
out.append(b)  
return out  
@staticmethod  
def _neg_days(bal):  
return bal is not None and any(bal[d] < 0 for d in (0, 1, 2, 3) if d < len(bal))  
def _recv_neg(self, rb, dates):  
if rb is None: return False  
if self._neg_days(self._balance(rb, dates)): return True  
return self._neg_days(rb.indicators.get("Конечный остаток", []))  
def _pack_info(self, name):  
w = []  
att = self.kb.pack_attrs(name)  
per_row = att["row"] if att and att["row"] else 0  
corr = self.kb.corr_of(name)  
gofra_name = corr[0] if corr else ""  
h = corr[1] if corr else 0  
if not corr: w.append(f"{name}: нет соответствия гофры — без проверки высоты")  
if not h and gofra_name:  
h = self.kb.gofra.get(norm(gofra_name), 0)  
if not h: w.append(f"{name}: гофра «{gofra_name}» не найдена в «Гофра высота»")  
if not h:  
if att and att["high"]: return per_row, att["high"], gofra_name, 0, w  
return per_row, 0, gofra_name, 0, w + [f"{name}: нет высоты гофры"]  
base_layers = max(1, self._layers(h) - 1)  
per_pallet = base_layers \* per_row if per_row else (att["high"] if att else 0)  
height_mm = base_layers \* h + PALLET_MM  
if height_mm > MAX_H_MM: w.append(f"{name}: высота {height_mm} > 2250")  
return per_row, per_pallet, gofra_name, height_mm, w  
def _manual_item(self, name, pallets):  
att = self.kb.pack_attrs(name)  
per_row = att["row"] if att and att["row"] else 0  
corr = self.kb.corr_of(name)  
gofra_name = corr[0] if corr else ""  
h = corr[1] if corr else 0  
if not h and gofra_name: h = self.kb.gofra.get(norm(gofra_name), 0)  
per_pallet = 0.0; height_mm = 0.0  
if h:  
base_layers = max(1, self._layers(h) - 1)  
per_pallet = base_layers \* per_row if per_row else (att["high"] if att else 0)  
height_mm = base_layers \* h + PALLET_MM  
elif att and att["high"]:  
per_pallet = att["high"]  
qty = int(round(pallets \* per_pallet)) if per_pallet else 0  
return LineItem(new_uid(), name, per_row, per_pallet, qty, pallets, gofra_name, height_mm, True)  
def plan(self, dates, blocks, manual_trips, profile) -> Plan:  
plan = Plan(dates=dates)  
plan.trips = {D_SKOP_MAIN: [], D_MAIN_SKOP: []}  
for dk in plan.trips: self._ensure_trips(plan, dk, 1, manual_trips, profile)  
kg_main = kg_skop = 0.0  
neg_main = neg_skop = 0  
for blk in blocks.values():  
if self.kb.is_service(blk.name): continue  
kg = self.kb.kg_pack(blk.name)  
bm = blk.wh.get(WH_MAIN); bs = blk.wh.get(WH_SKOP)  
e0m = bm.indicators.get("Конечный остаток", [0]) if bm else [0]  
e0s = bs.indicators.get("Конечный остаток", [0]) if bs else [0]  
if kg is not None:  
kg_main += max(0.0, e0m[0] if e0m else 0) \* kg  
kg_skop += max(0.0, e0s[0] if e0s else 0) \* kg  
if self._recv_neg(bm, dates): neg_main += 1  
if self._recv_neg(bs, dates): neg_skop += 1  
for recv, src, dkey in ((WH_MAIN, WH_SKOP, D_SKOP_MAIN), (WH_SKOP, WH_MAIN, D_MAIN_SKOP)):  
rb = blk.wh.get(recv)  
if not rb: continue  
sb = blk.wh.get(src)  
end_r = self._balance(rb, dates); end_s = self._balance(sb, dates)  
recv_neg = self._recv_neg(rb, dates); src_neg = self._recv_neg(sb, dates)  
moved = False; pin = 0.0  
if recv_neg:  
for dd in (0, 1, 2, 3):  
if dd >= len(dates): break  
if dd == 3 and rb.level not in (4,): continue  
need = max(0.0, -(end_r[dd] if dd < len(end_r) else 0) - pin)  
if need > 0: moved = True  
src_own = end_s[dd] if (end_s and dd < len(end_s)) else 0  
pin += max(0.0, min(need, max(0.0, src_own)))  
if not moved and dkey == D_SKOP_MAIN and blk.controlled > 0 and not src_neg:  
cap = max(0.0, 3.5 \* blk.controlled - ((end_r[0] if end_r else 0) + pin))  
src_own = end_s[1] if (end_s and len(end_s) > 1) else 0  
if min(blk.controlled, cap, max(0.0, src_own)) > 0.5: moved = True  
if not moved: continue  
plan.moved += 1  
per_row, per_pallet, gofra, h_mm, wl = self._pack_info(blk.name)  
plan.warnings += wl  
if per_pallet <= 0:  
plan.warnings.append(f"{blk.name}: нет данных паллетизации — позиция пропущена")  
plan.uncovered.add((norm(blk.name), dkey)); continue  
topup = (dkey != D_MAIN_SKOP)  
planned_in = planned_out = 0.0  
for dd in (0, 1, 2, 3):  
if dd >= len(dates): break  
if dd == 3 and rb.level not in (4,): continue  
need = max(0.0, -(end_r[dd] if dd < len(end_r) else 0) - planned_in)  
if need <= 0: continue  
src_own = end_s[dd] if (end_s and dd < len(end_s)) else 0  
avail = max(0.0, src_own) - planned_out  
if src_own < 0 and avail <= 0:  
plan.insufficient.append((blk.name, src, round(need, 2)))  
plan.uncovered.add((norm(blk.name), dkey)); continue  
raw = min(need, avail)  
if raw <= 0:  
plan.underload.append((blk.name, str(dates[dd]), round(need, 2)))  
plan.uncovered.add((norm(blk.name), dkey)); continue  
qty = math.ceil(raw - 1e-9)  
if topup and per_row:  
ql = math.ceil(qty / per_row) \* per_row  
if ql <= avail: qty = ql  
rest = (avail - qty) / per_pallet if per_pallet else 0  
q = 0.5 if rest > 5 else 0.1  
pal = math.ceil((qty / per_pallet) / q - 1e-9) \* q  
self._place(plan, dkey, LineItem(new_uid(), blk.name, per_row, per_pallet,  
qty, pal, gofra, h_mm, False, dd + 1),  
1 if dd in (0, 1, 3) else 2, manual_trips, profile)  
planned_in += qty; planned_out += qty  
ctrl = blk.controlled  
if dkey == D_SKOP_MAIN and ctrl > 0 and not src_neg:  
cap = max(0.0, 3.5 \* ctrl - ((end_r[0] if end_r else 0) + planned_in))  
src_own = end_s[1] if (end_s and len(end_s) > 1) else 0  
avail = max(0.0, src_own) - planned_out  
ins = min(ctrl, cap, avail)  
if ins > 0.5:  
t1_free = plan.trips[dkey][0].free() if plan.trips[dkey] else 0.0  
if recv_neg or t1_free > 0 or manual_trips.get(dkey, 0) > 1:  
qty = math.ceil(ins - 1e-9)  
if per_row:  
ql = math.ceil(qty / per_row) \* per_row  
if ql <= avail: qty = ql  
rest = (avail - qty) / per_pallet if per_pallet else 0  
q = 0.5 if rest > 5 else 0.1  
pal = math.ceil((qty / per_pallet) / q - 1e-9) \* q  
self._place_soft(plan, dkey, LineItem(new_uid(), blk.name,  
per_row, per_pallet, qty, pal, gofra, h_mm), manual_trips)  
plan.metrics = {"kg_main": round(kg_main, 1), "kg_skop": round(kg_skop, 1)}  
for dkey, trips in plan.trips.items():  
for ti, t in enumerate(trips):  
for m in profile.get("manual", {}).get(f"{dkey}#{ti}", []):  
t.items.append(self._manual_item(m["name"], m["pallets"]))  
if not any(t.items for t in plan.trips[D_MAIN_SKOP]) and neg_skop:  
plan.warnings.insert(0, f"Внимание: на Скопин РЦ позиций с дефицитом: {neg_skop}, но рейс Основной->Скопин пуст.")  
if not any(t.items for t in plan.trips[D_SKOP_MAIN]) and neg_main:  
plan.warnings.insert(0, f"Внимание: на Склад основной позиций с дефицитом: {neg_main}, но рейс Скопин->Основной пуст.")  
return plan  
def _ensure_trips(self, plan, dkey, n, manual_trips, profile):  
plan.trips.setdefault(dkey, [])  
while len(plan.trips[dkey]) < max(n, manual_trips.get(dkey, 0)):  
t = Trip(len(plan.trips[dkey]), dkey)  
t.capacity = profile.get("caps", {}).get(f"{dkey}#{t.index}", DEF_CAP)  
plan.trips[dkey].append(t)  
def _place(self, plan, dkey, item, trip_no, manual_trips, profile):  
forced = manual_trips.get(dkey, 0)  
self._ensure_trips(plan, dkey, trip_no if forced == 0 else 1, manual_trips, profile)  
trips = plan.trips[dkey]  
want = min(trip_no, len(trips)) - 1  
order = [want] + [i for i in range(len(trips)) if i != want]  
for i in order:  
if trips[i].free() >= item.pallets:  
trips[i].items.append(item); return  
if not forced:  
t = Trip(len(trips), dkey)  
t.capacity = profile.get("caps", {}).get(f"{dkey}#{t.index}", DEF_CAP)  
trips.append(t); t.items.append(item); return  
t = trips[want if want < len(trips) else 0]  
t.items.append(item)  
plan.warnings.append(f"{item.name}: рейс {t.index+1} ({dkey}) перегружен ({t.occupied()}/{t.capacity})")  
def _place_soft(self, plan, dkey, item, manual_trips):  
trips = plan.trips[dkey]  
forced = manual_trips.get(dkey, 0)  
limit = forced if forced > 0 else len(trips)  
for i in range(min(limit, len(trips))):  
if trips[i].free() >= item.pallets:  
trips[i].items.append(item); return  

# ============================== ЭКСПОРТ ==============================  
def export_excel(plan, path):  
wb = openpyxl.Workbook(); wb.remove(wb.active)  
for dkey, trips in plan.trips.items():  
for t in trips:  
if not t.items: continue  
ws = wb.create_sheet(f"{('Скопин->Основн' if dkey.startswith(WH_SKOP) else 'Основн->Скопин')} Р{t.index+1}")  
ws.append(["Рейс", t.index+1, "Направление", dkey,
"Вместимость", t.capacity,  
"Занято", t.occupied(), "Свободно", t.free(), "Макс. высота", t.max_height()])  
ws.append(["№", "Номенклатура", "Упак в слое", "Упак на паллете", "Упак общее",  
"Паллетомест", "В заказе", "Гофра", "Высота паллеты"])  
for n, i in enumerate(t.items, 1):  
ws.append([n, i.name, i.per_layer or "", i.per_pallet or "", i.qty or "",  
i.pallets, i.qty or "", i.gofra, i.height_mm or ""])  
ws = wb.create_sheet("Предупреждения")  
for w in plan.warnings: ws.append([w])  
ws = wb.create_sheet("Недостаточные остатки")  
for r in plan.insufficient: ws.append(list(r))  
ws = wb.create_sheet("Недогрузы")  
for r in plan.underload: ws.append(list(r))  
wb.save(path)  

# ============================== GUI ==============================  
STYLE = """  
QWidget{background:#f4f6f8;font-family:Segoe UI;font-size:12px}  
QFrame#card,QFrame#panel,QFrame#metrics{background:#fbfcfd;border:1px solid
#d7dde3;border-radius:12px}  
QLabel#h1{color:#1a6fd4;font-size:16px;font-weight:700}  
QLabel#tag{background:#dcebfb;color:#1a6fd4;border-radius:6px;padding:3px 8px;font-weight:600}  
QLabel#build{color:#8a94a0;font-size:10px}  
QPushButton{border-radius:8px;padding:6px 14px;border:1px solid #cfd6dd;background:#fff}  
QPushButton#accent{background:#e04a2f;color:#fff;font-weight:700;border:none;font-size:15px;padding:14px 30px;border-radius:10px}  
QFrame#impgroup{border:1.5px solid #3a3a3a;border-radius:14px;background:#fbfcfd}  
QLabel#impgrouplab{font-size:10px;color:#333333}  
QPushButton[doc="1"]{border:1.5px solid #9fb6cc;border-radius:10px;background:#ffffff;padding:10px 12px}  
QPushButton[doc="1"][st="ok"]{border:2px solid #2f9e44}  
QTreeWidget{border-radius:8px;border:1px solid #e1e6eb;background:#fff}  
QSpinBox{border-radius:6px;border:1px solid #cfd6dd;padding:2px}  
QTabWidget::pane{border:1px solid #d7dde3;border-radius:10px;background:#fbfcfd}  
"""  

class TripTree(QTreeWidget):  
def __init__(self, trip, owner):  
super().__init__(); self.trip, self.owner = trip, owner  
self.setColumnCount(7)  
self.setHeaderLabels(["№", "Номенклатура", "Уп/слой", "Упак к перемещению",  
"Паллет мест", "Ур. потребности", ""])  
self.setDragDropMode(QAbstractItemView.DragDropMode.DragDrop)  
self.setDefaultDropAction(Qt.MoveAction)  
self.setSelectionMode(QAbstractItemView.ExtendedSelection)  
self.setSelectionBehavior(QAbstractItemView.SelectItems)  
hdr = self.header()  
hdr.setSectionResizeMode(QHeaderView.Interactive)  
hdr.setStretchLastSection(False)  
for i, w in enumerate((40, 380, 70, 130, 90, 110, 30)):  
self.setColumnWidth(i, w)  
self.itemClicked.connect(self.owner.on_tree_click)  
self.refresh()  
@staticmethod  
def _align(it, ncols):  
for c in range(ncols):  
it.setTextAlignment(c, Qt.AlignLeft | Qt.AlignVCenter if c in (0, 1) else Qt.AlignCenter | Qt.AlignVCenter)  
def refresh(self):  
self.clear()  
unc = self.owner.plan.uncovered if self.owner.plan else set()  
tot_q = sum(i.qty for i in self.trip.items); tot_p = sum(i.pallets for i in self.trip.items)  
it = QTreeWidgetItem(["", "Итог", "", str(int(round(tot_q))) if tot_q else "0",  
str(round(tot_p, 2)) if tot_p else "0", "", ""])  
it.setFlags(Qt.ItemIsEnabled)  
for c in range(self.columnCount()):  
f = QFont(); f.setBold(True); it.setFont(c, f); it.setBackground(c, QColor("#e8eef4"))  
self._align(it, self.columnCount()); self.addTopLevelItem(it)  
for n, i in enumerate(self.trip.items, 1):  
closed = True  
if not i.manual: closed = (norm(i.name), self.trip.direction) not in unc  
it = QTreeWidgetItem([str(n), i.name,  
str(int(round(i.per_layer))) if i.per_layer else "",  
str(int(round(i.qty))) if i.qty else "",  
str(round(i.pallets, 2)) if i.pallets else "",  
"" if (i.manual or not i.level_min) else str(i.level_min), "✕"])  
if not i.manual:  
it.setBackground(5, QColor("#8fd694") if closed else QColor("#f2a2a2"))  
it.setForeground(6, QColor("#d43c3c"))  
it.setData(0, Qt.UserRole, i.uid)  
self._align(it, self.columnCount()); self.addTopLevelItem(it)  
def keyPressEvent(self, e):  
if e.matches(QKeySequence.Copy):  
idxs = self.selectedIndexes()  
if idxs:  
rows = {}  
for ix in idxs:  
if ix.column() >= self.columnCount() - 1: continue  
rows.setdefault(ix.row(), []).append((ix.column(), str(ix.data() or "")))  
QApplication.clipboard().setText("\\n".join(  
"\\t".join(t for _, t in sorted(rows[r])) for r in sorted(rows)))  
return  
super().keyPressEvent(e)  
def mimeTypes(self): return [MIME]  
def mimeData(self, items):  
md = QMimeData(); md.setData(MIME, QByteArray(f"{items[0].data(0, Qt.UserRole)}".encode())); return md  
def dragEnterEvent(self, e): e.accept() if MIME in e.mimeData().formats() else e.ignore()  
def dragMoveEvent(self, e): e.accept()  
def dropEvent(self, e):  
uid = int(bytes(e.mimeData().data(MIME)).decode())  
e.accept(); self.owner.on_drop(uid, self.trip)  

class TripCard(QFrame):  
def __init__(self, trip, owner):  
super().__init__(); self.setObjectName("card"); self.trip, self.owner = trip, owner  
v = QVBoxLayout(self); h = QHBoxLayout()  
t = QLabel(f"Рейс {trip.index+1} · {('Скопин РЦ → Основной' if trip.direction.startswith(WH_SKOP) else 'Основной → Скопин РЦ')}")  
t.setObjectName("tag"); h.addWidget(t); h.addWidget(QLabel("Паллет:"))  
self.cap = QSpinBox(); self.cap.setRange(0, MAX_CAP); self.cap.setValue(trip.capacity)  
self.cap.valueChanged.connect(self.on_cap); h.addWidget(self.cap)  
self.lbl = QLabel(); h.addWidget(self.lbl); h.addStretch(); v.addLayout(h)  
self.tree = TripTree(trip, owner)  
self.tree.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  
self.tree.setMinimumHeight(180)  
v.addWidget(self.tree)  
b = QPushButton("+ ручная строка"); b.clicked.connect(self.add_manual); v.addWidget(b)  
self.update_lbl()  
def on_cap(self, v): self.trip.capacity = v; self.owner.save_profile(); self.update_lbl()  
def update_lbl(self):  
self.lbl.setText(f"занято {self.trip.occupied()} / свободно {self.trip.free()} / макс. высота {int(self.trip.max_height())} мм")  
def add_manual(self):  
name, ok = QInputDialog.getText(self, "Ручная строка", "Наименование / комментарий:")  
if not ok or not name.strip(): return  
p, ok2 = QInputDialog.getDouble(self, "Ручная строка",
"Паллетоместа:", 1.0, 0.0, 36.0, 1)  
if not ok2: return  
self.trip.items.append(self.owner.planner._manual_item(name, p))  
self.owner.save_profile_manual(self.trip, name, p)  
self.tree.refresh(); self.update_lbl()  

class MainWindow(QMainWindow):  
def __init__(self):  
super().__init__()  
self.setWindowTitle(f"Order Fulfillment Planner [{BUILD}]"); self.setStyleSheet(STYLE)  
self.kb = KnowledgeBase(); self.planner = Planner(self.kb)  
self.data = SupplyData(); self.dates, self.blocks = [], {}  
self.manual_trips = {D_SKOP_MAIN: 1, D_MAIN_SKOP: 1}  
self.profile = self._load_profile()  
self._build_ui()  
self.plan = Plan(dates=[])  
self.plan.trips = {D_SKOP_MAIN: [], D_MAIN_SKOP: []}  
for dk in self.plan.trips:  
self.planner._ensure_trips(self.plan, dk, 1, self.manual_trips, self.profile)  
self._render_cards(); self.refresh_doc_states()  
def _load_profile(self):  
try: return json.load(open(PROFILE_PATH, encoding="utf-8"))  
except Exception: return {"caps": {}, "manual": {}}  
def save_profile(self):  
try:  
if self.plan:  
for dkey, trips in self.plan.trips.items():  
for t in trips: self.profile["caps"][f"{dkey}#{t.index}"] = t.capacity  
json.dump(self.profile, open(PROFILE_PATH, "w", encoding="utf-8"), ensure_ascii=False)  
except Exception as e: print("Не удалось сохранить профиль:", e)  
def save_profile_manual(self, trip, name, p):  
self.profile.setdefault("manual", {}).setdefault(f"{trip.direction}#{trip.index}",  
[]).append({"name": name, "pallets": p})  
self.save_profile()  
def refresh_doc_states(self):  
def setst(btn, ok):  
btn.setProperty("st", "ok" if ok else "")  
btn.style().unpolish(btn); btn.style().polish(btn)  
setst(self.btn_stock, bool(self.data.stock)); setst(self.btn_orders, bool(self.data.orders))  
setst(self.btn_prod, bool(self.data.prod)); setst(self.btn_nomen, bool(self.kb.nomen))  
setst(self.btn_gofra, bool(self.kb.gofra)); setst(self.btn_corr, bool(self.kb.corr))  
def _build_ui(self):  
cw = QWidget(); self.setCentralWidget(cw); root = QVBoxLayout(cw)  
h = QHBoxLayout()  
t = QLabel("Order Fulfillment Planner"); t.setObjectName("h1"); h.addWidget(t)  
self.date_lbl = QLabel(datetime.date.today().strftime("%d.%m.%Y"))  
self.date_lbl.setObjectName("tag"); h.addWidget(self.date_lbl)  
bl = QLabel(BUILD); bl.setObjectName("build"); h.addWidget(bl); h.addStretch()  
root.addLayout(h)  
tb = QHBoxLayout()  
self.imp_frame = QFrame(); self.imp_frame.setObjectName("impgroup")  
gv = QVBoxLayout(self.imp_frame); gv.setContentsMargins(10, 8, 10, 2); gv.setSpacing(2)  
grow = QHBoxLayout(); grow.setSpacing(8)  
self.btn_stock = QPushButton("Товары на складах");
self.btn_stock.clicked.connect(self.imp_stock)  
self.btn_orders = QPushButton("Заказы");
self.btn_orders.clicked.connect(self.imp_orders)  
self.btn_prod = QPushButton("Производство");
self.btn_prod.clicked.connect(self.imp_prod)  
self.btn_nomen = QPushButton("Номенклатура-справочник"); self.btn_nomen.clicked.connect(self.imp_nomen)  
self.btn_gofra = QPushButton("Гофра"); self.btn_gofra.clicked.connect(self.imp_gofra)  
self.btn_corr = QPushButton("Номенклатура");
self.btn_corr.clicked.connect(self.imp_corr)  
for b in (self.btn_stock, self.btn_orders, self.btn_prod, self.btn_nomen, self.btn_gofra, self.btn_corr):  
b.setProperty("doc", "1"); b.setProperty("st", ""); grow.addWidget(b)  
gv.addLayout(grow)  
glab = QLabel("Импорт"); glab.setObjectName("impgrouplab");
glab.setAlignment(Qt.AlignCenter)  
gv.addWidget(glab); tb.addWidget(self.imp_frame)  
self.status = QLabel(""); tb.addWidget(self.status); tb.addStretch()  
acc = QPushButton("Рассчитать"); acc.setObjectName("accent");
acc.clicked.connect(self.calc)  
tb.addWidget(acc)  
right = QVBoxLayout(); right.setSpacing(2)  
b_rules = QPushButton("Правила")  
b_rules.clicked.connect(lambda: QMessageBox.information(self, "База знаний", RULES_TEXT))  
right.addWidget(b_rules)  
exp = QPushButton("Экспорт Excel"); exp.clicked.connect(self.export); right.addWidget(exp)  
tb.addLayout(right); root.addLayout(tb)  
sp = QSplitter(Qt.Horizontal); self.trip_split = {}  
for dkey in (D_SKOP_MAIN, D_MAIN_SKOP):  
panel = QFrame(); panel.setObjectName("panel"); v = QVBoxLayout(panel)  
lab = QLabel("Скопин РЦ → Склад Основной" if
dkey.startswith(WH_SKOP) else "Склад Основной → Скопин РЦ")  
lab.setObjectName("tag"); v.addWidget(lab)  
hr = QHBoxLayout(); hr.addWidget(QLabel("Рейсов:"))  
sb = QSpinBox(); sb.setRange(0, 6); sb.setValue(1)  
sb.valueChanged.connect(lambda v, k=dkey: self.manual_trips.update({k: v}))  
hr.addWidget(sb); hr.addStretch(); v.addLayout(hr)  
sc = QScrollArea(); sc.setWidgetResizable(True)  
inner = QWidget(); lay = QVBoxLayout(inner); lay.setContentsMargins(0, 0, 0, 0)  
self.trip_split[dkey] = QSplitter(Qt.Vertical)  
lay.addWidget(self.trip_split[dkey])  
sc.setWidget(inner); v.addWidget(sc); sp.addWidget(panel)  
self.vsp = QSplitter(Qt.Vertical); self.vsp.addWidget(sp)  
m = QFrame(); m.setObjectName("metrics"); vm = QVBoxLayout(m)  
self.metrics_lbl = QLabel("Показатели: —"); vm.addWidget(self.metrics_lbl)  
self.tabs = QTabWidget()  
self.warn_tbl = QPlainTextEdit(); self.tabs.addTab(self.warn_tbl,
"Предупреждения")  
self.ins_tbl = QTableWidget(); self.ins_tbl.setColumnCount(3)  
self.ins_tbl.setHorizontalHeaderLabels(["Номенклатура", "Склад", "Дефицит"])  
self.tabs.addTab(self.ins_tbl, "Недостаточные остатки")  
self.und_tbl = QTableWidget(); self.und_tbl.setColumnCount(3)  
self.und_tbl.setHorizontalHeaderLabels(["Номенклатура", "День", "Непокрыто"])  
self.tabs.addTab(self.und_tbl, "Недогрузы")  
vm.addWidget(self.tabs); self.vsp.addWidget(m)  
self.vsp.setStretchFactor(0, 3); self.vsp.setStretchFactor(1, 1)  
root.addWidget(self.vsp); self.vsp.setSizes([680, 220])  
def _render_cards(self):  
for dkey, spl in self.trip_split.items():  
while spl.count():  
w = spl.widget(0); w.setParent(None); w.deleteLater()  
trips = self.plan.trips.get(dkey, [])  
for t in trips: spl.addWidget(TripCard(t, self))  
spacer = QWidget(); spl.addWidget(spacer)  
for i in range(spl.count()): spl.setStretchFactor(i, 0)  
if len(trips) == 1:  
spl.setStretchFactor(0, 1); spl.setStretchFactor(spl.count() - 1, 0)  
else:  
spl.setStretchFactor(spl.count() - 1, 1)  
def _open(self): return QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel (\*.xlsx \*.xls)")[0]  
def imp_stock(self):  
p = self._open()  
if not p: return  
try: self.status.setText(self.data.import_stock(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def imp_orders(self):  
p = self._open()  
if not p: return  
try: self.status.setText(self.data.import_orders(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def imp_prod(self):  
p = self._open()  
if not p: return  
try: self.status.setText(self.data.import_prod(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def imp_nomen(self):  
p = self._open()  
if p:  
try: self.status.setText(self.kb.import_nomen(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def imp_gofra(self):  
p = self._open()  
if p:  
try: self.status.setText(self.kb.import_gofra(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def imp_corr(self):  
p = self._open()  
if p:  
try: self.status.setText(self.kb.import_corr(p)); self.refresh_doc_states()  
except Exception as e: QMessageBox.critical(self, "Ошибка", str(e))  
def on_tree_click(self, it, col):  
if col != 6: return  
uid = it.data(0, Qt.UserRole)  
if uid is None: return  
name = it.text(1)  
ans = QMessageBox.question(self, "Удаление строки", f"Удалить строку «{name}» из рейса?",  
QMessageBox.Yes | QMessageBox.No, QMessageBox.No)  
if ans != QMessageBox.Yes: return  
removed = None  
for dkey, trips in self.plan.trips.items():  
for t in trips:  
for i in list(t.items):  
if i.uid == uid: t.items.remove(i); removed = (dkey, t.index, i)  
if removed is not None and removed[2].manual:  
dkey, ti, item = removed  
lst = self.profile.get("manual", {}).get(f"{dkey}#{ti}", [])  
for m in list(lst):  
if m["name"] == item.name: lst.remove(m); break  
for spl in self.trip_split.values():  
for i in range(spl.count()):  
w = spl.widget(i)  
if isinstance(w, TripCard): w.tree.refresh(); w.update_lbl()  
self.save_profile()  
def calc(self):  
if not self.data.ready():  
QMessageBox.warning(self, "Нет данных", "Загрузите «Товары на складах» и «Заказы» (и «Производство» при наличии)."); return  
today = datetime.date.today()  
dates, blocks = self.data.build(today)  
if not dates or not blocks:  
QMessageBox.critical(self, "Ошибка импорта", "В файлах не найдено дат/позиций."); return  
if dates[0] != today:  
QMessageBox.critical(self, "Ошибка даты",  
f"Первая дата файлов {dates[0]:%d.%m.%Y} не совпадает с текущей {today:%d.%m.%Y}."); return  
self.dates, self.blocks = dates, blocks  
try:  
self.plan = self.planner.plan(self.dates, self.blocks, self.manual_trips, self.profile)  
except Exception as e:  
QMessageBox.critical(self, "Ошибка расчёта",
f"{e}\\n\\n{traceback.format_exc()}"); return  
if self.plan.moved == 0 and self.blocks:  
neg = 0; neg_lines = []  
for b in self.blocks.values():  
for wname in (WH_MAIN, WH_SKOP):  
wblk = b.wh.get(wname)  
if not wblk: continue  
er = self.planner._balance(wblk, self.dates)  
st_er = wblk.indicators.get("Конечный остаток", [])  
if (er and any(d < len(er) and er[d] < 0 for d in (0, 1, 2, 3))) or \\  
any(d < len(st_er) and st_er[d] < 0 for d in (0, 1, 2, 3)):  
neg += 1  
if len(neg_lines) < 6:  
neg_lines.append(f"{b.name} | {wname} | конечный: {[round(x, 1) for x in (er or st_er)[:4]]}")  
msg = (f"Нет позиций с потребностью. Позиций: {len(self.blocks)}; "  
f"складов с отрицательным конечным остатком в днях 0-3: {neg}.")  
if neg_lines: msg += " Примеры: " + " // ".join(neg_lines)  
self.plan.warnings.insert(0, msg)  
self._render_cards()  
rows = sum(len(t.items) for tr in self.plan.trips.values() for t in tr)  
self.metrics_lbl.setText(  
f"Запасы, кг: Склад основной = {self.plan.metrics['kg_main']} | "  
f"Скопин РЦ = {self.plan.metrics['kg_skop']} | применён лимит высоты {MAX_H_MM} мм (поддон {PALLET_MM} мм)")  
self.status.setText(f"[{BUILD}] Позиций: {len(self.blocks)} | с
потребностью: {self.plan.moved} | "  
f"строк в рейсах: {rows} | предупреждений: {len(self.plan.warnings)}")  
self.warn_tbl.setPlainText("\\n".join(self.plan.warnings) or "Предупреждений нет")  
self.ins_tbl.setRowCount(len(self.plan.insufficient))  
for r, row in enumerate(self.plan.insufficient):  
for c, v in enumerate(row): self.ins_tbl.setItem(r, c, QTableWidgetItem(str(v)))  
self.und_tbl.setRowCount(len(self.plan.underload))  
for r, row in enumerate(self.plan.underload):  
for c, v in enumerate(row): self.und_tbl.setItem(r, c, QTableWidgetItem(str(v)))  
self.save_profile()  
def on_drop(self, uid, target):  
if not self.plan: return  
src_trip = item = None  
for trips in self.plan.trips.values():  
for t in trips:  
for i in list(t.items):  
if i.uid == uid: src_trip, item = t, i  
if not item or src_trip is target: return  
src_trip.items.remove(item); target.items.append(item)  
for spl in self.trip_split.values():  
for i in range(spl.count()):  
w = spl.widget(i)  
if isinstance(w, TripCard): w.tree.refresh(); w.update_lbl()  
self.save_profile()  
def export(self):  
if not self.plan: QMessageBox.warning(self, "Нет данных", "Сначала выполните расчёт."); return  
p, _ = QFileDialog.getSaveFileName(self, "Сохранить отчёт", "plan.xlsx", "Excel (\*.xlsx)")  
if p: export_excel(self.plan, p); self.status.setText(f"Экспорт: {p}")  

if __name__ == "__main__":  
app = QApplication([])  
w = MainWindow(); w.resize(1500, 900); w.show()  
app.exec()  
\`\`\`  

## 9. Критерии приёмки  
1. Запуск без ошибок; в заголовке виден \`[v21 …]\`.  
2. Импорт 6 файлов → зелёные рамки; статусы с числом строк.  
3. «Рассчитать»: обе панели наполняются по дефицитам (пример из данных: «Вареники с картофелем и луком "СП" 350 г (8)», Скопин −204 → строка в Основной→Скопин с «Ур. потребности» 1); в Основной→Скопин НЕТ страховых строк; при «Рейсов: 1» ровно по одной карточке.  
4. Крестик удаляет строку с подтверждением и не возвращает её после пересчёта; Ctrl+C копирует выделенное; «Итог» считает суммы.  
5. Экспорт Excel содержит листы рейсов + «Предупреждения» + «Недостаточные остатки» + «Недогрузы».
 
