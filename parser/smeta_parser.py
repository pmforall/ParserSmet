#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
smeta_parser.py — парсер смет, выпущенных в ПК «АтомСмета» (Смета.ру).

Назначение
----------
Собирает данные со вкладки «Source» всех смет *.xlsx из указанной папки и
формирует структурированный JSON для последующей встройки в автономный
HTML-инструмент (см. build.py) либо для любой внешней обработки.

Что читается из вкладки «Source»
-------------------------------
«Source» — машинный дамп модели документа АтомСметы: каждая строка имеет
код типа записи в столбце A.

    A=0/1   — заголовок документа (шифр, наименование, основание, объект, НБ)
    A=3/4/5 — открытие узла структуры (документ / раздел / подраздел)
    A=17    — позиция локальной сметы (расценка, материал, оборудование, …)
    A=18    — строка спецификации (без стоимостных показателей)
    A=19    — примечание/пояснение
    A=50    — показатель узла (набор из 27 числовых показателей)
    A=51    — итоговая строка узла (содержит id узла, вид, шифр, наименование)
    A=52    — открытие узла структуры (дублирует A=51 по составу полей)
    A=70    — условия расчёта и коэффициенты (РК, виды работ, налогообложение)

Числовые столбцы записей A=17/A=51/A=52 (одинаковая раскладка):
    O  — прямые затраты (ПЗ), руб.
    P  — стоимость материальных ресурсов (материалы + оборудование), руб.
    Q  — эксплуатация машин (ЭММ), руб.
    R  — ЗП машинистов (ЗПМ), руб.
    S  — основная ЗП рабочих (ОЗП), руб.
    U  — трудозатраты строителей, чел.-ч
    V  — трудозатраты машинистов, чел.-ч
    X  — накладные расходы (НР), руб.
    Y  — сметная прибыль (СП), руб.
    GM — всего с НР и СП по позиции, руб.

Номенклатура показателей узла задаётся столбцом K записи A=50 (см. INDICATORS);
именно она в точности соответствует таблице технического задания.

Все суммы — в базисном уровне цен 01.01.2000 (НБ ФЕР-2001).

Запуск
------
    python smeta_parser.py <папка_со_сметами> [-o smetas.json] [--no-items]
                           [--report checks.json]

Зависимости: openpyxl (поставляется со средой DSH).
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import math
import os
import re
import sys
import traceback
from typing import Any, Dict, Iterable, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    import reference as ref_mod          # справочник укрупненных физобъемов и техники
except Exception:                        # pragma: no cover — модуль необязателен
    ref_mod = None

try:
    from openpyxl import load_workbook
except Exception as exc:  # pragma: no cover
    sys.stderr.write("Не найден openpyxl: %s\n" % exc)
    raise

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import printed_form as pf      # чтение печатной формы «…11 граф…»
import printed_positions as pp

SOURCE_SHEET = "Source"

# ПК «ГРАНД-Смета»: книги выгружаются печатной формой локального сметного расчёта
# (без вкладки «Source»). Разбор — в модуле grand_parser; формат определяется по
# признакам формы («ЛОКАЛЬНЫЙ СМЕТНЫЙ РАСЧЕТ», «№ п/п» + «Итоги по смете:»).
try:
    import grand_parser as gp       # чтение печатной формы ГРАНД-Сметы
except Exception:                   # pragma: no cover
    gp = None
EPS = 0.01  # допуск сверки, руб.
# трудозатраты (показатели 207/208) — в чел.-ч, хранятся с точностью 4 знаков

# --------------------------------------------------------------------------
# Номенклатура показателей узла (столбец K записей A=50)
# --------------------------------------------------------------------------
# code: (semantic key, подпись для таблицы ТЗ, единица измерения)
INDICATORS: Dict[int, Tuple[str, str, str]] = {
    224: ("total",      "Всего с НР и СП",              "руб"),
    205: ("ozp",        "Основная ЗП рабочих",          "руб"),
    203: ("emm",        "Эксплуатация машин",           "руб"),
    204: ("zpm",        "ЗП машинистов",                "руб"),
    226: ("material",   "Стоимость материалов (всего)", "руб"),
    233: ("transport",  "Перевозка грузов",             "руб"),
    210: ("nr",         "Накладные расходы",            "руб"),
    211: ("sp",         "Сметная прибыль",              "руб"),
    217: ("other",      "Прочие работы с НР и СП",      "руб"),
    216: ("equipment",  "Стоимость оборудования (всего)", "руб"),
    208: ("tz_mash",    "Трудозатраты машинистов",      "чел-ч"),
    207: ("tz_workers", "Трудозатраты строителей",      "чел-ч"),
    # дополнительные показатели (используются для контроля и аналитики)
    201: ("pz",         "Прямые затраты",               "руб"),
    202: ("mat_ob",     "Стоимость материальных ресурсов (всего)", "руб"),
    222: ("mat_ob_cust",  "… материалы и оборудование заказчика", "руб"),
    225: ("mat_ob_contr", "… материалы и оборудование подрядчика", "руб"),
    227: ("mat_cust",   "… материалы заказчика",        "руб"),
    228: ("mat_contr",  "… материалы подрядчика",       "руб"),
    223: ("equip_cust", "… оборудование заказчика",     "руб"),
    229: ("equip_contr", "… оборудование подрядчика",    "руб"),
    231: ("emm_tsn",    "ЭММ по ТСН-2001.16",           "руб"),
    232: ("ozp_tsn",    "ОЗП по ТСН-2001.16",           "руб"),
    230: ("other_cost", "Прочие затраты по ТСН-2001.16", "руб"),
    206: ("return_mat", "Возврат материалов",           "руб"),
    209: ("transport_mat", "Транспорт материалов",      "руб"),
    214: ("smr_build",  "Строительные работы с НР и СП", "руб"),
    215: ("smr_install", "Монтажные работы с НР и СП",  "руб"),
}

# Порядок столбцов таблицы ТЗ: (key, заголовок)
TABLE_COLUMNS: List[Tuple[str, str]] = [
    ("number",     "№ сметного расчета"),
    ("name",       "Наименование сметного расчета"),
    ("osnovanie",  "Основание"),
]
for _c in (224, 205, 203, 204, 226, 233, 210, 211, 217, 216, 208, 207):
    TABLE_COLUMNS.append((INDICATORS[_c][0], INDICATORS[_c][1] + (", чел/час" if INDICATORS[_c][2] == "чел-ч" else ", руб")))

# Коды видов затрат АтомСметы (столбец EK записи A=17)
VID_MATERIAL = {1100, 7010, 500001, 500002, 500003, 500004, 500005}
VID_EQUIPMENT = {116}
VID_TRANSPORT = {700005}

KIND_RANK = {3: 0, 1: 1, 2: 2, 4: 2, 5: 3}
KIND_LABEL = {1: "Смета (ЛС)", 2: "Узел", 3: "Документ", 4: "Раздел", 5: "Подраздел"}

# --------------------------------------------------------------------------
# Вспомогательные функции
# --------------------------------------------------------------------------

def cell(ws, row: int, col: int) -> Any:
    return ws.cell(row=row, column=col).value


def fnum(v: Any) -> float:
    """Число или 0.0."""
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        f = float(v)
        return 0.0 if f != f else f  # NaN -> 0
    if isinstance(v, str):
        s = v.strip().replace("\xa0", "").replace(" ", "").replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0
    return 0.0


def sstr(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


_CODE_NOTE_RE = re.compile(r"[.,;:]?\s*[кkКK]\s*[=:]")
_CODE_MULT_RE = re.compile(r"\s*[×x*]\s*\d+(?:[.,]\d+)?\s*$", re.I)


def norm_code(v: Any) -> str:
    """Шифр сметной расценки без служебных пометок коэффициентов.

    АтомСмета дописывает к расценке коэффициент: «13-03-004-26. к=2» → «13-03-004-26»,
    «м08-01-067-08.к=1,5» → «м08-01-067-08», «13-03-004-26 ×2» → «13-03-004-26».
    """
    t = sstr(v).replace("\u00a0", " ").strip()
    if not t:
        return ""
    m = _CODE_NOTE_RE.search(t)
    if m:
        t = t[:m.start()]
    t = _CODE_MULT_RE.sub("", t)
    return re.sub(r"[\s.,;:()\[\]]+$", "", t).strip()


# ---- Укрупнённые группы строк сметы: работа + её материалы ----
# Стоимость одной «работы» складывается из нескольких строк Source: самой расценки
# (ФЕР/ТЕР/ГЭСН, столбец F) и следующих за ней строк материалов, оборудования и
# перевозки, у которых своего шифра норматива нет — в печатной форме они напечатаны
# под расценкой без номера «№ п/п» (либо «N,1», «N,2» у перевозки «исключен»).
# Материалы идут в Source сразу за своей работой, поэтому группа определяется ПОРЯДКОМ
# СЛЕДОВАНИЯ строк: строка с шифром сметной нормы открывает новую группу, следующие
# за ней строки без норматива относятся к ней же. То же правило в браузерном модуле
# (SourceParser.groupItems).
#
# Шифр сметной нормы: «06-01-001-01», «м08-01-062-02», «т03-21-01-030», «м39-02-012-01».
# Прайс-листы («Расчёт 1231Р»), коды материалов ГЭСН-сборников («14.4.01.20-0012»,
# «01.7.06.03-0003») и коды машин («91.14.01-011») нормой НЕ являются: у работ первая
# часть — номер сборника с ведущими нулями («06»), у материалов — номер группы без
# ведущего нуля («14», «01», «91»).
_NORM_RE = re.compile(r"^(?:(?:ФЕР|ТЕР|ГЭСН|ФСЭМ|ФСБЦ|ТСН)\s*)?(?:м|т|п)?\d{2}-\d{2}-\d{3}-\d{2}(?:-\d)?$",
                      re.IGNORECASE)


def is_norm_code(code: Any) -> bool:
    """Шифр похож на сметную норму (ФЕР/ТЕР/ГЭСН), а не на код материала или машины."""
    t = re.sub(r"\s+", "", sstr(code))
    return bool(t) and bool(_NORM_RE.match(t))


def row_kind_of(it: Dict[str, Any]) -> str:
    """Вид строки: ``work`` — расценка с шифром нормы, ``resource`` — материал/перевозка.

    Признак берётся ТОЛЬКО по шифру позиции (столбец F вкладки Source): основание
    для этого не годится — разбор подставляет туда «ФЕР <шифр>», когда основание
    в книге пустое, а у настоящих расценок основание начинается с «ФЕР-2001» / «ФСЭМ-2001».
    """
    return "work" if is_norm_code(it.get("code")) else "resource"


def group_items(items: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Разбиение позиций одной сметы на укрупнённые группы «работа + её материалы».

    Каждой позиции проставляются:
      ``group_no``   — «№ строки сметы»: номер ведущей строки группы, то есть номер
                       позиции, под которой её материалы напечатаны в печатной форме;
      ``group_code`` — шифр сметной нормы работы (ФЕР, ТЕР, ГЭСН, ФСЭМ, ТСН…);
      ``group_name`` — наименование ведущей работы;
      ``group_lead`` — строка сама является работой (ведущей строкой группы);
      ``group_size`` — число строк в группе (работа + её материалы).

    Позиции сортируются по номеру строки вкладки Source — так их отдаёт разбор книги.
    """
    lst = sorted(items or [], key=lambda x: (x.get("row") or 0))
    has_work = any(row_kind_of(it) == "work" for it in lst)
    lead: Optional[Dict[str, Any]] = None
    cur = ""
    lead_size: Dict[str, int] = {}
    marks: List[Dict[str, Any]] = []
    for it in lst:
        kind = row_kind_of(it)
        starts = (not has_work) or kind == "work" or lead is None
        if starts:
            lead = {"no": sstr(it.get("num")), "code": sstr(it.get("code")),
                    "name": sstr(it.get("name"))}
            cur = sstr(it.get("num")) or ("#" + str(it.get("row")))
            lead_size[cur] = 0
        lead_size[cur] = lead_size.get(cur, 0) + 1
        marks.append({"it": it, "starts": starts, "cur": cur, "lead": lead})
    out: List[Dict[str, Any]] = []
    for m in marks:
        it = m["it"]
        it["group_no"] = sstr(m["lead"]["no"])
        it["group_code"] = sstr(m["lead"]["code"])
        it["group_name"] = sstr(m["lead"]["name"])
        it["group_lead"] = bool(m["starts"])
        it["group_size"] = lead_size.get(m["cur"], 1)
        out.append(it)
    return out


_EPS_JS = 2.220446049250313e-16          # Number.EPSILON в JavaScript


def js_round(x: float) -> float:
    """Math.round: округление половины вверх (у питоновского round — к чётному)."""
    return math.floor(x + 0.5)


def round2(v: float) -> float:
    """Как Math.round((v + Number.EPSILON) * 100) / 100 в браузерном модуле."""
    return js_round((v + _EPS_JS) * 100) / 100


def round4(v: float) -> float:
    """Как Math.round((v + Number.EPSILON) * 10000) / 10000 — трудозатраты (чел.-ч)."""
    return js_round((v + _EPS_JS) * 10000) / 10000


def find_source_sheet(wb):
    for name in wb.sheetnames:
        if name.strip().lower() == SOURCE_SHEET.lower():
            return wb[name]
    return None


# --------------------------------------------------------------------------
# Разбор одного файла
# --------------------------------------------------------------------------

def synth_totals(items: Iterable[Dict[str, Any]]) -> Dict[str, float]:
    """Итоги сметы по её позициям — для книг без записей структуры (A=51/52)."""
    acc = {k: 0.0 for k in ("pz", "mat_ob", "material", "equipment", "transport", "other",
                            "ozp", "emm", "zpm", "nr", "sp", "tz_workers", "tz_mash", "total")}
    # прямые затраты (стоимость без НР и СП) по видам работ — как в узлах структуры
    kind_pz = {k: 0.0 for k in ("build", "install", "equipment", "other")}
    for it in items or []:
        values = it.get("values") or {}
        for k in acc:
            acc[k] += float(values.get(k, 0.0) or 0.0)
        kind_pz[it.get("work_kind") or "other"] += float(values.get("pz", 0.0) or 0.0)
    for k in acc:
        acc[k] = round4(acc[k]) if k in ("tz_workers", "tz_mash") else round2(acc[k])
    acc["kind_pz"] = kind_pz
    return acc


def merge_headers(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Заголовок сметы из всех записей A=1.

    В книгах с несколькими ЛС (и в книгах ЛСР-02-…) первая запись A=1 — это заголовок
    документа: в ней нет ни шифра сметы, ни основания, ни объекта. Поэтому основной считаем
    первую запись с шифром, а остальные используем для заполнения пустых полей.
    """
    recs = [r for r in (records or []) if r]
    if not recs:
        return {}
    numbered = [r for r in recs if r.get("number")]
    primary = numbered[0] if numbered else recs[0]
    out = dict(primary)
    for rec in recs:
        if rec is primary:
            continue
        for key, value in rec.items():
            if value and not out.get(key):
                out[key] = value
    return out


class WorkbookParse:
    def __init__(self, path: str):
        self.path = path
        self.file = os.path.basename(path)
        self.header: Dict[str, Any] = {}
        self.header_records: List[Dict[str, Any]] = []   # все записи A=1 (документ и локальные сметы)
        self.conditions: List[Dict[str, Any]] = []
        self.nodes: List[Dict[str, Any]] = []          # в порядке следования A=51 (post-order)
        self.node_indicators: Dict[int, Dict[str, float]] = {}
        self.items: List[Dict[str, Any]] = []
        self.notes: List[str] = []
        self.doc_summary: Dict[str, float] = {}        # сводка по документу (ПЕРЕВ, МАТЕРИАЛЫ, …)
        self.has_sections = False
        self.warnings: List[str] = []
        self.n_spec_rows = 0
        self.open_row: Dict[int, int] = {}     # id узла -> строка открытия (A=52)
        self.close_row: Dict[int, int] = {}    # id узла -> строка итога (A=51)
        self.summary_rows: List[Dict[str, Any]] = []   # сводные строки «Source» (D=2)
        self.resources: List[Dict[str, Any]] = []      # строки ресурсов (A=18)
        # В одной книге АтомСмета может хранить несколько расчётов (уровней цен): каждая
        # позиция A=17/18 повторяется по разу на расчёт, идентификатор — в столбце 27.
        # Уровни описаны записями A=75 («Базовый 2000», «Текущий 2017», …) — id в столбце 14.
        self.price_variants: Dict[int, Dict[str, Any]] = {}
        self.price_levels: List[Dict[str, Any]] = []
        self.main_variant: Optional[int] = None
        self.price_selection: Dict[str, Any] = {}

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "WorkbookParse":
        """Собрать объект разбора из словаря внешнего парсера (ГРАНД-Смета)."""
        wp = cls(d.get("path") or d.get("file", ""))
        for key in ("header", "conditions", "nodes", "items", "resources", "notes",
                    "doc_summary", "summary_rows", "open_row", "close_row",
                    "node_indicators", "price_variants", "price_levels",
                    "price_selection", "main_variant"):
            if key in d and d[key] is not None:
                setattr(wp, key, d[key])
        wp.has_sections = bool(d.get("has_sections"))
        wp.n_spec_rows = int(d.get("n_spec_rows") or 0)
        wp.warnings = list(d.get("warnings") or [])
        return wp

    # ---- основной разбор -------------------------------------------------
    def parse(self) -> "WorkbookParse":
        wb = load_workbook(self.path, data_only=True, read_only=False)
        try:
            ws = find_source_sheet(wb)
            if ws is None:
                raise ValueError("во вкладке «%s» нет листа «Source»" % self.file)
            max_row = ws.max_row
            footer_row_to_node: Dict[int, int] = {}
            ind_runs: List[Tuple[int, List[int]]] = []   # (row первого A=51 перед прогоном, строки A=50)
            cur_run: List[int] = []
            last_footer_row: Optional[int] = None
            cur_open_ref: Optional[int] = None
            header_rows: List[int] = []

            for r in range(1, max_row + 1):
                t = cell(ws, r, 1)
                if t is None:
                    if cur_run:
                        ind_runs.append((last_footer_row, cur_run)); cur_run = []
                    continue
                t = int(t) if isinstance(t, float) and float(t).is_integer() else t

                # A=1 и A=75 содержат реквизиты и описание уровня цен.
                if t not in (1, 75) and sstr(cell(ws, r, 2)) != "1":
                    if t == 51:
                        if cur_run:
                            ind_runs.append((last_footer_row, cur_run))
                            cur_run = []
                        last_footer_row = None
                    elif t in (3, 4, 5):
                        cur_open_ref = None
                    continue

                if t == 50:
                    cur_run.append(r)
                    continue
                if cur_run:
                    ind_runs.append((last_footer_row, cur_run))
                    cur_run = []

                if t == 1:
                    header_rows.append(r)
                    # Заголовочных записей A=1 в книге может быть несколько: сначала идёт
                    # заголовок документа (без шифра сметы), затем — заголовок локальной сметы
                    # с номером, основанием, объектом и нормативной базой. Значимые поля
                    # собираем из всех записей, предпочитая запись с номером сметы
                    # (иначе у книг ЛСР-02-… пропадало «Основание» и «Объект»).
                    self.header_records.append({
                        "row": r,
                        "number": sstr(cell(ws, r, 6)),
                        "name": sstr(cell(ws, r, 7)),
                        "osnovanie": sstr(cell(ws, r, 10)),
                        "norm_base": sstr(cell(ws, r, 21)),
                        "object": sstr(cell(ws, r, 23)),
                        "made_by_post": sstr(cell(ws, r, 28)),
                        "made_by": sstr(cell(ws, r, 29)),
                        "checked_by_post": sstr(cell(ws, r, 30)),
                        "checked_by": sstr(cell(ws, r, 31)),
                        "tr": sstr(cell(ws, r, 78)),
                        "corrections": sstr(cell(ws, r, 83)),
                    })
                elif t in (51, 52):
                    node = {
                        "id": int(fnum(cell(ws, r, 4))),
                        "kind": int(fnum(cell(ws, r, 3))) or 0,
                        "row": r,
                        "code": sstr(cell(ws, r, 6)),
                        "name": sstr(cell(ws, r, 7)),
                        "totals": self._money(ws, r),
                    }
                    if node["id"]:
                        if t == 52:
                            self.open_row.setdefault(node["id"], r)
                        else:
                            self.close_row[node["id"]] = r
                    if t == 51:
                        last_footer_row = r
                        if node["id"]:
                            footer_row_to_node[r] = node["id"]
                        self._merge_node(node)
                elif t in (3, 4, 5):
                    # открытие узла структуры; D ссылается на строку итоговой записи A=51
                    # (итоговая запись может идти позже — привязка разрешается в конце разбора)
                    cur_open_ref = int(fnum(cell(ws, r, 4))) or None
                elif t == 17:
                    item = self._item(ws, r)
                    if item is None:
                        continue
                    item["open_ref"] = cur_open_ref
                    self.items.append(item)
                elif t == 18:
                    # Строка A=18 бывает двух видов: строка ресурса (материал, машина) и
                    # «поправка» к позиции — в печатной форме она напечатана отдельной строкой
                    # со своим номером («13,1», «34,2») и стоимостью (обычно отрицательной).
                    # Стоимость поправок входит в итоги узлов (Σ позиций A=17 + Σ поправок =
                    # показатели узла), поэтому такие строки учитываем как позиции; строки
                    # ресурсов без стоимости остаются ресурсами (книги типовых книг.), иначе сумма
                    # позиций не сходилась бы с итогом сметы.
                    if self._is_correction(ws, r):
                        corr = self._item(ws, r)
                        if corr is not None:
                            corr["open_ref"] = cur_open_ref
                            corr["correction"] = True
                            self.items.append(corr)
                            continue
                    self.n_spec_rows += 1
                    rc_code = sstr(cell(ws, r, 6))
                    rc_name = sstr(cell(ws, r, 7))
                    if rc_code or rc_name:
                        self.resources.append({
                            "row": r,
                            "parent": int(fnum(cell(ws, r, 3))) or None,
                            "code": rc_code,
                            "name": rc_name,
                            "unit": sstr(cell(ws, r, 126)) or sstr(cell(ws, r, 8)),
                            "qty": round(fnum(cell(ws, r, 9)), 6),
                            "price": round2(fnum(cell(ws, r, 15))),
                            "cost": round2(fnum(cell(ws, r, 195))),
                            "vid_code": int(fnum(cell(ws, r, 141))),
                            "vid": sstr(cell(ws, r, 142)),
                            "group": sstr(cell(ws, r, 137)),
                            "kind": int(fnum(cell(ws, r, 140))),
                        })
                elif t == 75:
                    # описание расчёта (уровня цен): «Базовый 2000», «Текущий 2017» и т. п.
                    self.price_levels.append({
                        "row": r,
                        "name": sstr(cell(ws, r, 2)),
                        "year": sstr(cell(ws, r, 3)),
                        "variant": int(fnum(cell(ws, r, 14))) or 0,
                    })
                elif t == 19:
                    txt = sstr(cell(ws, r, 7))
                    if txt:
                        self.notes.append(txt)
                elif t == 70:
                    self.conditions.append({
                        "key": sstr(cell(ws, r, 5)),
                        "label": sstr(cell(ws, r, 6)) or sstr(cell(ws, r, 16)),
                        "title": sstr(cell(ws, r, 16)),
                        "on": bool(fnum(cell(ws, r, 7))),
                        "value": fnum(cell(ws, r, 7)),
                        "num": sstr(cell(ws, r, 8)),
                        "kind": fnum(cell(ws, r, 14)),
                    })
            if cur_run:
                ind_runs.append((last_footer_row, cur_run))
            self.header = merge_headers(self.header_records)

            # привязка позиций к узлам структуры (раздел/подраздел) по ссылке D
            kind_by_id = {n["id"]: n["kind"] for n in self.nodes}
            ls_id = next((n["id"] for n in self.nodes if n["kind"] == 1),
                         next((n["id"] for n in self.nodes if n["kind"] == 3), None))
            has_sections = any(n["kind"] in (4, 5) for n in self.nodes)
            self.has_sections = has_sections
            for it in self.items:
                open_node = footer_row_to_node.get(it.pop("open_ref", None) or -1)
                it["open_node"] = open_node
                node_id = open_node
                kind = kind_by_id.get(node_id)
                if node_id is None or kind not in (1, 2, 3, 4, 5):
                    node_id = ls_id
                it["node"] = node_id

            # Если в книге несколько расчётов (уровней цен), оставляем основной: «Базовый»
            # (таблица ТЗ — базисный уровень 2000 г.) либо, если его нет, тот, суммы позиций
            # которого совпадают с итогами узлов. Остальные расчёты в разбор не попадают.
            self._select_price_variant()

            # показатели узлов и сводка по документу
            for ref_row, rows in ind_runs:
                node_id = footer_row_to_node.get(ref_row) if ref_row else None
                values: Dict[str, float] = {}
                for r in rows:
                    if int(fnum(cell(ws, r, 4))) != 1:   # D=2 — сводка по документу
                        kk = cell(ws, r, 11)
                        gg = sstr(cell(ws, r, 7))
                        if kk is not None and int(fnum(kk)) == 212:
                            # ключ (G), отображаемое имя (H), краткое имя (N), значение (F, резерв — W)
                            fv = fnum(cell(ws, r, 6))
                            wv = fnum(cell(ws, r, 23))
                            self.summary_rows.append({
                                "row": r, "node": node_id, "key": gg,
                                "display": sstr(cell(ws, r, 8)), "label": sstr(cell(ws, r, 14)),
                                "value": round2(fv if fv != 0 else wv)})
                            if fnum(cell(ws, r, 13)) == 1 and gg:
                                self.doc_summary[gg] = round2(wv if wv != 0.0 else fv)
                        continue
                    k = cell(ws, r, 11)
                    if k is None:
                        continue
                    k = int(fnum(k))
                    if k not in INDICATORS:
                        continue
                    # Значение показателя: АтомСмета пишет его в столбец F, дублируя в W
                    # (в части выпусков W остаётся нулевым, например в ЛСР-02-24-01(2));
                    # нулевое значение F означает, что показатель действительно нулевой, — тогда
                    # берём дубль из W. Трудозатраты (207/208) хранятся в чел.-ч с точностью
                    # 4 знаков, остальные показатели — деньги, округляются до копеек.
                    fv = fnum(cell(ws, r, 6))
                    vv = fv if fv != 0.0 else fnum(cell(ws, r, 23))
                    values[INDICATORS[k][0]] = (round4(vv) if k in (207, 208) else round2(vv))
                if node_id is not None and values:
                    if not values.get("transport"):
                        alt = self.doc_summary.get("ПЕРЕВ", 0.0)
                        if alt:
                            values["transport_doc"] = alt
                            values["transport"] = alt
                    if not values.get("other"):
                        alt = self.doc_summary.get("ПРОЧ_КСт", 0.0)
                        if alt:
                            values["other_doc"] = alt
                            values["other"] = alt
                    self.node_indicators.setdefault(node_id, {}).update(values)
            return self
        finally:
            wb.close()

    # ---- выбор расчёта (уровня цен) --------------------------------------
    def _select_price_variant(self) -> None:
        """Оставить позиции одного расчёта, если книга содержит несколько уровней цен.

        Каждая позиция АтомСметы хранится по разу на расчёт: идентификатор расчёта — столбец 27
        записей A=17/A=18, описания уровней — записи A=75 («Базовый 2000», «Текущий 2017»).
        Без выбора строки разных уровней суммировались бы (итог завышался бы в разы).

        Правило: 1) единственный расчёт; 2) уровень «Базовый»/2000 (таблица ТЗ — базисный
        уровень цен); 3) расчёт, суммы позиций которого совпадают с итогами узлов; при равенстве
        оценок — расчёт с большим числом позиций, затем с меньшим идентификатором.
        """
        by_variant: Dict[int, Dict[str, Any]] = {}
        for it in self.items:
            v = int(it.get("price_variant") or 0)
            if v <= 0:
                continue
            d = by_variant.setdefault(v, {"id": v, "items": 0, "total": 0.0, "by_node": {},
                                          "levels": []})
            d["items"] += 1
            d["total"] = round2(d["total"] + (it["values"].get("total") or 0.0))
            nid = it.get("node")
            d["by_node"][nid] = round2(d["by_node"].get(nid, 0.0) + (it["values"].get("total") or 0.0))
        self.price_variants = by_variant
        for lv in self.price_levels:
            v = int(lv.get("variant") or 0)
            if v in by_variant:
                by_variant[v]["levels"].append("%s %s" % (lv.get("name") or "", lv.get("year") or ""))
        if not by_variant:
            return
        if len(by_variant) == 1:
            self.main_variant = next(iter(by_variant))
            self.price_selection = {"chosen": self.main_variant, "reason": "single",
                                    "variants": self._variant_info()}
            return

        node_total = {n["id"]: (n["totals"].get("total") or 0.0) for n in self.nodes if n.get("id")}
        for d in by_variant.values():
            # сравниваем с итогами только тех узлов, к которым позиции этого расчёта привязаны
            # напрямую: у сводных узлов (ЛС, документ) своих позиций нет
            score = 0.0
            for nid, got in d["by_node"].items():
                want = node_total.get(nid) or 0.0
                if not want:
                    continue
                score += abs(got - want)
            d["score"] = round2(score)

        base = [v for v, d in by_variant.items()
                if any("базов" in (x or "").lower() or "2000" in (x or "") for x in d["levels"])]
        reason = "base-level"
        if len(base) == 1:
            chosen = base[0]
        elif base:
            chosen = min(base, key=lambda v: (by_variant[v]["score"], -by_variant[v]["items"], v))
        else:
            reason = "node-totals"
            best_score = min(d["score"] for d in by_variant.values())
            cand = [v for v, d in by_variant.items() if abs(d["score"] - best_score) < 1e-9]
            if len(cand) > 1:
                reason = "node-totals+items"
            chosen = min(cand, key=lambda v: (-by_variant[v]["items"], v))
        self.main_variant = chosen
        self.price_selection = {"chosen": chosen, "reason": reason, "variants": self._variant_info()}
        kept = [it for it in self.items if int(it.get("price_variant") or 0) in (0, chosen)]
        dropped = len(self.items) - len(kept)
        if dropped:
            self.warnings.append(
                "книга содержит %d расчёта (уровня цен): оставлен расчёт %s%s; пропущено %d позиций"
                % (len(by_variant), chosen,
                   (" («%s»)" % ", ".join(by_variant[chosen]["levels"])) if by_variant[chosen]["levels"] else "",
                   dropped))
        self.items = kept

    def _variant_info(self) -> List[Dict[str, Any]]:
        """Сводка по расчётам книги — для отчёта и контроля (какой выбран, какой пропущен)."""
        out = []
        for v, d in sorted(self.price_variants.items()):
            out.append({"id": v, "items": d["items"], "total": round2(d["total"]),
                        "levels": list(d["levels"]), "score": d.get("score"),
                        "chosen": v == self.main_variant})
        return out

    # ---- служебное -------------------------------------------------------
    def _money(self, ws, r: int) -> Dict[str, float]:
        """Числовая раскладка записи A=17/A=51/A=52 (без показателей узла).

        Столбец GM («Всего с НР и СП») в части выпусков АтомСметы пуст (например в ЛСР-02-24-01(2)),
        тогда итог считаем по определению: ПЗ + НР + СП.
        """
        pz = round2(fnum(cell(ws, r, 15)))
        nr = round2(fnum(cell(ws, r, 24)))
        sp = round2(fnum(cell(ws, r, 25)))
        total = round2(fnum(cell(ws, r, 195)))
        if not total:
            total = round2(pz + nr + sp)
        return {
            "pz": pz,
            "mat_ob": round2(fnum(cell(ws, r, 16))),
            "emm": round2(fnum(cell(ws, r, 17))),
            "zpm": round2(fnum(cell(ws, r, 18))),
            "ozp": round2(fnum(cell(ws, r, 19))),
            "tz_workers": round4(fnum(cell(ws, r, 21))),
            "tz_mash": round4(fnum(cell(ws, r, 22))),
            "nr": nr,
            "sp": sp,
            "total": total,
        }

    def _merge_node(self, node: Dict[str, Any]) -> None:
        for n in self.nodes:
            if n["id"] == node["id"]:
                n.update({k: v for k, v in node.items() if v not in ("", None)})
                if node["totals"]:
                    n["totals"] = node["totals"]
                return
        self.nodes.append(node)

    # ---- позиция ---------------------------------------------------------
    def _is_correction(self, ws, r: int) -> bool:
        """True для строки A=18 — «поправки» к позиции (печатается отдельной строкой формы).

        Признаки: есть номер строки в столбце E («13,1») и ненулевая стоимость (итог, ПЗ,
        НР или СП). Строки ресурсов без стоимости (книги типовых книг., позиции вида «1,1»
        с нулями) к позициям не относятся.
        """
        if not sstr(cell(ws, r, 5)):
            return False
        money = self._money(ws, r)
        return any(abs(money[k]) > EPS for k in ("total", "pz", "nr", "sp"))

    def _item(self, ws, r: int) -> Optional[Dict[str, Any]]:
        num = sstr(cell(ws, r, 5))
        if not num:
            return None                      # внутренние/ресурсные строки без номера позиции
        values = self._money(ws, r)
        vid_code = int(fnum(cell(ws, r, 141)))
        vid_name = sstr(cell(ws, r, 142))
        group = sstr(cell(ws, r, 137))
        kind_flag = int(fnum(cell(ws, r, 140)))
        cat = classify(vid_code, vid_name, group, kind_flag)
        pz = values["pz"]
        mat_ob = values["mat_ob"]
        values["material"] = round2(mat_ob) if cat in ("material", "works") else 0.0
        values["equipment"] = round2(mat_ob) if cat == "equipment" else 0.0
        values["transport"] = round2(mat_ob) if cat == "transport" else 0.0
        values["other"] = round2(values["total"]) if cat == "other" else 0.0
        if not values["total"]:
            values["total"] = round2(pz + values["nr"] + values["sp"])
        shifr = norm_code(cell(ws, r, 6))
        basis = sstr(cell(ws, r, 62))
        if not basis:
            basis = ("ФЕР " + shifr) if shifr else ""
        unit = sstr(cell(ws, r, 126)) or sstr(cell(ws, r, 8))
        qty = fnum(cell(ws, r, 9))
        nr_pct = fnum(cell(ws, r, 46)) or fnum(cell(ws, r, 78)) or fnum(cell(ws, r, 180))
        sp_pct = fnum(cell(ws, r, 47)) or fnum(cell(ws, r, 79)) or fnum(cell(ws, r, 181))
        item = {
            "num": num,
            "row": r,
            "code": shifr,
            "name": sstr(cell(ws, r, 7)),
            "basis": basis,
            "unit": unit,
            "qty": round(qty, 6),
            "cat": cat,
            "vid": vid_name,
            "vid_code": vid_code,
            "group": group,
            "nr_pct": round(nr_pct, 2),
            "sp_pct": round(sp_pct, 2),
            "work_kind": work_kind_of(cat, vid_name, group),
            "values": values,
        }
        variant = int(fnum(cell(ws, r, 27))) or 0
        item["price_variant"] = variant if variant > 0 else 0
        # исходный шифр сохраняется только если он отличался от нормализованного
        raw_code = sstr(cell(ws, r, 6))
        if raw_code != shifr:
            item["code_raw"] = raw_code
        return item

    # ---- привязка узлов --------------------------------------------------
    def build_tree(self) -> None:
        """Иерархия узлов.

        Итоговые записи A=51 идут в порядке обхода «снизу вверх» (post-order),
        поэтому родителем узла считается ближайший следующий узел с меньшим
        рангом вида (3 — документ, 1 — смета, 4 — раздел, 5 — подраздел),
        а смета (вид 1) всегда вкладывается в документ (вид 3).
        """
        seq = self.nodes
        roots0 = [n["id"] for n in seq if KIND_RANK.get(n["kind"], 9) == 0]
        for i, n in enumerate(seq):
            rank = KIND_RANK.get(n["kind"], 9)
            parent: Optional[int] = None
            if rank == 1 and roots0:
                parent = roots0[0]
            else:
                best_rank = -1
                for m in seq[i + 1:]:
                    mr = KIND_RANK.get(m["kind"], 9)
                    if mr < rank and mr > best_rank:
                        best_rank = mr
                        parent = m["id"]
            n["parent"] = parent
            n["rank"] = rank
        by_id = {n["id"]: n for n in self.nodes}
        for n in self.nodes:
            depth = 0
            p = n.get("parent")
            seen = set()
            while p is not None and p in by_id and p not in seen:
                seen.add(p)
                depth += 1
                p = by_id[p].get("parent")
            n["depth"] = depth
        self.tree_sequence = [n["id"] for n in seq]


# Вид работ позиции — как АтомСмета делит стоимость: строительные работы (показатель 214),
# монтажные работы (215), оборудование (216) и прочие (217). Признак берётся из вида затрат
# и группы работ позиции (столбцы EL и EG вкладки «Source»). Правило проверено на контрольном
# наборе: сумма позиций по видам в точности совпадает с показателями узлов (214/215/216/217).
BUILD_MARKS = ("строительн", "общестроительн", "перевозк", "погрузочно-разгрузочн")
INSTALL_MARKS = ("монтаж", "санитарно-техническ")
WORK_KIND_LABELS = {
    "build": "Строительные работы",
    "install": "Монтажные работы",
    "equipment": "Оборудование",
    "other": "Прочие работы",
}


def work_kind_of(cat: str, vid_name: str, group: str) -> str:
    """Вид работ позиции: build (строительные), install (монтажные), equipment, other.

    Порядок проверок важен: монтажные признаки (группа «Монтажные работы», вид затрат
    «…монтаж … оборудования», «Санитарно-технические работы») смотрятся раньше категории
    позиции — АтомСмета относит такие позиции к монтажным работам, хотя их вид затрат
    помечен как оборудование (например «Связь: монтаж теле-радио оборудования», код 110004).
    """
    text = ("%s %s" % (group or "", vid_name or "")).lower()
    if any(m in text for m in INSTALL_MARKS):
        return "install"
    if cat == "equipment":
        return "equipment"
    if any(m in text for m in BUILD_MARKS):
        return "build"
    return "other"


def classify(vid_code: int, vid_name: str, group: str, kind_flag: int) -> str:
    """Категория позиции: works | material | equipment | transport | other."""
    blob = ("%s %s" % (vid_name, group)).lower()
    if kind_flag == 3 or vid_code in VID_EQUIPMENT or "оборудован" in blob:
        return "equipment"
    if vid_code in VID_TRANSPORT or "перевозк" in blob:
        return "transport"
    if "прочие" in blob or "проч" in blob.split()[0:1]:
        return "other"
    if vid_code in VID_MATERIAL or "материал" in blob:
        return "material"
    return "works"


# --------------------------------------------------------------------------
# Сверки (контроль корректности исходных данных)
# --------------------------------------------------------------------------

def add_check(checks: List[Dict[str, Any]], scope: str, ref: str, severity: str,
              code: str, message: str, delta: Optional[float] = None,
              ls: Any = None) -> None:
    item = {"scope": scope, "ref": ref, "severity": severity, "code": code, "message": message}
    if delta is not None:
        item["delta"] = round2(delta)
    if ls is not None:
        item["ls"] = ls
    checks.append(item)


def check_node(checks: List[Dict[str, Any]], ref: str, t: Dict[str, float],
               ls: Any = None) -> None:
    if not t:
        add_check(checks, "node", ref, "error", "NO_TOTALS",
                  "У узла отсутствуют числовые показатели", ls=ls)
        return
    d = t.get("pz", 0.0) - (t.get("mat_ob", 0.0) + t.get("ozp", 0.0) + t.get("emm", 0.0))
    if abs(d) > EPS:
        add_check(checks, "node", ref, "warn", "PZ_STRUCT",
                  "ПЗ ≠ МатОб + ОЗП + ЭММ", d, ls=ls)
    d = t.get("total", 0.0) - (t.get("pz", 0.0) + t.get("nr", 0.0) + t.get("sp", 0.0))
    if abs(d) > EPS:
        add_check(checks, "node", ref, "warn", "TOTAL_STRUCT",
                  "Всего с НР и СП ≠ ПЗ + НР + СП", d, ls=ls)
    if t.get("material", 0.0) + t.get("equipment", 0.0) - t.get("mat_ob", 0.0) > EPS:
        add_check(checks, "node", ref, "warn", "MAT_OB_STRUCT",
                  "Материалы + оборудование > стоимости материальных ресурсов",
                  t.get("material", 0.0) + t.get("equipment", 0.0) - t.get("mat_ob", 0.0), ls=ls)
    for key in ("total", "pz", "ozp", "emm", "material", "equipment", "nr", "sp", "tz_workers", "tz_mash"):
        if t.get(key, 0.0) < -EPS:
            add_check(checks, "node", ref, "warn", "NEGATIVE",
                      "Отрицательное значение показателя «%s»" % key, t.get(key), ls=ls)
    smr = t.get("total", 0.0) - t.get("equipment", 0.0)
    if smr > EPS and abs(t.get("nr", 0.0)) < EPS and abs(t.get("sp", 0.0)) < EPS:
        add_check(checks, "node", ref, "warn", "NO_NR_SP",
                  "Есть работы, но НР и СП равны нулю", ls=ls)


def check_items(checks: List[Dict[str, Any]], ref: str, items: List[Dict[str, Any]],
                t: Dict[str, float]) -> None:
    """Сверка суммы позиций с показателями узла."""
    acc = {k: 0.0 for k in ("pz", "mat_ob", "emm", "zpm", "ozp", "nr", "sp",
                            "tz_workers", "tz_mash")}
    for it in items:
        for k in acc:
            acc[k] += it["values"].get(k, 0.0)
    for k, label in (("pz", "Прямые затраты"), ("mat_ob", "Материальные ресурсы"),
                     ("emm", "Эксплуатация машин"), ("zpm", "ЗП машинистов"),
                     ("ozp", "Основная ЗП рабочих"), ("nr", "Накладные расходы"),
                     ("sp", "Сметная прибыль"), ("tz_workers", "Трудозатраты строителей"),
                     ("tz_mash", "Трудозатраты машинистов")):
        base = t.get(k)
        if base is None:
            continue
        d = acc[k] - base
        if abs(d) > max(EPS, abs(base) * 1e-9):
            add_check(checks, "smeta", ref, "error", "SUM_MISMATCH",
                      "Сумма позиций ≠ итог сметы по показателю «%s»" % label, d)
    neg = [it for it in items if it["values"].get("total", 0.0) < -EPS]
    for it in items:
        iref = "%s/№%s" % (ref, it["num"])
        if not it["name"]:
            add_check(checks, "item", iref, "warn", "NO_NAME", "Не заполнено наименование позиции")
        if it["qty"] < 0:
            add_check(checks, "item", iref, "info", "NEG_QTY",
                      "Отрицательное количество (возврат/корректировка)", it["qty"])
        if abs(it["qty"]) < 1e-12 and abs(it["values"].get("total", 0.0)) > EPS:
            add_check(checks, "item", iref, "warn", "QTY_ZERO",
                      "Нулевое количество при ненулевой стоимости", it["values"].get("total"))
    if neg:
        add_check(checks, "smeta", ref, "info", "NEG_ITEMS",
                  "Позиций с отрицательной стоимостью: %d" % len(neg))


def check_children(checks: List[Dict[str, Any]], ref: str, node: Dict[str, Any],
                   children: List[Dict[str, Any]]) -> None:
    """Сверка итога узла с суммой дочерних узлов."""
    if not children:
        return
    acc = {k: 0.0 for k in ("pz", "mat_ob", "emm", "ozp", "nr", "sp", "total")}
    for ch in children:
        for k in acc:
            acc[k] += ch["totals"].get(k, 0.0)
    for k, label in (("pz", "Прямые затраты"), ("total", "Всего с НР и СП"),
                     ("nr", "Накладные расходы"), ("sp", "Сметная прибыль")):
        base = node["totals"].get(k)
        if base is None:
            continue
        d = acc[k] - base
        if abs(d) > max(EPS, abs(base) * 1e-9):
            add_check(checks, "node", ref, "warn", "CHILD_SUM",
                      "Итог узла ≠ сумме вложенных узлов по показателю «%s»" % label, d)




# --------------------------------------------------------------------------
# Печатная форма («…11 граф…») — эталон для проверки корректности данных
# --------------------------------------------------------------------------

def read_printed(path: str) -> List[Dict[str, Any]]:
    """Прочитать листы печатной формы: с «граф» в названии, иначе — распознанные по содержанию."""
    out: List[Dict[str, Any]] = []
    wb = load_workbook(path, data_only=True)      # обычный режим: нужен произвольный доступ
    try:
        def rows_of(name):
            try:
                return [[c.value for c in row] for row in wb[name].iter_rows()]
            except Exception:          # pragma: no cover
                return None
        for name in pf.sheets_with_mark(wb, rows_of):
            positions = pp.read(wb[name])
            try:
                ws = wb[name]
                parsed = pf.parse(ws)
                parsed["position_rows"] = positions
                out.append(parsed)
            except Exception as exc:      # pragma: no cover
                out.append({"sheet": name, "error": str(exc), "totals": {}, "position_rows": positions})
    finally:
        wb.close()
    return out


def _norm(text: str) -> str:
    return "".join(ch for ch in str(text or "").lower() if ch.isalnum())


def pick_source_totals(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Строки сводки вкладки «Source»: итог с коэффициентом и базисный итог.

    «ВСЕГО ПО СМЕТЕ» — итог с коэффициентом условий производства работ,
    «Всего по Смете» — тот же итог в базисном уровне цен (без коэффициента).
    """
    with_k = None
    base = None
    itog = None
    for x in rows:
        lab, dsp = _norm(x.get("label")), _norm(x.get("display"))
        if dsp == "всегопосмете":
            if with_k is None or x["row"] > with_k["row"]:
                with_k = x
        if lab == "всегопосмете" or (dsp.startswith("всегопосмете") and len(dsp) > 12):
            if base is None or x["row"] > base["row"]:
                base = x
        # В книгах с итоговым блоком «17 граф» строки «ВСЕГО ПО СМЕТЕ» нет: базисный итог
        # напечатан как «ИТОГО по смете, в т.ч.» (ключ сводки «ИтогоПоСмете»).
        if dsp.startswith("итогопосмете") or lab.startswith("итогопосмете") \
                or _norm(x.get("key")) == "итогопосмете":
            if itog is None or x["row"] > itog["row"]:
                itog = x
    if base is None:
        base = itog
    return {"with_k": with_k, "base": base}


def match_printed(printed_list: List[Dict[str, Any]], ls_nodes: List[Dict[str, Any]],
                  ls: Dict[str, Any], idx: int, hdr: Dict[str, Any],
                  used: set) -> Optional[Dict[str, Any]]:
    """Сопоставить лист печатной формы с конкретной локальной сметой."""
    if not printed_list:
        return None
    cand = [p for p in printed_list if id(p) not in used and p.get("totals")]
    if not cand:
        return None
    num = _norm(ls.get("code") or (hdr.get("number") if len(ls_nodes) == 1 else ""))
    doc = _norm(hdr.get("number") and "" or "")
    for p in cand:
        if num and _norm(p.get("number")) == num:
            used.add(id(p))
            return p
    code = _norm(ls.get("code"))
    for p in cand:
        if code and code in _norm(p.get("number")):
            used.add(id(p))
            return p
    for p in cand:
        if _norm(p.get("name")) and _norm(hdr.get("name")) and _norm(p.get("name")) == _norm(hdr.get("name")):
            used.add(id(p))
            return p
    # по порядку следования
    for p in cand:
        used.add(id(p))
        return p
    return None


def _find_printed_section(prn: Dict[str, Any], node: Dict[str, Any], parent_name: str = "",
                          used: Optional[set] = None) -> Optional[Dict[str, Any]]:
    """Сопоставить узел (раздел/подраздел) со строкой «Итого по разделу (подразделу)» формы.

    Имена сопоставляются по вхождению, поэтому узел ищем по совокупности признаков:
    точное совпадение имени, тот же вид (раздел/подраздел), тот же раздел-родитель и самое
    длинное имя. Каждая строка формы используется только один раз: иначе раздел «Напорные
    подающие трубопроводы…» совпал бы с подразделом «Трубопроводы», а в книгах с двумя
    одноимёнными подразделами оба узла получили бы итог первого.
    """
    want = _norm(node.get("name"))
    if not want:
        return None
    used = used if used is not None else set()
    kind_want = "раздел" if node.get("kind") == 4 else "подраздел"
    parent_want = _norm(parent_name)
    cands = []
    for i, sec in enumerate(prn.get("sections") or []):
        if i in used:
            continue
        got = _norm(sec.get("name"))
        if not got:
            continue
        exact = (got == want)
        if not (exact or got in want or want in got):
            continue
        same_kind = (sec.get("kind") == kind_want)
        # вид должен совпадать: иначе раздел «Напорные подающие трубопроводы…» забрал бы
        # итог подраздела «Трубопроводы» (когда в форме нет строки самого раздела).
        # Исключение — точное совпадение имён.
        if not (same_kind or exact):
            continue
        same_parent = (not parent_want or sec.get("kind") != "подраздел"
                       or _norm(sec.get("parent")) == parent_want)
        cands.append((bool(exact), bool(same_kind), bool(same_parent), len(got), -i, sec))
    if not cands:
        return None
    cands.sort(key=lambda x: x[:5], reverse=True)
    best = cands[0]
    used.add(-best[4])
    return best[5]


def _printed_diagnostics(prn: Dict[str, Any], cross: List[Dict[str, Any]],
                         source_totals: Dict[str, float]) -> Dict[str, Any]:
    """Сохранить печатную форму для сверки с отобранными данными Source."""
    store: Dict[str, Any] = {}
    store["printed"] = {
        "sheet": prn.get("sheet"), "number": prn.get("number"), "name": prn.get("name"),
        "doc_code": prn.get("doc_code"), "osnovanie": prn.get("osnovanie"),
        "object": prn.get("object"), "norm_base": prn.get("norm_base"),
        "head": prn.get("head") or {}, "totals": prn.get("totals") or {},
        "totals_with_k": prn.get("totals_with_k") or {}, "rows": prn.get("rows") or {},
        "n_positions": prn.get("n_positions"), "sum_position_total": prn.get("sum_position_total"),
        "sections": prn.get("sections") or [],
        "self_check": pf.check_printed_self(prn),
    }
    store["cross_check"] = cross
    store["totals_source"] = dict(source_totals)
    return store



# --------------------------------------------------------------------------
# Сборка JSON
# --------------------------------------------------------------------------

def build_dataset(paths: Iterable[str], include_items: bool = True,
                  reference: Optional[str] = None) -> Dict[str, Any]:
    """Собрать данные по списку книг.

    reference — путь к книге справочника («Справочник новый.xlsx»). Если не задан,
    справочник ищется среди переданных книг автоматически (книга без вкладки «Source»,
    но с вкладками «Расценки ФО» / «Множитель» / «Расценки ТЕХНИКА и ТРЗ»).
    """
    smetas: List[Dict[str, Any]] = []
    checks: List[Dict[str, Any]] = []
    files_meta: List[Dict[str, Any]] = []
    warnings: List[str] = []
    ref_book = None
    ref_source = ""
    if reference and ref_mod is not None:
        try:
            ref_book = ref_mod.load(reference)
            ref_source = os.path.basename(reference)
        except Exception as exc:
            warnings.append("Справочник «%s» не прочитан: %s" % (os.path.basename(reference), exc))

    for path in paths:
        if ref_book is None and ref_mod is not None:
            try:
                cand = ref_mod.load(path)
            except Exception:
                cand = None
            if cand is not None:
                ref_book = cand
                ref_source = os.path.basename(path)
                files_meta.append({"file": os.path.basename(path), "path": path, "smetas": 0,
                                   "nodes": 0, "items": 0, "spec_rows": 0, "reference": True})
                continue
        grand = gp is not None and gp.looks_like_grand(path)
        try:
            if grand:
                # ГРАНД-Смета: печатная форма вместо вкладки «Source»; разбор даёт
                # ту же внутреннюю модель (nodes/items/resources/summary_rows).
                gd = gp.parse_book(path)
                gd["items"] = [it for it in gd["items"] if it.get("price_variant", 0) == 0]
                wp = WorkbookParse.from_dict(gd)
                wp.grand = gd.get("grand", {})
                wp.build_tree()
                wp.printed = []
            else:
                wp = WorkbookParse(path).parse()
                wp.build_tree()
                wp.printed = read_printed(path)
        except Exception as exc:
            warnings.append("Файл «%s» не разобран: %s" % (os.path.basename(path), exc))
            traceback.print_exc()
            continue

        hdr = wp.header
        printed_list = getattr(wp, "printed", [])
        by_id = {n["id"]: n for n in wp.nodes}
        ls_all = ([n["id"] for n in wp.nodes if n["kind"] == 1]
                  or [n["id"] for n in wp.nodes if n["kind"] == 3])
        ls_ids = set(ls_all)
        ls_id = ls_all[0] if ls_all else None
        if grand:
            # ГРАНД-Смета: узлы уже образуют дерево «смета → разделы», позиции и ресурсы
            # отнесены к разделам на этапе разбора печатной формы.
            for it in wp.items:
                it["ls"] = ls_all[0] if ls_all else None
            for rc in getattr(wp, "resources", []):
                rc["ls"] = ls_all[0] if ls_all else None
            nodes_out = [{
                "id": n["id"],
                "kind": n["kind"],
                "kind_label": KIND_LABEL.get(n["kind"], "Узел %s" % n["kind"]),
                "parent": n.get("parent"),
                "depth": n.get("depth", 0),
                "row": n["row"],
                "number": n["code"],
                "name": n["name"],
                "totals": dict(n["totals"]),
            } for n in sorted(wp.nodes, key=lambda x: (x.get("depth", 0), x["row"]))]
            items_by_node: Dict[int, List[Dict[str, Any]]] = {}
            for it in wp.items:
                items_by_node.setdefault(it["node"], []).append(it)
            file_ref = wp.file
            ls_nodes = [n for n in nodes_out if n["id"] in ls_ids]
            multi = False
            printed_sheets: List[Any] = []
            printed_info: Dict[Any, Dict[str, Any]] = {}

            def raw_number(node: Dict[str, Any]) -> str:
                return node["number"] or hdr.get("number") or ""

            def source_totals_for(ls: Dict[str, Any]) -> Dict[str, Any]:
                return pick_source_totals(list(getattr(wp, "summary_rows", [])))

            def unit_scale(row: Optional[Dict[str, Any]], reference: float) -> Optional[Dict[str, Any]]:
                return dict(row, unit_source="руб", scale=1.0) if row else None

            total_info: Dict[Any, Dict[str, Any]] = {}
            self_check = gd.get("self_check") or {}
            for ls in ls_nodes:
                tot = dict(ls["totals"])
                src = source_totals_for(ls)
                src_base = unit_scale(src.get("base"), float(tot.get("total", 0.0) or 0.0))
                src_k = unit_scale(src.get("with_k"), float(tot.get("total", 0.0) or 0.0))
                k_src = float(self_check.get("work_coefficient") or 1.0)
                selected = synth_totals([it for it in wp.items if it.get("ls") == ls["id"]])
                calc_k = round2(selected.get("total", 0.0))
                # печатная форма ГРАНД-Сметы печатает сводку итогов в двух уровнях
                # цен; позиции же напечатаны с индивидуальными индексами пересчёта в
                # текущий уровень, поэтому Σпозиций заведомо больше базисной строки
                # «Всего». Сверка ведётся по печатным блокам (footer/итогам разделов),
                # а расхождение уровней фиксируется справочно — не как ошибка парсинга.
                lvl_ratio = float(self_check.get("level_ratio") or 1.0)
                tol_k = 0.05 + abs(float(src_base["value"]) if src_base else 0.0) * abs(lvl_ratio - 1.0)
                tc: List[Dict[str, Any]] = []
                if src_base:
                    dl = round2(src_base["value"] - selected.get("total", 0.0))
                    st = "ok" if abs(dl) <= max(tol_k, abs(src_base["value"]) * 1e-9) else "mismatch"
                    tc.append({"code": "SOURCE_TOTAL_BASE",
                               "label": "Всего по смете (базисный уровень, без коэффициента)",
                               "row": src_base["row"], "unit": "руб", "source": round2(src_base["value"]),
                               "unit_source": "руб", "scale": 1.0,
                               "calc": round2(selected.get("total", 0.0)), "printed": None,
                               "delta": dl, "status": st})
                    if st == "mismatch":
                        add_check(checks, "smeta",
                                  "%s | %s | печатная форма ГРАНД («Всего»)" % (file_ref, hdr.get("number", "")),
                                  "warn", "GRAND_TOTAL_BASE_MISMATCH",
                                  "Сумма позиций (базисный уровень) ≠ строке «Всего» блока "
                                  "«Итоги по смете:»", dl, ls=ls["id"])
                if src_k:
                    # печатная форма ГРАНД-Сметы печатает «ВСЕГО по смете» в текущем
                    # уровне = базисный итог × индекс пересчёта; со сверяем Σпозиций
                    # с базисом, а текущий итог — с произведением base × ratio
                    base_val = float(src_base["value"]) if src_base else calc_k
                    ratio_cur = (float(src_k["value"]) / base_val) if base_val else 1.0
                    exp_cur = round2(base_val * ratio_cur)
                    d_calc = round2(float(src_k["value"]) - exp_cur)
                    ok_all = abs(d_calc) <= max(tol_k, abs(src_k["value"]) * 1e-9)
                    tc.append({"code": "SOURCE_TOTAL",
                               "label": "ВСЕГО ПО СМЕТЕ (итог печатной формы ГРАНД-Сметы)",
                               "row": src_k["row"], "unit": "руб", "source": round2(src_k["value"]),
                               "unit_source": "руб", "scale": 1.0, "calc": exp_cur,
                               "printed": round2(src_k["value"]), "delta": d_calc,
                               "delta_printed": 0.0, "k": k_src,
                               "k_from": "индекс пересчёта в текущий уровень (%.4f)" % ratio_cur,
                               "status": ("ok" if ok_all else "mismatch")})
                    if not ok_all:
                        add_check(checks, "smeta",
                                  "%s | %s | печатная форма ГРАНД («ВСЕГО по смете»)"
                                  % (file_ref, hdr.get("number", "")), "warn",
                                  "GRAND_TOTAL_MISMATCH",
                                  "Базисный итог × индекс пересчёта ≠ строке «ВСЕГО по смете»",
                                  d_calc, ls=ls["id"])
                report_base = round2(float(self_check.get("total_base")
                                          or (src_base["value"] if src_base else tot.get("total", 0.0))))
                report_total = round2(float(self_check.get("total_current")
                                            or (src_k["value"] if src_k else report_base)))
                total_info[ls["id"]] = {
                    "total_check": tc, "source_total_calc": calc_k, "source_work_coefficient": k_src,
                    "total": report_total, "total_base": report_base,
                    "source_total": ({"row": src_k["row"], "name": src_k["display"], "key": src_k["key"],
                                      "value": round2(src_k["value"]), "unit_source": "руб", "scale": 1.0}
                                     if src_k else None),
                    "source_total_base": ({"row": src_base["row"], "name": src_base["display"],
                                           "key": src_base["key"], "value": round2(src_base["value"]),
                                           "unit_source": "руб", "scale": 1.0}
                                          if src_base else None),
                }
            num_count: Dict[str, int] = {}
            for node in ls_nodes:
                num_count[raw_number(node)] = num_count.get(raw_number(node), 0) + 1
        else:
            for it in wp.items:                     # страховка: позиция всегда в смете
                owner = by_id.get(it["node"])
                if ls_id is not None and owner is None:
                    it["node"] = ls_id

        def ancestor_ls(node_id: Any) -> Any:
            """Ближайшая объемлющая локальная смета (kind=1) по иерархии узлов."""
            cur, seen = node_id, set()
            while cur is not None and cur in by_id and cur not in seen:
                seen.add(cur)
                if cur in ls_ids:
                    return cur
                cur = by_id[cur].get("parent")
            return None

        spans = sorted(
            ((wp.open_row.get(i, float("-inf")), wp.close_row.get(i, float("inf")), i) for i in ls_all),
            key=lambda x: x[1] - x[0])

        def ls_by_row(row: int) -> Any:
            """Самая «узкая» ЛС, диапазон строк которой содержит строку позиции."""
            for a, b, i in spans:
                if a <= row <= b:
                    return i
            return None

        def ls_owner_of(node_id: Any, row: int, open_node: Any) -> Any:
            ok = by_id.get(open_node, {}).get("kind")
            if open_node in ls_ids:
                return open_node
            if ok in (4, 5):
                return ancestor_ls(node_id) or ls_by_row(row) or ls_id
            return ls_by_row(row) or ancestor_ls(node_id) or ls_id

        for it in wp.items:
            it["ls"] = ls_owner_of(it["node"], it["row"], it.pop("open_node", None))

        # строки ресурсов относим к смете по номеру позиции-владельца, иначе — по диапазону строк
        item_ls_by_num = {str(it["num"]): it.get("ls") for it in wp.items if it.get("num")}
        for rc in getattr(wp, "resources", []):
            owner = item_ls_by_num.get(str(rc.get("parent")))
            rc["ls"] = owner if owner is not None else (ls_by_row(rc["row"]) or ls_id)

        nodes_out: List[Dict[str, Any]] = []
        for n in sorted(wp.nodes, key=lambda x: (x.get("depth", 0), x["row"])):
            nodes_out.append({
                "id": n["id"],
                "kind": n["kind"],
                "kind_label": KIND_LABEL.get(n["kind"], "Узел %s" % n["kind"]),
                "parent": n.get("parent"),
                "depth": n.get("depth", 0),
                "row": n["row"],
                "number": n["code"],
                "name": n["name"],
                "totals": {**n["totals"], **wp.node_indicators.get(n["id"], {})},
            })

        items_by_node: Dict[int, List[Dict[str, Any]]] = {}
        for it in wp.items:
            items_by_node.setdefault(it["node"], []).append(it)

        children_of: Dict[Any, List[Dict[str, Any]]] = {}
        for n in nodes_out:
            children_of.setdefault(n["parent"], []).append(n)

        file_ref = wp.file
        for n in nodes_out:
            ref = "%s | %s | %s#%s" % (file_ref, hdr.get("number", ""), n["kind_label"], n["id"])
            check_node(checks, ref, n["totals"], ls_owner_of(n["id"], n["row"], n["id"]))

        # --- иерархическая сверка: позиции узла + вложенные узлы = итог узла ----
        category_keys = ("material", "equipment", "transport", "other")
        acc_keys = ("pz", "mat_ob", "emm", "zpm", "ozp", "nr", "sp",
                    "tz_workers", "tz_mash", "total") + category_keys
        # суммы по видам работ (строительные / монтажные / оборудование / прочие) —
        # для контроля разбивки СМР по показателям узлов 214/215/216/217
        kind_keys = ("build", "install", "equipment", "other")
        kind_acc: Dict[Any, Dict[str, float]] = {}
        kind_pz_acc: Dict[Any, Dict[str, float]] = {}
        eff: Dict[Any, Dict[str, float]] = {}
        by_depth_desc = sorted(nodes_out, key=lambda x: -x["depth"])
        for n in by_depth_desc:
            s = {k: 0.0 for k in acc_keys}
            for it in items_by_node.get(n["id"], []):
                for k in acc_keys:
                    s[k] += it["values"].get(k, 0.0)
            for ch in children_of.get(n["id"], []):
                cs = eff.get(ch["id"], {})
                for k in acc_keys:
                    s[k] += cs.get(k, 0.0)
            eff[n["id"]] = s
            for key in category_keys:
                n["totals"].setdefault(key, round2(s[key]))
            sk = {k: 0.0 for k in kind_keys}
            for it in items_by_node.get(n["id"], []):
                sk[it.get("work_kind") or "other"] += it["values"].get("total", 0.0)
            for ch in children_of.get(n["id"], []):
                cs = kind_acc.get(ch["id"], {})
                for k in kind_keys:
                    sk[k] += cs.get(k, 0.0)
            kind_acc[n["id"]] = sk
            # прямые затраты (стоимость без НР и СП) по видам работ — сумма по позициям
            sp = {k: 0.0 for k in kind_keys}
            for it in items_by_node.get(n["id"], []):
                sp[it.get("work_kind") or "other"] += it["values"].get("pz", 0.0)
            for ch in children_of.get(n["id"], []):
                cp = kind_pz_acc.get(ch["id"], {})
                for k in kind_keys:
                    sp[k] += cp.get(k, 0.0)
            kind_pz_acc[n["id"]] = sp
            # «стоимость без НР и СП» по видам работ хранится в итогах узла: по ней отчёты
            # выводят прямые затраты в разрезе строительных, монтажных, оборудования и прочих
            n["totals"]["kind_pz"] = {k: sp[k] for k in kind_keys}
            pz_kind = sum(sp.values())
            if n["totals"].get("pz") is not None:
                dp = round2(pz_kind - float(n["totals"]["pz"]))
                if abs(dp) > max(EPS, abs(float(n["totals"]["pz"])) * 1e-9):
                    add_check(checks, "node", ref, "warn", "NRSP_SPLIT",
                              "Сумма прямых затрат по видам работ ≠ «Прямым затратам» узла "
                              "(стоимость без НР и СП)", dp,
                              ls=ls_owner_of(n["id"], n["row"], n["id"]))
            ref = "%s | %s | %s#%s" % (file_ref, hdr.get("number", ""), n["kind_label"], n["id"])
            # СМР = строительные + монтажные; вместе с оборудованием и прочими они дают итог
            tot_kind = sum(sk.values())
            tot_node = n["totals"].get("total")
            if tot_node:
                dk = round2(tot_kind - tot_node)
                if abs(dk) > max(EPS, abs(tot_node) * 1e-9):
                    add_check(checks, "node", ref, "warn", "SMR_TOTAL",
                              "Строительные + монтажные + оборудование + прочие ≠ «Всего с НР и СП»",
                              dk, ls=ls_owner_of(n["id"], n["row"], n["id"]))
            for k, ind, label in (("build", "smr_build", "Строительные работы с НР и СП"),
                                  ("install", "smr_install", "Монтажные работы с НР и СП"),
                                  ("equipment", "equipment", "Стоимость оборудования"),
                                  ("other", "other", "Прочие работы с НР и СП")):
                base_ind = n["totals"].get(ind)
                if base_ind is None:
                    continue
                d = round2(sk[k] - base_ind)
                if abs(d) > max(EPS, abs(base_ind) * 1e-9):
                    add_check(checks, "node", ref, "warn", "SMR_SPLIT",
                              "Сумма позиций вида «%s» ≠ показателю узла" % label, d,
                              ls=ls_owner_of(n["id"], n["row"], n["id"]))
            for k, label in (("pz", "Прямые затраты"), ("mat_ob", "Материальные ресурсы"),
                             ("emm", "Эксплуатация машин"), ("zpm", "ЗП машинистов"),
                             ("ozp", "Основная ЗП рабочих"), ("nr", "Накладные расходы"),
                             ("sp", "Сметная прибыль"), ("total", "Всего с НР и СП"),
                             ("tz_workers", "Трудозатраты строителей"),
                             ("tz_mash", "Трудозатраты машинистов")):
                base = n["totals"].get(k)
                if base is None:
                    continue
                d = s[k] - base
                if abs(d) > max(EPS, abs(base) * 1e-9):
                    add_check(checks, "node", ref, "error", "SUM_MISMATCH",
                              "Позиции + вложенные узлы ≠ итог по показателю «%s»" % label, d,
                              ls=ls_owner_of(n["id"], n["row"], n["id"]))

        for it in wp.items:
            n = by_id.get(it["node"])
            iref = "%s | %s | №%s" % (file_ref, hdr.get("number", ""), it["num"])
            if not it["name"]:
                add_check(checks, "item", iref, "warn", "NO_NAME",
                          "Не заполнено наименование позиции", ls=it.get("ls"))
            if it["qty"] < 0:
                add_check(checks, "item", iref, "info", "NEG_QTY",
                          "Отрицательное количество (возврат/корректировка)", it["qty"],
                          ls=it.get("ls"))
            if abs(it["qty"]) < 1e-12 and abs(it["values"].get("total", 0.0)) > EPS:
                add_check(checks, "item", iref, "warn", "QTY_ZERO",
                          "Нулевое количество при ненулевой стоимости", it["values"].get("total"),
                          ls=it.get("ls"))
            if abs(it["values"].get("total", 0.0)) < EPS and abs(it["values"].get("pz", 0.0)) < EPS:
                add_check(checks, "item", iref, "info", "ZERO_COST",
                          "Нулевая стоимость позиции", ls=it.get("ls"))
            if wp.has_sections and n is not None and n["kind"] not in (4, 5):
                add_check(checks, "item", iref, "info", "NO_SECTION",
                          "Позиция не отнесена ни к одному разделу сметы", ls=it.get("ls"))

        ls_nodes = [n for n in nodes_out if n["id"] in ls_ids]
        if not ls_nodes:
            # В книге может не оказаться записей структуры (A=51/52) — тогда сводную строку
            # сметы строить не из чего, и данные пропадали бы целиком. Собираем одну смету
            # по заголовку и позициям, помечая файл предупреждением NO_STRUCTURE.
            add_check(checks, "file", "%s | структура" % file_ref, "warn", "NO_STRUCTURE",
                      "В книге нет записей структуры (A=51/A=52) — смета собрана по заголовку "
                      "и позициям")
            ls_nodes = [{
                "id": None, "kind": 1, "kind_label": KIND_LABEL.get(1, "Смета (ЛС)"),
                "parent": None, "depth": 0, "row": 0, "synthetic": True,
                "number": hdr.get("number", ""), "name": hdr.get("name", ""),
                "totals": synth_totals(wp.items),
            }]

        # --- сверка с печатной формой: листы, содержащие «11 граф» ---
        printed_sheets = [p.get("sheet") for p in printed_list]
        if printed_sheets and len(printed_sheets) != len(ls_nodes):
            add_check(checks, "file", "%s | печатная форма" % file_ref, "warn", "PRINTED_COUNT",
                      "Листов с печатной формой «%s» (%d) не столько же, сколько локальных смет "
                      "в файле (%d)" % (pf.MARK, len(printed_sheets), len(ls_nodes)),
                      len(printed_sheets) - len(ls_nodes))
        printed_info: Dict[Any, Dict[str, Any]] = {}
        used_printed: set = set()
        for idx, ls in enumerate(ls_nodes):
            prn = match_printed(printed_list, ls_nodes, ls, idx, hdr, used_printed)
            if prn is None:
                add_check(checks, "smeta", "%s | %s | печатная форма" % (file_ref, hdr.get("number", "")),
                          "warn", "NO_PRINTED",
                          "Печатная форма «...%s...» не найдена — сверка с печатной формой не выполнена"
                          % pf.MARK)
                continue
            source_totals = dict(ls["totals"])
            cross = pf.compare(source_totals, prn["totals"])
            # «Перевозка грузов»: в «Source» показатель 233 включает погрузо-разгрузочные работы,
            # а в печатной форме они напечатаны отдельной строкой (сводка «ПОГР-РАЗГР»).
            # Расхождение ровно на эту величину — не ошибка, объясняем его в примечании.
            prr = float(wp.doc_summary.get("ПОГР-РАЗГР", 0.0) or 0.0)
            if prr:
                for row in cross:
                    if (row.get("key") == "transport" and row.get("status") == "mismatch"
                            and abs(abs(float(row.get("delta") or 0.0)) - prr) <= max(0.05, prr * 1e-9)):
                        row["status"] = "ok"
                        row["note"] = ("в «Source» перевозка грузов включает погрузо-разгрузочные "
                                       "работы (%.2f руб.) — они напечатаны в форме отдельно" % prr)
                        add_check(checks, "smeta",
                                  "%s | %s | печатная форма (%s)" % (file_ref, hdr.get("number", ""),
                                                                     prn["sheet"]),
                                  "info", "TRANSPORT_PRR", str(row["note"]), prr, ls=ls["id"])
            if pf.source_missing(source_totals):
                add_check(checks, "smeta",
                          "%s | %s | вкладка Source" % (file_ref, hdr.get("number", "")),
                          "warn", "SOURCE_VALUES_MISSING",
                          "Во вкладке «Source» нет значений показателей (книга пересохранена без "
                          "пересчёта формул)")
            for row in cross:
                if row["status"] == "mismatch":
                    add_check(checks, "smeta",
                              "%s | %s | печатная форма (%s)" % (file_ref, hdr.get("number", ""), prn["sheet"]),
                              "error", "PRINTED_MISMATCH",
                              "Расхождение с печатной формой по показателю «%s»" % row["label"],
                              row["delta"])
                elif row["status"] == "kind_diff":
                    add_check(checks, "smeta",
                              "%s | %s | печатная форма (%s)" % (file_ref, hdr.get("number", ""), prn["sheet"]),
                              "warn", "PRINTED_KIND_DIFF",
                              "Разбивка стоимости по видам работ: печатная форма относит часть позиций "
                              "к другим видам работ, чем вкладка «Source» (показатель «%s»). "
                              "Для отчёта принято деление по данным «Source»" % row["label"],
                              row["delta"])
            for row in pf.check_printed_self(prn):
                if row.get("severity") == "warn":
                    add_check(checks, "smeta",
                              "%s | %s | печатная форма (%s)" % (file_ref, hdr.get("number", ""), prn["sheet"]),
                              "warn", row["code"], row["message"], row.get("delta"))
            printed_info[ls["id"]] = _printed_diagnostics(prn, cross, source_totals)
            # итоги разделов и подразделов сверяем с печатной формой по наименованию
            used_sections: set = set()
            for node in nodes_out:
                if node["kind"] not in (4, 5):
                    continue
                par = by_id.get(node.get("parent")) or {}
                sec = _find_printed_section(prn, node, par.get("name") or "", used_sections)
                if sec is None:
                    continue
                node["printed_section"] = sec
                d = (sec.get("pz") or 0.0) - node["totals"].get("pz", 0.0)
                if abs(d) > max(pf.TOL, abs(sec.get("pz") or 0.0) * 1e-9):
                    add_check(checks, "node",
                              "%s | %s | печатная форма (%s)" % (file_ref, hdr.get("number", ""),
                                                                 prn.get("sheet")),
                              "error", "PRINTED_SECTION_MISMATCH",
                              "Итог «%s» в печатной форме ≠ разбору «Source» по показателю «Прямые затраты»"
                              % (node["name"] or node["number"]), d)

        multi = len(ls_nodes) > 1
        summary_all = list(getattr(wp, "summary_rows", []))

        def source_totals_for(ls: Dict[str, Any]) -> Dict[str, Any]:
            close = ls.get("close_row", wp.close_row.get(ls["id"]))
            block = [x for x in summary_all if close is None or x["row"] < close]
            return pick_source_totals(block or summary_all)

        def raw_number(node: Dict[str, Any]) -> str:
            return ((node["number"] or hdr.get("number")) if multi
                    else (hdr.get("number") or node["number"])) or ""

        num_count: Dict[str, int] = {}
        for node in ls_nodes:
            num_count[raw_number(node)] = num_count.get(raw_number(node), 0) + 1

        # --- контроль суммы всех затрат по строке «ВСЕГО ПО СМЕТЕ» вкладки «Source» ---
        total_info: Dict[Any, Dict[str, Any]] = {}
        def unit_scale(row: Optional[Dict[str, Any]], reference: float) -> Optional[Dict[str, Any]]:
            """Привести строку «Source» к рублям: в части книг сводка напечатана в тыс. руб.

            Признак — расхождение ровно в 1000 раз с расчётом по показателям узлов
            (например, в fixture-014 «ВСЕГО ПО СМЕТЕ» = 41 741,78 вместо 41 741 780).
            """
            if not row:
                return None
            ratio = (abs(float(row.get("value") or 0.0)) / abs(reference)) if reference else 0.0
            if 0.0005 < ratio < 0.002:
                return dict(row, value=float(row["value"]) * 1000.0,
                            unit_source="тыс. руб.", scale=1000.0)
            return dict(row, unit_source="руб", scale=1.0)

        for ls in ls_nodes:
            inst = printed_info.get(ls["id"], {})
            src = source_totals_for(ls)
            tot = dict(inst.get("totals_source") or ls["totals"])
            src_base = unit_scale(src.get("base"), float(tot.get("total", 0.0) or 0.0))
            src_k = unit_scale(src.get("with_k"), float(tot.get("total", 0.0) or 0.0))
            labor = (tot.get("ozp", 0.0) + tot.get("emm", 0.0) + tot.get("nr", 0.0) + tot.get("sp", 0.0))
            prn = inst.get("printed") or {}
            k_condition = next((c["value"] for c in reversed(wp.conditions)
                                if c["key"] in ("КСт_ОС", "КСт_МР", "КСт_СР")
                                and 0 < c["value"] < 3), None)
            k_derived = None
            if src_base and src_k and labor > 0:
                kk = (src_k["value"] - src_base["value"]) / labor + 1.0
                if 0.5 < kk < 3.0:
                    k_derived = round(kk, 6)
            k_src = k_derived or k_condition or 1.0
            k_from = ("строка «Source»" if k_derived else
                      "условия работ «Source»" if k_condition else "не определён (принят 1)")
            selected = synth_totals([it for it in wp.items if it.get("ls") == ls["id"]])
            selected_labor = sum(selected.get(k, 0.0) for k in ("ozp", "emm", "nr", "sp"))
            calc_k = round2(selected.get("total", 0.0) + (k_src - 1.0) * selected_labor)
            tol_k = 0.005  # любое расхождение от копейки видно в контроле
            tc: List[Dict[str, Any]] = []
            if src_base:
                dl = round2(src_base["value"] - selected.get("total", 0.0))
                st = "ok" if abs(dl) <= tol_k else "mismatch"
                tc.append({"code": "SOURCE_TOTAL_BASE",
                           "label": "Всего по смете (базисный уровень, без коэффициента)",
                           "row": src_base["row"], "unit": "руб", "source": round2(src_base["value"]),
                           "unit_source": src_base.get("unit_source"), "scale": src_base.get("scale"),
                           "calc": round2(selected.get("total", 0.0)), "printed": None,
                           "delta": dl, "status": st})
                if st == "mismatch":
                    add_check(checks, "smeta",
                              "%s | %s | вкладка Source (строка «Всего по Смете»)"
                              % (file_ref, hdr.get("number", "")), "error",
                              "SOURCE_TOTAL_BASE_MISMATCH",
                              "Сумма всех затрат (базисный уровень) ≠ строке «Всего по Смете» "
                              "вкладки «Source»", dl, ls=ls["id"])
            if src_k:
                prn_total_k = (prn.get("totals_with_k") or {}).get("total")
                d_prn = round2(src_k["value"] - float(prn_total_k)) if prn_total_k is not None else None
                d_calc = round2(src_k["value"] - calc_k)
                ok_all = (d_calc is None or abs(d_calc) <= tol_k) and \
                         (d_prn is None or abs(d_prn) <= max(0.05, abs(d_prn) * 1e-9))
                tc.append({"code": "SOURCE_TOTAL",
                           "label": "ВСЕГО ПО СМЕТЕ (с коэффициентом условий производства работ)",
                           "row": src_k["row"], "unit": "руб", "source": round2(src_k["value"]),
                           "unit_source": src_k.get("unit_source"), "scale": src_k.get("scale"),
                           "calc": calc_k,
                           "printed": (round2(float(prn_total_k)) if prn_total_k is not None else None),
                           "delta": d_calc, "delta_printed": d_prn, "k": k_src, "k_from": k_from,
                           "status": ("ok" if ok_all else "mismatch")})
                if abs(d_calc) > tol_k:
                    add_check(checks, "smeta",
                              "%s | %s | вкладка Source (строка «ВСЕГО ПО СМЕТЕ»)"
                              % (file_ref, hdr.get("number", "")), "error",
                              "SOURCE_TOTAL_MISMATCH",
                              "Сумма всех затрат с коэффициентом условий работ ≠ строке "
                              "«ВСЕГО ПО СМЕТЕ» вкладки «Source»", d_calc, ls=ls["id"])
                if d_prn is not None and abs(d_prn) > max(0.05, abs(d_prn) * 1e-9):
                    add_check(checks, "smeta",
                              "%s | %s | вкладка Source ↔ печатная форма" % (file_ref, hdr.get("number", "")),
                              "error", "SOURCE_TOTAL_FORM_MISMATCH",
                              "Строка «ВСЕГО ПО СМЕТЕ» вкладки «Source» ≠ «ВСЕГО ПО СМЕТЕ» "
                              "печатной формы «%s»" % pf.MARK, d_prn, ls=ls["id"])
            if not src_k and not src_base:
                add_check(checks, "smeta", "%s | %s | вкладка Source" % (file_ref, hdr.get("number", "")),
                          "warn", "NO_SOURCE_TOTAL",
                          "Во вкладке «Source» не найдена строка «ВСЕГО ПО СМЕТЕ» — контроль суммы "
                          "всех затрат по ней невозможен", ls=ls["id"])
            report_base = round2(src_base["value"] if src_base else float(tot.get("total", 0.0) or 0.0))
            report_total = round2(src_k["value"] if src_k else report_base + (k_src - 1.0) * labor)
            total_info[ls["id"]] = {
                "total_check": tc, "source_total_calc": calc_k, "source_work_coefficient": k_src,
                "total": report_total, "total_base": report_base,
                "source_total": ({"row": src_k["row"], "name": src_k["display"], "key": src_k["key"],
                                  "value": round2(src_k["value"]),
                                  "unit_source": src_k.get("unit_source"), "scale": src_k.get("scale")}
                                 if src_k else None),
                "source_total_base": ({"row": src_base["row"], "name": src_base["display"],
                                       "key": src_base["key"], "value": round2(src_base["value"]),
                                       "unit_source": src_base.get("unit_source"),
                                       "scale": src_base.get("scale")}
                                      if src_base else None),
            }

        for ls in ls_nodes:
            inst = printed_info.get(ls["id"], {})
            number = raw_number(ls)
            # если у нескольких ЛС файла нет собственного шифра — различаем их номером узла
            if multi and (num_count.get(number, 0) > 1 or not number):
                number = "%s (ЛС #%s)" % (number or "ЛС", ls["id"])
            name = (ls["name"] or hdr.get("name")) if multi else (hdr.get("name") or ls["name"])
            ls_items = [it for it in wp.items if it.get("ls") == ls["id"]]
            # укрупнённые группы «работа + её материалы»: порядок строк вкладки «Source»
            group_items(ls_items)
            smeta = {
                "id": "%s#%s" % (os.path.splitext(wp.file)[0], ls["id"]),
                "file": wp.file,
                "path": path,
                "doc_code": (wp.nodes[0]["code"] if wp.nodes else ""),
                "node_id": ls["id"],
                "number": number,
                "name": name,
                "ls_name": ls["name"],
                "osnovanie": hdr.get("osnovanie", ""),
                "object": hdr.get("object", ""),
                "norm_base": hdr.get("norm_base", ""),
                "tr": hdr.get("tr", ""),
                "corrections": hdr.get("corrections", ""),
                "made_by": " ".join(x for x in (hdr.get("made_by_post"), hdr.get("made_by")) if x),
                "checked_by": " ".join(x for x in (hdr.get("checked_by_post"), hdr.get("checked_by")) if x),
                "conditions": [c for c in wp.conditions if c.get("value") or c.get("on")],
                # какой расчёт (уровень цен) выбран, если в книге их несколько
                "price_selection": dict(getattr(wp, "price_selection", {}) or {}),
                "price_variants": list(getattr(wp, "price_variants", {}).values()),
                "n_items": len(ls_items),
                "n_spec_rows": wp.n_spec_rows,
                "notes": wp.notes[:40],
                "doc_summary": dict(wp.doc_summary),
                "totals": dict(ls["totals"]),
                "totals_source": dict(inst.get("totals_source") or ls["totals"]),
                "printed": inst.get("printed"),
                "printed_sheets": printed_sheets,
                "n_smetas_in_file": len(ls_nodes),
                "cross_check": inst.get("cross_check") or [],
                "resources": [rc for rc in getattr(wp, "resources", []) if rc.get("ls") == ls["id"]],
                "total_check": (total_info.get(ls["id"], {}) or {}).get("total_check") or [],
                "source_total": (total_info.get(ls["id"], {}) or {}).get("source_total"),
                "source_total_base": (total_info.get(ls["id"], {}) or {}).get("source_total_base"),
                "source_work_coefficient": total_info[ls["id"]]["source_work_coefficient"],
                "source_total_calc": (total_info.get(ls["id"], {}) or {}).get("source_total_calc"),
                # итоги сметы для отчёта по укрупнённым физобъёмам: «total» — уровень строки
                # «ВСЕГО ПО СМЕТЕ» (с коэффициентом условий работ), «total_base» — базисный
                # уровень «Всего по смете» (паритет с браузерным buildDataset)
                "total": (total_info.get(ls["id"], {}) or {}).get("total"),
                "total_base": (total_info.get(ls["id"], {}) or {}).get("total_base"),
                "nodes": nodes_out,
                "items": ls_items,
            }
            smeta["checks"] = [c for c in checks
                               if c["ref"].startswith(file_ref + " | ")
                               and c.get("ls") in (None, ls["id"])]
            smetas.append(smeta)

        file_smetas = smetas[-len(ls_nodes):] if ls_nodes else []
        pp.apply(file_smetas, printed_list, printed_sheets)
        for smeta in file_smetas:
            checks.extend(c for c in smeta["checks"] if c["code"] in (
                "PRINTED_POSITION_SELECTION", "POSITIONS_NO_GRAPH", "POSITION_AMBIGUOUS",
                "POSITION_NOT_IN_PRINTED", "PRINTED_SELECTION_TOTAL"))
            if not include_items:
                smeta["items"] = []

        files_meta.append({
            "file": wp.file,
            "path": path,
            "smetas": len(ls_nodes),
            "nodes": len(nodes_out),
            "items": len(wp.items),
            "spec_rows": wp.n_spec_rows,
            "printed_sheets": printed_sheets,
            "price_selection": dict(getattr(wp, "price_selection", {}) or {}),
        })
        if (getattr(wp, "price_selection", {}) or {}).get("variants") and \
                len(wp.price_selection["variants"]) > 1:
            sel = wp.price_selection
            chosen = next((v for v in sel["variants"] if v.get("chosen")), {})
            add_check(checks, "file", "%s | уровни цен" % file_ref, "info", "PRICE_VARIANT",
                      "В книге несколько расчётов (уровней цен): оставлен %s%s, пропущены %s"
                      % (chosen.get("id"),
                         (" («%s»)" % ", ".join(chosen.get("levels") or [])) if chosen.get("levels") else "",
                         ", ".join("%s%s (%d позиций, Σ %.2f)"
                                   % (v.get("id"),
                                      (" «%s»" % ", ".join(v.get("levels") or [])) if v.get("levels") else "",
                                      v.get("items", 0), v.get("total", 0.0))
                                   for v in sel["variants"] if not v.get("chosen"))))

    reference_report = None
    if ref_book is not None and ref_mod is not None:
        try:
            # отчёты считаются по показателям уровня строки «ВСЕГО ПО СМЕТЕ» —
            # так же, как это делает браузерный модуль интерфейса (scenarioSmetas)
            smetas_k = ref_mod.scenario_smetas(smetas)
            reference_report = {
                "fo": ref_mod.build_fo_report(ref_book, smetas_k),
                "tech": ref_mod.build_tech_report(ref_book, smetas_k),
            }
        except Exception as exc:
            warnings.append("Отчёты по справочнику не сформированы: %s" % exc)

    meta = {
        "app": "Парсер смет АтомСмета / Смета.ру",
        "version": "1.0",
        "generated": _dt.datetime.now().strftime("%d.%m.%Y %H:%M"),
        "plugin": "https://www.smeta.ru/ (ПК «АтомСмета»)",
        "source_sheet": "Source",
        "printed_mark": pf.MARK,
        "verification": "контроль данных выполняется по печатной форме — листам, содержащим «%s»" % pf.MARK,
        "price_level": "Базисный уровень цен 01.01.2000 (НБ ФЕР-2001)",
        "indicators": [{"code": c, "key": INDICATORS[c][0], "label": INDICATORS[c][1],
                        "unit": INDICATORS[c][2]} for c in sorted(INDICATORS)],
        "table_columns": [{"key": k, "label": l} for k, l in TABLE_COLUMNS],
        "category_labels": {
            "works": "Работы (расценки)",
            "material": "Материалы",
            "equipment": "Оборудование",
            "transport": "Перевозка грузов",
            "other": "Прочие работы",
        },
        "files": files_meta,
        "warnings": warnings,
        "checks_total": len(checks),
    }
    if ref_book is not None:
        meta["reference"] = {"source": ref_source, "sheets": ref_book.get("sheets", {}),
                             "fo": len(ref_book.get("fo", [])), "mult": len(ref_book.get("mult", [])),
                             "tech": len(ref_book.get("tech", []))}
        meta["reference_source"] = ref_source
    return {"meta": meta, "smetas": smetas, "checks": checks, "reference_report": reference_report}


def discover(folder: str) -> List[str]:
    out: List[str] = []
    for root, _dirs, files in os.walk(folder):
        for fn in sorted(files):
            if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                out.append(os.path.join(root, fn))
    return sorted(out, key=lambda p: os.path.basename(p))


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Парсер смет АтомСмета (вкладка Source)")
    ap.add_argument("input", help="папка со сметами *.xlsx (или отдельный файл)")
    ap.add_argument("-o", "--out", default="smetas.json", help="файл результата (JSON)")
    ap.add_argument("--report", default="", help="файл отчёта о сверках (JSON)")
    ap.add_argument("--no-items", action="store_true", help="не выгружать позиции смет")
    ap.add_argument("--reference", default="", help="книга справочника («Справочник новый.xlsx»)")
    ap.add_argument("--no-reference", action="store_true", help="не подключать справочник")
    args = ap.parse_args(argv)

    if os.path.isdir(args.input):
        paths = discover(args.input)
    else:
        paths = [args.input]
    if not paths:
        sys.stderr.write("Не найдено ни одного файла *.xlsx в «%s»\n" % args.input)
        return 2

    print("Найдено файлов: %d" % len(paths))
    data = build_dataset(paths, include_items=not args.no_items,
                         reference=(None if args.no_reference else (args.reference or None)))

    with open(args.out, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
    print("Записано: %s (%.1f КБ)" % (args.out, os.path.getsize(args.out) / 1024.0))

    if args.report:
        with open(args.report, "w", encoding="utf-8") as fh:
            json.dump(data["checks"], fh, ensure_ascii=False, indent=1)
        print("Отчёт о сверках: %s" % args.report)

    print("\n%-46s %8s %7s %8s %16s" % ("Файл", "смет", "узлов", "позиций", "Всего с НР и СП"))
    for s in data["smetas"]:
        print("%-46s %8d %7d %8d %16.2f" % (
            s["file"][:46], 1, len(s["nodes"]), s["n_items"], s["totals"].get("total", 0.0)))
    if data["meta"].get("reference"):
        r = data["meta"]["reference"]
        print("\nСправочник: %s — ФО %d, множителей %d, техники %d"
              % (r.get("source", "?"), r.get("fo", 0), r.get("mult", 0), r.get("tech", 0)))
        rep = data.get("reference_report") or {}
        if rep.get("fo"):
            fs = rep["fo"]["stats"]
            print("  физобъёмы: групп %d, позиций %d (в справочнике %d), ЕИ нормализованы %d, стоимость %.2f"
                  % (fs["rows"], fs["positions"], fs["matched"], fs["unit_known"], fs["cost"]))
        if rep.get("tech"):
            ts = rep["tech"]["stats"]
            print("  техника: машин %d, вхождений %d, стоимость %.2f"
                  % (ts["machines"], ts["mentions"], ts["cost"]))

    by_sev: Dict[str, int] = {}
    for c in data["checks"]:
        by_sev[c["severity"]] = by_sev.get(c["severity"], 0) + 1
    print("\nСверки: " + (", ".join("%s=%d" % kv for kv in sorted(by_sev.items())) or "замечаний нет"))
    for c in data["checks"][:20]:
        print("  [%s] %s: %s%s" % (c["severity"], c["ref"], c["message"],
                                   "" if "delta" not in c else " (Δ=%.2f)" % c["delta"]))
    if data["meta"]["warnings"]:
        print("\nПредупреждения:")
        for w in data["meta"]["warnings"]:
            print("  " + w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
