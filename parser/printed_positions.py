"""Состав листа Excel «Позиции»: сопоставление печатной формы с Source.

Числа из печатной формы не используются. Один номер без шифра/наименования
не подтверждает позицию; неоднозначная связь между ЛС остаётся в контроле.
Аналог браузерного модуля PrintedPositions в app/template.html.
"""
from __future__ import annotations

import re

from printed_form import round2


def text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def norm(value):
    return re.sub(r"[^0-9a-zа-я]", "", text(value).lower().replace("ё", "е"))


def number(value):
    value = re.sub(r"\s", "", text(value)).removeprefix("№").replace(",", ".")
    return re.sub(r"^0+(?=\d)", "", value) if re.fullmatch(r"\d+(?:\.\d+)?", value) else ""


def codes(value):
    value = re.sub(r"[‐‑–—−]", "-", text(value).lower())
    value = re.sub(r"(?:ферм|терм|гэснм)\s+(?=\d{2}[.-])", "м", value)
    value = re.sub(r"(?:фссцпг|тссцпг)\s+(?=\d{2}[.-])", "т", value)
    return re.findall(r"[мт]?\d{2}(?:[.-]\d{1,6}){2,}", re.sub(r"\s*([.-])\s*", r"\1", value))


def same_name(left, right, unit=None):
    right = text(right)
    if "," in right and norm(unit) and norm(right.rsplit(",", 1)[1]) == norm(unit):
        right = right.rsplit(",", 1)[0]
    left, right = norm(left), norm(right)
    return bool(left and right and (left == right or (min(len(left), len(right)) >= 32
                and (left.startswith(right) or right.startswith(left)))))


def score(item, printed):
    left, right = codes(item.get("code")), codes(printed.get("basis"))
    name = same_name(item.get("name"), printed.get("name"), item.get("unit"))
    if left and right:
        compatible = set(left).intersection(right) or any(c.startswith("т") and c[1:] in right
            and name and re.search(r"фссц|тссц", text(printed.get("basis")).lower()) for c in left)
        return (3 if name else 2) if compatible else 0
    if norm(item.get("code")) and norm(item.get("code")) == norm(printed.get("basis")):
        return 3 if name else 2
    return 1 if name else 0


def read(ws):
    result = []
    for row, cells in enumerate(ws.iter_rows(min_col=1, max_col=3, values_only=True), 1):
        a, b, c = cells
        if not number(a) or not text(c):
            continue
        if number(a) == "1" and text(b) == "2" and text(c) == "3":
            continue
        result.append({"sheet": ws.title, "row": row, "num": text(a),
                       "basis": text(b), "name": text(c)})
    return result


# пометка спецресурса, который смета исключает из стоимости («Исключен Перевозка грузов…»)
EXCLUDED_TEXT = re.compile(r"^\s*исключ", re.IGNORECASE)


def scope(smetas, printed):
    num = norm(printed.get("number"))
    pool = [i for i, s in enumerate(smetas) if num and norm(s.get("number")) == num]
    if pool:
        return set(pool)
    name = norm(printed.get("name"))
    pool = [i for i, s in enumerate(smetas) if name and
            name in (norm(s.get("name")), norm(s.get("ls_name")))]
    return set(pool or range(len(smetas)))


def apply(smetas, printed, names):
    """Добавить position_selection и проверки, сохранив items и итоги Source."""
    records, by_number, matched, ambiguous = [], {}, {}, set()
    for si, smeta in enumerate(smetas):
        for item in smeta.get("items", []):
            ri = len(records)
            records.append((si, item))
            if num := number(item.get("num")):
                by_number.setdefault(num, []).append(ri)
    for sheet in printed:
        allowed, consumed, proposals = scope(smetas, sheet), set(), []
        for position in sheet.get("position_rows", []):
            candidates = [(ri, score(records[ri][1], position))
                          for ri in by_number.get(number(position["num"]), [])
                          if records[ri][0] in allowed]
            best = max((value for _, value in candidates), default=0)
            proposals.append((position, best, [ri for ri, value in candidates if value == best and value > 0]))
        proposals.sort(key=lambda p: (-p[1], len(p[2]), p[0]["row"]))
        for position, _, candidates in proposals:
            pool = [ri for ri in candidates if ri not in consumed]
            if not pool:
                continue
            if len({records[ri][0] for ri in pool}) > 1:
                ambiguous.update(pool)
                continue
            ri = pool[0]
            consumed.add(ri)
            matched.setdefault(ri, {"source_row": records[ri][1]["row"],
                                    "printed_sheet": position["sheet"], "printed_row": position["row"]})
    # Номера позиций, которые форма печатает только строкой №N.1 (спецресурс «исключен»):
    # такая строка Source подтверждена печатью, хотя отдельного номера в форме не имеет.
    excluded_numbers = set()
    for sheet in printed:
        for position in sheet.get("position_rows", []):
            if not EXCLUDED_TEXT.match(text(position.get("name"))):
                continue
            m = re.match(r"^(\d+)(?:[.,]\d+)?$", text(position.get("num")).replace(" ", "").replace(",", "."))
            if m:
                excluded_numbers.add(m.group(1))
    for si, smeta in enumerate(smetas):
        own = [ri for ri, (owner, _) in enumerate(records) if owner == si]
        kept = [ri for ri in own if ri in matched]
        excluded = [ri for ri in own if ri not in matched]
        excluded_nums = {records[ri][1]["num"]: True for ri in excluded
                         if number(records[ri][1].get("num")) in excluded_numbers}
        selection = smeta["position_selection"] = {
            "sheets": list(names), "source_count": len(own), "matches": [matched[ri] for ri in kept],
            "excluded_rows": [records[ri][1]["row"] for ri in excluded],
            "excluded_nums": excluded_nums,
            "ambiguous_rows": [records[ri][1]["row"] for ri in excluded if ri in ambiguous]}
        ref = f'{smeta["file"]} | {smeta["number"]} | лист «Позиции»'

        def add(code, severity, message, delta=None, suffix=""):
            smeta.setdefault("checks", []).append(dict(scope="smeta", ref=ref + suffix,
                severity=severity, code=code, delta=delta, message=message))

        extra = (f" (из них {len(excluded_nums)} подтверждены строкой «исключен» формы "
                 f"и оставлены)") if excluded_nums else ""
        add("PRINTED_POSITION_SELECTION", "warn" if excluded else "info",
            f'На лист «Позиции» включено {len(kept)} из {len(own)} отмеченных позиций Source. '
            f'Исключено: {len(excluded)}{extra}. Состав проверен по листам с «граф» в названии.')
        if not names:
            add("POSITIONS_NO_GRAPH", "warn", 'В книге нет листа с «граф» в названии. '
                'Лист «Позиции» не содержит позиций этой книги.')
        else:
            for ri in excluded:
                item = records[ri][1]
                reason = ('Неоднозначное совпадение между локальными сметами.' if ri in ambiguous
                          else 'Соответствующая строка печатной формы не найдена.')
                add("POSITION_AMBIGUOUS" if ri in ambiguous else "POSITION_NOT_IN_PRINTED", "warn",
                    f'Не включена в лист «Позиции»: № {item["num"]}, шифр {item.get("code") or "—"}. {reason}',
                    suffix=f' | Source!{item["row"]}')
        if excluded and smeta.get("source_total"):
            total, labor = 0.0, 0.0
            for ri in kept:
                values = records[ri][1].get("values", {})
                total += values.get("total", 0.0)
                labor += sum(values.get(k, 0.0) for k in ("ozp", "emm", "nr", "sp"))
            calc = round2(total + ((smeta.get("source_work_coefficient") or 1) - 1) * labor)
            source = smeta["source_total"]["value"]
            delta = round2(source - calc)
            selection["total"] = calc
            if abs(delta) > 0.005:
                add("PRINTED_SELECTION_TOTAL", "warn", f'Итог Source: {source:g} руб.; '
                    f'сумма позиций листа «Позиции» с коэффициентом условий работ: {calc:g} руб. '
                    'В сводке сохранён итог Source.', delta)
