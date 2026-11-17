"""Read attendance templates, maintain personal rules, and prepare verifiable writes."""

import argparse
import collections
import contextlib
import copy
import datetime as dt
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import json
import os
from pathlib import Path
import random
import re
import secrets
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
import posixpath


NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REST = {"休息", "休", "事", "事假", "病", "病假", "缺", "缺勤", "调休"}
WEEKDAYS = "一二三四五六日"


def default_profile_path():
    codex_home = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    return codex_home / "skill-data" / "attendance-fill" / "profile.json"


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def atomic_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=path.parent)
    try:
        with os.fdopen(handle, "w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        os.replace(temporary, path)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


@contextlib.contextmanager
def file_lock(path):
    """Fail clearly on concurrent writers instead of silently losing an update."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    lock = path.with_name(path.name + ".lock")
    try:
        descriptor = os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError as exc:
        raise ValueError(f"配置或本期记录正在被其他进程更新，请稍后重试。锁文件为 {lock}") from exc
    try:
        os.write(descriptor, str(os.getpid()).encode("ascii"))
        os.close(descriptor)
        yield
    finally:
        lock.unlink(missing_ok=True)


def defaults():
    return read_json(Path(__file__).resolve().parents[1] / "assets" / "default_rules.json")


def load_profile(path, require_name=True):
    if not Path(path).exists():
        if require_name:
            raise ValueError("首次使用需要询问并保存使用者姓名，请运行 profile --name 姓名。")
        return {**defaults(), "name": None}
    profile = read_json(path)
    for key, value in defaults().items():
        profile.setdefault(key, copy.deepcopy(value))
    if require_name and not profile.get("name"):
        raise ValueError("个人配置缺少姓名，请首次确认使用者姓名后保存。")
    return profile


def column_number(label):
    result = 0
    for character in label:
        result = result * 26 + ord(character) - 64
    return result


def column_label(number):
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def split_address(address):
    match = re.fullmatch(r"([A-Z]+)(\d+)", address)
    if not match:
        raise ValueError(f"无效单元格地址 {address}")
    return column_number(match.group(1)), int(match.group(2))


def address(column, row):
    return f"{column_label(column)}{row}"


def scalar(value):
    try:
        number = Decimal(value)
        return int(number) if number == number.to_integral_value() else float(number)
    except Exception:
        return value


def read_workbook(path):
    """Read only serialized cells, avoiding inflated Excel used-range dimensions."""
    path = Path(path).resolve()
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        properties = workbook.find("m:workbookPr", NS)
        date_1904 = properties is not None and properties.get("date1904") in {"1", "true"}
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(node.itertext()) for node in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        relationships = {
            node.get("Id"): node.get("Target")
            for node in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        styles = ET.fromstring(archive.read("xl/styles.xml")) if "xl/styles.xml" in archive.namelist() else None
        sheets = []
        for sheet in workbook.findall("m:sheets/m:sheet", NS):
            target = relationships[sheet.get(f"{{{REL}}}id")]
            part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
            root = ET.fromstring(archive.read(part))
            cells = {}
            rows = {}
            for row in root.findall("m:sheetData/m:row", NS):
                rows[int(row.get("r"))] = dict(row.attrib)
                for cell in row.findall("m:c", NS):
                    value_node = cell.find("m:v", NS)
                    value = value_node.text if value_node is not None else None
                    cell_type = cell.get("t")
                    if cell_type == "s" and value is not None:
                        value = strings[int(value)]
                    elif cell_type == "inlineStr":
                        inline = cell.find("m:is", NS)
                        value = "".join(inline.itertext()) if inline is not None else ""
                    elif value is not None and cell_type == "b":
                        value = value == "1"
                    elif value is not None and cell_type not in {"str", "e"}:
                        value = scalar(value)
                    formula = cell.find("m:f", NS)
                    cells[cell.get("r")] = {
                        "value": value,
                        "formula": formula.text if formula is not None else None,
                        "style": int(cell.get("s", "0")),
                    }
            sheets.append({
                "name": sheet.get("name"), "cells": cells, "root": root, "rows": rows,
                "merges": [node.get("ref") for node in root.findall("m:mergeCells/m:mergeCell", NS)],
                "columns": [dict(node.attrib) for node in root.findall("m:cols/m:col", NS)],
            })
        return {"path": str(path), "sheets": sheets, "styles": styles, "date_1904": date_1904}


def value_at(sheet, cell):
    return sheet["cells"].get(cell, {}).get("value")


def date_value(value, date_1904):
    if isinstance(value, (int, float)) and 20000 <= value <= 100000:
        epoch = dt.date(1904, 1, 1) if date_1904 else dt.date(1899, 12, 30)
        return epoch + dt.timedelta(days=int(value))
    if isinstance(value, str):
        match = re.fullmatch(r"\s*(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})日?\s*", value)
        if match:
            return dt.date(*map(int, match.groups()))
    return None


def merge_anchor(sheet, cell):
    column, row = split_address(cell)
    for merged in sheet["merges"]:
        start, end = merged.split(":") if ":" in merged else (merged, merged)
        first_column, first_row = split_address(start)
        last_column, last_row = split_address(end)
        if first_column <= column <= last_column and first_row <= row <= last_row:
            return start, (first_column, first_row, last_column, last_row)
    return cell, (column, row, column, row)


def find_layout(workbook, name=None):
    candidates = []
    for sheet in workbook["sheets"]:
        for cell, item in sheet["cells"].items():
            if item["value"] != "日期":
                continue
            date_column, header_row = split_address(cell)
            headers = {split_address(ref)[0]: entry["value"] for ref, entry in sheet["cells"].items()
                       if split_address(ref)[1] == header_row and entry["value"] is not None}
            weekday_columns = [column for column, label in headers.items() if label == "星期"]
            if not weekday_columns:
                continue
            people = []
            for column, label in sorted(headers.items()):
                if isinstance(label, str) and headers.get(column + 1) == "当日总工时":
                    people.append({"name": label.strip(), "period_column": column_label(column),
                                   "hours_column": column_label(column + 1), "column_number": column})
            person = next((item for item in people if item["name"] == name), None)
            dates = []
            total_row = None
            weekday_errors = []
            for ref, entry in sheet["cells"].items():
                column, row = split_address(ref)
                if column != date_column or row <= header_row:
                    continue
                if str(entry["value"]).strip() == "合计总工时":
                    total_row = row
                day = date_value(entry["value"], workbook["date_1904"])
                if day is None:
                    continue
                _, bounds = merge_anchor(sheet, ref)
                block_length = bounds[3] - bounds[1] + 1
                if block_length != 3:
                    raise ValueError(f"{sheet['name']}!{ref} 日期区块应为三行，实际为 {block_length} 行。")
                weekday_cell = address(weekday_columns[0], row)
                weekday = value_at(sheet, weekday_cell)
                expected = WEEKDAYS[day.weekday()]
                if weekday not in {f"星期{expected}", f"周{expected}", f"星期{day.weekday() + 1}", f"周{day.weekday() + 1}"}:
                    weekday_errors.append({"cell": weekday_cell, "date": day.isoformat(), "value": weekday,
                                           "expected": f"星期{expected}"})
                record = {"date": day.isoformat(), "weekday": day.weekday(), "row": row,
                          "date_cell": ref, "weekday_cell": weekday_cell}
                if person:
                    record["period_cells"] = [f"{person['period_column']}{row + slot}" for slot in range(3)]
                    record["hours_cell"] = merge_anchor(sheet, f"{person['hours_column']}{row}")[0]
                dates.append(record)
            dates.sort(key=lambda record: record["date"])
            if not dates:
                continue
            if len({record["date"] for record in dates}) != len(dates):
                raise ValueError("模板中存在重复日期。")
            total_cell = None
            if person and total_row:
                total_cell = merge_anchor(sheet, f"{person['period_column']}{total_row}")[0]
            candidates.append({"sheet": sheet["name"], "header_row": header_row, "people": people,
                               "person": person, "days": dates, "total_cell": total_cell,
                               "weekday_errors": weekday_errors})
    if name:
        candidates = [candidate for candidate in candidates if candidate["person"]]
    if not candidates:
        raise ValueError(f"找不到姓名为 {name} 的考勤栏目。" if name else "找不到符合格式的考勤工作表。")
    if len(candidates) > 1:
        raise ValueError("找到多个匹配的考勤表，请先提供单一目标周期的模板。")
    return candidates[0]


def category(weekday):
    return "saturday" if weekday == 5 else "sunday" if weekday == 6 else "weekday"


def minutes_text(minutes):
    return f"{minutes // 60}:{minutes % 60:02d}"


def format_hours(minutes):
    value = (Decimal(minutes) / Decimal(60)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return format(value, "f").rstrip("0").rstrip(".") if value else "0"


def parse_period(value):
    if value is None or str(value).strip() == "":
        return None
    text = str(value).strip()
    if text in REST:
        return "rest"
    match = re.fullmatch(r"(\d{1,2})[:：](\d{2})\s*[-–—~～至]\s*(\d{1,2})[:：](\d{2})", text)
    if not match:
        raise ValueError(f"无法识别出勤时段 {value}")
    start_hour, start_minute, end_hour, end_minute = map(int, match.groups())
    if start_hour > 23 or start_minute > 59 or end_hour > 24 or end_minute > 59 or (end_hour == 24 and end_minute):
        raise ValueError(f"出勤时段超出当日范围 {value}")
    start, end = start_hour * 60 + start_minute, end_hour * 60 + end_minute
    if end <= start:
        raise ValueError(f"出勤时段结束应晚于开始 {value}")
    return [start, end]


def validate_periods(periods):
    last_end = None
    total = 0
    for period in periods:
        if period in (None, "rest"):
            continue
        start, end = period
        if not (0 <= start < end <= 1440):
            raise ValueError("出勤时段超出当日范围或时长无效。")
        if last_end is not None and start < last_end:
            raise ValueError("同一天的出勤时段相互重叠。")
        last_end = end
        total += end - start
    return total


def parse_hours(value):
    if value is None:
        return None
    match = re.fullmatch(r"\s*(.*?)(\d+(?:\.\d+)?)\s*(?:h|小时|时)?\s*", str(value), re.I)
    if not match:
        return None
    return {"location": match.group(1).strip(), "hours": Decimal(match.group(2))}


def command_profile(args):
    with file_lock(args.profile):
        profile = load_profile(args.profile, require_name=False)
        if args.name is not None:
            name = args.name.strip()
            if not name:
                raise ValueError("姓名不能为空。")
            profile["name"] = name
            atomic_json(args.profile, profile)
        if not profile.get("name"):
            return {"ok": False, "pending": ["首次使用，请询问使用者姓名后运行 profile --name 姓名。"]}
        return {"ok": True, "profile_path": str(Path(args.profile).resolve()), "profile": profile}


def command_inspect(args):
    profile = load_profile(args.profile, require_name=False)
    workbook = read_workbook(args.template)
    layout = find_layout(workbook, args.name or profile.get("name"))
    return {"ok": True, "template": workbook["path"], **layout,
            "weekends": [day["date"] for day in layout["days"] if day["weekday"] >= 5]}


def command_learn(args):
    with file_lock(args.profile):
        profile = load_profile(args.profile)
        patterns = collections.defaultdict(list)
        locations = collections.Counter()
        conflicts, skipped, sources = [], [], []
        for path in args.history:
            workbook = read_workbook(path)
            layout = find_layout(workbook, profile["name"])
            sheet = next(item for item in workbook["sheets"] if item["name"] == layout["sheet"])
            for day in layout["days"]:
                source = {"file": workbook["path"], "sheet": sheet["name"], "date": day["date"],
                          "range": f"{day['period_cells'][0]}:{column_label(split_address(day['hours_cell'])[0])}{day['row'] + 2}"}
                labels = [value_at(sheet, cell) for cell in day["period_cells"]]
                hours_text = value_at(sheet, day["hours_cell"])
                hours = parse_hours(hours_text)
                try:
                    periods = [parse_period(label) for label in labels]
                    total = validate_periods(periods)
                except ValueError as exc:
                    skipped.append({**source, "reason": str(exc)})
                    continue
                if any(period in (None, "rest") for period in periods) or hours is None:
                    skipped.append({**source, "reason": "含空白、休息或缺少日工时"})
                    if hours and hours["location"]:
                        locations[hours["location"]] += 1
                    continue
                if abs(Decimal(total) - hours["hours"] * 60) > Decimal("0.31"):
                    conflicts.append({**source, "periods": labels, "computed_minutes": total,
                                      "computed_hours": format_hours(total), "written_hours": hours_text})
                    continue
                if hours["location"]:
                    locations[hours["location"]] += 1
                    skipped.append({**source, "reason": "出差日，仅提取地点格式"})
                    continue
                group = category(day["weekday"])
                pattern = tuple(tuple(period) for period in periods)
                patterns[group].append((pattern, day["date"], source))
        changes = {}
        for group, observations in patterns.items():
            counts = collections.Counter(item[0] for item in observations)
            best_count = max(counts.values())
            tied = {pattern for pattern, count in counts.items() if count == best_count}
            previous = tuple(tuple(period) for period in profile["schedules"][group])
            selected = previous if previous in tied else max(
                (item for item in observations if item[0] in tied), key=lambda item: item[1]
            )[0]
            learned = [list(period) for period in selected]
            changes[group] = {"before": profile["schedules"][group], "after": learned,
                              "count": best_count, "eligible_days": len(observations)}
            profile["schedules"][group] = learned
            sources.extend(item[2] for item in observations if item[0] == selected)
        if not patterns:
            raise ValueError("历史记录中没有可用于更新常规时段的完整非出差出勤日。")
        if locations:
            best_locations = [location for location, count in locations.items() if count == max(locations.values())]
            current_location = profile.get("default_trip_location")
            profile["default_trip_location"] = current_location if current_location in best_locations else sorted(best_locations)[0]
        profile["history_sources"] = sources
        profile["history_conflicts"] = conflicts
        atomic_json(args.profile, profile)
        return {"ok": True, "name": profile["name"], "changes": changes, "conflicts": conflicts,
                "skipped": skipped, "sources": sources,
                "default_trip_location": profile.get("default_trip_location")}


def parse_confirmations(raw, dates, profile):
    pending = []
    known_dates = {day["date"] for day in dates}
    weekends = raw.get("weekends", {})
    if not isinstance(weekends, dict):
        raise ValueError("weekends 应为日期到是否出勤的映射。")
    for day in dates:
        if day["weekday"] >= 5 and type(weekends.get(day["date"])) is not bool:
            pending.append(f"{day['date']} 周末是否出勤尚未确认")
    if "trips" not in raw:
        pending.append("本期出差情况尚未确认，请提供 trips，确认无出差时使用空列表")
    if raw.get("other_exceptions_confirmed") is not True:
        pending.append("其他例外情况尚未确认，请提供 other_exceptions_confirmed=true")
    trips, exceptions = {}, {}
    for item in raw.get("trips", []):
        date = item.get("date")
        if date not in known_dates:
            raise ValueError(f"出差日期 {date} 超出本期范围。")
        if date in trips:
            raise ValueError(f"出差日期 {date} 重复。")
        if type(item.get("attend")) is not bool:
            pending.append(f"{date} 出差日是否出勤尚未确认")
        location = str(item.get("location") or profile.get("default_trip_location") or "").strip()
        if item.get("attend") is True and not location:
            pending.append(f"{date} 出差日缺少地点")
        trips[date] = {**item, "location": location}
    for item in raw.get("exceptions", []):
        date = item.get("date")
        if date not in known_dates:
            raise ValueError(f"例外日期 {date} 超出本期范围。")
        if date in exceptions:
            raise ValueError(f"例外日期 {date} 重复。")
        if "periods" in item:
            if not isinstance(item["periods"], list) or len(item["periods"]) != 3:
                raise ValueError(f"{date} 例外时段需要上午、下午、晚上三项。")
            periods = [parse_period(value) for value in item["periods"]]
            if any(period is None for period in periods):
                raise ValueError(f"{date} 例外时段中空白请明确填写休息。")
            validate_periods(periods)
            exceptions[date] = {**item, "parsed_periods": periods}
        elif item.get("status") in REST:
            exceptions[date] = item
        else:
            raise ValueError(f"{date} 例外应提供有效 status 或三段 periods。")
    unknown_weekends = set(weekends) - known_dates
    if unknown_weekends:
        raise ValueError(f"周末确认含本期范围外日期 {sorted(unknown_weekends)}")
    for date, trip in trips.items():
        weekday = dt.date.fromisoformat(date).weekday()
        if weekday >= 5 and type(weekends.get(date)) is bool and type(trip.get("attend")) is bool:
            if weekends[date] != trip["attend"]:
                pending.append(f"{date} 周末与出差的出勤确认互相冲突")
    return pending, weekends, trips, exceptions


def adjusted_periods(base, variation, rng):
    step = int(variation["step_minutes"])
    maximum = int(variation["max_shift_minutes"])
    if step <= 0 or maximum < step:
        raise ValueError("随机步长必须大于零，最大偏移至少为一个步长。")
    offsets = list(range(-(maximum // step) * step, maximum + 1, step))
    for _ in range(100):
        periods = copy.deepcopy(base)
        for slot in rng.sample(range(3), rng.choice([1, 2])):
            shifts = [rng.choice(offsets), rng.choice(offsets)]
            if shifts == [0, 0]:
                shifts[0] = step
            periods[slot] = [base[slot][0] + shifts[0], base[slot][1] + shifts[1]]
        try:
            validate_periods(periods)
            if any(periods[index + 1][0] - periods[index][1] < step for index in range(2)):
                continue
            return periods
        except ValueError:
            continue
    raise ValueError("当前作息与随机范围无法生成合理时段，请检查常规时段或随机范围。")


def build_plan(args, workbook, layout, profile, raw):
    pending, weekends, trips, exceptions = parse_confirmations(raw, layout["days"], profile)
    if layout["weekday_errors"]:
        raise ValueError(f"模板日期与星期不符 {layout['weekday_errors']}")
    if not layout["total_cell"]:
        raise ValueError("找不到本人期间合计总工时单元格。")
    if pending:
        return {"ok": False, "pending": pending}
    schedules = profile["schedules"]
    for base in schedules.values():
        if len(base) != 3:
            raise ValueError("常规作息必须包含上午、下午、晚上三段。")
        validate_periods(base)
    first, last = layout["days"][0]["date"], layout["days"][-1]["date"]
    state_key = hashlib.sha256(json.dumps([profile["name"], layout["sheet"], first, last], ensure_ascii=False).encode()).hexdigest()[:24]
    state_directory = Path(args.state_dir) if args.state_dir else Path(args.profile).parent / "periods"
    state_path = state_directory / f"{state_key}.json"
    rules_digest = hashlib.sha256(json.dumps({key: profile[key] for key in ("schedules", "variation")}, sort_keys=True).encode()).hexdigest()
    with file_lock(state_path):
        state = read_json(state_path) if state_path.exists() and not args.regenerate else {"seed": secrets.randbits(64), "days": {}}
        rng = random.Random(state["seed"])
        attendance = {}
        for day in layout["days"]:
            date = day["date"]
            attend = weekends.get(date, True)
            if date in trips:
                attend = trips[date]["attend"]
            if date in exceptions:
                exception = exceptions[date]
                attend = "parsed_periods" in exception and validate_periods(exception["parsed_periods"]) > 0
            attendance[date] = attend
        distribution = {}
        rebalanced_dates = []
        for group in ("weekday", "saturday", "sunday"):
            eligible = [day["date"] for day in layout["days"] if category(day["weekday"]) == group
                        and attendance[day["date"]] and day["date"] not in exceptions]
            existing = {date: state["days"][date] for date in eligible
                        if state["days"].get(date, {}).get("rules_digest") == rules_digest}
            new_dates = [date for date in eligible if date not in existing]
            count = len(eligible)
            target = min(count - 2, max(1, round(count * float(profile["variation"]["fraction"])))) if count >= 3 else 0
            new_variations = max(0, min(len(new_dates), target - sum(day["varied"] for day in existing.values())))
            chosen = set(rng.sample(new_dates, new_variations))
            for date in new_dates:
                base = schedules[group]
                periods = adjusted_periods(base, profile["variation"], rng) if date in chosen else copy.deepcopy(base)
                state["days"][date] = {"rules_digest": rules_digest, "periods": periods, "varied": date in chosen}
            changed = sum(bool(state["days"][date]["varied"]) for date in eligible)
            if count >= 3 and changed == 0:
                date = rng.choice(eligible)
                state["days"][date]["periods"] = adjusted_periods(schedules[group], profile["variation"], rng)
                state["days"][date]["varied"] = True
                rebalanced_dates.append(date)
                changed = 1
            if count >= 3 and count - changed < 2:
                variations = [date for date in eligible if state["days"][date]["varied"]]
                for date in rng.sample(variations, 2 - (count - changed)):
                    state["days"][date]["periods"] = copy.deepcopy(schedules[group])
                    state["days"][date]["varied"] = False
                    rebalanced_dates.append(date)
                    changed -= 1
            distribution[group] = {"attendance_days": count, "base_days": count - changed, "varied_days": changed}
        days, writes = [], []
        total_minutes = 0
        for day in layout["days"]:
            date, group = day["date"], category(day["weekday"])
            trip = trips.get(date)
            exception = exceptions.get(date)
            if exception and "parsed_periods" in exception:
                periods = exception["parsed_periods"]
                labels = [profile["rest_label"] if period == "rest" else f"{minutes_text(period[0])}-{minutes_text(period[1])}" for period in periods]
                kind = "exception"
            elif not attendance[date]:
                periods = ["rest"] * 3
                labels = [exception["status"] if exception else profile["rest_label"]] * 3
                kind = "exception" if exception else "rest"
            else:
                periods = state["days"][date]["periods"]
                labels = [f"{minutes_text(period[0])}-{minutes_text(period[1])}" for period in periods]
                kind = "trip" if trip else group
            minutes = validate_periods(periods)
            prefix = trip["location"] if trip and minutes else ""
            hours_text = f"{prefix}{format_hours(minutes)}h"
            total_minutes += minutes
            record = {"date": date, "periods": labels, "minutes": minutes, "hours_text": hours_text,
                      "kind": kind, "attend": minutes > 0, "location": prefix,
                      "period_cells": day["period_cells"], "hours_cell": day["hours_cell"]}
            days.append(record)
            writes.extend({"cell": cell, "value": value} for cell, value in zip(day["period_cells"], labels))
            writes.append({"cell": day["hours_cell"], "value": hours_text})
        total_text = f"{format_hours(total_minutes)}h"
        writes.append({"cell": layout["total_cell"], "value": total_text})
        if len({item["cell"] for item in writes}) != len(writes):
            raise ValueError("目标单元格重叠，请检查模板的人员栏目及合并范围。")
        plan = {"ok": True, "schema_version": 1, "template": workbook["path"], "sheet": layout["sheet"],
                "name": profile["name"], "start": first, "end": last, "writes": writes, "days": days,
                "total_minutes": total_minutes, "total_cell": layout["total_cell"], "total_text": total_text,
                "rules_digest": rules_digest, "state_path": str(state_path.resolve()),
                "summary": {"days": len(days), "attendance_days": sum(day["attend"] for day in days),
                            "trip_attendance_days": sum(day["attend"] and bool(day["location"]) for day in days),
                            "distribution": distribution, "rebalanced_dates": rebalanced_dates}}
        atomic_json(args.output, plan)
        state.update({"name": profile["name"], "sheet": layout["sheet"], "start": first, "end": last})
        atomic_json(state_path, state)
    return {"ok": True, "plan_path": str(Path(args.output).resolve()), "name": profile["name"],
            "total_minutes": total_minutes, "total_text": total_text, "summary": plan["summary"]}


def command_plan(args):
    if Path(args.output).resolve() == Path(args.template).resolve() or Path(args.output).suffix.lower() != ".json":
        raise ValueError("计划输出需要独立的 .json 文件路径。")
    profile = load_profile(args.profile)
    workbook = read_workbook(args.template)
    layout = find_layout(workbook, profile["name"])
    return build_plan(args, workbook, layout, profile, read_json(args.confirmations))


def canonical(node):
    if node is None:
        return None
    # Explicit Unicode font names determine rendering. Legacy code-page metadata
    # can be normalized away by XLSX engines without changing the visible font.
    children = [canonical(child) for child in node if child.tag.rsplit("}", 1)[-1] != "charset"]
    return (node.tag.rsplit("}", 1)[-1], tuple(sorted(node.attrib.items())),
            (node.text or "").strip(), tuple(sorted(children)))


def style_signatures(workbook):
    styles = workbook["styles"]
    if styles is None:
        return [None]
    fonts = styles.find("m:fonts", NS)
    fills = styles.find("m:fills", NS)
    borders = styles.find("m:borders", NS)
    custom_formats = {int(node.get("numFmtId")): node.get("formatCode") for node in styles.findall("m:numFmts/m:numFmt", NS)}
    result = []
    for xf in styles.findall("m:cellXfs/m:xf", NS):
        number_format = int(xf.get("numFmtId", "0"))
        result.append({
            "font": canonical(fonts[int(xf.get("fontId", "0"))]) if fonts is not None else None,
            "fill": canonical(fills[int(xf.get("fillId", "0"))]) if fills is not None else None,
            "border": canonical(borders[int(xf.get("borderId", "0"))]) if borders is not None else None,
            "number_format": custom_formats.get(number_format, f"builtin:{number_format}"),
            "alignment": canonical(xf.find("m:alignment", NS)),
            "protection": canonical(xf.find("m:protection", NS)),
        })
    return result


def dimensions(sheet, maximum_column, maximum_row):
    formatting = sheet["root"].find("m:sheetFormatPr", NS)
    attributes = formatting.attrib if formatting is not None else {}
    default_width = float(attributes.get("defaultColWidth", "8.43"))
    default_height = float(attributes.get("defaultRowHeight", "15"))
    columns = {column: {"width": default_width, "hidden": False} for column in range(1, maximum_column + 1)}
    for item in sheet["columns"]:
        for column in range(max(1, int(item["min"])), min(maximum_column, int(item["max"])) + 1):
            columns[column] = {"width": float(item.get("width", default_width)), "hidden": item.get("hidden") in {"1", "true"}}
    rows = {row: {"height": float(sheet["rows"].get(row, {}).get("ht", default_height)),
                  "hidden": sheet["rows"].get(row, {}).get("hidden") in {"1", "true"}}
            for row in range(1, maximum_row + 1)}
    return columns, rows


def effective_style(sheet, column, row):
    item = sheet["cells"].get(address(column, row))
    if item is not None:
        return item["style"]
    row_info = sheet["rows"].get(row, {})
    if row_info.get("customFormat") in {"1", "true"}:
        return int(row_info.get("s", "0"))
    for item in sheet["columns"]:
        if int(item["min"]) <= column <= int(item["max"]):
            return int(item.get("style", "0"))
    return 0


def command_verify(args):
    original, filled = read_workbook(args.template), read_workbook(args.filled)
    plan = read_json(args.plan)
    issues = collections.defaultdict(list)
    original_names = [sheet["name"] for sheet in original["sheets"]]
    filled_names = [sheet["name"] for sheet in filled["sheets"]]
    if original_names != filled_names:
        issues["sheets"].append({"before": original_names, "after": filled_names})
    expected = {item["cell"]: item["value"] for item in plan["writes"]}
    original_styles, filled_styles = style_signatures(original), style_signatures(filled)
    for source in original["sheets"]:
        target = next((sheet for sheet in filled["sheets"] if sheet["name"] == source["name"]), None)
        if target is None:
            continue
        active = source["name"] == plan["sheet"]
        for cell in set(source["cells"]) | set(target["cells"]):
            before = source["cells"].get(cell, {})
            after = target["cells"].get(cell, {})
            if active and cell in expected:
                if after.get("value") != expected[cell] or after.get("formula") is not None:
                    issues["writes"].append({"sheet": source["name"], "cell": cell,
                                              "expected": expected[cell], "actual": after.get("value")})
            elif before.get("formula") != after.get("formula") or (
                not before.get("formula") and before.get("value") != after.get("value")
            ):
                issues["other_content"].append({"sheet": source["name"], "cell": cell,
                                                 "before": before, "after": after})
        if set(source["merges"]) != set(target["merges"]):
            issues["merges"].append({"sheet": source["name"],
                                       "missing": sorted(set(source["merges"]) - set(target["merges"])),
                                       "added": sorted(set(target["merges"]) - set(source["merges"]))})
        content_cells = [ref for ref, item in source["cells"].items() if item["value"] is not None or item["formula"]]
        maximum_column = max((split_address(ref)[0] for ref in content_cells), default=1)
        maximum_row = max((split_address(ref)[1] for ref in content_cells), default=1)
        for merged in source["merges"]:
            last = merged.split(":")[-1]
            column, row = split_address(last)
            maximum_column, maximum_row = max(maximum_column, column), max(maximum_row, row)
        before_columns, before_rows = dimensions(source, maximum_column, maximum_row)
        after_columns, after_rows = dimensions(target, maximum_column, maximum_row)
        for column in before_columns:
            before, after = before_columns[column], after_columns[column]
            if abs(before["width"] - after["width"]) > 0.0001 or before["hidden"] != after["hidden"]:
                issues["column_dimensions"].append({"sheet": source["name"], "column": column_label(column),
                                                    "before": before, "after": after})
        for row in before_rows:
            before, after = before_rows[row], after_rows[row]
            if abs(before["height"] - after["height"]) > 0.0001 or before["hidden"] != after["hidden"]:
                issues["row_dimensions"].append({"sheet": source["name"], "row": row, "before": before, "after": after})
        for row in range(1, maximum_row + 1):
            for column in range(1, maximum_column + 1):
                before = original_styles[effective_style(source, column, row)]
                after = filled_styles[effective_style(target, column, row)]
                if before != after:
                    issues["styles"].append({"sheet": source["name"], "cell": address(column, row),
                                               "components": [key for key in before if before[key] != after[key]]})
    minutes = 0
    for day in plan["days"]:
        computed = validate_periods([parse_period(period) for period in day["periods"]])
        if computed != day["minutes"]:
            issues["arithmetic"].append({"date": day["date"], "expected": computed, "plan": day["minutes"]})
        expected_hours = f"{day.get('location', '')}{format_hours(computed)}h"
        if expected_hours != day["hours_text"]:
            issues["arithmetic"].append({"date": day["date"], "expected": expected_hours, "plan": day["hours_text"]})
        minutes += computed
    if minutes != plan["total_minutes"] or f"{format_hours(minutes)}h" != plan["total_text"]:
        issues["arithmetic"].append({"expected_minutes": minutes, "plan_total_minutes": plan["total_minutes"]})
    try:
        layout = find_layout(filled, plan["name"])
        if layout["weekday_errors"]:
            issues["weekday"].extend(layout["weekday_errors"])
        if [day["date"] for day in layout["days"]] != [day["date"] for day in plan["days"]]:
            issues["dates"].append("保存后日期与计划不一致")
    except ValueError as exc:
        issues["layout"].append(str(exc))
    charsets_before = sorted({node.get("val") for node in original["styles"].findall("m:fonts/m:font/m:charset", NS)}) if original["styles"] is not None else []
    charsets_after = sorted({node.get("val") for node in filled["styles"].findall("m:fonts/m:font/m:charset", NS)}) if filled["styles"] is not None else []
    metadata_changes = []
    if charsets_before != charsets_after:
        metadata_changes.append({"field": "legacy_font_charset", "before": charsets_before, "after": charsets_after,
                                 "meaning": "显式字体名称和可见字体属性另行校验，代码页元数据归一化不计为可见样式变化"})
    return {"ok": not issues, "checked_writes": len(expected), "total_minutes": minutes,
            "checks": {key: len(value) for key, value in issues.items()},
            "issues": {key: value[:20] for key, value in issues.items()}, "metadata_changes": metadata_changes}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    profile = subparsers.add_parser("profile", help="Save an explicitly supplied name or read the personal profile")
    profile.add_argument("--name")
    inspect = subparsers.add_parser("inspect", help="Inspect the template and confirmation dates")
    inspect.add_argument("--template", required=True)
    inspect.add_argument("--name")
    learn = subparsers.add_parser("learn", help="Update personal rules from uploaded attendance history")
    learn.add_argument("--history", nargs="+", required=True)
    plan = subparsers.add_parser("plan", help="Generate stable attendance writes after confirmations")
    plan.add_argument("--template", required=True)
    plan.add_argument("--confirmations", required=True)
    plan.add_argument("--output", required=True)
    plan.add_argument("--state-dir")
    plan.add_argument("--regenerate", action="store_true")
    verify = subparsers.add_parser("verify", help="Verify saved workbook values, arithmetic, and template preservation")
    verify.add_argument("--template", required=True)
    verify.add_argument("--plan", required=True)
    verify.add_argument("--filled", required=True)
    for command_parser in (profile, inspect, learn, plan):
        command_parser.add_argument("--profile", default=str(default_profile_path()))
    arguments = parser.parse_args()
    try:
        result = globals()[f"command_{arguments.command}"](arguments)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result.get("ok") else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        return 1


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())
