from __future__ import annotations

import argparse
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook


MAIL_SHEETS = ["Mail", "Drops", "Summary"]

def read_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    """Read a sheet treating only truly empty cells as missing.

    pandas would otherwise turn the *text* "NA", "NULL", "None", "nan"... into NaN,
    e.g. turning the surname "Na" / "NA" into a blank last name.
    """
    return pd.read_excel(path, sheet_name=sheet_name, keep_default_na=False, na_values=[""])


def _is_missing(value: Any) -> bool:
    return value is None or bool(pd.isna(value))


def normalize_name(value: Any) -> str:
    if _is_missing(value):
        return ""
    return re.sub(r"[^A-Z]", "", str(value).upper().strip())


def normalize_zip(value: Any) -> str | None:
    if _is_missing(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool) and float(value).is_integer():
        text = str(int(value))
    else:
        text = str(value).strip()
    if not text:
        return None
    if "-" in text:
        text = text.split("-", 1)[0].strip()
    if not re.fullmatch(r"\d{3,5}", text):
        return None
    return text.zfill(5)


def normalize_address(value: Any, suffix_map: dict[str, str]) -> tuple[str | None, str | None]:
    if _is_missing(value):
        return None, None
    text = str(value).upper().replace(".", " ").replace(",", " ")
    tokens = text.split()
    if not tokens:
        return None, None
    house_number = tokens[0]
    if not house_number.isdigit():
        return None, None

    street_tokens = tokens[1:]
    if street_tokens:
        last = street_tokens[-1]
        if last in suffix_map:
            last = suffix_map[last]
        street_tokens[-1] = last
        street = " ".join(street_tokens)
    else:
        street = ""
    return house_number, street


def normalize_unit(value: Any) -> str:
    if _is_missing(value):
        return ""
    text = str(value).upper().replace(".", "").replace(",", "").replace("#", "")
    pieces = [piece for piece in text.split() if piece not in {"APT", "APARTMENT", "UNIT", "STE", "SUITE"}]
    return "".join(pieces)


def canonical_name(name: Any, nickname_map: dict[str, str]) -> str:
    cleaned = normalize_name(name)
    if not cleaned:
        return ""
    return nickname_map.get(cleaned, cleaned)


def _normalize_record_like(record: dict[str, Any], suffix_map: dict[str, str]) -> dict[str, Any]:
    if "HOUSE_NUMBER" in record and "STREET" in record:
        first_name = normalize_name(record.get("FIRST_NAME"))
        last_name = normalize_name(record.get("LAST_NAME"))
        suffix = normalize_name(record.get("SUFFIX"))
        zip_value = normalize_zip(record.get("ZIP"))
        house_number = str(record.get("HOUSE_NUMBER", ""))
        street = str(record.get("STREET", ""))
        unit = "" if _is_missing(record.get("UNIT")) else str(record.get("UNIT"))  # already normalized
        return {
            "FIRST_NAME": first_name,
            "LAST_NAME": last_name,
            "SUFFIX": suffix,
            "ZIP": zip_value,
            "HOUSE_NUMBER": house_number,
            "STREET": street,
            "UNIT": unit,
        }

    first_name = normalize_name(record.get("FIRST_NAME") or record.get("FIRST"))
    last_name = normalize_name(record.get("LAST_NAME") or record.get("LAST"))
    suffix = normalize_name(record.get("SUFFIX"))
    zip_value = normalize_zip(record.get("ZIP"))
    house_number, street = normalize_address(record.get("ADDR_1") or record.get("ADDRESS_1"), suffix_map)
    unit = normalize_unit(record.get("ADDRESS_2") or record.get("ADDR_2"))
    return {
        "FIRST_NAME": first_name,
        "LAST_NAME": last_name,
        "SUFFIX": suffix,
        "ZIP": zip_value,
        "HOUSE_NUMBER": house_number or "",
        "STREET": street or "",
        "UNIT": unit,
    }


def records_match(
    a: dict[str, Any],
    b: dict[str, Any],
    nickname_map: dict[str, str],
    suffix_map: dict[str, str],
) -> bool:
    a_norm = _normalize_record_like(a, suffix_map)
    b_norm = _normalize_record_like(b, suffix_map)

    if a_norm["ZIP"] != b_norm["ZIP"]:
        return False
    if a_norm["HOUSE_NUMBER"] != b_norm["HOUSE_NUMBER"]:
        return False
    if a_norm["STREET"] != b_norm["STREET"]:
        return False

    a_unit = a_norm["UNIT"]
    b_unit = b_norm["UNIT"]
    if a_unit == "" and b_unit == "":
        pass
    elif a_unit == "" or b_unit == "":
        return False
    elif a_unit != b_unit:
        return False

    if a_norm["LAST_NAME"] != b_norm["LAST_NAME"]:
        return False

    a_suffix = a_norm["SUFFIX"]
    b_suffix = b_norm["SUFFIX"]
    if a_suffix and b_suffix and a_suffix != b_suffix:
        return False

    a_first = a_norm["FIRST_NAME"]
    b_first = b_norm["FIRST_NAME"]
    if not a_first or not b_first:
        return False

    if len(a_first) == 1 and a_first == b_first[:1]:
        return True
    if len(b_first) == 1 and b_first == a_first[:1]:
        return True

    if canonical_name(a_first, nickname_map) == canonical_name(b_first, nickname_map):
        return True

    return False


def load_plan(path: Path) -> tuple[dict[str, dict[str, Any]], list[str]]:
    plan_df = read_sheet(path, "Plan")
    plan_map: dict[str, dict[str, Any]] = {}
    order: list[str] = []
    for _, row in plan_df.iterrows():
        key_code = str(row["KEY_CODE"]).strip()
        if not key_code:
            continue
        priority_value = row["PRIORITY"]
        priority = int(priority_value) if pd.notna(priority_value) else 999999
        plan_map[key_code] = {
            "KEY_CODE": key_code,
            "TYPE": str(row["TYPE"]).strip(),
            "PRIORITY": priority,
            "LIST": row["LIST"],
        }
        order.append(key_code)
    return plan_map, order


def load_nicknames(path: Path) -> dict[str, str]:
    nicknames = read_sheet(path, "Nicknames")
    mapping: dict[str, str] = {}
    for _, row in nicknames.iterrows():
        nickname = normalize_name(row["NICKNAME"])
        full_name = normalize_name(row["FULL_NAME"])
        if nickname and full_name:
            mapping[nickname] = full_name
    return mapping


def load_suffixes(path: Path) -> dict[str, str]:
    suffixes = read_sheet(path, "StreetSuffixes")
    mapping: dict[str, str] = {}
    for _, row in suffixes.iterrows():
        key = normalize_name(row["WORD"])
        value = normalize_name(row["STANDARD"])
        if key and value:
            mapping[key] = value
    return mapping


def main() -> None:
    parser = argparse.ArgumentParser(description="Merge and purge mailing records by duplicate-person rules.")
    parser.add_argument("input_path", help="Path to merge_test_input.xlsx")
    parser.add_argument("output_path", help="Path to output.xlsx")
    args = parser.parse_args()

    input_path = Path(args.input_path)
    output_path = Path(args.output_path)

    plan_map, plan_order = load_plan(input_path)
    nickname_map = load_nicknames(input_path)
    suffix_map = load_suffixes(input_path)

    records_df = read_sheet(input_path, "Records")

    valid_records: list[dict[str, Any]] = []
    dropped_rows: list[dict[str, Any]] = []
    mail_input_counts: defaultdict[str, int] = defaultdict(int)
    reject_counts: defaultdict[str, int] = defaultdict(int)
    suppress_counts: defaultdict[str, int] = defaultdict(int)
    dup_inter_counts: defaultdict[str, int] = defaultdict(int)
    dup_intra_counts: defaultdict[str, int] = defaultdict(int)
    net_counts: defaultdict[str, int] = defaultdict(int)

    for _, row in records_df.iterrows():
        key_code = str(row["KEY_CODE"]).strip()
        if key_code in plan_map and plan_map[key_code]["TYPE"].lower() in {"house", "outside"}:
            mail_input_counts[key_code] += 1

        file_meta = plan_map.get(key_code)
        if not file_meta:
            continue

        file_type = str(file_meta["TYPE"]).strip()
        priority = int(file_meta["PRIORITY"])
        list_name = file_meta["LIST"]

        first_name = normalize_name(row["FIRST_NAME"])
        last_name = normalize_name(row["LAST_NAME"])
        suffix_value = normalize_name(row["SUFFIX"])
        zip_value = normalize_zip(row["ZIP"])

        house_number, street = normalize_address(row["ADDRESS_1"], suffix_map)
        if street:
            street = str(street).upper()
        if house_number is not None and street is not None:
            house_number = str(house_number)
        unit_value = normalize_unit(row["ADDRESS_2"])

        is_valid = bool(last_name) and bool(zip_value) and house_number is not None and street is not None
        if file_type.lower() == "suppress":
            if not is_valid:
                continue
            valid_records.append(
                {
                    "REC_ID": int(row["REC_ID"]),
                    "KEY_CODE": key_code,
                    "FIRST_NAME": first_name,
                    "LAST_NAME": last_name,
                    "SUFFIX": suffix_value,
                    "ZIP": zip_value,
                    "HOUSE_NUMBER": house_number,
                    "STREET": street,
                    "UNIT": unit_value,
                    "FILE_TYPE": file_type,
                    "PRIORITY": priority,
                    "LIST": list_name,
                    "MAIL": False,
                }
            )
            continue

        if not is_valid:
            reject_counts[key_code] += 1
            dropped_rows.append(
                {
                    "REC_ID": int(row["REC_ID"]),
                    "KEY_CODE": key_code,
                    "REASON": "REJECT",
                    "SURVIVOR_REC_ID": "",
                }
            )
            continue

        valid_records.append(
            {
                "REC_ID": int(row["REC_ID"]),
                "KEY_CODE": key_code,
                "FIRST_NAME": first_name,
                "LAST_NAME": last_name,
                "SUFFIX": suffix_value,
                "ZIP": zip_value,
                "HOUSE_NUMBER": house_number,
                "STREET": street,
                "UNIT": unit_value,
                "FILE_TYPE": file_type,
                "PRIORITY": priority,
                "LIST": list_name,
                "MAIL": True,
            }
        )

    parent = list(range(len(valid_records)))

    def find(index: int) -> int:
        while parent[index] != index:
            parent[index] = parent[parent[index]]
            index = parent[index]
        return index

    def union(a: int, b: int) -> None:
        root_a = find(a)
        root_b = find(b)
        if root_a != root_b:
            parent[root_b] = root_a

    for i in range(len(valid_records)):
        for j in range(i + 1, len(valid_records)):
            if records_match(valid_records[i], valid_records[j], nickname_map, suffix_map):
                union(i, j)

    groups: dict[int, list[int]] = defaultdict(list)
    for index in range(len(valid_records)):
        groups[find(index)].append(index)

    final_survivors: set[int] = set()
    survivor_map: dict[int, int] = {}
    for group_indices in groups.values():
        mail_group = [idx for idx in group_indices if valid_records[idx]["MAIL"]]
        has_suppress = any(not valid_records[idx]["MAIL"] for idx in group_indices)

        if not mail_group:
            continue

        if has_suppress:
            for idx in mail_group:
                key_code = valid_records[idx]["KEY_CODE"]
                suppress_counts[key_code] += 1
                dropped_rows.append(
                    {
                        "REC_ID": valid_records[idx]["REC_ID"],
                        "KEY_CODE": key_code,
                        "REASON": "SUPPRESS",
                        "SURVIVOR_REC_ID": "",
                    }
                )
            continue

        winner = min(mail_group, key=lambda idx: (valid_records[idx]["PRIORITY"], valid_records[idx]["REC_ID"]))
        survivor_map[winner] = winner
        final_survivors.add(winner)

        for idx in mail_group:
            if idx == winner:
                continue
            reason = "DUP_INTRA" if valid_records[idx]["KEY_CODE"] == valid_records[winner]["KEY_CODE"] else "DUP_INTER"
            if reason == "DUP_INTER":
                dup_inter_counts[valid_records[idx]["KEY_CODE"]] += 1
            else:
                dup_intra_counts[valid_records[idx]["KEY_CODE"]] += 1
            dropped_rows.append(
                {
                    "REC_ID": valid_records[idx]["REC_ID"],
                    "KEY_CODE": valid_records[idx]["KEY_CODE"],
                    "REASON": reason,
                    "SURVIVOR_REC_ID": valid_records[winner]["REC_ID"],
                }
            )

    mail_rows: list[dict[str, Any]] = []
    for idx in sorted(final_survivors, key=lambda i: valid_records[i]["REC_ID"]):
        record = valid_records[idx]
        key_code = record["KEY_CODE"]
        net_counts[key_code] += 1
        group_members = [
            valid_records[j]
            for j in groups[find(idx)]
            if valid_records[j]["MAIL"]
        ]
        outside_lists = {
            valid_records[j]["LIST"]
            for j in groups[find(idx)]
            if valid_records[j]["MAIL"] and valid_records[j]["FILE_TYPE"].lower() == "outside"
        }
        multi_count = len(outside_lists) if record["FILE_TYPE"].lower() == "outside" else 0
        if record["FILE_TYPE"].lower() == "outside" and multi_count == 0:
            multi_count = 1
        if record["FILE_TYPE"].lower() == "house":
            multi_count = 0

        row = {
            "REC_ID": record["REC_ID"],
            "KEY_CODE": key_code,
            "FIRST_NAME": record["FIRST_NAME"],
            "LAST_NAME": record["LAST_NAME"],
            "SUFFIX": record["SUFFIX"],
            "ADDRESS_1": "",
            "ADDRESS_2": "",
            "CITY": "",
            "STATE": "",
            "ZIP5": record["ZIP"],
            "MULTI_COUNT": multi_count,
        }
        mail_rows.append(row)

    # Re-read source rows for original values so output shows original strings exactly.
    original_records = records_df.set_index("REC_ID").to_dict(orient="index")
    for row in mail_rows:
        source = original_records.get(row["REC_ID"])
        if source is not None:
            row["FIRST_NAME"] = source.get("FIRST_NAME")
            row["LAST_NAME"] = source.get("LAST_NAME")
            row["SUFFIX"] = source.get("SUFFIX")
            row["ADDRESS_1"] = source.get("ADDRESS_1")
            row["ADDRESS_2"] = source.get("ADDRESS_2")
            row["CITY"] = source.get("CITY")
            row["STATE"] = source.get("STATE")

    summary_rows: list[dict[str, Any]] = []
    for key_code in plan_order:
        if plan_map[key_code]["TYPE"].lower() not in {"house", "outside"}:
            continue
        input_count = mail_input_counts.get(key_code, 0)
        reject_count = reject_counts.get(key_code, 0)
        suppress_count = suppress_counts.get(key_code, 0)
        dup_inter_count = dup_inter_counts.get(key_code, 0)
        dup_intra_count = dup_intra_counts.get(key_code, 0)
        net = net_counts.get(key_code, 0)
        multis = 0
        for row in mail_rows:
            if row["KEY_CODE"] == key_code and row["MULTI_COUNT"] >= 2:
                multis += 1
        summary_rows.append(
            {
                "KEY_CODE": key_code,
                "INPUT": input_count,
                "REJECT": reject_count,
                "SUPPRESS": suppress_count,
                "DUP_INTER": dup_inter_count,
                "DUP_INTRA": dup_intra_count,
                "NET": net,
                "MULTIS": multis,
            }
        )

    total_input = sum(r["INPUT"] for r in summary_rows)
    total_reject = sum(r["REJECT"] for r in summary_rows)
    total_suppress = sum(r["SUPPRESS"] for r in summary_rows)
    total_dup_inter = sum(r["DUP_INTER"] for r in summary_rows)
    total_dup_intra = sum(r["DUP_INTRA"] for r in summary_rows)
    total_net = sum(r["NET"] for r in summary_rows)
    total_multis = sum(r["MULTIS"] for r in summary_rows)
    summary_rows.append(
        {
            "KEY_CODE": "TOTAL",
            "INPUT": total_input,
            "REJECT": total_reject,
            "SUPPRESS": total_suppress,
            "DUP_INTER": total_dup_inter,
            "DUP_INTRA": total_dup_intra,
            "NET": total_net,
            "MULTIS": total_multis,
        }
    )

    mail_df = pd.DataFrame(mail_rows, columns=["REC_ID", "KEY_CODE", "FIRST_NAME", "LAST_NAME", "SUFFIX", "ADDRESS_1", "ADDRESS_2", "CITY", "STATE", "ZIP5", "MULTI_COUNT"]).sort_values("REC_ID").reset_index(drop=True)

    drops_df = pd.DataFrame(dropped_rows, columns=["REC_ID", "KEY_CODE", "REASON", "SURVIVOR_REC_ID"]).sort_values("REC_ID").reset_index(drop=True)
    summary_df = pd.DataFrame(summary_rows, columns=["KEY_CODE", "INPUT", "REJECT", "SUPPRESS", "DUP_INTER", "DUP_INTRA", "NET", "MULTIS"])

    for df in [mail_df, drops_df, summary_df]:
        for column in df.columns:
            df[column] = df[column].apply(lambda value: "" if pd.isna(value) else value)

    wb = Workbook()
    wb.remove(wb.active)
    for title, data in [("Mail", mail_df), ("Drops", drops_df), ("Summary", summary_df)]:
        ws = wb.create_sheet(title)
        ws.append(list(data.columns))
        for _, row in data.iterrows():
            ws.append([row[col] for col in data.columns])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


if __name__ == "__main__":
    main()
