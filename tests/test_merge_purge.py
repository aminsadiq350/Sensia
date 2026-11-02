import pytest

import merge_purge as mp


def test_normalize_name_and_zip():
    assert mp.normalize_name("  michael ") == "MICHAEL"
    assert mp.normalize_name("o'brien") == "OBRIEN"
    assert mp.normalize_name("Smith-Jones") == "SMITHJONES"
    assert mp.normalize_name("Jr.") == "JR"
    assert mp.normalize_zip("02134-1102") == "02134"
    assert mp.normalize_zip("2134") == "02134"
    assert mp.normalize_zip("N/A") is None


def test_missing_cells_normalize_as_empty():
    missing = float("nan")
    assert mp.normalize_name(missing) == ""
    assert mp.normalize_unit(missing) == ""


def test_numeric_and_spaced_zip_values_normalize():
    assert mp.normalize_zip(2134.0) == "02134"
    assert mp.normalize_zip("02134 -1102") == "02134"


def test_lookup_tables_must_be_passed_explicitly():
    with pytest.raises(TypeError):
        mp.normalize_address("14 Maple Street")
    with pytest.raises(TypeError):
        mp.canonical_name("Bob")
    with pytest.raises(TypeError):
        mp.records_match({}, {})


def test_normalize_address_and_unit():
    suffix_map = {"STREET": "ST", "ST": "ST", "DR": "DR"}
    house, street = mp.normalize_address("14 Maple Street", suffix_map)
    assert house == "14"
    assert street == "MAPLE ST"

    house, street = mp.normalize_address("101 Forest Dr.", suffix_map)
    assert house == "101"
    assert street == "FOREST DR"

    assert mp.normalize_unit("Apt. 12b") == "12B"
    assert mp.normalize_unit("#12B") == "12B"
    assert mp.normalize_unit("") == ""


def test_duplicate_rules_match_transitively():
    nickname_map = {"BOB": "ROBERT"}
    suffix_map = {"STREET": "ST", "ST": "ST"}
    rec_a = {
        "FIRST_NAME": "Robert",
        "LAST_NAME": "King",
        "SUFFIX": "",
        "ZIP": "30301",
        "ADDR_1": "14 Maple St",
        "ADDR_2": "",
    }
    rec_b = {
        "FIRST_NAME": "Bob",
        "LAST_NAME": "KING",
        "SUFFIX": "",
        "ZIP": "30301",
        "ADDR_1": "14 MAPLE STREET",
        "ADDR_2": "",
    }
    rec_c = {
        "FIRST_NAME": "R.",
        "LAST_NAME": "King",
        "SUFFIX": "",
        "ZIP": "30301",
        "ADDR_1": "14 Maple St",
        "ADDR_2": "",
    }

    assert mp.records_match(rec_a, rec_b, nickname_map, suffix_map)
    assert mp.records_match(rec_a, rec_c, nickname_map, suffix_map)
    assert not mp.records_match(rec_b, rec_c, nickname_map, suffix_map)


def test_missing_suffix_can_match_nonempty_suffix():
    record_without_suffix = {
        "FIRST_NAME": "Daniel",
        "LAST_NAME": "Price",
        "SUFFIX": float("nan"),
        "ZIP": "78701",
        "ADDR_1": "42 Quarry Rd",
        "ADDR_2": "",
    }
    record_with_suffix = dict(record_without_suffix, SUFFIX="JR")

    assert mp.records_match(record_without_suffix, record_with_suffix, {}, {})


# ---------------------------------------------------------------------------
# Regression tests and end-to-end checks
# ---------------------------------------------------------------------------
import openpyxl


def _rec(first, last, addr="14 Maple St", unit="", suffix="", zip_="30301"):
    return {"FIRST_NAME": first, "LAST_NAME": last, "SUFFIX": suffix, "ZIP": zip_,
            "ADDR_1": addr, "ADDR_2": unit}


def test_unit_that_looks_like_filler_word_is_not_dropped_twice():
    # "St E" normalizes to "STE"; it must not be normalized again to "" and match a no-unit record.
    assert mp.normalize_unit("St E") == "STE"
    suffix_map = {"ST": "ST"}
    a = {"FIRST_NAME": "ROBERT", "LAST_NAME": "KING", "SUFFIX": "", "ZIP": "30301",
         "HOUSE_NUMBER": "14", "STREET": "MAPLE ST", "UNIT": "STE"}
    b = dict(a, UNIT="")
    assert not mp.records_match(a, b, {}, suffix_map)
    assert mp.records_match(a, dict(a), {}, suffix_map)


def test_unit_rule_empty_vs_nonempty():
    assert mp.records_match(_rec("Ann", "Lee"), _rec("Ann", "Lee"), {}, {"ST": "ST"})
    assert not mp.records_match(_rec("Ann", "Lee"), _rec("Ann", "Lee", unit="5"), {}, {"ST": "ST"})


def test_suffix_conflict_blocks_match():
    assert not mp.records_match(_rec("Tom", "Ray", suffix="Jr."), _rec("Tom", "Ray", suffix="Sr"), {}, {})


def _write_input(path, records):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Plan")
    ws.append(["KEY_CODE", "TYPE", "PRIORITY", "LIST", "SEGMENT"])
    for row in [("H1", "House", 10, "Cust", "a"), ("R1", "Outside", 50, "L1", "a"),
                ("R2", "Outside", 50, "L2", "a"), ("S1", "Suppress", None, "DNM", "a")]:
        ws.append(row)
    ws = wb.create_sheet("Records")
    ws.append(["REC_ID", "KEY_CODE", "CUSTOMER_ID", "FIRST_NAME", "LAST_NAME", "SUFFIX", "ADDRESS_1",
               "ADDRESS_2", "CITY", "STATE", "ZIP", "EMAIL", "PHONE", "GENDER", "LAST_ORDER_DATE", "ORDER_TOTAL"])
    for rec_id, key, first, last, addr, zip_ in records:
        ws.append([rec_id, key, None, first, last, None, addr, None, "X", "ST", zip_, None, None, None, None, None])
    ws = wb.create_sheet("Nicknames")
    ws.append(["NICKNAME", "FULL_NAME"])
    ws.append(["BOB", "ROBERT"])
    ws = wb.create_sheet("StreetSuffixes")
    ws.append(["WORD", "STANDARD"])
    ws.append(["STREET", "ST"])
    ws.append(["ST", "ST"])
    wb.save(path)


def _run(tmp_path, records, monkeypatch):
    src, out = tmp_path / "in.xlsx", tmp_path / "out.xlsx"
    _write_input(src, records)
    monkeypatch.setattr("sys.argv", ["merge_purge.py", str(src), str(out)])
    mp.main()
    wb = openpyxl.load_workbook(out)
    return {ws.title: [r for r in ws.iter_rows(values_only=True)] for ws in wb}


def test_worked_example_end_to_end(tmp_path, monkeypatch):
    data = _run(tmp_path, [
        (11, "H1", "Robert", "King", "14 Maple Street", 30301),
        (12, "R1", "bob", "KING", "14 MAPLE ST.", "30301-4410"),
        (13, "R2", "R.", "King", "14 Maple St", "30301"),
        (14, "R2", "Mary", "King", "14 Maple St", "30301"),
    ], monkeypatch)
    assert [r[0] for r in data["Mail"][1:]] == [11, 14]
    assert data["Mail"][1][-1] == 0 and data["Mail"][2][-1] == 1
    assert data["Drops"][1:] == [(12, "R1", "DUP_INTER", 11), (13, "R2", "DUP_INTER", 11)]


def test_na_surname_is_not_rejected(tmp_path, monkeypatch):
    data = _run(tmp_path, [(1, "R1", "Min", "NA", "5 Oak St", "30301")], monkeypatch)
    assert [r[0] for r in data["Mail"][1:]] == [1]
    assert data["Drops"][1:] == []


def test_reject_suppress_and_summary_balance(tmp_path, monkeypatch):
    data = _run(tmp_path, [
        (1, "H1", "Ann", "Lee", "1 Elm St", "30301"),
        (2, "R1", "Ann", "Lee", "1 Elm Street", "30301"),   # suppressed with 1
        (3, "S1", "Ann", "Lee", "1 Elm St", "30301"),
        (4, "R1", "Joe", "Fox", "Elm St", "30301"),         # no house number -> REJECT
        (5, "R1", "Joe", "Fox", "2 Elm St", "N/A"),         # bad ZIP -> REJECT
        (6, "R1", "", "", "3 Elm St", "30301"),             # empty last name -> REJECT
        (7, "R1", "Sam", "Roe", "9 Elm St", "30301"),
        (8, "R2", "Sam", "Roe", "9 Elm St", "30301"),       # DUP_INTER of 7 (tie -> lowest REC_ID)
    ], monkeypatch)
    reasons = {r[0]: r[2] for r in data["Drops"][1:]}
    assert reasons == {1: "SUPPRESS", 2: "SUPPRESS", 4: "REJECT", 5: "REJECT", 6: "REJECT", 8: "DUP_INTER"}
    assert [r[0] for r in data["Mail"][1:]] == [7]
    assert data["Mail"][1][-1] == 2          # MULTI_COUNT: lists L1 and L2
    for row in data["Summary"][1:]:
        key, inp, rej, sup, inter, intra, net, multis = row
        assert inp == rej + sup + inter + intra + net
    assert data["Summary"][-1][0] == "TOTAL" and data["Summary"][-1][-1] == 1
