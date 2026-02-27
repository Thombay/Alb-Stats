import pandas as pd
from datetime import date

from app import (
    UNKNOWN_SEMESTER,
    build_intensity_per_person,
    canonicalize_override_map,
    chargen_override_key,
    classify_chargen_kind,
    count_chargen_units,
    dashboard_data_from_records,
    default_status_selection,
    extract_entry_semesters,
    extract_semester,
    parse_chargen_entries,
    semester_sort_key,
)


def test_parse_chargen_entries_splits_pipe_and_newline() -> None:
    value = "Alb:Senior (SS 2023) | Alb:Consenior (WS 2023/24)\nAlb:Barwart (SS 2024)"
    parsed = parse_chargen_entries(value)
    assert parsed == [
        "Alb:Senior (SS 2023)",
        "Alb:Consenior (WS 2023/24)",
        "Alb:Barwart (SS 2024)",
    ]


def test_extract_semester_and_unknown() -> None:
    assert extract_semester("Alb:Senior (ss   2022)") == "SS 2022"
    assert extract_semester("Alb:Consenior (WS 2023/24)") == "WS 2023/24"
    assert extract_semester("Alb:Nichts ohne Semester") == UNKNOWN_SEMESTER


def test_semester_sort_key_orders_semesters() -> None:
    semesters = ["WS 2024/25", "SS 2024", "SS 2023", UNKNOWN_SEMESTER, "WS 2023/24"]
    ordered = sorted(semesters, key=semester_sort_key)
    assert ordered == ["SS 2023", "WS 2023/24", "SS 2024", "WS 2024/25", UNKNOWN_SEMESTER]


def test_classify_chargen_kind() -> None:
    phil_date = pd.Timestamp("2024-03-01")
    assert classify_chargen_kind("SS 2023", phil_date, "Alb:Senior (SS 2023)") == "Aktivenchargen"
    assert classify_chargen_kind("WS 2024/25", phil_date, "OV-Präsident (Grazer Cartellverband)") == (
        "Aktivenchargen"
    )
    assert classify_chargen_kind("WS 2024/25", phil_date, "Alb:Bierkassier (WS 2024/25)") == "Aktivenchargen"
    assert classify_chargen_kind("SS 2023", phil_date, "Phil-xx (Philisterconsenior 1)") == "Philisterchargen"
    assert classify_chargen_kind("WS 2024/25", phil_date, "Alb:Scriptor (WS 2024/25)") == "Philisterchargen"
    assert classify_chargen_kind(UNKNOWN_SEMESTER, phil_date, "Alb:NoSemester") == "Unklare Chargen"


def test_dashboard_data_builds_type_counts_per_person() -> None:
    member_records = [
        {
            "Couleurname": "Stoiber",
            "Mitgliedstatus": "UP",
            "AgeYears": 35.0,
            "PhilistrierungDate": "2024-03-01",
            "ChargenEntries": [
                "Alb:Senior (SS 2023)",
                "Alb:Consenior (WS 2024/25)",
                "Alb:NoSemester",
            ],
            "TotalChargen": 3,
        }
    ]
    data = dashboard_data_from_records(member_records)
    type_row = data["kind_person"].set_index("Couleurname").loc["Stoiber"]
    assert int(type_row["Aktivenchargen"]) == 2
    assert int(type_row["Philisterchargen"]) == 0
    assert int(type_row["Unklare Chargen"]) == 1


def test_default_status_selection_is_empty() -> None:
    assert default_status_selection(["BU", "UP", "EM"]) == []


def test_funktionaere_are_not_counted_as_chargen() -> None:
    member_records = [
        {
            "Couleurname": "Beispiel",
            "Mitgliedstatus": "UP",
            "AgeYears": 30.0,
            "PhilistrierungDate": "2024-03-01",
            "ChargenEntries": [
                "Chef-Red. (Chefredakteur)",
                "Contaktforum:Zirkel-Vorsitzender (Ab 16. Jan. 2025)",
                "Alb:Senior (SS 2023)",
            ],
            "TotalChargen": 99,
        }
    ]
    data = dashboard_data_from_records(member_records)
    per_person = data["per_person_total"].set_index("Couleurname").loc["Beispiel"]
    assert int(per_person["TotalChargen"]) == 1


def test_date_range_counts_as_multiple_semesters() -> None:
    entry = "GCV:OV-Praesident (Von 01. Jul. 2009 bis 30. Jun. 2010)"
    semesters = extract_entry_semesters(entry)
    assert semesters == ["WS 2009/10", "SS 2010"]
    assert count_chargen_units([entry]) == 2


def test_manual_override_can_reclassify_unclear_or_exclude() -> None:
    member_records = [
        {
            "Couleurname": "Test",
            "Mitgliedstatus": "UP",
            "AgeYears": 40.0,
            "PhilistrierungDate": None,
            "ChargenEntries": ["XYZ:Amt ohne Semester"],
            "TotalChargen": 1,
        }
    ]

    unresolved = dashboard_data_from_records(member_records, {})
    row = unresolved["kind_person"].set_index("Couleurname").loc["Test"]
    assert int(row["Unklare Chargen"]) == 1
    assert len(unresolved["unknown_entries"]) == 1

    as_active = dashboard_data_from_records(member_records, {"XYZ:Amt ohne Semester": "aktiven"})
    row_active = as_active["kind_person"].set_index("Couleurname").loc["Test"]
    assert int(row_active["Aktivenchargen"]) == 1

    as_funktionaer = dashboard_data_from_records(member_records, {"XYZ:Amt ohne Semester": "funktionaere"})
    total = as_funktionaer["per_person_total"].set_index("Couleurname").loc["Test"]
    assert int(total["TotalChargen"]) == 0


def test_override_key_removes_date_part() -> None:
    key = chargen_override_key("GCV:OV-Praesident (Von 01. Jul. 2009 bis 30. Jun. 2010)")
    assert key == "OV-Praesident"


def test_manual_override_applies_to_same_role_without_date() -> None:
    member_records = [
        {
            "Couleurname": "Test",
            "Mitgliedstatus": "UP",
            "AgeYears": 40.0,
            "PhilistrierungDate": None,
            "ChargenEntries": [
                "GCV:OV-Praesident (Von 01. Jul. 2009 bis 30. Jun. 2010)",
                "GCV:OV-Praesident (Von 01. Jul. 2011 bis 30. Jun. 2012)",
            ],
            "TotalChargen": 4,
        }
    ]
    key = "OV-Praesident"
    as_funktionaer = dashboard_data_from_records(member_records, {key: "funktionaere"})
    total = as_funktionaer["per_person_total"].set_index("Couleurname").loc["Test"]
    assert int(total["TotalChargen"]) == 0


def test_manual_override_with_old_prefixed_key_still_applies() -> None:
    member_records = [
        {
            "Couleurname": "Test",
            "Mitgliedstatus": "UP",
            "AgeYears": 40.0,
            "PhilistrierungDate": None,
            "ChargenEntries": ["GCV:OV-Praesident (Von 01. Jul. 2009 bis 30. Jun. 2010)"],
            "TotalChargen": 2,
        }
    ]
    as_funktionaer = dashboard_data_from_records(member_records, {"GCV:OV-Praesident": "funktionaere"})
    total = as_funktionaer["per_person_total"].set_index("Couleurname").loc["Test"]
    assert int(total["TotalChargen"]) == 0


def test_canonicalize_override_map_repairs_mojibake_keys() -> None:
    fixed = canonicalize_override_map({"OV-SchriftfÃ¼hrer": "verband_aktiven"})
    assert "OV-Schriftführer" in fixed
    assert fixed["OV-Schriftführer"] == "verband_aktiven"


def test_category_person_contains_funktionaere_but_total_excludes_them() -> None:
    member_records = [
        {
            "Couleurname": "Test",
            "Mitgliedstatus": "UP",
            "AgeYears": 40.0,
            "PhilistrierungDate": "2024-03-01",
            "ChargenEntries": [
                "Alb:Senior (SS 2023)",
                "Alb:IT-Beauftragter (Ab 01. Jul. 2023)",
            ],
            "TotalChargen": 2,
        }
    ]
    data = dashboard_data_from_records(member_records, {})
    category = data["category_person"]
    assert not category[category["ChargenCategory"] == "Funktionaere"].empty
    total = data["per_person_total"].set_index("Couleurname").loc["Test"]
    assert int(total["TotalChargen"]) == 1


def test_intensity_uses_reception_to_today() -> None:
    included = pd.DataFrame(
        [
            {"Couleurname": "A", "Semester": "SS 2023"},
            {"Couleurname": "A", "Semester": "WS 2023/24"},
            {"Couleurname": "A", "Semester": "SS 2024"},
            {"Couleurname": "A", "Semester": "WS 2024/25"},
        ]
    )
    reception = pd.DataFrame([{"Couleurname": "A", "ReceptionDate": pd.Timestamp("2020-01-01")}])
    out = build_intensity_per_person(included, reception, today=date(2025, 1, 1))
    row = out.set_index("Couleurname").loc["A"]
    expected_years = (date(2025, 1, 1) - date(2020, 1, 1)).days / 365.2425
    assert abs(float(row["BasisYears"]) - expected_years) < 0.01
    assert abs(float(row["AvgPerYear"]) - (4.0 / expected_years)) < 0.01


def test_gaius_ov_kassier_and_vizepraesident_count_as_one_per_semester() -> None:
    entry1 = "OV-Vizepraesident (Albertina) (Von 01.10.1981 bis 30.09.1983)"
    entry2 = "OV-Kassier (Albertina) (Von 01.10.1981 bis 30.09.1983)"
    member_records = [
        {
            "Couleurname": "Gaius",
            "Mitgliedstatus": "UP",
            "AgeYears": 70.0,
            "PhilistrierungDate": "1990-01-01",
            "ChargenEntries": [entry1, entry2],
            "TotalChargen": 4,
        }
    ]
    data = dashboard_data_from_records(member_records, {})
    total = data["per_person_total"].set_index("Couleurname").loc["Gaius"]
    expected_single = count_chargen_units([entry1])
    assert int(total["TotalChargen"]) == expected_single
