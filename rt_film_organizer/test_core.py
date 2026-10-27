from pathlib import Path

from core import FilmRecord, parse_marking, render_filename


def test_parse_spaced_marking():
    parsed = parse_marking("KDHC JY JA2026 013 015 04 100A S01 SIT SPPS380 4.9T W23 A27 261021 F")
    assert parsed["line_no"] == "JA2026-013-015-04-100A"
    assert parsed["joint_no"] == "S01"
    assert parsed["welder_no"] == "W-2023-A-27"
    assert parsed["section"] == ""


def test_parse_section():
    parsed = parse_marking("JA2026-013-08-0-0-300A R02 SEC.08 W-2023-A-12")
    assert parsed["joint_no"] == "R02"
    assert parsed["section"] == "Sec.08"
    assert parsed["welder_no"] == "W-2023-A-12"


def test_filename():
    record = FilmRecord(source=str(Path("024019.jpg")), joint_no="S01", sequence=2)
    assert render_filename("{joint}-{seq}", record) == "S01-2.JPG"


def test_parse_realistic_ocr_errors():
    parsed = parse_marking("KDHCJY5A2026/01301504 SO1 SISPPS380-4.9T/W23427/261021F 100")
    assert parsed == {
        "line_no": "JA2026-013-015-04-100A",
        "joint_no": "S01",
        "welder_no": "W-2023-A-27",
        "section": "",
    }


def test_assign_three_film_sections():
    from core import assign_sequences
    records = [
        FilmRecord(source=f"{index}.jpg", line_no="JA2026-013-015-04-100A", joint_no="S01")
        for index in range(3)
    ]
    assign_sequences(records)
    assert [record.sequence for record in records] == [1, 2, 3]
    assert [record.section for record in records] == ["(1-2)", "(2-3)", "(3-1)"]


def test_reconcile_ocr_variations_in_triplet():
    from core import assign_sequences
    records = [
        FilmRecord(source="1.jpg", line_no="JA2026-013-015-04-100A", joint_no="", welder_no="W-2023-A-27"),
        FilmRecord(source="2.jpg", line_no="", joint_no="S01", welder_no="W-2023-A-27"),
        FilmRecord(source="3.jpg", line_no="", joint_no="S1", welder_no="W-2023-A-27"),
    ]
    assign_sequences(records)
    assert [record.sequence for record in records] == [1, 2, 3]
    assert [record.section for record in records] == ["(1-2)", "(2-3)", "(3-1)"]
    assert [record.line_no for record in records] == ["JA2026-013-015-04-100A"] * 3
    assert [record.joint_no for record in records] == ["S01"] * 3


def test_repair_joint_suffix():
    assert parse_marking("JA2026 013 015 04 100A R06R W23 A27")["joint_no"] == "R06R"
    assert parse_marking("JA2026 013 015 04 100A RO6R W23 A27")["joint_no"] == "R06R"
    assert parse_marking("JA2026 013 015 04 100A S1R W23 A27")["joint_no"] == "S01R"


def test_original_three_plus_variable_repair_photos():
    from core import assign_sequences
    records = [
        FilmRecord(source="1.jpg", joint_no="R06"),
        FilmRecord(source="2.jpg", joint_no="R06"),
        FilmRecord(source="3.jpg", joint_no="R06"),
        FilmRecord(source="4.jpg", joint_no="R06R", section="(2-3)", section_source="납마킹 OCR"),
        FilmRecord(source="5.jpg", joint_no="R06R"),
    ]
    assign_sequences(records)
    assert [record.section for record in records[:3]] == ["(1-2)", "(2-3)", "(3-1)"]
    assert records[3].section == "(2-3)"
    assert records[3].section_source == "보수구간 OCR"
    assert records[4].section == ""
    assert records[4].section_source == "확인필요(보수구간 선택)"

