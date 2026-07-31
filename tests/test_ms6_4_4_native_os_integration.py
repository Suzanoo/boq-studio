from pathlib import Path

from boq_studio.native_os import (
    add_recent_file,
    is_macos,
    load_recent_files,
    normalize_recent_files,
    recent_files_path,
    save_recent_files,
    user_config_dir,
)


def test_platform_detection_and_config_locations(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path / "Roaming"))
    assert is_macos("Darwin") is True
    assert is_macos("Windows") is False
    assert user_config_dir("Windows", tmp_path) == tmp_path / "Roaming" / "BOQ Studio"
    assert user_config_dir("Darwin", tmp_path) == tmp_path / "Library" / "Application Support" / "BOQ Studio"
    assert recent_files_path("Linux", tmp_path).name == "recent_files.json"


def test_recent_files_are_existing_xlsx_unique_and_limited(tmp_path):
    files = []
    for index in range(10):
        path = tmp_path / f"book-{index}.xlsx"
        path.write_bytes(b"")
        files.append(path)
    ignored = tmp_path / "notes.txt"
    ignored.write_text("x")
    result = normalize_recent_files([files[0], files[0], ignored, *files[1:]])
    assert result == files[:8]


def test_recent_files_round_trip_and_invalid_json(tmp_path):
    first = tmp_path / "first.xlsx"; first.write_bytes(b"")
    second = tmp_path / "second.xlsx"; second.write_bytes(b"")
    store = tmp_path / "recent.json"
    saved = save_recent_files(store, [first, second])
    assert saved == [first, second]
    assert load_recent_files(store) == [first, second]
    store.write_text("not-json")
    assert load_recent_files(store) == []


def test_add_recent_file_moves_workbook_to_front(tmp_path):
    first = tmp_path / "first.xlsx"; first.write_bytes(b"")
    second = tmp_path / "second.xlsx"; second.write_bytes(b"")
    assert add_recent_file(second, [first, second]) == [second, first]
