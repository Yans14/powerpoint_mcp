from utils.units import to_emu


def test_to_emu_inches() -> None:
    assert to_emu("1in") == 914400


def test_to_emu_points() -> None:
    assert to_emu("24pt") == 304800


def test_to_emu_raw_number() -> None:
    assert to_emu(12345) == 12345
