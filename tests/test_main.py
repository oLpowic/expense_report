"""Testy main.py: komendy, komunikaty błędów i automatyczny raport roczny w styczniu."""
import datetime
import shutil

import pytest

import main
from src.report_yearly import create_yearly_report
from tests.conftest import EXAMPLE


@pytest.fixture
def run(project, monkeypatch, capsys):
    """Uruchamia main() w tymczasowym projekcie z podanymi argumentami."""
    monkeypatch.setattr(main, "BASE", project)

    def _run(*args):
        monkeypatch.setattr("sys.argv", ["main.py", *args])
        main.main()
        return capsys.readouterr().out
    return _run


def test_help_prints_usage(run):
    for flag in ("help", "-h", "--help"):
        out = run(flag)
        assert "python main.py year" in out


def test_missing_config_gives_hint(tmp_path, monkeypatch):
    monkeypatch.setattr(main, "BASE", tmp_path)
    monkeypatch.setattr("sys.argv", ["main.py"])
    with pytest.raises(SystemExit) as err:
        main.main()
    assert "config/template" in str(err.value)


def test_default_run_creates_report_for_latest_month(run, project):
    out = run()
    assert (project / "output" / "2026" / "2026-11.xlsx").exists()
    assert "Okres 2026-11 (kalendarzowy" in out


def test_month_argument_and_summary_numbers(run, project):
    out = run("2026-10")
    assert (project / "output" / "2026" / "2026-10.xlsx").exists()
    assert "wydatki 1600.00 PLN" in out


def test_payday_flag_uses_payday_periods(run, project):
    out = run("2026-11", "--payday")
    assert "od wypłaty" in out and "2026-11-03" in out        # okres zaczyna się od wypłaty z 3.11
    assert (project / "output" / "2026" / "2026-11.xlsx").exists()


def test_flag_can_come_before_month_and_calendar_flag_overrides_settings(run, project):
    import json
    cfg = project / "config" / "settings.json"
    data = json.loads(cfg.read_text(encoding="utf-8"))
    data["period_mode"] = "payday"
    cfg.write_text(json.dumps(data), encoding="utf-8")
    assert "od wypłaty" in run("2026-11")                     # tryb z settings.json
    assert "kalendarzowy" in run("--calendar", "2026-11")     # flaga wygrywa z ustawieniem


def test_default_month_in_payday_mode_is_latest_period(run):
    assert "Okres 2026-11 (od wypłaty" in run("--payday")


def test_unknown_or_double_flag_exits(run):
    for args in (("--weekly",), ("--payday", "--calendar")):
        with pytest.raises(SystemExit) as err:
            run(*args)
        assert "flag" in str(err.value)


def test_payday_mode_without_payday_exits_with_hint(run, project):
    import json
    cfg = project / "config" / "settings.json"
    data = json.loads(cfg.read_text(encoding="utf-8"))
    data["payday"] = {"min_amount": 999999}
    cfg.write_text(json.dumps(data), encoding="utf-8")
    with pytest.raises(SystemExit) as err:
        run("--payday")
    assert "nie znaleziono żadnej wypłaty" in str(err.value)


def test_yearly_report_ignores_payday_flag(run, project):
    run("year", "2026", "--payday")
    assert (project / "output" / "2026_roczny.xlsx").exists()


def test_month_with_filter(run, project):
    run("2026-11", "everyday")
    assert (project / "output" / "2026" / "2026-11_everyday.xlsx").exists()


def test_unknown_filter_lists_available_ones(run):
    with pytest.raises(FileNotFoundError) as err:
        run("2026-11", "nie_ma_takiego")
    assert "everyday" in str(err.value)


def test_month_without_transactions_exits_with_message(run):
    with pytest.raises(SystemExit) as err:
        run("2025-01")
    assert "Brak transakcji" in str(err.value)


def test_year_command(run, project):
    out = run("year", "2026")
    assert (project / "output" / "2026_roczny.xlsx").exists()
    assert "Rok 2026" in out


def test_year_command_defaults_to_year_of_last_transaction(run, project):
    run("year")
    assert (project / "output" / "2026_roczny.xlsx").exists()


def test_year_command_with_filter(run, project):
    run("year", "2026", "everyday")
    assert (project / "output" / "2026_roczny_everyday.xlsx").exists()


# ---------------------------------------------------------------- automat w styczniu
class FakeDate(datetime.date):
    fake = datetime.date(2027, 1, 10)

    @classmethod
    def today(cls):
        return cls.fake


@pytest.fixture
def january(monkeypatch, project, pipeline):
    monkeypatch.setattr(main, "BASE", project)
    monkeypatch.setattr(main, "date", FakeDate)
    FakeDate.fake = datetime.date(2027, 1, 10)
    return pipeline


def auto(p, settings=None):
    main.maybe_auto_yearly(p["df"], p["categories"], settings or p["settings"], p["data_start"], p["data_end"])


def test_january_creates_previous_year_report_once(january, project, capsys):
    auto(january)
    path = project / "output" / "2026_roczny.xlsx"
    assert path.exists()
    assert "automatycznie utworzono" in capsys.readouterr().out
    modified = path.stat().st_mtime_ns
    auto(january)                                     # drugi raz nic nie robi
    assert path.stat().st_mtime_ns == modified


def test_other_months_do_not_create_yearly_report(january, project):
    FakeDate.fake = datetime.date(2026, 11, 24)
    auto(january)
    assert not (project / "output" / "2026_roczny.xlsx").exists()


def test_auto_yearly_can_be_disabled(january, project):
    auto(january, dict(january["settings"], auto_yearly_in_january=False))
    assert not (project / "output" / "2026_roczny.xlsx").exists()


def test_january_without_previous_year_data_does_nothing(january, project):
    FakeDate.fake = datetime.date(2028, 1, 5)         # poprzedni rok = 2027, nie ma danych
    auto(january)
    assert not list((project / "output").glob("*_roczny.xlsx"))


def test_manual_yearly_report_is_overwritten_not_duplicated(january, project):
    """Raport roczny wygenerowany wcześniej (rok niepełny) nie blokuje automatu, jeśli został usunięty,
    a istniejący plik nie jest nadpisywany przez automat."""
    create_yearly_report(project, january["df"], january["categories"], january["settings"], 2026, None,
                         january["data_start"], january["data_end"])
    path = project / "output" / "2026_roczny.xlsx"
    modified = path.stat().st_mtime_ns
    auto(january)
    assert path.stat().st_mtime_ns == modified
