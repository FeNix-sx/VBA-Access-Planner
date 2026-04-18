# -*- coding: utf-8 -*-
"""
Импорт новых стандартных модулей из DB_VBA/VBA_Modules в Microsoft Access.

Целевая база по умолчанию: D:\\Planner\\Planner.accdb

Требования:
  pip install pywin32

В Access: File → Options → Trust Center → Trust Center Settings → Macro Settings
  → включить «Trust access to the VBA project object model».

Закройте базу в других экземплярах Access перед запуском.

Если ошибка «требует повышения» (HRESULT 0x800702E4): запустите терминал **от имени администратора**
и снова выполните скрипт, либо импортируйте .bas вручную через VBE (File → Import).
"""
import argparse
import sys
from pathlib import Path
from typing import Tuple

# vbext_ct_StdModule — стандартный модуль (VBE)
VBEXT_CT_STDMODULE = 1

NEW_MODULE_NAMES: Tuple[str, ...] = (
    "modProjectAnalysisCore",
    "modProjectAnalysisDeep",
    "modProjectAnalysisTables",
    "modProjectAnalysisVbe",
    "modFrmDemoConstants",
    "modFrmDemoLogic",
    "modDailyPlannerSettings",
    "modDailyPlannerGames",
)


def _repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _modules_dir() -> Path:
    return _repo_root() / "DB_VBA" / "VBA_Modules"


def _remove_component(project, name: str) -> None:
    try:
        project.VBComponents.Remove(project.VBComponents(name))
    except Exception:
        pass


def _add_module_from_file(project, name: str, bas_path: Path) -> None:
    _remove_component(project, name)
    comp = project.VBComponents.Add(VBEXT_CT_STDMODULE)
    comp.Name = name
    comp.CodeModule.AddFromFile(str(bas_path))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--db",
        type=Path,
        default=Path(r"D:\Planner\Planner.accdb"),
        help="Путь к .accdb (frontend с VBA-проектом)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Только проверить наличие .bas, не открывать Access",
    )
    args = parser.parse_args()
    db_path: Path = args.db
    modules_dir = _modules_dir()

    missing = [n for n in NEW_MODULE_NAMES if not (modules_dir / f"{n}.bas").exists()]
    if missing:
        print("Не найдены файлы:", file=sys.stderr)
        for n in missing:
            print(f"  {modules_dir / (n + '.bas')}", file=sys.stderr)
        return 2

    if args.dry_run:
        print("dry-run OK, файлы на месте:", len(NEW_MODULE_NAMES))
        for n in NEW_MODULE_NAMES:
            print(" ", modules_dir / f"{n}.bas")
        return 0

    try:
        import pythoncom  # type: ignore[import-untyped]
        from win32com.client import DispatchEx  # type: ignore[import-untyped]
    except ImportError:
        print("Установите: pip install pywin32", file=sys.stderr)
        return 3

    if not db_path.is_file():
        print("База не найдена:", db_path, file=sys.stderr)
        return 4

    co_initialized = False
    pythoncom.CoInitialize()
    co_initialized = True
    access = None
    try:
        access = DispatchEx("Access.Application")
        access.Visible = True
        access.OpenCurrentDatabase(str(db_path.resolve()), False)
        project = access.VBE.ActiveVBProject

        for name in NEW_MODULE_NAMES:
            bas_path = modules_dir / f"{name}.bas"
            print("Импорт", name, "...")
            _add_module_from_file(project, name, bas_path)
            print("  OK")

        access.CloseCurrentDatabase()
        access.Quit()
        access = None
        print("Готово. Проверьте проект в VBE и сохраните базу (Ctrl+S).")
        return 0
    except Exception as e:
        print("Ошибка COM/Access:", e, file=sys.stderr)
        hr = None
        if e.args and isinstance(e.args[0], int):
            hr = e.args[0]
        # CO_E_ELEVATION_REQUIRED / «Запрошенная операция требует повышения»
        if hr == -2147024156:
            print(
                "Это UAC: COM не смог запустить Access с текущими правами.",
                "Запустите PowerShell/cmd «От имени администратора» и повторите команду,",
                "или откройте Access вручную и импортируйте модули через VBE.",
                file=sys.stderr,
            )
        else:
            print(
                "Проверьте: установлен Access, путь к .accdb, база не открыта в другом окне,",
                "включён «Trust access to the VBA project object model».",
                file=sys.stderr,
            )
        return 1
    finally:
        if access is not None:
            try:
                access.Quit()
            except Exception:
                pass
        if co_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
