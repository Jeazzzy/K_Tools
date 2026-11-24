# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
import runpy

from PyInstaller.utils.hooks import collect_all, collect_delvewheel_libs_directory


app_dir = Path(SPECPATH)
version_info = runpy.run_path(str(app_dir / "version.py"))
executable_name = version_info["EXECUTABLE_BASENAME"]

# pyogrio contains Cython extensions whose dependencies are invisible to
# PyInstaller's static analysis (most importantly _io -> _geometry). Collect
# every package module explicitly, as well as GDAL/PROJ data and the DLLs from
# the sibling pyogrio.libs directory created by delvewheel.
pyogrio_datas, pyogrio_binaries, pyogrio_hiddenimports = collect_all(
    "pyogrio",
    include_py_files=False,
    filter_submodules=lambda name: not name.startswith("pyogrio.tests"),
)
pyogrio_datas, pyogrio_binaries = collect_delvewheel_libs_directory(
    "pyogrio",
    datas=pyogrio_datas,
    binaries=pyogrio_binaries,
)
pyogrio_datas.append((str(app_dir / "icon.ico"), "."))
pyogrio_datas.append(
    (str(app_dir / "resources" / "addresses_lo.csv"), "resources")
)

a = Analysis(
    [str(app_dir / "main.py")],
    pathex=[str(app_dir)],
    binaries=pyogrio_binaries,
    datas=pyogrio_datas,
    hiddenimports=pyogrio_hiddenimports + [
        "pythoncom",
        "pywintypes",
        "win32com",
        "win32com.client",
        "win32com.client.dynamic",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[str(app_dir / "hooks" / "pyi_rth_pyogrio.py")],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=executable_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[str(app_dir / "icon.ico")],
)
