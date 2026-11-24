"""Smoke-test native and packaged dependencies in Python and the frozen EXE."""

from __future__ import annotations

import importlib
import os
from pathlib import Path
import sys
import tempfile
import traceback


PYOGRIO_NATIVE_MODULES = (
    "pyogrio._err",
    "pyogrio._geometry",
    "pyogrio._io",
    "pyogrio._ogr",
    "pyogrio._vsi",
)


def _write_report(report_path: str | os.PathLike[str] | None, text: str) -> None:
    if report_path:
        path = Path(report_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")

    stream = getattr(sys, "stdout", None)
    if stream is not None:
        stream.write(text)
        stream.flush()


def run(report_path: str | os.PathLike[str] | None = None) -> int:
    """Return zero only when GIS and Word integration dependencies are packaged."""

    try:
        import geopandas as gpd
        import pyogrio
        from shapely.geometry import Point

        for module_name in PYOGRIO_NATIVE_MODULES:
            importlib.import_module(module_name)
        importlib.import_module("pythoncom")
        importlib.import_module("win32com.client")

        address_file = Path(__file__).resolve().parent / "resources" / "addresses_lo.csv"
        if not address_file.is_file() or address_file.stat().st_size == 0:
            raise RuntimeError("The Lenoblast address hint resource is missing")

        with tempfile.TemporaryDirectory(prefix="k_tools_gis_test_") as temp_dir:
            tab_path = Path(temp_dir) / "probe.tab"
            source = gpd.GeoDataFrame(
                {"name": ["probe"]},
                geometry=[Point(37.6176, 55.7558)],
                crs="EPSG:4326",
            )
            source.to_file(
                tab_path,
                driver="MapInfo File",
                engine="pyogrio",
                layer_options={"BOUNDS": "-1000000,-1000000,19000000,19000000"},
            )
            restored = gpd.read_file(tab_path, engine="pyogrio")

            if len(restored) != 1 or restored.loc[0, "name"] != "probe":
                raise RuntimeError("MapInfo TAB round-trip returned unexpected data")

        report = (
            "OK: GIS self-test passed\n"
            "OK: Word COM dependencies and address hints are packaged\n"
            f"pyogrio={pyogrio.__version__}\n"
            f"GDAL={pyogrio.__gdal_version_string__}\n"
        )
        _write_report(report_path, report)
        return 0
    except BaseException:
        _write_report(report_path, "FAILED: GIS self-test\n" + traceback.format_exc())
        return 1


if __name__ == "__main__":
    output = sys.argv[1] if len(sys.argv) > 1 else None
    raise SystemExit(run(output))
