# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_submodules


server_dir = Path(SPECPATH)
project_dir = server_dir.parent
algorithm_dir = server_dir / "pycatia_regulation_reflection_point_detection"

datas = [
    (str(project_dir / "web"), "web"),
    (str(algorithm_dir / "main.py"), "pycatia_regulation_reflection_point_detection"),
    (str(algorithm_dir / "resources"), "pycatia_regulation_reflection_point_detection/resources"),
]

for catpart in algorithm_dir.glob("*.CATPart"):
    datas.append((str(catpart), "pycatia_regulation_reflection_point_detection"))

hiddenimports = (
    collect_submodules("uvicorn")
    + collect_submodules("win32com")
    + [
        "run_algorithm_worker",
        "pythoncom",
        "pywintypes",
        "win32timezone",
        "win32com.client",
        "multipart",
    ]
)

a = Analysis(
    [str(server_dir / "server.py")],
    pathex=[str(server_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name="rearview_server",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
