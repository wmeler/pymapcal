# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


project_root = Path.cwd().resolve()
build_support_dir = project_root / "build_support"
datas = [
    (str(project_root / "README.md"), "."),
    (str(project_root / "README.pl.md"), "."),
    (str(project_root / "README.en.md"), "."),
]
binaries = []

for candidate_name in ("imgkap.exe", "imgkap", "FreeImage.dll"):
    candidate = build_support_dir / candidate_name
    if candidate.exists():
        binaries.append((str(candidate), "."))


main_analysis = Analysis(
    ["main.py"],
    pathex=[str(project_root)],
    binaries=binaries,
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
main_pyz = PYZ(main_analysis.pure)

main_exe = EXE(
    main_pyz,
    main_analysis.scripts,
    main_analysis.binaries,
    main_analysis.datas,
    [],
    name="pymapcal",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

cli_analysis = Analysis(
    ["download_wz_pdfs.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
cli_pyz = PYZ(cli_analysis.pure)

cli_exe = EXE(
    cli_pyz,
    cli_analysis.scripts,
    cli_analysis.binaries,
    cli_analysis.datas,
    [],
    name="download_wz_pdfs",
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
