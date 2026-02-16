# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for JakeGResumeBuilder macOS App
Builds a standalone .app bundle with all dependencies included
"""
a = Analysis(
    ['JakeGResumeBuilder_GUI_v.1.0.2_macOS.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'pyspellchecker',
        'reportlab',
        'docx',
        'PIL',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludedimports=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='JakeGResumeBuilder',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

app = BUNDLE(
    exe,
    name='JakeGResumeBuilder v.1.0.2.app',
    icon=None,
    bundle_identifier='com.jakegresumebuilder.app',
    info_plist={
        'NSPrincipalClass': 'NSApplication',
        'NSHighResolutionCapable': 'True',
        'CFBundleExecutable': 'JakeGResumeBuilder',
        'CFBundleName': 'JakeGResumeBuilder',
        'CFBundleVersion': '1.0.2',
        'CFBundleShortVersionString': '1.0.2',
        'CFBundleIdentifier': 'com.jakegresumebuilder.app',
    },
)
