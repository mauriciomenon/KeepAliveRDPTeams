# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['..\\keep-alive-app.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['PyQt5', 'PySide6', 'qtpy', 'tkinter', 'matplotlib', 'numpy', 'pandas', 'scipy', 'PIL', 'cv2', 'tensorflow', 'torch', 'django', 'flask', 'requests', 'urllib3', 'cryptography', 'sqlite3', 'unittest', 'doctest', 'pdb', 'turtle', 'PyQt6.QtWebEngine', 'PyQt6.QtWebEngineWidgets', 'PyQt6.QtWebEngineCore', 'PyQt6.QtMultimedia', 'PyQt6.QtMultimediaWidgets', 'PyQt6.QtOpenGL', 'PyQt6.QtOpenGLWidgets', 'PyQt6.QtSql', 'PyQt6.QtTest', 'PyQt6.QtDesigner'],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [('O', None, 'OPTION'), ('O', None, 'OPTION')],
    name='KeepAliveRDP_Icon',
    debug=False,
    bootloader_ignore_signals=False,
    strip=True,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['simple_icon.ico'],
)
