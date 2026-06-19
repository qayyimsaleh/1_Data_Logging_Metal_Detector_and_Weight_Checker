# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['web_report.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static',    'static'),
    ],
    hiddenimports=[
        'flask', 'werkzeug', 'jinja2', 'markupsafe', 'click',
        'itsdangerous', 'pyodbc', 'openpyxl',
        'reportlab', 'reportlab.graphics', 'shared_config',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # heavy unused scientific stack
        'numpy', 'pandas', 'scipy', 'matplotlib',
        'PIL', 'Pillow', 'cv2', 'sklearn', 'skimage',
        # GUI toolkits not needed
        'tkinter', '_tkinter', 'tkinter.ttk',
        'PyQt5', 'PyQt6', 'PySide2', 'PySide6', 'wx',
        # jupyter / IPython
        'IPython', 'jupyter', 'notebook', 'ipykernel',
        'ipywidgets', 'nbformat', 'nbconvert',
        # other heavy unused libs
        'sqlalchemy', 'pydantic', 'aiohttp', 'tornado',
        'boto3', 'botocore', 'cryptography', 'paramiko',
        'docutils', 'sphinx', 'setuptools', 'pkg_resources',
        # test frameworks
        'unittest', 'pytest', 'nose',
        # unused stdlib
        'curses', 'lib2to3', 'test',
    ],
    noarchive=False,
    optimize=2,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='web_report',
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
