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
        'itsdangerous', 'pyodbc', 'openpyxl', 'shared_config',
        # reportlab — must list submodules explicitly (dynamic imports)
        'reportlab',
        'reportlab.lib', 'reportlab.lib.colors', 'reportlab.lib.enums',
        'reportlab.lib.pagesizes', 'reportlab.lib.styles', 'reportlab.lib.units',
        'reportlab.lib.utils', 'reportlab.lib.fonts', 'reportlab.lib.geomutils',
        'reportlab.platypus', 'reportlab.platypus.tables',
        'reportlab.platypus.paragraph', 'reportlab.platypus.flowables',
        'reportlab.platypus.doctemplate', 'reportlab.platypus.frames',
        'reportlab.pdfgen', 'reportlab.pdfgen.canvas',
        'reportlab.pdfbase', 'reportlab.pdfbase.pdfmetrics',
        'reportlab.pdfbase.ttfonts', 'reportlab.pdfbase._fontdata',
        'reportlab.graphics', 'reportlab.graphics.shapes',
        'reportlab.rl_config', 'reportlab.rl_settings',
        # Pillow — required by reportlab
        'PIL', 'PIL.Image', 'PIL.ImageFont', 'PIL.ImageDraw',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # heavy unused scientific stack
        'numpy', 'pandas', 'scipy', 'matplotlib',
        'cv2', 'sklearn', 'skimage',
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
