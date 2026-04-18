from PyInstaller.utils.hooks import collect_all
import os
curr_dir = os.getcwd()
datas = [
    ('assets', 'assets/'), 
]

# Zero-Installation: Add tesseract folder if it exists in root
if os.path.exists(os.path.join(curr_dir, 'tesseract')):
    datas += [('tesseract', 'tesseract/')]

binaries = []
hiddenimports = []
tmp_ret = collect_all('customtkinter')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('tkinterdnd2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
# Heavy AI modules excluded to keep EXE size manageable
# tmp_ret = collect_all('easyocr')
# datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
# tmp_ret = collect_all('torch')
# datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
# tmp_ret = collect_all('torchvision')
# datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

tmp_ret = collect_all('pyhanko')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('pypdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('reportlab')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# Explicitly add modules imported inside obfuscated folders
hiddenimports += [
    'fitz',
    'pytesseract',
    'PIL',
    'PIL.Image',
    'numpy',
    'tkinter.colorchooser',
    'tkinter.filedialog',
    'tkinter.messagebox',
    'tkinter.simpledialog',
    'tkinter.commondialog',
    'oscrypto',
    'cryptography',
    'asn1crypto',
    'pyhanko_certvalidator',
    'matplotlib',
    'docx'
]

# Windows-only hidden imports
if os.name == 'nt':
    hiddenimports += [
        'win32crypt',
        'win32api',
        'win32timezone',
        'win32ctypes',
        'pywintypes',
        'pythoncom'
    ]



a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'torchvision', 'easyocr', 'PIL.ImageQt', 'notebook', 'ui.views.bank_statement_view'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CA_Office_PDF_Utility',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[os.path.join('assets', 'logo.png')],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='CA_Office_PDF_Utility',
)
