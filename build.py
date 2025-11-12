import PyInstaller.__main__
import os
import sys
import shutil

# Central app version (keep in sync with manifest and main.py)
VERSION = "2.1.0"

def _write_version_file(path: str, version: str):
    """Generate a Windows version info file for PyInstaller."""
    parts = [int(p) for p in (version.split('.'))]
    while len(parts) < 4:
        parts.append(0)
    filevers = tuple(parts[:4])
    prodvers = tuple(parts[:4])
    content = f"""
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers={filevers},
    prodvers={prodvers},
    mask=0x3f,
    flags=0x0,
    OS=0x4,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', 'Advantage Conveyor'),
        StringStruct('FileDescription', 'Advantage Multi-Vendor Price Scraper'),
        StringStruct('FileVersion', '{version}'),
        StringStruct('InternalName', 'AdvantageScraper'),
        StringStruct('OriginalFilename', 'AdvantageScraper.exe'),
        StringStruct('ProductName', 'Advantage Price Scraper'),
        StringStruct('ProductVersion', '{version}')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""
    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)

def build_exe():
    """Build standalone executable"""
    # Clean previous builds
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    # Ensure version file exists and matches VERSION
    version_file = 'version_info.txt'
    _write_version_file(version_file, VERSION)

    data_sep = ';' if os.name == 'nt' else ':'
    add_icon_png = f"--add-data=icon.png{data_sep}."
    add_icon_ico = f"--add-data=icon.ico{data_sep}."

    args = [
        'main.py',
        '--name=AdvantageScraper',
        '--onefile',
        '--windowed',
        '--icon=icon.ico',
        '--clean',
        f'--version-file={version_file}',
        add_icon_png,
        add_icon_ico,
        # Hidden imports (some are defensive to avoid rare hook gaps)
        '--hidden-import=flask',
        '--hidden-import=flask_cors',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.font',
        '--hidden-import=socket',
        '--hidden-import=webbrowser',
        '--hidden-import=threading',
        '--hidden-import=queue',
        '--hidden-import=et_xmlfile',
        '--hidden-import=jdcal',
        # Collect packages with data and submodules
        '--collect-all=flask',
        '--collect-all=flask_cors',
        '--collect-all=openpyxl',
        '--noconfirm'
    ]

    print('Running PyInstaller with args:')
    for a in args:
        print(' ', a)
    PyInstaller.__main__.run(args)

    print("\n" + "="*50)
    print("Build complete!")
    print(f"Executable location: {os.path.abspath('dist/AdvantageScraper.exe')}")
    print("="*50)

if __name__ == "__main__":
    build_exe()
