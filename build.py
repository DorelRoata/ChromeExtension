import PyInstaller.__main__
import os
import shutil

def build_exe():
    """Build standalone executable"""
    
    # Clean previous builds
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    if os.path.exists('build'):
        shutil.rmtree('build')
    
    PyInstaller.__main__.run([
        'main.py',
        '--name=PriceScraper',
        '--onefile',
        '--windowed',
        '--icon=icon.png',
        '--hidden-import=flask',
        '--hidden-import=flask_cors',
        '--hidden-import=openpyxl',
        '--hidden-import=tkinter',
        '--hidden-import=tkinter.font',
        '--hidden-import=socket',
        '--hidden-import=webbrowser',
        '--hidden-import=threading',
        '--hidden-import=queue',
        '--collect-all=flask',
        '--collect-all=flask_cors',
        '--noconfirm'
    ])
    
    print("\n" + "="*50)
    print("Build complete!")
    print(f"Executable location: {os.path.abspath('dist/PriceScraper.exe')}")
    print("="*50)

if __name__ == "__main__":
    build_exe()