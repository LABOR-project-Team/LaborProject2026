# build_exe.py
import PyInstaller.__main__
import sys
import os

if __name__ == "__main__":
    PyInstaller.__main__.run([
        'run.py',  # Use run.py as entry point
        '--onefile',
        '--add-data=app/templates;app/templates',
        '--add-data=app/static;app/static',
        '--add-data=app/data;app/data',
        '--add-data=Loopbaan onderzoek 5.0.xlsx;.',
        '--hidden-import=uvicorn',
        '--hidden-import=uvicorn.lifespan.on',
        '--hidden-import=jinja2',
        '--hidden-import=docx',
        '--hidden-import=openpyxl',
        '--name=LoopbaanOnderzoek',
        '--clean',
        '--noconfirm'
    ])