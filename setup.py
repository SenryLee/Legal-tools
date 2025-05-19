from setuptools import setup

APP = ['律师函测试.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'includes': ['pandas', 'docx', 'tkinter', 'openpyxl'],
    'packages': ['pandas', 'docx', 'tkinter', 'openpyxl'],
    'plist': {
        'CFBundleName': 'LawLetterGen',
        'CFBundleDisplayName': 'LawLetterGen',
        'CFBundleIdentifier': 'com.senry.lawlettergen',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
    }
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)