from setuptools import setup

APP = ['main.py']
OPTIONS = {'includes': ['pyautogui', 'python-docx']}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)