"""
py2app/py2exe build script for stats.

Will automatically ensure that all build prerequisites are available
via ez_setup

Usage (Mac OS X):
    python3 setup.py py2app

Usage (Windows):
    python3 setup.py py2exe
"""
import sys
from setuptools import setup

APP = ['stats.py']
DATA_FILES = []
OPTIONS = {}

if sys.platform == 'darwin':
    setup(
        app=APP,
        data_files=DATA_FILES,
        options={'py2app': OPTIONS},
        setup_requires=['py2app'],
    )
elif sys.platform == 'win32':
    setup(
        app=APP,
        data_files=DATA_FILES,
        options={'py2exe': OPTIONS},
        setup_requires=['py2exe'],
    )

