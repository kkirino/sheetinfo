from setuptools import setup

setup(
    name="sheetinfo",
    version="0.0.1",
    description="get sheet information in xls/xlsx files without MS Office application",
    license="MIT",
    author="kkirino",
    install_requires=["xlrd", "openpyxl"],
    entry_points={"console_scripts": ["sheetinfo = sheetinfo:main"]},
)
