from setuptools import setup

setup(
    # Application name:
    name="SLI-Tools",

    # Version number:
    version="0.1.0",

    # Application author details:
    author="Mark Rutherford",
    author_email="mcr5801@g.rit.edu",

    # Details
    license="LICENSE",
    description="Easily unzip and grade student assignments.",

    long_description=open("README.md").read(),

    # Dependent packages (distributions)
    install_requires=[
        "openpyxl",
    ],
)
