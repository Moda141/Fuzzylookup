from setuptools import setup, find_packages

setup(
    name="fuzzylookup",
    version="0.1.0",
    description="Fuzzy matching lookup for CSV/Excel datasets (Arabic + English)",
    license="MIT",  
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.3",
        "openpyxl>=3.0",
        "rapidfuzz>=3.0",
    ],
)
