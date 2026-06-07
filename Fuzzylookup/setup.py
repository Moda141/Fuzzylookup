from setuptools import setup, find_packages

setup(
    name="fuzzylookup",
    version="0.2.0",
    description="Fuzzy matching lookup for CSV/Excel/SQL datasets (Arabic + English)",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    author="Mohamed",
    url="https://github.com/Moda141/Fuzzylookup",
    license="MIT",
    packages=find_packages(),
    python_requires=">=3.8",
    install_requires=[
        "pandas>=1.3",
        "openpyxl>=3.0",
        "rapidfuzz>=3.0",
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Text Processing :: Linguistic",
        "Natural Language :: Arabic",
    ],
    keywords="fuzzy matching arabic nlp lookup merge deduplication",
)
