# FuzzyLookup Documentation

## Overview
`fuzzylookup` (version 0.1.0) is a Python package designed for fuzzy matching and lookups against CSV or Excel datasets [cite: 1, 2, 3]. It explicitly supports both Arabic and English text [cite: 2, 3]. Under the hood, it leverages `pandas` for data handling and `rapidfuzz` for high-performance string matching [cite: 2, 3].

## Installation
The package requires Python 3.8 or higher [cite: 3]. 
Dependencies include:
* `pandas>=1.3` [cite: 3]
* `openpyxl>=3.0` (for Excel support) [cite: 3]
* `rapidfuzz>=3.0` [cite: 3]

## Core Features

### 1. Arabic Text Normalization
By default, `fuzzylookup` applies normalization to Arabic text (`normalize_arabic=True`) [cite: 2]. This process improves match quality by:
* Removing *tashkeel* (diacritics) [cite: 2].
* Normalizing Alef variants (أ, إ, آ, ٱ) to a standard Alef (ا) [cite: 2].
* Normalizing Teh marbuta (ة) to Heh (ه) [cite: 2].
* Normalizing Alef maqsura (ى) to Yeh (ي) [cite: 2].

### 2. Positional Name-Aware Scoring
A standout feature is the `name_aware` scoring algorithm [cite: 2]. Standard fuzzy matching algorithms might score "محمد كمال" and "كمال محمد" similarly, but `fuzzylookup` can punish wrong token orders when `name_aware=True` is enabled [cite: 2]. 
* It compares names token-by-token in order [cite: 2].
* The first token accounts for 60% of the score, and the remaining tokens account for 40% [cite: 2].
* It blends this positional score with `WRatio` to handle both typos and correct word order gracefully [cite: 2].
* If WRatio exceeds the positional score by more than 15 points, it indicates token reordering is inflating the score, triggering a penalty [cite: 2].

## API Reference

### `FuzzyLookup` Class
The primary entry point is the `FuzzyLookup` class [cite: 1, 2].

**Parameters:**
* `source` (str, Path, or pd.DataFrame): Path to the CSV/Excel file or an existing DataFrame [cite: 2].
* `column` (str): The column name to match against [cite: 2].
* `scorer` (str): Matching algorithm to use (`ratio`, `partial`, `token_sort`, `token_set`, `wratio`). Defaults to `wratio` [cite: 2].
* `normalize_arabic` (bool): Whether to strip diacritics and normalize characters. Defaults to `True` [cite: 2].
* `name_aware` (bool): Enables positional name scoring. Recommended for full person names. Defaults to `False` [cite: 2].
* `encoding` (str): File encoding for CSVs. Defaults to `utf-8` [cite: 2].

### Methods

* `lookup(query: str, top_n: int = 5, min_score: float = 0.0, columns: list = None)`: Returns the top `N` best matches for the search string [cite: 2]. Returns a list of dictionaries with row data, the calculated `score`, and `_index` [cite: 2].
* `lookup_best(query: str, min_score: float = 0.0, columns: list = None)`: Returns only the single best match, or `None` if no match meets the minimum score [cite: 2].
* `lookup_many(queries: list, top_n: int = 1, min_score: float = 0.0, columns: list = None)`: Batch lookup for multiple queries. Returns a dictionary mapping each query to its list of results [cite: 2].

### Properties
* `columns`: Returns a list of the columns available in the loaded dataframe [cite: 2].
* `shape`: Returns a tuple representing the shape of the underlying dataframe [cite: 2].

## Workflow Example

Below is a typical workflow demonstrating how to instantiate the lookup object and perform queries [cite: 2]:

```python
from fuzzylookup import FuzzyLookup

# 1. Initialize the lookup instance with a dataset
fl = FuzzyLookup("names.csv", column="name", name_aware=True)

# 2. Perform a standard lookup for the top 3 matches
results = fl.lookup("محمد كمال", top_n=3)

# 3. Perform a strict lookup requiring a high match score
best_match = fl.lookup_best("كمال محمد", min_score=85.0)

# 4. Batch processing multiple names
batch_results = fl.lookup_many(["أحمد علي", "محمود حسن"], top_n=1)
```
