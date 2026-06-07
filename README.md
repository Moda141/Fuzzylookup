# FuzzyLookup

**Fuzzy string matching for CSV, Excel, and SQL datasets — built for Arabic and English names.**

```bash
pip install fuzzylookup
```

---

## Features

- Arabic-aware normalization — strips diacritics, unifies alef variants, teh marbuta, alef maqsura
- Positional name scoring — `"محمد كمال"` and `"كمال محمد"` score differently (`name_aware=True`)
- Multiple sources — CSV, Excel, Parquet, Feather, pandas DataFrame, SQL (sqlite3 / SQLAlchemy)
- `fuzzy_merge()` — fuzzy join between two DataFrames, like `pd.merge()` with a score threshold
- ~10x faster on large datasets via a blocking index (first-token prefix bucketing)
- Five scorers: `ratio`, `partial`, `token_sort`, `token_set`, `wratio`

---

## Quick Start

### Lookup from a file

```python
from fuzzylookup import FuzzyLookup

fl = FuzzyLookup("customers.csv", column="name", name_aware=True)

# Single lookup
fl.lookup("محمد كمال", top_n=3, min_score=70)
# [{'name': 'محمد كمال عبد الرحمن', 'score': 83.4, '_index': 0}, ...]

# Best match only
fl.lookup_best("احمد سعيد", min_score=70)

# Batch lookup
fl.lookup_many(["محمد", "أحمد", "علي"], top_n=1, min_score=70)
```

### From SQL

```python
import sqlite3
from fuzzylookup import FuzzyLookup

con = sqlite3.connect("customers.db")
fl = FuzzyLookup(
    source=None,
    column="name",
    connection=con,
    sql_query="SELECT * FROM customers WHERE active = 1",
    name_aware=True,
)
fl.lookup("محمد كمال", top_n=3)
```

### Fuzzy merge — join two DataFrames

```python
from fuzzylookup import fuzzy_merge

result = fuzzy_merge(
    crm_df, master_df,
    left_on="cust_name",
    right_on="name",
    min_score=80,
    name_aware=True,
)
```

Or from a `FuzzyLookup` instance — uses the blocking index automatically:

```python
master = FuzzyLookup("master.csv", column="name", name_aware=True)

result = master.merge(
    crm_df,
    other_on="cust_name",
    min_score=80,
    return_columns=["account_no", "cust_name"],
)
```

---

## API Reference

### `FuzzyLookup(source, column, ...)`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `source` | str / Path / DataFrame / None | — | File path, DataFrame, or None for SQL |
| `column` | str | — | Column to match against |
| `scorer` | str | `"wratio"` | `ratio` / `partial` / `token_sort` / `token_set` / `wratio` |
| `normalize_arabic` | bool | `True` | Strip diacritics, normalize alef/teh marbuta/alef maqsura |
| `name_aware` | bool | `False` | Positional name scoring |
| `encoding` | str | `"utf-8"` | CSV encoding |
| `sql_query` | str | `None` | SQL SELECT (required when `connection=` is used) |
| `connection` | connection | `None` | sqlite3 or SQLAlchemy connection |
| `use_blocking` | bool | `True` | Enable blocking index (~10x speedup) |
| `block_prefix_len` | int | `2` | Prefix length for blocking buckets |

### `.lookup(query, top_n, min_score, columns)`

Returns a list of dicts, each with row data + `score` (0–100) + `_index`.

### `.lookup_best(query, min_score, columns)`

Returns the single best match dict, or `None` if below `min_score`.

### `.lookup_many(queries, top_n, min_score, columns)`

Batch lookup — returns `dict[query → list[match]]`.

### `.merge(other, other_on, min_score, top_n, return_columns, return_score)`

Fuzzy-join the reference dataset against `other` DataFrame.

---

### `fuzzy_merge(left, right, left_on, right_on, ...)`

| Parameter | Default | Description |
|-----------|---------|-------------|
| `min_score` | `80.0` | Minimum score threshold |
| `scorer` | `"wratio"` | Matching algorithm |
| `normalize_arabic` | `True` | Arabic normalization |
| `name_aware` | `False` | Positional scoring |
| `top_n` | `1` | Top N matches per left row |
| `suffixes` | `("_left","_right")` | Suffix for overlapping columns |
| `return_score` | `True` | Add `fuzzy_score` column |
| `use_blocking` | `True` | Enable blocking index |

---

## Arabic Name Matching

```python
fl = FuzzyLookup("names.csv", column="name", name_aware=True)

# Normalized automatically before matching:
# أحمد  →  احمد   (alef variants)
# فاطمة →  فاطمه  (teh marbuta)
# موسى  →  موسي   (alef maqsura)
# مُحَمَّد → محمد   (diacritics removed)

# Positional scoring:
# "محمد كمال" vs "محمد كمال"  →  100   ✓ exact
# "محمد كمال" vs "كمال محمد"  →  ~55   ✗ wrong order penalized
# "محمد كمال" vs "محمد علي"   →  ~65   ~ first token matches
```

---

## Performance

The blocking index reduces the candidate pool per query from the full dataset
to ~10% by bucketing on the first 2 characters of the first name token.

| Dataset | Without blocking | With blocking | Speedup |
|---------|-----------------|---------------|---------|
| 500 queries × 10,000 rows | 26s | 2.1s | **12x** |
| 2,000 queries × 10,000 rows | ~104s | ~8s | **~12x** |

Disable if first tokens are very inconsistent: `use_blocking=False`

---

## Requirements

- Python ≥ 3.8
- pandas ≥ 1.3
- rapidfuzz ≥ 3.0
- openpyxl ≥ 3.0

---

## License

MIT
