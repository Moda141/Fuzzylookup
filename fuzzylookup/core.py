"""
fuzzylookup - Fuzzy matching lookup for CSV/Excel datasets
Supports Arabic and English text, with positional name-aware scoring.
"""

from __future__ import annotations

import re
import unicodedata
from pathlib import Path
from typing import Any, Optional, Union

try:
    import pandas as pd
except ImportError:
    raise ImportError("pandas is required: pip install pandas openpyxl")

try:
    from rapidfuzz import fuzz, process
except ImportError:
    raise ImportError("rapidfuzz is required: pip install rapidfuzz")


# ---------------------------------------------------------------------------
# Text normalization
# ---------------------------------------------------------------------------

def _normalize_arabic(text: str) -> str:
    """Normalize Arabic text for better matching."""
    if not isinstance(text, str):
        return str(text)
    text = unicodedata.normalize("NFC", text)
    # Remove tashkeel (diacritics)
    text = re.sub(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06DC\u06DF-\u06E4\u06E7\u06E8\u06EA-\u06ED]", "", text)
    text = re.sub(r"[أإآٱ]", "ا", text)   # Alef variants
    text = re.sub(r"ة", "ه", text)          # Teh marbuta
    text = re.sub(r"ى", "ي", text)          # Alef maqsura
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _normalize(text: str, arabic: bool = True) -> str:
    if not isinstance(text, str):
        text = str(text)
    text = text.strip().lower()
    if arabic:
        text = _normalize_arabic(text)
    return text


# ---------------------------------------------------------------------------
# Scorer aliases
# ---------------------------------------------------------------------------

SCORERS = {
    "ratio":      fuzz.ratio,
    "partial":    fuzz.partial_ratio,
    "token_sort": fuzz.token_sort_ratio,
    "token_set":  fuzz.token_set_ratio,
    "wratio":     fuzz.WRatio,
}


# ---------------------------------------------------------------------------
# Positional Name Scoring
# ---------------------------------------------------------------------------

def _tokenize(name: str) -> list[str]:
    tokens = name.strip().split()
    return tokens if tokens else [""]


def _positional_name_score(
    query: str,
    candidate: str,
    first_weight: float = 0.6,
    rest_weight: float = 0.4,
) -> float:
    """
    Compare names token-by-token in order.

    - First token vs first token  → 60% of score
    - Remaining tokens joined     → 40% of score

    Example:
        "محمد كمال" vs "محمد كمال"  → ~100
        "محمد كمال" vs "كمال محمد"  → ~50  (first tokens مختلفين)
        "محمد كمال" vs "محمد علي"   → ~65  (first token متطابق، الباقي مختلف)
    """
    q_tokens = _tokenize(query)
    c_tokens = _tokenize(candidate)

    # Single token on either side → plain ratio
    if len(q_tokens) == 1 or len(c_tokens) == 1:
        return fuzz.ratio(query, candidate)

    # First token score
    first_score = fuzz.ratio(q_tokens[0], c_tokens[0])

    # Rest tokens (join and compare)
    q_rest = " ".join(q_tokens[1:])
    c_rest = " ".join(c_tokens[1:])
    rest_score = fuzz.token_sort_ratio(q_rest, c_rest)

    return (first_score * first_weight) + (rest_score * rest_weight)


def _smart_name_score(query: str, candidate: str) -> float:
    """
    Blend positional + WRatio for best of both worlds:
    - Positional punishes wrong token order  (محمد كمال ≠ كمال محمد)
    - WRatio handles typos and partial names

    If WRatio >> positional by >15 pts → token reordering is inflating WRatio
    → apply penalty so swapped names don't score the same as correct order.
    """
    positional = _positional_name_score(query, candidate)
    wratio = fuzz.WRatio(query, candidate)

    diff = wratio - positional
    if diff > 15:
        # WRatio is inflating because of token reordering — penalize
        return (positional * 0.7) + (wratio * 0.3)
    else:
        return (positional * 0.5) + (wratio * 0.5)


# ---------------------------------------------------------------------------
# FuzzyLookup
# ---------------------------------------------------------------------------

class FuzzyLookup:
    """
    Fuzzy lookup over a CSV or Excel dataset.

    Parameters
    ----------
    source : str | Path | pd.DataFrame
        Path to CSV/Excel file, or an already-loaded DataFrame.
    column : str
        The column to match against.
    scorer : str
        Matching algorithm: ratio, partial, token_sort, token_set, wratio (default).
    normalize_arabic : bool
        Strip diacritics & normalize Arabic characters (default True).
    name_aware : bool
        Enable positional name scoring so that "محمد كمال" and "كمال محمد"
        score differently. First token is weighted higher than the rest.
        Recommended when the column contains full person names (default False).
    encoding : str
        File encoding for CSV (default 'utf-8').

    Examples
    --------
    >>> fl = FuzzyLookup("names.csv", column="name", name_aware=True)
    >>> fl.lookup("محمد كمال", top_n=3)
    """

    def __init__(
        self,
        source: Union[str, Path, "pd.DataFrame"],
        column: str,
        scorer: str = "wratio",
        normalize_arabic: bool = True,
        name_aware: bool = False,
        encoding: str = "utf-8",
    ):
        self.column = column
        self.scorer = SCORERS.get(scorer, fuzz.WRatio)
        self.normalize_arabic = normalize_arabic
        self.name_aware = name_aware

        # Load data
        if isinstance(source, (str, Path)):
            path = Path(source)
            if path.suffix.lower() in {".xlsx", ".xls"}:
                self._df = pd.read_excel(path)
            else:
                self._df = pd.read_csv(path, encoding=encoding)
        elif isinstance(source, pd.DataFrame):
            self._df = source.copy()
        else:
            raise TypeError("source must be a file path or a pandas DataFrame")

        if column not in self._df.columns:
            raise ValueError(
                f"Column '{column}' not found. Available: {list(self._df.columns)}"
            )

        self._choices: list[str] = (
            self._df[column].fillna("").astype(str).tolist()
        )
        self._normalized_choices: list[str] = [
            _normalize(c, arabic=self.normalize_arabic) for c in self._choices
        ]

    # ------------------------------------------------------------------
    # Internal scoring
    # ------------------------------------------------------------------

    def _score(self, query: str, candidate: str) -> float:
        """Score a query against a candidate string."""
        if self.name_aware:
            return _smart_name_score(query, candidate)
        return self.scorer(query, candidate)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def lookup(
        self,
        query: str,
        top_n: int = 5,
        min_score: float = 0.0,
        columns: Optional[list[str]] = None,
    ) -> list[dict[str, Any]]:
        """
        Return the top-N best matches for *query*.

        Parameters
        ----------
        query : str       — search string
        top_n : int       — max results (default 5)
        min_score : float — minimum score 0–100 (default 0)
        columns : list    — which columns to return (default: all)
        """
        norm_query = _normalize(query, arabic=self.normalize_arabic)
        cols = columns or list(self._df.columns)

        if self.name_aware:
            # Manual scoring loop (rapidfuzz process doesn't support custom scorers easily)
            scored = [
                (i, self._score(norm_query, cand))
                for i, cand in enumerate(self._normalized_choices)
            ]
            scored = [(i, s) for i, s in scored if s >= min_score]
            scored.sort(key=lambda x: x[1], reverse=True)
            scored = scored[:top_n]

            results = []
            for idx, score in scored:
                row = self._df.iloc[idx][cols].to_dict()
                row["score"] = round(score, 2)
                row["_index"] = int(idx)
                results.append(row)
        else:
            matches = process.extract(
                norm_query,
                self._normalized_choices,
                scorer=self.scorer,
                limit=top_n,
                score_cutoff=min_score,
            )
            results = []
            for _matched_str, score, idx in matches:
                row = self._df.iloc[idx][cols].to_dict()
                row["score"] = round(score, 2)
                row["_index"] = int(idx)
                results.append(row)
            results.sort(key=lambda r: r["score"], reverse=True)

        return results

    def lookup_best(
        self,
        query: str,
        min_score: float = 0.0,
        columns: Optional[list[str]] = None,
    ) -> Optional[dict[str, Any]]:
        """Return only the single best match, or None if below min_score."""
        results = self.lookup(query, top_n=1, min_score=min_score, columns=columns)
        return results[0] if results else None

    def lookup_many(
        self,
        queries: list[str],
        top_n: int = 1,
        min_score: float = 0.0,
        columns: Optional[list[str]] = None,
    ) -> dict[str, list[dict[str, Any]]]:
        """Batch lookup for multiple queries."""
        return {
            q: self.lookup(q, top_n=top_n, min_score=min_score, columns=columns)
            for q in queries
        }

    # ------------------------------------------------------------------
    # Convenience
    # ------------------------------------------------------------------

    @property
    def columns(self) -> list[str]:
        return list(self._df.columns)

    @property
    def shape(self) -> tuple[int, int]:
        return self._df.shape

    def __repr__(self) -> str:
        mode = "name_aware" if self.name_aware else self.scorer.__name__
        return (
            f"FuzzyLookup(column='{self.column}', "
            f"rows={self._df.shape[0]}, "
            f"mode='{mode}')"
        )
