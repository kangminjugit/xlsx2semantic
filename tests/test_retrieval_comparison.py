"""Retrieval performance comparison: semantic XML vs raw XML.

Parses a real XLSX file both ways, then measures how well each
representation retrieves relevant chunks for natural language queries
using TF-IDF cosine similarity.

Requires: scikit-learn (pip install scikit-learn)
"""

from __future__ import annotations

import re
import textwrap
from pathlib import Path

import pytest

from xlsx2semantic import parse_file

# ── Config ──

XLSX_PATH = Path(__file__).parent / "docs" / "SCH-0005-Overall-Enrollment.xlsx"
CHUNK_MAX_CHARS = 800  # max characters per chunk

# Ground-truth query → expected answer pairs.
# Each query simulates a user asking a natural language question.
# `expected` is a substring that MUST appear in the retrieved chunk for it to count as a hit.
QUERIES = [
    {
        "query": "How many students are enrolled in Alabama?",
        "expected": "Alabama",
    },
    {
        "query": "What is the total enrollment in California?",
        "expected": "California",
    },
    {
        "query": "How many Asian students are in New York?",
        "expected": "New York",
    },
    {
        "query": "What percentage of students in Texas are Hispanic?",
        "expected": "Texas",
    },
    {
        "query": "How many Black students are enrolled in Georgia?",
        "expected": "Georgia",
    },
    {
        "query": "What is the total number of students in Florida?",
        "expected": "Florida",
    },
    {
        "query": "English language learners in Illinois",
        "expected": "Illinois",
    },
    {
        "query": "Students with disabilities in Ohio",
        "expected": "Ohio",
    },
]

TOP_K = 3  # number of chunks to retrieve per query


# ── Helpers ──


def _chunk_by_records(xml: str, max_chars: int = CHUNK_MAX_CHARS) -> list[str]:
    """Split XML into chunks, one per <record> element.

    If a record exceeds max_chars, it is still kept as a single chunk.
    Adjacent small records are merged up to max_chars.
    """
    # Extract schema/title context to prepend to each chunk
    context_parts: list[str] = []
    title_match = re.search(r"<title>.*?</title>", xml, re.DOTALL)
    if title_match:
        context_parts.append(title_match.group())
    schema_match = re.search(r"<schema>.*?</schema>", xml, re.DOTALL)
    if schema_match:
        context_parts.append(schema_match.group())
    context = "\n".join(context_parts)

    # Split by records
    records = re.findall(r"<record\b.*?</record>", xml, re.DOTALL)
    if not records:
        # Fallback: fixed-size chunking
        return _chunk_fixed(xml, max_chars)

    chunks: list[str] = []
    current = context + "\n" if context else ""

    for rec in records:
        if len(current) + len(rec) > max_chars and current.strip():
            chunks.append(current.strip())
            current = (context + "\n") if context else ""
        current += rec + "\n"

    if current.strip():
        chunks.append(current.strip())

    return chunks


def _chunk_fixed(xml: str, max_chars: int = CHUNK_MAX_CHARS) -> list[str]:
    """Fallback: split XML into fixed-size chunks by lines."""
    lines = xml.splitlines()
    chunks: list[str] = []
    current: list[str] = []
    current_len = 0

    for line in lines:
        if current_len + len(line) > max_chars and current:
            chunks.append("\n".join(current))
            current = []
            current_len = 0
        current.append(line)
        current_len += len(line) + 1

    if current:
        chunks.append("\n".join(current))

    return chunks


def _retrieve_top_k(query: str, chunks: list[str], k: int = TOP_K) -> list[str]:
    """Retrieve top-k chunks most similar to query using TF-IDF cosine similarity."""
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import cosine_similarity

    if not chunks:
        return []

    corpus = [query] + chunks
    vectorizer = TfidfVectorizer(stop_words="english", token_pattern=r"(?u)\b\w[\w.]+\b")
    tfidf = vectorizer.fit_transform(corpus)

    similarities = cosine_similarity(tfidf[0:1], tfidf[1:]).flatten()

    top_indices = similarities.argsort()[::-1][:k]
    return [chunks[i] for i in top_indices]


def _hit_rate(queries: list[dict], chunks: list[str], k: int = TOP_K) -> tuple[float, list[dict]]:
    """Calculate hit rate: fraction of queries where expected answer is in top-k.

    Returns (hit_rate, details).
    """
    hits = 0
    details: list[dict] = []

    for q in queries:
        retrieved = _retrieve_top_k(q["query"], chunks, k)
        combined = " ".join(retrieved)
        found = q["expected"].lower() in combined.lower()

        if found:
            hits += 1

        details.append({
            "query": q["query"],
            "expected": q["expected"],
            "found": found,
        })

    return hits / len(queries) if queries else 0.0, details


# ── Tests ──


@pytest.fixture(scope="module")
def parsed_result():
    """Parse the real XLSX file with layout hints."""
    if not XLSX_PATH.exists():
        pytest.skip(f"Test fixture not found: {XLSX_PATH}")

    return parse_file(
        XLSX_PATH,
        title_range="B2:*2",
        header_range="B4:*6",
        row_meta_col="B",
    )


@pytest.fixture(scope="module")
def semantic_chunks(parsed_result) -> list[str]:
    """Chunk semantic XML output."""
    all_chunks: list[str] = []
    for xml in parsed_result.semantic_xml.values():
        all_chunks.extend(_chunk_by_records(xml))
    return all_chunks


@pytest.fixture(scope="module")
def raw_chunks(parsed_result) -> list[str]:
    """Chunk raw XML output."""
    all_chunks: list[str] = []
    for name, xml in parsed_result.xml_entries.items():
        if "worksheets" in name:
            all_chunks.extend(_chunk_fixed(xml))
    return all_chunks


def test_semantic_has_better_hit_rate(semantic_chunks, raw_chunks):
    """Semantic XML should have a higher retrieval hit rate than raw XML."""
    semantic_hit_rate, semantic_details = _hit_rate(QUERIES, semantic_chunks)
    raw_hit_rate, raw_details = _hit_rate(QUERIES, raw_chunks)

    # Print comparison report
    report = []
    report.append("")
    report.append("=" * 70)
    report.append("  RETRIEVAL PERFORMANCE COMPARISON")
    report.append("=" * 70)
    report.append(f"  Queries: {len(QUERIES)}  |  Top-K: {TOP_K}")
    report.append(f"  Semantic chunks: {len(semantic_chunks)}  |  Raw chunks: {len(raw_chunks)}")
    report.append("-" * 70)
    report.append(f"  {'Query':<50} {'Semantic':>8} {'Raw':>8}")
    report.append("-" * 70)

    for s, r in zip(semantic_details, raw_details):
        s_mark = "HIT" if s["found"] else "MISS"
        r_mark = "HIT" if r["found"] else "MISS"
        query_short = s["query"][:48]
        report.append(f"  {query_short:<50} {s_mark:>8} {r_mark:>8}")

    report.append("-" * 70)
    report.append(f"  {'HIT RATE':<50} {semantic_hit_rate:>7.0%} {raw_hit_rate:>8.0%}")
    report.append("=" * 70)

    print("\n".join(report))

    assert semantic_hit_rate >= raw_hit_rate, (
        f"Semantic ({semantic_hit_rate:.0%}) should be >= Raw ({raw_hit_rate:.0%})"
    )


def test_semantic_hit_rate_above_threshold(semantic_chunks):
    """Semantic XML should retrieve correct answers for at least 75% of queries."""
    hit_rate, _ = _hit_rate(QUERIES, semantic_chunks)

    assert hit_rate >= 0.75, (
        f"Semantic hit rate {hit_rate:.0%} is below 75% threshold"
    )


def test_chunk_quality_semantic(semantic_chunks):
    """Semantic chunks should contain readable record structures."""
    record_chunks = [c for c in semantic_chunks if "<record" in c]

    assert len(record_chunks) > 0, "No chunks contain <record> elements"

    # Each record chunk should have state attribute or tag-based data
    readable = 0
    for chunk in record_chunks:
        has_state = bool(re.search(r'state\w*="[A-Z]', chunk, re.IGNORECASE))
        has_tags = bool(re.search(r"<total_number>|<race_ethnicity_", chunk))
        if has_state or has_tags:
            readable += 1

    readability = readable / len(record_chunks)
    assert readability >= 0.8, (
        f"Only {readability:.0%} of record chunks are readable (expected >= 80%)"
    )


def test_raw_chunks_are_opaque(raw_chunks):
    """Raw XML chunks should be harder to read — no semantic tag names."""
    opaque = 0
    for chunk in raw_chunks:
        # Raw XML uses <c r="B7" t="s"><v>74</v> — no meaningful tag names
        has_cell_refs = bool(re.search(r'<c r="[A-Z]+\d+"', chunk))
        has_no_semantic = "<record" not in chunk and "<total_number>" not in chunk
        if has_cell_refs and has_no_semantic:
            opaque += 1

    if raw_chunks:
        opacity = opaque / len(raw_chunks)
        assert opacity >= 0.5, (
            f"Only {opacity:.0%} of raw chunks are opaque (expected >= 50%)"
        )
