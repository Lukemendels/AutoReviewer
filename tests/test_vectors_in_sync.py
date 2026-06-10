"""The committed vector files must be exactly what the twin generates.

If this fails, someone hand-edited a vector file or changed the twin without
running tests/vectors/regenerate.py.
"""

import os

from tests.vectors.regenerate import (
    build_fingerprint_vectors,
    build_parser_vectors,
    build_session_vectors,
    build_coverage_vectors,
    vector_escape,
    vector_unescape,
)

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vectors")


def _read(name):
    with open(os.path.join(HERE, name), "r", encoding="ascii", newline="") as fh:
        return fh.read().replace("\r\n", "\n")


def test_fingerprint_vectors_in_sync():
    assert _read("fingerprint_vectors.txt") == build_fingerprint_vectors()


def test_parser_vectors_in_sync():
    assert _read("parser_vectors.txt") == build_parser_vectors()


def test_session_vectors_in_sync():
    assert _read("session_vectors.txt") == build_session_vectors()


def test_coverage_vectors_in_sync():
    assert _read("coverage_vectors.txt") == build_coverage_vectors()


def test_escape_round_trip():
    cases = [
        "",
        "plain",
        "tab\there",
        "nl\nhere",
        "cr\rhere",
        "back\\slash",
        "em—dash 安全 \U0001F600",
        "\\u0041 literal escape text",
    ]
    for s in cases:
        assert vector_unescape(vector_escape(s)) == s
        vector_escape(s).encode("ascii")  # escaped form is always pure ASCII
