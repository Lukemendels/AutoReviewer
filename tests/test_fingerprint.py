from ref.fingerprint import fingerprint, utf16_code_units, P1, P2, H1_SEED, H2_SEED


def manual_fingerprint(units):
    """Independent re-derivation straight from the VBA listing."""
    h1, h2 = H1_SEED, H2_SEED
    for i, ch in enumerate(units, start=1):
        h1 = (h1 * 131 + ch + 1) % P1
        h2 = (h2 * 137 + ch + (i % 251) + 1) % P2
    return f"{h1:08X}{h2:08X}"


def test_empty_string():
    assert fingerprint("") == f"{H1_SEED:08X}{H2_SEED:08X}"


def test_format_is_16_uppercase_hex():
    fp = fingerprint("hello world")
    assert len(fp) == 16
    assert fp == fp.upper()
    int(fp, 16)  # must be valid hex


def test_deterministic():
    s = 'a JSON {"line": "with\ttabs"} and unicode — dash'
    assert fingerprint(s) == fingerprint(s)


def test_matches_manual_derivation_ascii():
    s = "The quick brown fox"
    assert fingerprint(s) == manual_fingerprint(ord(c) for c in s)


def test_single_char_difference_changes_output():
    assert fingerprint("abc") != fingerprint("abd")


def test_reordering_changes_output():
    # Lane 2 folds the position in, so same multiset of chars must differ.
    assert fingerprint("ab") != fingerprint("ba")


def test_utf16_units_bmp():
    # em dash U+2014 is one unit
    assert list(utf16_code_units("—")) == [0x2014]


def test_utf16_units_surrogate_pair():
    # U+1F600 GRINNING FACE encodes as the surrogate pair D83D DE00 in
    # UTF-16; VBA's Mid$/AscW walk exactly those two units.
    assert list(utf16_code_units("\U0001F600")) == [0xD83D, 0xDE00]


def test_emoji_matches_manual_over_units():
    s = "ok \U0001F600 done"
    assert fingerprint(s) == manual_fingerprint(utf16_code_units(s))


def test_cjk():
    s = "安全运输"  # CJK, all BMP single units
    assert fingerprint(s) == manual_fingerprint(ord(c) for c in s)


def test_long_string_stays_exact():
    # The Double-safety claim: intermediates < 2^38. Long inputs must still
    # match the manual modular derivation exactly.
    s = "A" * 5000 + "—" + "z" * 5000
    assert fingerprint(s) == manual_fingerprint(utf16_code_units(s))


def test_position_salt_period():
    # i Mod 251: positions 1 and 252 carry the same salt; the strings still
    # differ through lane accumulation. Just pin that nothing throws and
    # results differ for a targeted pair.
    a = "x" * 251 + "Q"
    b = "Q" + "x" * 251
    assert fingerprint(a) != fingerprint(b)


def test_crlf_vs_lf_differ():
    assert fingerprint("a\r\nb") != fingerprint("a\nb")
