"""Bit-exact Python twin of ArContentFingerprint in modSysUtils.bas.

The VBA computes two 31-bit polynomial-hash lanes over the string's UTF-16
code units (VBA strings ARE UTF-16; Mid$/AscW walk code units, not code
points), using Double arithmetic. Every intermediate value stays below 2^38
(< 2^53), so VBA's Double math is exact integer math and Python's arbitrary-
precision ints reproduce it bit for bit.

VBA source (modSysUtils.bas):

    h1 = 2166136261# - P1          ' P1 = 2147483647 (2^31 - 1)
    h2 = 1099511628#               ' P2 = 2147483629
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1)) And &HFFFF&
        h1 = h1 * 131 + ch + 1
        h1 = h1 - Int(h1 / P1) * P1
        h2 = h2 * 137 + ch + (i Mod 251) + 1
        h2 = h2 - Int(h2 / P2) * P2
    Next i
    ArContentFingerprint = Hex31(h1) & Hex31(h2)

Notes that force specific choices here:
- `AscW(...) And &HFFFF&` yields the unsigned UTF-16 code unit (AscW returns a
  signed 16-bit reading for units >= 0x8000; the mask undoes that). We
  therefore iterate UTF-16-LE code units, NOT code points: an emoji outside
  the BMP contributes TWO units (its surrogate pair), exactly as in VBA.
- `i` is the 1-based code-unit index (the position salt in lane 2).
- Hex$ output is uppercase; Hex31 zero-pads to 8 chars.
"""

P1 = 2147483647  # 2^31 - 1, prime
P2 = 2147483629
B1 = 131
B2 = 137
H1_SEED = 2166136261 - P1  # 18652614
H2_SEED = 1099511628


def utf16_code_units(s: str):
    """Yield the string's UTF-16 code units as unsigned ints (VBA's view)."""
    b = s.encode("utf-16-le")  # raises on lone surrogates; vectors avoid them
    for j in range(0, len(b), 2):
        yield b[j] | (b[j + 1] << 8)


def fingerprint(s: str) -> str:
    h1 = H1_SEED
    h2 = H2_SEED
    i = 0
    for ch in utf16_code_units(s):
        i += 1
        h1 = (h1 * B1 + ch + 1) % P1
        h2 = (h2 * B2 + ch + (i % 251) + 1) % P2
    return f"{h1:08X}{h2:08X}"
