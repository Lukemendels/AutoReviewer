"""Independent verification of zip/writer.js's output, using Python's zipfile module
instead of this project's own zip reader -- the point is a check that doesn't share any
code with the thing being verified (spec §13 invariant #5).

Usage: python3 scripts/verify_zip_roundtrip.py <fixtures_dir> <roundtrip_dir>

For every fixture, asserts the round-tripped copy: opens cleanly, testzip() reports no bad
CRCs, has the same entry names/order as the original, every entry's bytes are
byte-identical to the original's, and it passes a schema-lite structural check (required
parts present, well-formed XML, content-types complete).
"""

import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

# Schema-lite structural check (spec §13 invariant #2): a Word .docx that opens without a
# repair prompt has, at minimum, these parts present and well-formed. This can't be a
# substitute for actually opening the file in Word -- there's no Word in this environment,
# and LibreOffice (the initially-planned automated proxy) turns out not to load *any* .docx
# in this sandboxed environment, including pristine, never-touched fixtures, so it isn't a
# usable check here either. This structural check is what's left as an automated signal;
# treat "no repair prompt" as still needing a manual check in real Word.
REQUIRED_PARTS = ["[Content_Types].xml", "_rels/.rels", "word/document.xml"]
XML_PARTS_TO_VALIDATE = ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/comments.xml", "word/commentsExtended.xml"]


def verify_structure(roundtrip_path: Path) -> list[str]:
    errors = []
    with zipfile.ZipFile(roundtrip_path) as rt:
        names = set(rt.namelist())
        for part in REQUIRED_PARTS:
            if part not in names:
                errors.append(f"required part missing: {part}")

        for part in XML_PARTS_TO_VALIDATE:
            if part not in names:
                continue
            try:
                ET.fromstring(rt.read(part))
            except ET.ParseError as e:
                errors.append(f"{part} does not parse as well-formed XML: {e}")

        if "[Content_Types].xml" in names and "word/document.xml" in names:
            content_types = rt.read("[Content_Types].xml").decode("utf-8", errors="replace")
            if "word/document.xml" not in content_types:
                errors.append("[Content_Types].xml has no Override entry for word/document.xml")
    return errors


def verify_one(original_path: Path, roundtrip_path: Path) -> list[str]:
    errors = []
    with zipfile.ZipFile(original_path) as orig, zipfile.ZipFile(roundtrip_path) as rt:
        bad = rt.testzip()
        if bad is not None:
            errors.append(f"testzip() reported a bad CRC for entry: {bad}")

        orig_names = orig.namelist()
        rt_names = rt.namelist()
        if orig_names != rt_names:
            errors.append(f"entry order/names differ:\n  original:   {orig_names}\n  round-trip: {rt_names}")

        for name in orig_names:
            if name not in rt_names:
                continue  # already reported above
            orig_bytes = orig.read(name)
            rt_bytes = rt.read(name)
            if orig_bytes != rt_bytes:
                errors.append(f"entry '{name}' content differs ({len(orig_bytes)} vs {len(rt_bytes)} bytes)")

    errors.extend(verify_structure(roundtrip_path))
    return errors


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(2)
    fixtures_dir = Path(sys.argv[1])
    roundtrip_dir = Path(sys.argv[2])

    fixtures = sorted(fixtures_dir.glob("*.docx"))
    if not fixtures:
        print(f"no .docx fixtures found in {fixtures_dir}")
        sys.exit(2)

    failures = 0
    for original_path in fixtures:
        roundtrip_path = roundtrip_dir / original_path.name
        if not roundtrip_path.exists():
            print(f"FAIL  {original_path.name}: no round-tripped copy at {roundtrip_path}")
            failures += 1
            continue
        errors = verify_one(original_path, roundtrip_path)
        if errors:
            print(f"FAIL  {original_path.name}:")
            for e in errors:
                print(f"      {e}")
            failures += 1
        else:
            print(f"PASS  {original_path.name}")

    print()
    if failures:
        print(f"{failures} fixture(s) failed independent zipfile verification.")
        sys.exit(1)
    print(f"All {len(fixtures)} fixture(s) verified byte-identical and structurally valid (Python zipfile).")


if __name__ == "__main__":
    main()
