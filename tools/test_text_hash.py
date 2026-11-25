"""The arithmetic of CORE_Util.TextHash (#771), modelled in IEEE double precision.

TextHash adds the UTF-16 bytes of a text three at a time and takes the modulus
2^31 - 1 once per three; before #771 it took it after every byte. VBA's Double is
an IEEE double, as Python's float is, so this model shows that both methods give
the same fingerprint: every intermediate value is a whole number below 2^53, and
the rounded quotient never crosses a whole number. It also finds the edge cases
TEST_MainState.CaseTextHashUnchanged runs in Excel, and checks the values that
case expects.
"""

from __future__ import annotations

import math
import random
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MODULUS = 2147483647.0
BASE = 131.0
GROUP_MAX = 255 * 131 * 131 + 255 * 131 + 255


def per_byte(data: bytes) -> float:
    """The method before #771: the modulus after every byte."""
    value = 0.0
    for byte in data:
        value = value * BASE + byte
        value = value - math.floor(value / MODULUS) * MODULUS
    return value


def per_group(data: bytes) -> float:
    """TextHash since #771: whole groups of three bytes, then the last one or two."""
    value = 0.0
    group_end = len(data) - len(data) % 3
    for i in range(0, group_end, 3):
        value = ((value * BASE + data[i]) * BASE + data[i + 1]) * BASE + data[i + 2]
        value = value - math.floor(value / MODULUS) * MODULUS
    for i in range(group_end, len(data)):
        value = value * BASE + data[i]
        value = value - math.floor(value / MODULUS) * MODULUS
    return value


def hex8(value: float) -> str:
    return f"{int(value):08X}"


def text_hash(text: str) -> str:
    return hex8(per_byte(text.encode("utf-16-le")))


# Sequences whose last group of three reaches a multiple of the modulus, or one
# less, from a hash above 0.9 times the modulus: quotients near 2.2 million,
# where a rounded quotient is most at risk. CaseTextHashUnchanged runs them.
EDGES = {
    "459ABE7E9E24D4123F": "00000000",
    "20BE26A2B54CB55366": "00000000",
    "2B9539906EF2713631": "00000000",
    "459ABE7E9E24D4123E": "7FFFFFFE",
    "20BE26A2B54CB55365": "7FFFFFFE",
    "2B9539906EF2713630": "7FFFFFFE",
}


class TextHashTests(unittest.TestCase):
    def test_a_group_stays_below_two_to_the_53(self) -> None:
        largest = (MODULUS - 1) * BASE ** 3 + GROUP_MAX
        self.assertLess(largest, 2.0 ** 53)
        self.assertEqual(((MODULUS - 1) * BASE + 255) * BASE * BASE + 255 * BASE + 255, largest)

    def test_the_quotient_error_is_below_the_gap_to_a_whole_number(self) -> None:
        largest_quotient = ((MODULUS - 1) * BASE ** 3 + GROUP_MAX) / MODULUS
        self.assertLess(largest_quotient * 2.0 ** -53, 1 / MODULUS)

    def test_both_methods_agree_on_random_texts_and_every_short_length(self) -> None:
        rng = random.Random(771)
        for _ in range(20000):
            data = bytes(rng.randrange(256) for _ in range(rng.randrange(0, 40)))
            self.assertEqual(per_group(data), per_byte(data), data.hex())
        for fill in (0x00, 0xFF, 0x7F):
            for length in range(13):
                data = bytes([fill]) * length
                self.assertEqual(per_group(data), per_byte(data), data.hex())

    def test_the_edge_groups_reach_a_multiple_and_both_methods_agree(self) -> None:
        for digits, expected in EDGES.items():
            data = bytes.fromhex(digits)
            state = per_group(data[:6])
            self.assertGreater(state, 0.9 * MODULUS)
            group = ((state * BASE + data[6]) * BASE + data[7]) * BASE + data[8]
            self.assertGreater(group / MODULUS, 2.0e6)
            self.assertIn(group % MODULUS, (0.0, MODULUS - 1))
            self.assertEqual(hex8(per_group(data)), expected)
            self.assertEqual(hex8(per_byte(data)), expected)

    def test_the_values_the_excel_case_expects(self) -> None:
        self.assertEqual(text_hash(""), "00000000")
        self.assertEqual(text_hash("A"), "00002143")
        self.assertEqual(text_hash("K-SACCR"), "6F1505D8")
        self.assertEqual(text_hash("EUR|2026-09-30|€"), "0BCB808F")
        self.assertEqual(hex8(per_byte(bytes(range(256)))), "4DACC373")
        self.assertEqual(hex8(per_byte(bytes(range(255, -1, -1)))), "7BD9AFBA")

    def test_the_vba_source_uses_this_arithmetic(self) -> None:
        source = (ROOT / "src" / "core" / "CORE_Util.bas").read_text(encoding="cp1252")
        body = source[source.index("Public Function TextHash("):]
        body = body[:body.index("End Function")]
        self.assertIn("HASH_MODULUS As Double = 2147483647#", body)
        self.assertIn("HASH_BASE    As Double = 131#", body)
        self.assertRegex(body, re.escape(
            "Hash = ((Hash * HASH_BASE + Bytes(i)) * HASH_BASE + Bytes(i + 1)) * HASH_BASE + Bytes(i + 2)"))
        self.assertEqual(body.count("Hash = Hash - Int(Hash / HASH_MODULUS) * HASH_MODULUS"), 2)
        case = (ROOT / "tests" / "modules" / "TEST_MainState.bas").read_text(encoding="cp1252")
        for digits, expected in EDGES.items():
            self.assertIn(f'"{digits}", "{expected}"', case)


if __name__ == "__main__":
    unittest.main()
