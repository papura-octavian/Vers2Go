"""Teste pentru citirea secțiunilor și ordinea de cântare.

Rulare:  python -m unittest test_vers2go
"""

import os
import tempfile
import unittest
import zipfile
from xml.dom import minidom

from vers2go import build_plan, generate_presentation, parse_header


def song(*blocks):
    """Versurile din blocuri separate prin câte un rând gol."""
    return "\n\n".join(blocks)


def codes(lyrics, order=""):
    """Codurile slide-urilor (fără titlu), de ex. „S1 R S2 R”."""
    return " ".join(s.code for s in build_plan(lyrics, "", order).slides)


class HeaderTests(unittest.TestCase):
    def test_labels(self):
        self.assertEqual(parse_header("S1"), ("S", 1, ""))
        self.assertEqual(parse_header("Strofa 2"), ("S", 2, ""))
        self.assertEqual(parse_header("strofă 3:"), ("S", 3, ""))
        self.assertEqual(parse_header("R"), ("R", None, ""))
        self.assertEqual(parse_header("Refren 2"), ("R", 2, ""))
        self.assertEqual(parse_header("[Punte]"), ("B", None, ""))
        self.assertEqual(parse_header("Bridge:"), ("B", None, ""))
        self.assertEqual(parse_header("Final"), ("C", None, ""))
        self.assertEqual(parse_header("Pre-refren"), ("P", None, ""))
        self.assertEqual(parse_header("R: Primul vers"), ("R", None, "Primul vers"))

    def test_lyric_lines_are_not_labels(self):
        self.assertIsNone(parse_header("Punte de lumină peste ape"))
        self.assertIsNone(parse_header("Refrenul vieții mele"))
        self.assertIsNone(parse_header("Slavă Ție, Doamne"))

    def test_numbered_verse_only_at_block_start(self):
        self.assertEqual(parse_header("2. Cântăm", at_block_start=True), ("S", 2, "Cântăm"))
        self.assertIsNone(parse_header("2. Cântăm"))


class OrderTests(unittest.TestCase):
    def test_unlabeled_stanzas_like_before(self):
        self.assertEqual(codes(song("unu", "doi", "trei")), "S1 S2 S3")

    def test_old_chorus_marker(self):
        lyrics = song("unu\nunu bis", "R: cor\ncor bis", "doi\ndoi bis")
        self.assertEqual(codes(lyrics), "S1 R S2 R")
        self.assertEqual(build_plan(lyrics).slides[1].text, "cor\ncor bis")

    def test_chorus_after_verses_and_bridge(self):
        lyrics = song("S1\nunu", "R\ncor", "S2\ndoi", "B\npeste ape", "C\namin")
        self.assertEqual(codes(lyrics), "S1 R S2 R B R C")

    def test_written_out_chorus_is_detected(self):
        # refrenul lipit de mai multe ori, fără etichete: nu se dublează
        self.assertEqual(codes(song("unu", "cor", "doi", "cor")), "S1 S2 S3 S2")
        self.assertEqual(codes(song("unu", "R\ncor", "doi", "R")), "S1 R S2 R")

    def test_two_choruses(self):
        lyrics = song("S1\nunu", "R1\ncor unu", "S2\ndoi", "R2\ncor doi", "S3\ntrei")
        # după S2 e scris R2, deci de acolo încolo se cântă R2
        self.assertEqual(codes(lyrics), "S1 R1 S2 R2 S3 R2")

    def test_pre_chorus(self):
        lyrics = song("S1\nunu", "P\nurcăm", "R\ncor", "S2\ndoi")
        self.assertEqual(codes(lyrics), "S1 P R S2 R")

    def test_explicit_order(self):
        lyrics = song("S1\nunu", "R\ncor", "S2\ndoi", "B\npeste ape")
        self.assertEqual(codes(lyrics, "S1 R S2 R B Rx2"), "S1 R S2 R B R R")
        self.assertEqual(codes(lyrics, "Strofa 1, Refren, Punte, R1"), "S1 R B R")

    def test_explicit_order_errors(self):
        plan = build_plan(song("S1\nunu", "R\ncor"), "", "S1 R S4 X S")
        self.assertEqual(len(plan.errors), 3)
        self.assertEqual(" ".join(s.code for s in plan.slides), "S1 R")

    def test_reference_before_definition(self):
        self.assertEqual(codes(song("R", "S1\nunu", "R\ncor")), "R S1 R")

    def test_title_slide(self):
        plan = build_plan("unu", "Titlu")
        self.assertEqual(plan.slides[0].text, "Titlu")
        self.assertEqual(plan.slides[0].label, "")


class PptxTests(unittest.TestCase):
    def test_generated_file_is_well_formed(self):
        lyrics = song("S1\nunu & <doi>", "R\ncor", "S2\ndoi")
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "test.pptx")
            for show_labels in (False, True):
                n = generate_presentation(lyrics, "Titlu", path, "", show_labels)
                self.assertEqual(n, 5)
                with zipfile.ZipFile(path) as zf:
                    for name in zf.namelist():
                        if name.endswith((".xml", ".rels")):
                            minidom.parseString(zf.read(name))
                    slide = zf.read("ppt/slides/slide2.xml").decode()
                    self.assertEqual("Strofa 1" in slide, show_labels)

    def test_bad_order_raises(self):
        with self.assertRaises(ValueError):
            generate_presentation("unu", "", os.devnull, "S9")


if __name__ == "__main__":
    unittest.main()
