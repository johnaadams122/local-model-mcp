"""Tests for local_llm.docx_text.

Builds minimal .docx files on disk (real zip archives with a hand-crafted
word/document.xml) so extraction can be tested without a python-docx
dependency.

The battery covers containment, public document seams and content hardening.
Rows are grouped by behavior, after six basic extraction rows.
"""

import io
import os
import tempfile
import time
import unittest
import xml.parsers.expat
import zipfile
import zlib
from unittest import mock

from local_llm import docx_text
from local_llm.docx_text import _assert_no_dtd, _assert_within_cap, extract_text
from local_llm.docx_text import (DEFAULT_MAX_BYTES, _assert_no_dtd,
                                 _assert_within_cap, extract_text)

_DOCUMENT_XML_TEMPLATE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body>{body}</w:body></w:document>"
)

_ENTITY_LADDER = (
    '<!ENTITY a "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA">'
    '<!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">'
    '<!ENTITY c "&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;">'
    '<!ENTITY d "&c;&c;&c;&c;&c;&c;&c;&c;&c;&c;">'
    '<!ENTITY e "&d;&d;&d;&d;&d;&d;&d;&d;&d;&d;">'
)


def _entity_bomb_xml(encoding="UTF-8"):
    """A five-rung internal-DTD entity ladder. Every entity is declared INSIDE
    the document, so nothing is fetched from the network. Measured:
    a 358-byte archive, 468 declared bytes, 2.07:1 -- and 500,000 characters of
    visible text."""
    return (
        '<?xml version="1.0" encoding="%s" standalone="yes"?>'
        '<!DOCTYPE w:document [%s]>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&e;</w:t>'
        '</w:r></w:p></w:body></w:document>' % (encoding, _ENTITY_LADDER)
    )


def _bulky_xml(root_tag, count):
    """Filler whose compression ratio lands in the same band as real Word parts
    (~10:1), NOT the ~217:1 of a repeated single byte. A fixture that
    compressed unrealistically well would trip the ratio guard and would prove
    nothing about real documents.

    Deliberately NOT namespace-declared: w:-prefixed tags with no xmlns:w, so
    it is well FORMED in shape but would not parse standalone. That is
    intentional -- these members carry realistic BYTES past the metadata
    screen, and extract_text never decompresses or parses anything except
    word/document.xml."""
    parts = ["<%s>" % root_tag]
    for index in range(count):
        parts.append(
            '<w:style w:styleId="S%d"><w:name w:val="Style %d"/>'
            '<w:rPr><w:sz w:val="%d"/><w:color w:val="%06X"/></w:rPr>'
            "</w:style>"
            % (index, index, 16 + (index % 40), (index * 7919) % 0xFFFFFF))
    parts.append("</%s>" % root_tag)
    return "".join(parts)


# A REPRESENTATIVE subset of a real Word document's members, in real write
# order: _rels/.rels first, [Content_Types].xml LAST, document.xml neither. A
# real document inspected for reference carried 26 members; this fixture
# carries 14 and is deliberately not the full set. Any positional assertion
# against infolist() fails here by design.
_WORD_MEMBER_ORDER = (
    "_rels/.rels",
    "word/_rels/document.xml.rels",
    "word/document.xml",
    "word/styles.xml",
    "word/settings.xml",
    "word/webSettings.xml",
    "word/numbering.xml",
    "word/fontTable.xml",
    "word/theme/theme1.xml",
    "word/endnotes.xml",
    "word/footnotes.xml",
    "docProps/core.xml",
    "docProps/app.xml",
    "[Content_Types].xml",
)

def _make_docx(body_xml):
    handle, path = tempfile.mkstemp(suffix=".docx")
    os.close(handle)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("word/document.xml", _DOCUMENT_XML_TEMPLATE.format(body=body_xml))
    return path



class TestExtractText(unittest.TestCase):
    def setUp(self):
        pass                      # no per-test state; see _track for cleanup

    def _track(self, path):
        """Remove `path` at cleanup time, NOT in tearDown.

        Registered BEFORE any handle opened on it, so LIFO ordering closes the
        handle first. Do not add a tearDown that deletes these paths.
        """
        self.addCleanup(os.remove, path)
        return path

    def _docx(self, body_xml):
        # Cleanup goes through _track (addCleanup), not a per-test list read
        # by tearDown, so every row built on _docx removes its temp file.
        return self._track(_make_docx(body_xml))

    def _zip_path(self):
        """A tracked temp path for hand-built archives that are NOT valid Word
        documents (the _docx helper always writes valid body XML)."""
        descriptor, path = tempfile.mkstemp(suffix=".docx")
        os.close(descriptor)
        self._track(path)
        return path

    def _raw_docx(self, document_bytes, compression=zipfile.ZIP_DEFLATED):
        """An archive whose word/document.xml is EXACTLY the given bytes.

        compression is a parameter because the three metadata branches can
        only be isolated by controlling it: ZIP_STORED forces a 1.0:1 ratio,
        which is the only way to reach the declared-size branch without the
        ratio branch also firing. Defaults to DEFLATED."""
        path = self._zip_path()
        with zipfile.ZipFile(path, "w", compression) as archive:
            archive.writestr("word/document.xml", document_bytes)
        return path

    def _member_payload(self, name, body_xml):
        if name == "word/document.xml":
            return _DOCUMENT_XML_TEMPLATE.format(body=body_xml)
        if name in ("word/styles.xml", "word/numbering.xml"):
            return _bulky_xml("w:styles", 400)
        return '<?xml version="1.0" encoding="UTF-8"?><root name="%s"/>' % name

    def _word_shaped_docx(self, body_xml, extra_members=(),
                          member_overrides=None):
        """A Word-SHAPED archive: a representative member subset, real write
        order, realistic member sizes and compression ratios.

        member_overrides REPLACES the payload of a member already in the
        standard order. Use it rather than extra_members for such a name:
        zipfile will happily write a second entry under the same name, and
        which of the two a reader returns is not something a test should
        depend on."""
        overrides = member_overrides or {}
        path = self._zip_path()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for name in _WORD_MEMBER_ORDER:
                archive.writestr(name, overrides.get(
                    name, self._member_payload(name, body_xml)))
            for name, payload in extra_members:
                archive.writestr(name, payload)
        return path

    # ---------------------------------------------- basic extraction rows
    def test_single_paragraph(self):
        path = self._docx("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        self.assertEqual(extract_text(path), "Hello world")

    def test_multiple_paragraphs_joined_with_newline(self):
        path = self._docx(
            "<w:p><w:r><w:t>Line one</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Line two</w:t></w:r></w:p>"
        )
        self.assertEqual(extract_text(path), "Line one\nLine two")

    def test_multiple_runs_in_one_paragraph_concatenated(self):
        path = self._docx(
            "<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>"
        )
        self.assertEqual(extract_text(path), "Hello world")

    def test_empty_paragraph_skipped(self):
        path = self._docx(
            "<w:p><w:r><w:t>Real text</w:t></w:r></w:p>"
            "<w:p><w:pPr/></w:p>"
        )
        self.assertEqual(extract_text(path), "Real text")

    def test_table_cell_paragraphs_included(self):
        path = self._docx(
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell text</w:t></w:r></w:p>"
            "</w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(extract_text(path), "Cell text")

    def test_missing_file_raises_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            extract_text("does_not_exist.docx")


    # ------------------------------------------------- Word-shaped acceptance
    def test_word_shaped_document_is_accepted_under_the_real_defaults(self):
        # ACCEPTANCE: the guard must not reject a realistically shaped Word
        # archive. Runs with the SHIPPED defaults, no relaxed arguments.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>Procurement summary</w:t></w:r></w:p>")
        self.assertEqual(extract_text(path), "Procurement summary")

    def test_the_SHIPPED_default_bounds_are_exactly_these_values(self):
        # Every cap-rejection and boundary row below supplies max_*
        # EXPLICITLY, and the acceptance row above only proves that a small
        # archive is accepted. Without this row nothing pins the numbers:
        # multiplying all three constants -- and the signature defaults with
        # them -- would leave every other row green. That is not academic:
        # file_tools passes no max_ratio, so every MCP .docx read relies on
        # DEFAULT_MAX_RATIO, and any caller that passes no cap keywords rides
        # entirely on these defaults.
        #
        # Pin the constants AND the signature. They can move independently --
        # rebinding a default to a new literal leaves the constant untouched,
        # and vice versa -- so pinning either alone leaves a way through.
        self.assertEqual(docx_text.DEFAULT_MAX_BYTES, 384000)
        self.assertEqual(docx_text.DEFAULT_MAX_RATIO, 100)
        self.assertEqual(docx_text.DEFAULT_MAX_TEXT_CHARS, 384000)
        # __defaults__ rather than inspect: no new import, and it pins the
        # ORDER of the three parameters, which a positional caller depends on.
        self.assertEqual(extract_text.__defaults__, (384000, 100, 384000))

    def test_the_DEFAULT_path_rejects_an_over_cap_archive_with_no_arguments(self):
        # The companion to the row above. Pinning the constants proves the
        # numbers; it does not prove the default path ENFORCES them. This row
        # passes no max_* argument at all -- the call shape of any caller
        # relying on the defaults -- so it fails if the defaults are ever
        # wired to something other than the screen.
        #
        # Declared 400,000 > DEFAULT_MAX_BYTES (384,000), and the declared-size
        # branch is screened BEFORE the ratio branch, so this row is about the
        # bytes default specifically. The payload is a single repeated byte, so
        # the fixture compresses to almost nothing on disk.
        path = self._raw_docx(b"A" * 400000)
        with self.assertRaisesRegex(ValueError, "declared size"):
            extract_text(path)

    def test_a_LARGE_MEDIA_part_is_accepted_and_never_opened(self):
        # By design the suffix screen covers .xml/.rels parts, not every
        # member, and this row is the mechanical guard on that decision. A
        # real Word document embeds media, and a 400,000-byte image exceeds
        # DEFAULT_MAX_BYTES while never being parsed -- so screening every
        # member would reject legitimate documents to guard bytes nobody
        # decompresses.
        #
        # Without this row the acceptance fixture is XML and .rels parts only,
        # so deleting the suffix filter leaves every other test green and the
        # regression surfaces only on a real document with a photo in it. The
        # payload is a single repeated byte, so it compresses to almost
        # nothing: a widened screen would fail it on declared size AND ratio.
        #
        # All three source forms are driven, because the path form is NOT
        # what production uses: file_tools._read_text hands .docx bytes down
        # as extract_text(io.BytesIO(payload), ...), so BytesIO is the route
        # summarize_file and extract_json_file take. An implementation that
        # drops the suffix filter ONLY on the file-like branch would pass a
        # path-only row and then REJECT any real Word document containing a
        # photo, through the exact route the MCP tools use -- a false
        # rejection rather than a hole.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>Procurement summary</w:t></w:r></w:p>",
            extra_members=(("word/media/image1.png",
                            b"\x89PNG\r\n\x1a\n" + b"Z" * 400000),))

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, make_source in (("path", lambda p: p),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            source = make_source(path)

            def run():
                self.assertEqual(extract_text(source),
                                 "Procurement summary", label)

            self.assertEqual(
                self._opens_during(run), ["word/document.xml"],
                "%s: a non-XML media part was decompressed: only "
                "word/document.xml may ever be opened" % label)

    def test_an_UPPERCASE_xml_or_rels_member_is_still_screened(self):
        # Every hostile member name elsewhere in this suite is lowercase, so
        # deleting the `.lower()` from the suffix predicate leaves the entire
        # file green while an uppercase
        # ".XML" or ".RELS" part walks straight past the metadata screen.
        # OPC does not require lowercase part names -- only the reserved names
        # are fixed -- so this is a reachable bypass, not a curiosity.
        for member in ("word/HEADER1.XML", "word/_rels/HEADER1.XML.RELS"):
            path = self._word_shaped_docx(
                "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
                extra_members=((member, b"A" * 200000),))
            with self.assertRaisesRegex(
                    ValueError, "compression ratio|declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

    def test_LATER_non_target_members_use_the_CALLERS_limits(self):
        # Every other hostile later-member row in this block passes
        # 384000/100 -- which ARE DEFAULT_MAX_BYTES and DEFAULT_MAX_RATIO -- so
        # an archive sweep hard-coded to the library defaults, with only the
        # immediate target rescreen honouring the arguments, would pass all of
        # them. Both cases below sit BETWEEN a caller limit and the default,
        # so only a sweep that reads the caller values refuses them.
        #
        # DECLARED SIZE, isolated: os.urandom is incompressible, so the ratio
        # branch cannot fire first and this can only be the size branch.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            extra_members=(("word/later.xml", os.urandom(200000)),))

        def run_size():
            with self.assertRaisesRegex(ValueError, "declared size"):
                extract_text(path, max_bytes=100000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run_size), [],
            "a member was opened before the caller-limit size screen refused")

        # RATIO, isolated: forged 200000/4000 = 50:1, over a caller cap of 25
        # and under the library default of 100. file_size stays below both size
        # caps, so the size branch cannot fire first.
        ratio_path, forged = self._docx_with_forged_ratio(200000, 4000)

        def run_ratio():
            with forged:
                with self.assertRaisesRegex(ValueError, "compression ratio"):
                    extract_text(ratio_path, max_bytes=384000, max_ratio=25)

        self.assertEqual(
            self._opens_during(run_ratio), [],
            "a member was opened before the caller-limit ratio screen refused")

    def test_a_hostile_docProps_member_is_screened_with_zero_opens(self):
        # Every other hostile fixture in this block lives under
        # word/, or is _rels/.rels, or is [Content_Types].xml. A predicate
        # limited to those three locations passes the whole battery while
        # skipping docProps/core.xml and docProps/app.xml -- both written by the
        # representative fixture, both ending in .xml, and therefore both inside
        # the screen this module claims to apply.
        #
        # member_overrides, not extra_members: the name is already in the
        # standard write order, and a duplicate entry leaves which one a reader
        # returns undefined.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            member_overrides={"docProps/core.xml": b"Z" * 200000})

        def run():
            with self.assertRaisesRegex(
                    ValueError, "compression ratio|declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "docProps/core.xml escaped the metadata screen")

    def test_a_CORRUPT_archive_FAILS_CLOSED_and_returns_no_text(self):
        # A non-ZIP or truncated archive. A regression that catches
        # zipfile.BadZipFile and returns "" passes every other row here and
        # hands an EMPTY digest to the model -- a silent wrong answer rather
        # than a refusal, on the production read path, which is strictly worse
        # than an exception the caller can see. Either exception type is
        # acceptable: what is asserted is that this RAISES, not RETURNS.
        #
        # The failure here is raised by the ZipFile CONSTRUCTOR, before any
        # member is opened, which is a different code path from the mid-read
        # corruption row -- so an implementation swallowing constructor-time
        # BadZipFile for BytesIO and handles ONLY, and returning "", would pass
        # a path-only row while feeding an empty digest to the model on exactly
        # the BytesIO route production uses. All three source forms are run.
        path = self._zip_path()
        with open(path, "wb") as handle:
            handle.write(b"not a zip archive at all, not even close")

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, source in (("path", path),
                              ("BytesIO", self._bytesio_of(path)),
                              ("handle", handle_of(path))):
            with self.assertRaises(
                    (ValueError, zipfile.BadZipFile),
                    msg="%s: a non-ZIP archive RETURNED instead of raising -- "
                        "the model receives an empty digest" % label):
                extract_text(source)

    def test_the_ancestor_walk_is_LINEAR_not_quadratic_in_depth(self):
        # A parent-map-plus-walk-UP design is QUADRATIC in nesting depth
        # against input that stays inside every cap. Measured, one paragraph,
        # n runs at depths 1..n beneath w:sdt wrappers:
        #
        #   n=800   322,000 ancestor steps  0.045s
        #   n=3200  5,128,000 steps         0.716s
        #   n=6400  20,496,000 steps        3.703s   input 243,379 B, out 6,400 chars
        #
        # 243 KB of legitimate-looking XML yielding 6 KB of text cost 3.7
        # seconds inside summarize_file. Neither the byte cap nor the character
        # cap bounds it, because neither bytes nor characters are what grows.
        # The one-pass traversal measures 0.011s at the same n.
        #
        # ZIP_STORED IS LOAD-BEARING. This body is 6,400 identical openers
        # followed by 6,400 identical closers, which deflate compresses to
        # almost nothing. With _raw_docx's ZIP_DEFLATED default, measured:
        #   declared 243,379  compressed 756  ratio 321.9:1
        # and _screen_member refuses it at the 100:1 limit --
        #   ValueError: docx member word/document.xml compression ratio
        #   321.9:1 exceeds the 100:1 limit
        # -- BEFORE the traversal runs, so every assertion below would be
        # unreachable and this row could never pass. ZIP_STORED makes
        # declared == compressed == 243,379, ratio 1.0:1, so the ONLY thing
        # under test is the traversal.
        depth = 6400
        body = ""
        for _ in range(depth):
            body = "<w:sdt><w:r><w:t>x</w:t></w:r>%s</w:sdt>" % body
        path = self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(
                body="<w:p>%s</w:p>" % body).encode("utf-8"),
            compression=zipfile.ZIP_STORED)

        # ASSERT THE PRECONDITION rather than trusting it. If a future edit
        # reintroduces compression here, this fails LOUDLY instead of the
        # row silently going green for the wrong reason -- a refusal would
        # otherwise look like "the traversal never ran".
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo("word/document.xml")
        self.assertEqual(
            info.file_size, info.compress_size,
            "the traversal fixture must be STORED: a compressed one is "
            "refused by the ratio screen before the traversal runs")

        started = time.perf_counter()
        text = extract_text(path)
        elapsed = time.perf_counter() - started

        self.assertEqual(len(text), depth,
                         "every run must still be extracted exactly once")
        # TRIPWIRE, not a benchmark. The bound sits ~180x above the one-pass
        # traversal and ~2x below the quadratic one it replaced, so it cannot
        # flake on a loaded box and cannot pass against a reintroduced walk-up.
        self.assertLess(elapsed, 2.0,
                        "traversal took %.2fs for %d nested runs -- the "
                        "quadratic ancestor walk is back" % (elapsed, depth))

    def test_word_shaped_fixture_really_is_multi_member_and_realistic(self):
        # Guards the FIXTURE itself. If this ever collapses to one member or to
        # a toy ratio, the acceptance test above silently stops meaning
        # anything. Members are selected BY NAME, never by position.
        path = self._word_shaped_docx("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
        with zipfile.ZipFile(path) as archive:
            infos = {info.filename: info for info in archive.infolist()}
            names = [info.filename for info in archive.infolist()]
        self.assertGreaterEqual(len(infos), 14)
        self.assertEqual(names[0], "_rels/.rels")
        self.assertEqual(names[-1], "[Content_Types].xml")
        self.assertNotEqual(names[0], "word/document.xml")
        styles = infos["word/styles.xml"]
        self.assertGreater(styles.file_size, 40000)
        ratio = styles.file_size / styles.compress_size
        self.assertGreater(ratio, 5)
        self.assertLess(ratio, 100)

    # ------------------------------------------------------ metadata screens
    def test_member_declaring_more_than_the_cap_is_rejected(self):
        path = self._raw_docx(b"A" * 5000)
        with self.assertRaisesRegex(ValueError, "declared size"):
            extract_text(path, max_bytes=1000)

    def test_member_over_the_compression_ratio_is_rejected(self):
        path = self._raw_docx(b"A" * 200000)
        with self.assertRaisesRegex(ValueError, "compression ratio"):
            extract_text(path, max_bytes=384000, max_ratio=100)

    def test_non_document_member_of_a_word_shaped_archive_is_screened(self):
        # The bomb rides in a member extract_text never decompresses, inside an
        # otherwise entirely valid Word-shaped archive.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            extra_members=(("word/header1.xml", b"A" * 200000),))
        with self.assertRaisesRegex(ValueError, "compression ratio|declared size"):
            extract_text(path, max_bytes=384000, max_ratio=100)

    def test_a_LATER_hostile_member_is_screened_before_ANY_decompression(self):
        # The row above and the single-member ordering rows do not, between
        # them, pin archive-WIDE ordering.
        # The ordering rows use ONE-member archives whose word/document.xml is
        # itself the hostile member; the row above uses a later member but
        # asserts only the eventual ValueError. An implementation that screens
        # word/document.xml, DECOMPRESSES it, and only then screens the rest of
        # infolist() satisfies both -- and that is exactly the read-before-
        # screen bug the ordering rows exist to catch, merely moved one member
        # to the right.
        #
        # _word_shaped_docx writes extra_members AFTER every name in
        # _WORD_MEMBER_ORDER, so word/header1.xml is the LAST entry and
        # word/document.xml is the third. Zero opens is therefore only
        # achievable by completing the screen over the whole central directory
        # before the first decompression -- not 'fewer' opens, ZERO.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            extra_members=(("word/header1.xml", b"A" * 200000),))

        def run():
            with self.assertRaisesRegex(
                    ValueError, "compression ratio|declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "a member was decompressed before the archive-wide screen "
            "reached the hostile member, which is written last")

    # The row above closes ONE cell of the ordering matrix -- the ratio
    # branch, on a later lowercase ".xml" member. The single-member
    # declared-size and zero-compressed-size ordering rows use
    # word/document.xml itself, and the ".rels" and uppercase rows assert only
    # the eventual ValueError. A two-phase implementation that screens
    # word/document.xml, decompresses it, and screens the rest afterwards
    # passes every one of those. The three rows below fill the remaining
    # branch x route cells; all four together are what pin the ordering.

    def test_a_LATER_declared_size_violation_is_screened_with_zero_opens(self):
        # Branch 1 (declared size) on the ".rels" route. 400,000 > max_bytes,
        # and declared size is screened before ratio, so this is unambiguously
        # the declared-size branch.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            extra_members=(("word/_rels/header1.xml.rels", b"A" * 400000),))

        def run():
            with self.assertRaisesRegex(ValueError, "declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(self._opens_during(run), [],
                         "a later .rels member was screened only after a "
                         "decompression")

    def test_a_LATER_uppercase_member_is_screened_with_zero_opens(self):
        # Ratio branch on the UPPERCASE route: 200,000 declared, highly
        # compressible, so the ratio is far past 100:1 while the declared size
        # stays under the cap.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            extra_members=(("word/HEADER1.XML", b"A" * 200000),))

        def run():
            with self.assertRaisesRegex(ValueError, "compression ratio"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(self._opens_during(run), [],
                         "an uppercase later member was screened only after a "
                         "decompression")

    def test_a_LATER_zero_compressed_size_member_is_screened_with_zero_opens(self):
        # Branch 2 (infinite ratio) on a later member. zipfile's writer cannot
        # produce compress_size == 0 for a non-empty member, so the entry is
        # FORGED -- the central directory is attacker-controlled data, which is
        # the threat being modelled, not a convenience. Appended after every
        # real member, so word/document.xml is screened long before it.
        path = self._word_shaped_docx("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
        forged = zipfile.ZipInfo("word/forged.xml")
        forged.file_size, forged.compress_size = 5000, 0
        real_infolist = zipfile.ZipFile.infolist

        def infolist_with_forged(self_archive):
            return list(real_infolist(self_archive)) + [forged]

        def run():
            with mock.patch.object(zipfile.ZipFile, "infolist",
                                   infolist_with_forged):
                with self.assertRaisesRegex(ValueError, "infinite"):
                    extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(self._opens_during(run), [],
                         "the zero-compressed-size screen ran after a member "
                         "had already been decompressed")

    # Ratio boundary rows. The high-ratio rejections above all sit far past
    # the limit, so INTEGER division --
    # file_size // compress_size -- passes every one of them and the exact
    # boundary too, while silently accepting everything between 100:1 and
    # 101:1. Both rows below use forged metadata so the ratio is exact.

    def _docx_with_forged_ratio(self, file_size, compress_size):
        path = self._word_shaped_docx("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
        forged = zipfile.ZipInfo("word/ratio.xml")
        forged.file_size, forged.compress_size = file_size, compress_size
        real_infolist = zipfile.ZipFile.infolist

        def infolist_with_forged(self_archive):
            return list(real_infolist(self_archive)) + [forged]

        return path, mock.patch.object(
            zipfile.ZipFile, "infolist", infolist_with_forged)

    def test_a_ratio_EXACTLY_at_the_limit_is_ACCEPTED(self):
        # 100000 / 1000 == 100.0, and the screen rejects only ratio > max_ratio.
        path, patched = self._docx_with_forged_ratio(100000, 1000)
        with patched:
            self.assertEqual(extract_text(path, max_bytes=384000,
                                          max_ratio=100), "x")

    def test_a_ratio_JUST_OVER_the_limit_is_REJECTED(self):
        # 100100 / 1000 == 100.1, which must be refused. Under integer
        # division it is 100, indistinguishable from the row above -- this is
        # the only row that separates them.
        path, patched = self._docx_with_forged_ratio(100100, 1000)
        with patched:
            with self.assertRaisesRegex(ValueError, "compression ratio"):
                extract_text(path, max_bytes=384000, max_ratio=100)

    # File-like sources. Every metadata-rejection row above passes a PATH. An
    # implementation that sweeps only when it opened the archive itself, and
    # skips the sweep for a file-like source, would pass all of them. The
    # rows below drive size and ratio violations through a BytesIO and an
    # open handle.
    #
    # That is not a hypothetical shape: file_tools._read_text calls
    # extract_text(io.BytesIO(payload), ...), so the file-like route IS the
    # production route for every MCP .docx read.

    def _bytesio_of(self, path):
        with open(path, "rb") as handle:
            return io.BytesIO(handle.read())

    def test_a_declared_size_violation_is_screened_through_BYTESIO(self):
        path = self._raw_docx(b"A" * 5000, compression=zipfile.ZIP_STORED)

        def run():
            with self.assertRaisesRegex(ValueError, "declared size"):
                extract_text(self._bytesio_of(path), max_bytes=1000)

        self.assertEqual(self._opens_during(run), [],
                         "a BytesIO source reached a decompression: the "
                         "metadata sweep is path-only")

    def test_a_ratio_violation_is_screened_through_BYTESIO(self):
        path = self._raw_docx(b"A" * 200000)

        def run():
            with self.assertRaisesRegex(ValueError, "compression ratio"):
                extract_text(self._bytesio_of(path), max_bytes=384000,
                             max_ratio=100)

        self.assertEqual(self._opens_during(run), [],
                         "a BytesIO source skipped the ratio screen")

    def test_a_LATER_NON_TARGET_hostile_member_is_screened_through_BYTESIO(
            self):
        # Both file-like metadata rows above make word/document.xml ITSELF
        # hostile -- and the mandatory immediate rescreen of the target
        # catches that even when the archive-wide .xml/.rels sweep is skipped
        # entirely. So a "sweep runs on the path branch only" impostor does
        # NOT turn those rows red. This row observes the file-like route's
        # archive-wide sweep directly.
        #
        # [Content_Types].xml is not the target, and _WORD_MEMBER_ORDER writes
        # it LAST, so only an archive-wide sweep reaches it before the target
        # is opened.
        #
        # Both branches are isolated below, each with its own message. A
        # single compressible payload with an ALTERNATION regex would only
        # exercise the RATIO branch and could not say which branch fired, so
        # a sweep that honours ratio but not declared size on file-like
        # sources would pass.
        #
        # ZERO-COMPRESSED-SIZE is not repeated here; its file-like coverage
        # lives in test_a_ZERO_COMPRESSED_member_is_screened_through_file_like_
        # routes. Its fixture forges the central directory by patching
        # ZipFile.infolist, which is source-independent, but source-specific
        # screening logic can still skip that branch on one route.
        #
        # The declared-size case sits BETWEEN the caller's limit and the
        # default: 200000 is over the caller's 100000 and under the library's
        # 384000, so a default-hardcoded sweep ACCEPTS the member, the tiny
        # document part then passes its rescreen, no ValueError is raised, and
        # this row goes RED. Limits equal to the defaults (384000/100) could
        # not tell a file-like branch hard-coding DEFAULT_MAX_* apart from one
        # that honours the caller.
        for label, payload, limits, expected in (
                ("ratio", b"A" * 200000,
                 dict(max_bytes=384000, max_ratio=100), "compression ratio"),
                ("declared size", os.urandom(200000),
                 dict(max_bytes=100000, max_ratio=100), "declared size")):
            path = self._word_shaped_docx(
                "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
                member_overrides={"[Content_Types].xml": payload})

            def run():
                with self.assertRaisesRegex(ValueError, expected):
                    extract_text(self._bytesio_of(path), **limits)

            self.assertEqual(
                self._opens_during(run), [],
                "%s: a hostile NON-TARGET member reached a decompression "
                "through BytesIO -- the archive-wide sweep is path-only"
                % label)

    def test_a_LATER_NON_TARGET_hostile_member_is_screened_through_a_HANDLE(
            self):
        # The open-handle twin of the row above, with the same reasoning: the
        # other handle rows are all target-hostile, so the target rescreen
        # alone satisfies them.
        #
        # Two properties make this row discriminating:
        # (a) A SPECIFIC REGEX. "compression ratio|declared size" cannot say
        #     which branch fired, so a sweep implementing only one of the two
        #     screens on the handle route would pass. This row is isolated to
        #     the declared-size branch with its own message.
        # (b) NON-DEFAULT LIMITS. 384000/100 ARE DEFAULT_MAX_BYTES and
        #     DEFAULT_MAX_RATIO, so a handle branch hard-coded to the defaults
        #     would be indistinguishable. 200000 sits between the caller's
        #     100000 and the default 384000, and os.urandom is incompressible
        #     so the ratio branch cannot fire first.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            member_overrides={"[Content_Types].xml": os.urandom(200000)})

        def run():
            with open(path, "rb") as handle:
                with self.assertRaisesRegex(ValueError, "declared size"):
                    extract_text(handle, max_bytes=100000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "a hostile NON-TARGET member reached a decompression through an "
            "open handle: the archive-wide sweep is path-only")

    def test_a_declared_size_violation_is_screened_through_an_open_HANDLE(self):
        path = self._raw_docx(b"A" * 5000, compression=zipfile.ZIP_STORED)

        def run():
            with open(path, "rb") as handle:
                with self.assertRaisesRegex(ValueError, "declared size"):
                    extract_text(handle, max_bytes=1000)

        self.assertEqual(self._opens_during(run), [],
                         "an open handle skipped the metadata screen")

    def test_the_cap_is_PER_MEMBER_and_NOT_cumulative(self):
        # extract_text's docstring promises per-member
        # screening, but every fixture's screened members sum to well under
        # max_bytes -- so an implementation keeping a RUNNING TOTAL and
        # rejecting once the sum passes the cap is indistinguishable on every
        # other row, while falsely rejecting an archive whose members are each
        # individually valid.
        filler = _bulky_xml("hdr", 40).encode("utf-8")
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>Procurement summary</w:t></w:r></w:p>",
            extra_members=(("word/header1.xml", filler),
                           ("word/header2.xml", filler),
                           ("word/header3.xml", filler)))
        # The cap is DERIVED from the fixture, never hardcoded: sit one byte
        # above the largest single screened member and below their total, so
        # the two implementations are forced apart whatever _bulky_xml emits.
        with zipfile.ZipFile(path) as archive:
            sizes = [info.file_size for info in archive.infolist()
                     if info.filename.lower().endswith((".xml", ".rels"))]
        cap = max(sizes) + 1
        self.assertLess(
            cap, sum(sizes),
            "fixture cannot separate per-member from cumulative screening")
        self.assertEqual(
            extract_text(path, max_bytes=cap, max_ratio=100),
            "Procurement summary")
        # The same fixture and derived cap through BytesIO, the route
        # file_tools uses: a file-like branch accumulating a running total
        # would pass a path-only check while falsely rejecting valid Word
        # documents in production.
        self.assertEqual(
            extract_text(self._bytesio_of(path), max_bytes=cap, max_ratio=100),
            "Procurement summary")

    def test_the_target_is_SCREENED_IMMEDIATELY_BEFORE_it_is_opened(self):
        # By design the target is screened STRUCTURALLY, on the way to
        # archive.open, so "what we open, we screened" holds by construction
        # rather than by an assertion that could never fire. Deleting the
        # second _screen_member(target, ...) call would leave every other row
        # green, because the suffix sweep already covers word/document.xml:
        # the guarantee would silently degrade from structural back to
        # coincidental.
        #
        # Record BOTH events in order and assert the open is immediately
        # preceded by a screen of the member being opened. Recording only the
        # FILENAME is not enough: an implementation can screen a harmless
        # dummy ZipInfo that merely CARRIES the name "word/document.xml", with
        # permissive limits, immediately before opening the real target --
        # and pass, because the earlier sweep supplies the actual protection.
        # So record the ZipInfo OBJECT and the limits, and assert the screened
        # object IS the object opened, with the CALLER's limits.
        #
        # Run ALL THREE source forms: an implementation that skips the
        # immediate re-screen ONLY on the file-like or handle branch would
        # otherwise leave the pre-open guarantee unverified on that route. The
        # handle is opened through a closure so that addCleanup closes it even
        # when an assertion raises -- a bare open() here would leak the
        # descriptor on every failure.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, make_source in (("path", lambda p: p),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            events = []
            real_screen = docx_text._screen_member
            real_open = zipfile.ZipFile.open

            def recording_screen(info, max_bytes, max_ratio):
                events.append(("screen", info, max_bytes, max_ratio))
                return real_screen(info, max_bytes, max_ratio)

            def recording_open(self_archive, name, *args, **kwargs):
                events.append(("open", name))
                return real_open(self_archive, name, *args, **kwargs)

            path = self._word_shaped_docx("<w:p><w:r><w:t>x</w:t></w:r></w:p>")
            with mock.patch.object(docx_text, "_screen_member",
                                   recording_screen):
                with mock.patch.object(zipfile.ZipFile, "open",
                                       recording_open):
                    # These limits are DELIBERATELY NOT the defaults. Passing
                    # 384000/100 and asserting the screen saw 384000/100
                    # cannot discriminate, because that is exactly what
                    # DEFAULT_MAX_BYTES/DEFAULT_MAX_RATIO are: an
                    # implementation that rescreens the target with the
                    # LIBRARY DEFAULTS instead of the CALLER'S limits would
                    # pass. The custom-limit rejection rows do not catch it
                    # either, because the earlier suffix sweep rejects the
                    # target first. The guarantee is that what we open, we
                    # screened WITH THE CALLER'S LIMITS. The fixture is
                    # comfortably inside 300000 bytes and 50:1 (its worst
                    # member ratio is 20.8:1), so it must still be ACCEPTED.
                    self.assertEqual(
                        extract_text(make_source(path), max_bytes=300000,
                                     max_ratio=50), "x", label)

            opens = [e for e in events if e[0] == "open"]
            self.assertEqual(len(opens), 1,
                             "%s: exactly one member may be opened: %r"
                             % (label, events))
            previous = events[events.index(opens[0]) - 1]
            self.assertEqual(previous[0], "screen",
                             "%s: the open was not immediately preceded by a "
                             "screen: %r" % (label, events))
            # IDENTITY, not name: the screened ZipInfo must BE the one opened,
            # so a same-named decoy cannot stand in for it.
            self.assertIs(previous[1], opens[0][1],
                          "%s: a different ZipInfo object was screened than "
                          "was opened" % label)
            self.assertEqual((previous[2], previous[3]), (300000, 50),
                             "%s: the CALLER's limits did not reach the "
                             "pre-open screen: %r" % (label, previous))

    def test_the_bounded_read_and_the_guard_call_hold_through_BYTESIO(self):
        # The bounded `read(max_bytes + 1)` and the exact _assert_within_cap
        # call must be observed on the BytesIO branch, not just the PATH
        # branch. A file-like-only impostor doing `member.read()` unbounded
        # and skipping the post-read guard passes every other BytesIO row:
        # metadata violations stop before the read, and the accepted fixtures
        # all sit below their declared sizes. file_tools routes production
        # DOCX reads through io.BytesIO, so these assertions have to hold on
        # that route too.
        document_bytes = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(document_bytes)
        read_sizes, guard_calls = [], []
        real_open = zipfile.ZipFile.open
        real_guard = docx_text._assert_within_cap

        def recording_open(self_archive, name, *args, **kwargs):
            handle = real_open(self_archive, name, *args, **kwargs)
            real_read = handle.read

            def recording_read(size=-1):
                read_sizes.append(size)
                return real_read(size)

            handle.read = recording_read
            return handle

        def recording_guard(payload, name, max_bytes):
            guard_calls.append((payload, name, max_bytes))
            return real_guard(payload, name, max_bytes)

        with mock.patch.object(zipfile.ZipFile, "open", recording_open):
            with mock.patch.object(
                    docx_text, "_assert_within_cap", recording_guard):
                extract_text(self._bytesio_of(path), max_bytes=1000,
                             max_ratio=100)

        self.assertEqual(read_sizes, [1001],
                         "the BytesIO route did not bound the read: %s"
                         % read_sizes)
        self.assertEqual(
            guard_calls, [(document_bytes, "word/document.xml", 1000)],
            "the BytesIO route skipped or mis-called the post-read guard")

    def test_a_hostile_ROOT_LEVEL_member_is_screened_with_zero_opens(self):
        # The ROOT-level _rels/.rels is the FIRST member of a real Word file,
        # and the most convenient place to park a bomb, but the .rels row
        # puts its payload in word/_rels/document.xml.rels. An implementation
        # screening only entries under "word/" would pass that row while
        # skipping _rels/.rels, [Content_Types].xml and everything under
        # docProps/.
        #
        # member_overrides, not extra_members: _rels/.rels is already in
        # _WORD_MEMBER_ORDER, and writing a second entry under the same name
        # makes which one a reader returns undefined.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            member_overrides={"_rels/.rels": b"A" * 200000})

        def run():
            with self.assertRaisesRegex(
                    ValueError, "compression ratio|declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(self._opens_during(run), [],
                         "a ROOT-level member escaped the metadata screen")

    def test_a_hostile_CONTENT_TYPES_member_written_LAST_is_screened(self):
        # The row above makes _rels/.rels hostile, which _WORD_MEMBER_ORDER
        # writes FIRST. A predicate screening "_rels/.rels plus anything under
        # word/" would pass that row while skipping [Content_Types].xml and
        # everything under docProps/ -- both of which are .xml parts the sweep
        # promises to cover. [Content_Types].xml is written LAST in the
        # fixture, so this also re-tests archive-wide ordering on that route.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            member_overrides={"[Content_Types].xml": b"A" * 200000})

        def run():
            with self.assertRaisesRegex(
                    ValueError, "compression ratio|declared size"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "[Content_Types].xml escaped the sweep: the predicate is scoped "
            "to _rels/.rels plus word/, not to the .xml/.rels suffixes")

    def test_the_bounded_read_and_guard_call_hold_through_an_open_HANDLE(self):
        # The open-handle twin of the BytesIO call-level row. Behavioural rows
        # alone are not enough here: stdlib ZipExtFile masks both an
        # unbounded read and an omitted post-read guard, so an implementation
        # skipping them only for real handles would pass every other row.
        document_bytes = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(document_bytes)
        read_sizes, guard_calls = [], []
        real_open = zipfile.ZipFile.open
        real_guard = docx_text._assert_within_cap

        def recording_open(self_archive, name, *args, **kwargs):
            handle = real_open(self_archive, name, *args, **kwargs)
            real_read = handle.read

            def recording_read(size=-1):
                read_sizes.append(size)
                return real_read(size)

            handle.read = recording_read
            return handle

        def recording_guard(payload, name, max_bytes):
            guard_calls.append((payload, name, max_bytes))
            return real_guard(payload, name, max_bytes)

        with open(path, "rb") as source_handle:
            with mock.patch.object(zipfile.ZipFile, "open", recording_open):
                with mock.patch.object(
                        docx_text, "_assert_within_cap", recording_guard):
                    extract_text(source_handle, max_bytes=1000, max_ratio=100)

        self.assertEqual(read_sizes, [1001],
                         "the open-handle route did not bound the read: %s"
                         % read_sizes)
        self.assertEqual(
            guard_calls, [(document_bytes, "word/document.xml", 1000)],
            "the open-handle route skipped or mis-called the post-read guard")

    def test_the_public_annotations_are_the_RETURN_annotation_only(self):
        # Keep `-> str` and leave `path` unannotated, because `path: str`
        # would be FALSE: the parameter also accepts an open handle or a
        # BytesIO. This catches either mistake: adding a `str` annotation to
        # `path`, or dropping the return one.
        self.assertEqual(extract_text.__annotations__, {"return": str})

    def test_an_archive_with_NO_word_document_xml_opens_NOTHING(self):
        # The getinfo lookup raises KeyError when the member is absent, which
        # extract_text converts to ValueError. A
        # fallback that opens "the first .xml member" when the target is
        # missing would pass every other row while parsing a member this
        # module is scoped never to touch.
        path = self._zip_path()
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("_rels/.rels", "<Relationships/>")
            archive.writestr("word/styles.xml", "<w:styles/>")

        # All three source forms: BytesIO is the production route (file_tools
        # validates once and hands the bounded bytes down), and every other
        # BytesIO and handle fixture contains word/document.xml, so a
        # file-like-specific fallback to "the first .xml member" would
        # otherwise go unnoticed.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, make_source in (("path", lambda p: p),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            def run():
                with self.assertRaises(ValueError):
                    extract_text(make_source(path))

            self.assertEqual(
                self._opens_during(run), [],
                "%s: a member was opened for an archive carrying no "
                "word/document.xml" % label)

    def test_a_RELS_member_is_screened_even_though_it_is_not_named_xml(self):
        # OPC relationship parts ARE XML, but they end in ".rels", so a screen
        # written as endswith(".xml") skips every one of them. That is not a
        # corner case: _rels/.rels is the FIRST member of a real Word file
        # (see test_word_shaped_fixture_really_is_multi_member_and_realistic),
        # which makes it the most convenient place in the archive to park a
        # bomb.
        #
        # THIS ROW IS THE DISCRIMINATING ONE FOR THE SUFFIX. Every other
        # lowercase member row above uses a ".xml" name, so all of them stay
        # green against an ".xml"-only predicate. Measured: reverting the
        # predicate to endswith((".xml",)) reds THREE rows -- this one plus
        # the uppercase and file-like rels rows.
        #
        # member_overrides, NOT extra_members: this name is already in
        # _WORD_MEMBER_ORDER, so extra_members would write a SECOND entry
        # under it. Measured: the archive then carries both an 81-byte and a
        # 200,000-byte copy, zipfile emits "UserWarning: Duplicate name", and
        # the row passes only because infolist() happens to yield both --
        # exactly what _word_shaped_docx's docstring forbids depending on.
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            member_overrides={"word/_rels/document.xml.rels": b"A" * 200000})
        with self.assertRaisesRegex(ValueError, "compression ratio|declared size"):
            extract_text(path, max_bytes=384000, max_ratio=100)

    def test_zero_compress_size_member_is_rejected_as_infinite_ratio(self):
        # A member with file_size > 0 and compress_size == 0 has an effectively
        # infinite ratio. A non-document member is never opened, so nothing
        # else catches it.
        #
        # THE DECLARED SIZE MUST STAY UNDER max_bytes. The declared-size screen
        # runs FIRST, so a lie of 500000 against a 384000 cap raises "declared
        # size ... exceeds the cap" and this test would assert on a message the
        # infinite-ratio branch never got to produce.
        path = self._zip_path()
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr(
                "word/document.xml",
                _DOCUMENT_XML_TEMPLATE.format(body="<w:p><w:r><w:t>x</w:t></w:r></w:p>"))
            archive.writestr("word/styles.xml", "unused")
        real_infolist = zipfile.ZipFile.infolist

        def lying_infolist(self_archive):
            entries = real_infolist(self_archive)
            for info in entries:
                if info.filename == "word/styles.xml":
                    info.file_size = 5000      # UNDER max_bytes, see above
                    info.compress_size = 0
            return entries

        with mock.patch.object(zipfile.ZipFile, "infolist", lying_infolist):
            with self.assertRaisesRegex(ValueError, "infinite"):
                extract_text(path, max_bytes=384000, max_ratio=100)

    def test_assert_within_cap_accepts_at_the_boundary_and_rejects_past_it(self):
        # The post-read guard is unreachable through stdlib zipfile
        # (ZipExtFile slices output to the DECLARED size), so it is
        # exercised DIRECTLY rather than
        # through a forged archive that cannot exist.
        _assert_within_cap(b"A" * 1000, "word/document.xml", 1000)
        with self.assertRaisesRegex(ValueError, "expanded past"):
            _assert_within_cap(b"A" * 1001, "word/document.xml", 1000)

    # The metadata screen has THREE independent rejection branches, and each
    # one has to be proved to run pre-decompression on its own. A single row
    # with a 200,000-byte run of "A" covers only the RATIO branch: it is under
    # the 384,000 declared-size cap (so that branch cannot fire) and
    # compresses at 943.4:1 (so the ratio branch always does). The
    # declared-size and zero-compressed-size branches could then both be
    # moved after open/read with the whole suite still green, and a regex of
    # "compression ratio|declared size" would hide it: the alternation
    # matches whichever branch fires, so the row cannot tell them apart.
    #
    # Each row below pins ONE branch, with the other two made unreachable BY
    # CONSTRUCTION, and each asserts ZERO opens. Measured: all three are green
    # against the correct implementation and RED against the
    # read-before-screen impostor.
    #
    # ZipFile.read() routes through ZipFile.open(), so patching open() catches
    # both paths. Reading the central directory does NOT go through it, so a
    # rejected archive must produce ZERO calls -- not 'fewer'.

    def _opens_during(self, callable_):
        """Run callable_, returning the member names ZipFile.open was asked
        for. Shared by every zero-opens ordering row."""
        opened = []
        real_open = zipfile.ZipFile.open

        def recording_open(self_archive, name, *args, **kwargs):
            opened.append(getattr(name, "filename", name))
            return real_open(self_archive, name, *args, **kwargs)

        with mock.patch.object(zipfile.ZipFile, "open", recording_open):
            callable_()
        return opened

    def test_DECLARED_SIZE_rejection_happens_before_any_decompression(self):
        r"""Branch 1 of 3. ZIP_STORED is what isolates it: the member is
        written uncompressed, so declared == compressed == 1001 and the ratio
        is exactly 1.0:1, far under the 100:1 limit. The ONLY branch that can
        reject this archive is the declared-size one. Measured: declared=1001,
        compressed=1001, ratio=1.0:1."""
        path = self._raw_docx(b"A" * 1001, compression=zipfile.ZIP_STORED)

        def run():
            with self.assertRaisesRegex(ValueError, "declared size"):
                extract_text(path, max_bytes=1000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "the declared-size screen ran AFTER a member was decompressed")

    def test_ZERO_COMPRESSED_SIZE_rejection_happens_before_decompression(self):
        r"""Branch 2 of 3 -- the infinite-ratio branch, which the ratio branch
        below cannot reach because it divides by compress_size and so must
        guard against zero separately.

        This shape CANNOT be produced by zipfile's writer, which computes
        compress_size from what it actually wrote; a stored empty member gets
        file_size 0 too, which the `file_size > 0` guard excludes. But the
        central directory is attacker-controlled data, so the shape is real --
        forging infolist() is the only way to reach the branch, and it is the
        threat being modelled, not a convenience."""
        path = self._raw_docx(b"<w:document/>")

        def forged_infolist(_self):
            forged = zipfile.ZipInfo("word/document.xml")
            forged.file_size = 500
            forged.compress_size = 0
            return [forged]

        def run():
            with mock.patch.object(zipfile.ZipFile, "infolist",
                                   forged_infolist):
                with self.assertRaisesRegex(ValueError, "infinite"):
                    extract_text(path, max_bytes=1000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "the zero-compressed-size screen ran AFTER a member was "
            "decompressed")

    def test_RATIO_rejection_happens_before_any_decompression(self):
        r"""Branch 3 of 3. 200,000 declared bytes is under the 384,000 cap, so
        the declared-size branch cannot fire, and compress_size is non-zero, so
        the infinite-ratio branch cannot either. Measured: declared=200000,
        compressed=212, ratio=943.4:1.

        The regex names ONLY the ratio message. Do not use an alternation
        here -- it would let this row stand in for all three branches."""
        path = self._raw_docx(b"A" * 200000)

        def run():
            with self.assertRaisesRegex(ValueError, "compression ratio"):
                extract_text(path, max_bytes=384000, max_ratio=100)

        self.assertEqual(
            self._opens_during(run), [],
            "the ratio screen ran AFTER a member was decompressed")

    def test_the_read_is_BOUNDED_and_the_post_read_guard_is_CALLED(self):
        r"""`_assert_within_cap` is also covered as an isolated helper
        (test_assert_within_cap_accepts_at_the_boundary_and_rejects_past_it),
        but that says nothing about whether extract_text ever calls it -- or
        whether the read that feeds it is bounded.

        Both call-site protections could be dropped together and every other
        row stays green, because through stdlib zipfile the guard is
        UNREACHABLE by construction: ZipExtFile slices each read to the
        member's declared remaining length, so len(payload) can never exceed a
        declared size that already passed the metadata screen. An unreachable
        guard cannot be observed through behavior. So observe the CALLS.

        Measured against the real implementation: read is called
        once with 1001 (= max_bytes + 1) and the guard is called once. Against
        an impostor doing member.read() unbounded with the guard deleted:
        read(-1) and zero guard calls -- the row goes red on both assertions."""
        document_bytes = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(document_bytes)
        read_sizes, guard_calls = [], []
        real_open = zipfile.ZipFile.open
        real_guard = docx_text._assert_within_cap

        def recording_open(self_archive, name, *args, **kwargs):
            handle = real_open(self_archive, name, *args, **kwargs)
            real_read = handle.read

            def recording_read(size=-1):
                read_sizes.append(size)
                return real_read(size)

            handle.read = recording_read
            return handle

        def recording_guard(payload, name, max_bytes):
            # Record all three arguments and assert them exactly: recording
            # only `name` would let a call on EMPTY or TRUNCATED bytes, or
            # with the wrong cap, pass.
            guard_calls.append((payload, name, max_bytes))
            return real_guard(payload, name, max_bytes)

        with mock.patch.object(zipfile.ZipFile, "open", recording_open):
            with mock.patch.object(
                    docx_text, "_assert_within_cap", recording_guard):
                extract_text(path, max_bytes=1000, max_ratio=100)

        self.assertEqual(
            read_sizes, [1001],
            "word/document.xml was not read with a bound of max_bytes + 1: %s"
            % read_sizes)
        self.assertEqual(
            guard_calls, [(document_bytes, "word/document.xml", 1000)],
            "_assert_within_cap was not called ONCE with the exact bytes just "
            "read and the caller's max_bytes: %r" % (guard_calls,))

    def test_a_member_exactly_at_the_declared_size_cap_is_ACCEPTED(self):
        r"""The rejection rows exercise both screens only PAST their limits,
        so `>=` in place of `>` would reject a member sitting exactly at
        max_bytes while every one of them stayed green. A false refusal of a
        legitimate document is the failure a cap is most likely to have."""
        body = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>x</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(body)
        with zipfile.ZipFile(path) as archive:
            exact = archive.getinfo("word/document.xml").file_size
        # Accepted AT the cap...
        self.assertEqual(extract_text(path, max_bytes=exact), "x")
        # ...and refused one byte under it. The pair is what pins the
        # comparison to the boundary rather than merely proving permissiveness.
        with self.assertRaisesRegex(ValueError, "declared size"):
            extract_text(path, max_bytes=exact - 1)

    def test_a_member_exactly_at_the_compression_ratio_cap_is_ACCEPTED(self):
        r"""The ratio half of the boundary row above.

        FLOAT EQUALITY IS DELIBERATE AND FRAGILE. This computes the ratio the
        same way the implementation does (file_size / compress_size) so the
        two are bit-identical and `ratio > max_ratio` is False at the
        boundary. If the implementation ever changes that expression -- rounds
        it, reorders it, uses a tolerance -- this row must be updated with it.
        Recorded here rather than discovered later as a flake."""
        body = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>x</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(body)
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo("word/document.xml")
        exact = info.file_size / info.compress_size
        self.assertEqual(
            extract_text(path, max_bytes=384000, max_ratio=exact), "x")
        with self.assertRaisesRegex(ValueError, "compression ratio"):
            extract_text(path, max_bytes=384000, max_ratio=exact / 2)

    # ------------------------------------------------------- non-path input
    def test_accepts_an_open_binary_handle(self):
        # CHARACTERIZATION: zipfile.ZipFile accepts a file object, so
        # extract_text does too. Pinned so the file-like call sites (file_tools
        # passes a BytesIO) cannot regress.
        path = self._docx("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        with open(path, "rb") as handle:
            self.assertEqual(extract_text(handle), "Hello world")

    def test_the_SUPPLIED_handle_is_CONSUMED_not_reopened_by_name(self):
        # The row above proves a handle is ACCEPTED. It does not prove the
        # handle is USED. An implementation that does
        # `zipfile.ZipFile(path.name)` for a real file handle -- reopening the
        # PATHNAME instead of consuming the validated object -- passes every
        # other handle row, and reintroduces exactly the TOCTOU window the
        # file-like interface exists to close: a caller that validates
        # containment and identity ONCE and hands the handle down would have
        # the name re-pointed in between.
        #
        # Proven by SWAPPING what the pathname resolves to after the handle is
        # open. A conforming implementation reads the ORIGINAL bytes through
        # the descriptor it was given; the impostor reads the replacement.
        # This is the load-bearing guarantee of the borrowed-handle interface.
        # _docx uses tempfile.mkstemp, so two calls give two distinct paths;
        # it takes ONE argument and has no `name` keyword.
        original = self._docx("<w:p><w:r><w:t>ORIGINAL</w:t></w:r></w:p>")
        decoy = self._docx("<w:p><w:r><w:t>DECOY</w:t></w:r></w:p>")

        class HandleWithDecoyName(object):
            """Delegates the file protocol to the REAL descriptor while
            reporting a DIFFERENT .name.

            Deliberately not a BufferedReader subclass: `name` is a getset
            descriptor on the C type, so assigning to it raises
            AttributeError and the row would error instead of discriminating.
            zipfile.ZipFile needs only read/seek/tell/seekable.
            """

            def __init__(self, real, name):
                self._real, self.name = real, name

            def read(self, *args):
                return self._real.read(*args)

            def seek(self, *args):
                return self._real.seek(*args)

            def tell(self):
                return self._real.tell()

            def seekable(self):
                return True

            def readable(self):
                return True

        real = open(original, "rb")
        self.addCleanup(real.close)
        self.assertEqual(
            extract_text(HandleWithDecoyName(real, decoy)), "ORIGINAL",
            "extract_text re-opened .name instead of consuming the "
            "supplied handle -- the TOCTOU window is back")
    def test_accepts_a_bytesio_view(self):
        path = self._docx("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        with open(path, "rb") as handle:
            payload = handle.read()
        self.assertEqual(extract_text(io.BytesIO(payload)), "Hello world")
    # ------------------------------------------------------ DTD / expansion
    def test_utf8_entity_bomb_is_refused_naming_the_dtd(self):
        path = self._raw_docx(_entity_bomb_xml().encode("utf-8"))
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_visible_text_cap_fires_independently_of_the_dtd_screen(self):
        # No DTD anywhere: this document is simply long. Proves the second
        # bound is real and is not the DTD screen under another name.
        body = "".join("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("x" * 100)
                       for _ in range(50))
        path = self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(body=body).encode("utf-8"))
        with self.assertRaisesRegex(ValueError, "visible text exceeded"):
            extract_text(path, max_text_chars=1000)

    # ------------------------------------------- nested paragraph ownership
    def _nested_docx(self, depth, chunk):
        """<w:p> elements nested `depth` deep, each carrying `chunk` chars. No
        DTD anywhere -- this vector needs none."""
        body = ""
        for index in range(depth):
            body = ("<w:p><w:r><w:t>%s</w:t></w:r>%s</w:p>"
                    % (("%04d" % index) * (chunk // 4), body))
        return self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(body=body).encode("utf-8"))

    def test_nested_paragraphs_do_not_re_emit_their_descendants(self):
        # THE SECOND EXPANSION VECTOR, and it needs no DTD. Element.iter()
        # yields self plus ALL descendants, so joining every descendant <w:t>
        # per <w:p> emits nested text once per ANCESTOR paragraph -- quadratic
        # in depth. Measured at the default caps: this fixture is 452 bytes
        # of archive and produced 164,039 characters that way.
        path = self._nested_docx(40, 200)
        text = extract_text(path)
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo("word/document.xml")
        # Non-vacuity: the fixture must actually pass the byte-level screens,
        # or this row would prove nothing about the traversal.
        self.assertLess(info.file_size, 384000)
        self.assertLess(info.file_size / info.compress_size, 100)
        # Each paragraph's own text appears EXACTLY once.
        self.assertEqual(len(text.splitlines()), 40)
        self.assertLessEqual(
            len(text), info.file_size,
            "visible text (%d) exceeded the bytes that carried it (%d): "
            "nested paragraphs are being re-emitted" % (len(text), info.file_size))


    def test_the_public_first_parameter_is_STILL_NAMED_path(self):
        # Every other call is POSITIONAL and the __defaults__ assertion
        # inspects values, not names -- so a rename of the first parameter
        # would pass every other row unnoticed. By design the name stays
        # `path` even though the parameter also accepts a handle or a BytesIO:
        # this row proves `extract_text(path=...)` works and that a later
        # revision cannot rename the parameter by accident, which would be a
        # silent public break for any out-of-repo caller using the keyword.
        fixture = self._docx("<w:p><w:r><w:t>Hello world</w:t></w:r></w:p>")
        self.assertEqual(extract_text(path=fixture), "Hello world")
        code = extract_text.__code__
        self.assertEqual(
            code.co_varnames[:code.co_argcount],
            ("path", "max_bytes", "max_ratio", "max_text_chars"))

    def test_the_post_read_guard_runs_BEFORE_EITHER_PARSER(self):
        r"""test_the_read_is_BOUNDED_and_the_post_read_guard_is_CALLED
        observes that the read is bounded and that the guard is CALLED -- but
        not WHERE in the sequence it runs. Moving `_assert_within_cap` to
        after the DTD prescreen and after `ET.fromstring` would keep every
        other row green, because through stdlib zipfile the guard's
        rejecting branch is UNREACHABLE:
        `ZipExtFile` slices every read to the member's declared remaining
        length, so `len(payload)` never exceeds a declared size that already
        passed the metadata screen. A guard that never fires cannot have its
        position observed -- so this row forces it to fire, by handing back a
        read one byte past the cap, and then asserts that neither parser was
        reached.

        The two `must_not_run` callables raise `AssertionError`, which is NOT
        a `ValueError` and therefore is not swallowed by `assertRaises`
        below: if either parser runs first, this test fails loudly rather
        than passing on the wrong exception.

        The BytesIO and open-handle call-level rows prove the guard is
        eventually CALLED on those routes, but not its POSITION -- so a
        file-like branch that parses before the guard would pass them. All
        three documented source forms are driven here. The handle is opened
        through a closure with addCleanup rather than a bare open(), so the
        descriptor is closed even when an assertion raises."""
        document_bytes = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>").encode("utf-8")
        path = self._raw_docx(document_bytes)

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        guard_calls = []
        real_open = zipfile.ZipFile.open
        real_guard = docx_text._assert_within_cap

        def open_returning_oversized(self_archive, name, *args, **kwargs):
            member = real_open(self_archive, name, *args, **kwargs)
            member.read = lambda *_a, **_k: b"B" * 1001
            return member

        def recording_guard(payload, name, max_bytes):
            guard_calls.append(len(payload))
            return real_guard(payload, name, max_bytes)

        def dtd_must_not_run(*args, **kwargs):
            raise AssertionError(
                "the DTD prescreen ran BEFORE the post-read size guard")

        def parse_must_not_run(*args, **kwargs):
            raise AssertionError(
                "ET.fromstring ran BEFORE the post-read size guard")

        for label, make_source in (("path", lambda q: q),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            guard_calls[:] = []
            with mock.patch.object(zipfile.ZipFile, "open",
                                   open_returning_oversized):
                with mock.patch.object(docx_text, "_assert_within_cap",
                                       recording_guard):
                    with mock.patch.object(docx_text, "_assert_no_dtd",
                                           dtd_must_not_run):
                        with mock.patch.object(docx_text.ET, "fromstring",
                                               parse_must_not_run):
                            with self.assertRaises(ValueError):
                                extract_text(make_source(path),
                                             max_bytes=1000)

            # The guard saw the oversized payload, which is what makes the two
            # must-not-run assertions above meaningful rather than vacuous.
            self.assertEqual(guard_calls, [1001], label)

    # The VISIBLE-TEXT cap through file-like sources. An implementation
    # applying max_text_chars on the path branch alone would pass every path
    # row, and file_tools._read_text calls
    # extract_text(io.BytesIO(payload), max_text_chars=MAX_FILE_BYTES) --
    # so the file-like branch is exactly where this cap has to hold.

    def _long_docx(self):
        body = "".join("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("x" * 100)
                       for _ in range(20))          # 2,000 chars of text
        return self._docx(body)

    def test_the_TEXT_cap_is_enforced_through_BYTESIO(self):
        path = self._long_docx()
        with self.assertRaisesRegex(ValueError, "visible text exceeded"):
            extract_text(self._bytesio_of(path), max_text_chars=500)

    def test_the_TEXT_cap_is_enforced_through_an_open_HANDLE(self):
        path = self._long_docx()
        with open(path, "rb") as handle:
            with self.assertRaisesRegex(ValueError, "visible text exceeded"):
                extract_text(handle, max_text_chars=500)

    def test_the_DTD_claim_is_SCOPED_in_both_docstrings(self):
        # The DTD refusal is scoped to word/document.xml, and
        # test_a_dtd_in_an_unparsed_member_is_inert verifies that BEHAVIOUR
        # only -- so a docstring asserting an archive-level guarantee would be
        # a PROSE overclaim nothing else catches. Public guards cover
        # file_tools and README; these two docstrings are checked here.
        #
        # Requiring "word/document.xml" to appear SOMEWHERE and forbidding a
        # generic phrase such as "the archive is refused" would be a no-op:
        # the retired texts, "refusal of any DTD" in the module and "Refuse a
        # member that declares a DTD" in the helper, both mention
        # word/document.xml elsewhere and contain no such phrase. So name the
        # RETIRED strings and require the REPLACEMENT wording.
        module_doc, helper_doc = docx_text.__doc__, _assert_no_dtd.__doc__

        self.assertNotIn("refusal of any DTD", module_doc,
                         "the module docstring restored the retired overclaim")
        self.assertNotIn("Refuse a member that declares a DTD", helper_doc,
                         "the helper docstring restored the retired overclaim")

        self.assertIn("scoped to word/document.xml", module_doc)
        self.assertIn("Refuse word/document.xml", helper_doc)

        for label, doc in (("module", module_doc), ("helper", helper_doc)):
            self.assertIn("word/document.xml", doc, label)
            # Both must say WHY other members need no DTD screen: they are
            # never decompressed. A docstring that narrows without the reason
            # invites the next author to "fix" the narrowing.
            self.assertTrue(
                "never" in doc and "decompress" in doc,
                "the %s docstring narrows the DTD claim without saying that "
                "other members are never decompressed" % label)

    def test_the_DTD_refusal_holds_through_an_open_HANDLE(self):
        # The open-binary-handle twin of the path and BytesIO DTD rows. An
        # implementation that skips _assert_no_dtd specifically for real file
        # handles would pass those while allowing pre-cap entity expansion
        # through a documented, supported source type.
        #
        # Asserts the refusal AND its position: ET.fromstring must not be
        # reached on this route either.
        path = self._raw_docx(_entity_bomb_xml().encode("utf-8"))

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for a DTD-declaring member through "
                "an open handle: the screen is path/BytesIO-only")

        with open(path, "rb") as handle:
            with mock.patch.object(docx_text.ET, "fromstring",
                                   must_not_be_called):
                with self.assertRaisesRegex(ValueError, "declares a DTD"):
                    extract_text(handle)

    def test_a_DTD_AFTER_a_comment_in_the_prolog_is_still_refused(self):
        # The other DTD fixtures in this suite put <!DOCTYPE immediately
        # after the XML declaration. A prescreen that feeds expat
        # only a fixed PREFIX -- Parse(payload[:N], isfinal=False) -- passes
        # every one of them, passes both order rows, and passes the fake-parser
        # row too, while a perfectly legal COMMENT or processing instruction
        # pushes the DOCTYPE past N and lets ElementTree expand it.
        #
        # The padding is a comment, which is legal in the prolog, and the whole
        # document stays well under the byte cap, so no other screen can refuse
        # it. Only the DTD screen can -- and only if it reads the whole prolog.
        padded = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            + "<!-- %s -->" % ("p" * 8000)
            + '<!DOCTYPE w:document [%s]>' % _ENTITY_LADDER
            + '<w:document xmlns:w="http://schemas.openxmlformats.org/'
              'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&e;</w:t>'
              '</w:r></w:p></w:body></w:document>')
        path = self._raw_docx(padded.encode("utf-8"))

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for a late-prolog DTD: the "
                "prescreen reads only a prefix of the payload")

        with mock.patch.object(docx_text.ET, "fromstring", must_not_be_called):
            with self.assertRaisesRegex(ValueError, "declares a DTD"):
                extract_text(path)

    def test_a_DTD_after_a_NEAR_CAP_prolog_comment_is_still_refused(self):
        # The row above pads the prolog with 8,000 bytes, so an
        # implementation feeding expat only, say, the first 16 KiB passes it
        # -- and still lets a DOCTYPE hide behind a longer, perfectly legal
        # comment. The only bound on how far down the prolog a DTD can be
        # pushed is the byte cap itself, so the fixture has to sit near that
        # cap rather than at some multiple of the shorter padding.
        #
        # The padding must ALSO resist compression. A 250,000-byte run of one
        # character is refused by the RATIO screen long before the DTD screen
        # is consulted, which would leave this row green while proving
        # nothing. A 32-bit LCG over a 64-character alphabet gives
        # deterministic, high-entropy ASCII that zlib shrinks by well under
        # 2:1, needs no import, and contains no "-", so it cannot close the
        # comment early.
        alphabet = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                    "abcdefghijklmnopqrstuvwxyz0123456789_.")

        def document(doctype, padding):
            return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    + "<!-- " + padding + " -->"
                    + doctype
                    + '<w:document xmlns:w="http://schemas.openxmlformats.org'
                      '/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>'
                      'ok</w:t></w:r></w:p></w:body></w:document>')

        # THE OFFSET IS DERIVED FROM THE DEFAULT CAP, NEVER HARD-CODED. A
        # fixed offset such as 250,000 would let a prescreen that reads a
        # 300,000-byte PREFIX of the payload pass this row while a DOCTYPE
        # placed legally at ~380 KB -- still under the 384,000-byte member
        # cap -- slipped straight through to ET.fromstring. Deriving it from
        # DEFAULT_MAX_BYTES means the row re-aims itself if the cap moves.
        doctype = "<!DOCTYPE w:document [%s]>" % _ENTITY_LADDER
        overhead = len(document(doctype, "").encode("utf-8"))
        padding_len = docx_text.DEFAULT_MAX_BYTES - overhead - 1024
        self.assertGreater(
            padding_len, 0,
            "the entity ladder no longer leaves room under DEFAULT_MAX_BYTES")

        state, chars = 12345, []
        for _ in range(padding_len):
            state = (((110351 * 10000) + 5245) * state + 12345) % ((214748 * 10000) + 3648)
            chars.append(alphabet[(state >> 16) & 63])
        padding = "".join(chars)

        # NON-VACUITY, part one: the same document WITHOUT a DOCTYPE extracts
        # cleanly. So whatever refuses the hostile twin below, it is not the
        # padding, the declared size, or the compression ratio.
        clean_path = self._raw_docx(document("", padding).encode("utf-8"))
        self.assertEqual(extract_text(clean_path), "ok")

        hostile_path = self._raw_docx(
            document(doctype, padding).encode("utf-8"))

        # NON-VACUITY, part two: the hostile member passes both metadata
        # screens on its own numbers, so the DTD screen is the only thing
        # left that can refuse it.
        with zipfile.ZipFile(hostile_path) as archive:
            info = archive.getinfo("word/document.xml")
        self.assertLess(info.file_size, docx_text.DEFAULT_MAX_BYTES)
        self.assertLess(info.file_size / info.compress_size, 100)

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for a DTD sitting just under "
                "DEFAULT_MAX_BYTES into the prolog: the prescreen reads only "
                "a PREFIX of the payload")

        # ALL THREE SOURCE FORMS. Every OTHER DTD row in this battery puts its
        # declaration near the START of the payload, so a prescreen that reads
        # a full prefix on the path route but only a short one on the
        # file-like routes would pass all of them plus a path-only version of
        # this row -- on exactly the file-like route file_tools uses.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, source in (("path", hostile_path),
                              ("BytesIO", self._bytesio_of(hostile_path)),
                              ("handle", handle_of(hostile_path))):
            with mock.patch.object(
                    docx_text.ET, "fromstring", must_not_be_called):
                with self.assertRaisesRegex(
                        ValueError, "declares a DTD",
                        msg="%s: the near-cap DTD was not refused" % label):
                    extract_text(source)







    # Every entity-bomb fixture in this suite -- including all 10 hostile
    # cells of the encoding matrix below -- is ONE internal entity ladder. An
    # implementation that refuses only DTDs CARRYING ENTITIES, or that infers
    # "a DTD is present" from evidence of expansion, passes every one of them
    # while violating the unconditional refusal of any DTD declared in
    # word/document.xml. The five rows below carry nothing to expand, so the
    # DOCTYPE screen is the only thing that can refuse them -- which is
    # exactly why the encoding matrix does NOT replace them.

    _DOCTYPE_SHELL = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '%s'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>clean body</w:t>'
        '</w:r></w:p></w:body></w:document>')

    def _docx_declaring(self, doctype):
        """A well-formed, ENTITY-FREE document carrying exactly one DOCTYPE
        declaration and nothing that expands."""
        return self._raw_docx((self._DOCTYPE_SHELL % doctype).encode("utf-8"))

    def test_a_minimal_entity_free_DTD_is_refused(self):
        # No internal subset at all -- the smallest legal DOCTYPE.
        path = self._docx_declaring('<!DOCTYPE w:document>')
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_an_entity_free_INTERNAL_SUBSET_is_refused(self):
        # An internal subset that declares no ENTITY: refusal must key on the
        # DECLARATION, not on the presence of entity definitions inside it.
        path = self._docx_declaring(
            '<!DOCTYPE w:document [<!ELEMENT w:document ANY>]>')
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_an_EXTERNAL_system_dtd_is_refused(self):
        # An external identifier has no internal subset and expands nothing
        # locally, so a screen keyed on expansion sees an ordinary document.
        # Port 9 is discard: if a future regression ever DID fetch it, the row
        # fails fast rather than hanging the suite.
        path = self._docx_declaring(
            '<!DOCTYPE w:document SYSTEM "http://127.0.0.1:9/evil.dtd">')
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_an_EXTERNAL_public_dtd_is_refused(self):
        path = self._docx_declaring(
            '<!DOCTYPE w:document PUBLIC "-//X//DTD x//EN" "evil.dtd">')
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_the_refusal_of_an_ENTITY_FREE_dtd_also_precedes_elementtree(self):
        # The order row further down uses the entity bomb, whose expansion is
        # itself a reason to stop early. This one carries nothing to expand,
        # so it pins POSITION alone: a screen relocated to after ET.fromstring
        # would still raise ValueError eventually, and no row above notices.
        path = self._docx_declaring('<!DOCTYPE w:document>')

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for an entity-free DTD: the "
                "screen ran too late, or not at all")

        with mock.patch.object(docx_text.ET, "fromstring", must_not_be_called):
            with self.assertRaisesRegex(ValueError, "declares a DTD"):
                extract_text(path)

    def test_a_dtd_followed_by_MALFORMED_content_still_names_the_DTD(self):
        # This row does NOT prove the handler ABORTS: an implementation can
        # record the DTD, let expat expand the internal subset, CATCH the
        # ExpatError this fixture then provokes, and raise the expected
        # ValueError -- passing this row while performing the expansion. **It
        # is NOT proof of position.** It is a characterization row -- the DTD
        # refusal must win over a well-formedness error occurring later in the
        # same document -- and the row below is the actual proof.
        malformed_after_dtd = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<!DOCTYPE w:document [%s]>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&e;'
            % _ENTITY_LADDER)
        path = self._raw_docx(malformed_after_dtd.encode("utf-8"))
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            extract_text(path)

    def test_the_DOCTYPE_handler_RAISES_so_the_parse_CANNOT_continue(self):
        # Every other DTD row -- including both ET.fromstring order rows and
        # the malformed row above -- proves
        # only that _assert_no_dtd RETURNS before something else happens. None
        # of them can see an implementation that installs the handler, lets
        # expat.Parse run to COMPLETION, and raises afterwards, expanding the
        # entity subset on the way. That expansion is the entire attack.
        #
        # Position inside a parse is only observable from INSIDE the parse, so
        # ask the parser. This fake stands in for expat: it calls whatever the
        # module installed as StartDoctypeDeclHandler and then records whether
        # control came BACK to it. A correct handler raises, so `resumed` stays
        # empty; a handler that merely records returns normally, and this is
        # the only row in the suite that notices.
        resumed = []

        class FakeParser(object):
            def __init__(self):
                self.StartDoctypeDeclHandler = None

            def Parse(self, data, isfinal=False):
                if self.StartDoctypeDeclHandler is None:
                    resumed.append("no StartDoctypeDeclHandler was installed")
                    return 1
                self.StartDoctypeDeclHandler("w:document", None, None, 1)
                resumed.append("handler returned instead of raising")
                return 1

        with mock.patch("xml.parsers.expat.ParserCreate",
                        lambda *a, **k: FakeParser()):
            with self.assertRaisesRegex(ValueError, "declares a DTD"):
                _assert_no_dtd(b"<w:document/>", "word/document.xml")

        self.assertEqual(
            resumed, [],
            "StartDoctypeDeclHandler returned control to the parser: the "
            "screen RECORDS the DTD rather than aborting, so expat would "
            "continue into the internal subset")






    def test_entity_bomb_through_a_bytesio_view_is_refused(self):
        # The exact shape file_tools' read path hands down.
        #
        # Asserting ONLY the eventual ValueError is not enough: a BytesIO
        # branch that calls ET.fromstring BEFORE _assert_no_dtd expands the
        # bomb and then raises the very message asserted here, so the row
        # would stay green while the parser it exists to keep away from the
        # payload had already run. Like the path and handle rows, this one
        # carries a must-not-run guard, on the production route.
        path = self._raw_docx(_entity_bomb_xml().encode("utf-8"))
        with open(path, "rb") as handle:
            payload = handle.read()

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for a BytesIO entity bomb: the "
                "DTD refusal runs AFTER the parse on this source form")

        with mock.patch.object(docx_text.ET, "fromstring", must_not_be_called):
            with self.assertRaisesRegex(ValueError, "declares a DTD"):
                extract_text(io.BytesIO(payload))

    def test_the_bomb_really_does_pass_every_byte_level_screen(self):
        # NON-VACUITY GUARD for the four rows above. If the fixture ever became
        # something the declared-size or ratio screen catches on its own, those
        # rows would pass without the DTD guard existing at all.
        path = self._raw_docx(_entity_bomb_xml().encode("utf-8"))
        with zipfile.ZipFile(path) as archive:
            info = archive.getinfo("word/document.xml")
        self.assertLess(info.file_size, 384000)
        self.assertLess(info.file_size / info.compress_size, 100)

    def test_the_word_DOCTYPE_in_visible_text_is_not_a_dtd(self):
        # Pins the screen to the DECLARATION, not the vocabulary.
        path = self._docx("<w:p><w:r><w:t>the DOCTYPE keyword is discussed"
                          "</w:t></w:r></w:p>")
        self.assertEqual(extract_text(path), "the DOCTYPE keyword is discussed")

    def test_cdata_wrapping_a_doctype_is_not_a_dtd_either(self):
        # A raw-byte screen FALSE-REJECTS this; the structural screen does not,
        # because expat knows CDATA content is character data.
        path = self._raw_docx(_DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t><![CDATA[<!DOCTYPE x>]]></w:t></w:r></w:p>"
        ).encode("utf-8"))
        self.assertEqual(extract_text(path), "<!DOCTYPE x>")

    def test_assert_no_dtd_accepts_a_clean_member_and_rejects_a_declared_one(self):
        _assert_no_dtd(
            _DOCUMENT_XML_TEMPLATE.format(
                body="<w:p><w:r><w:t>x</w:t></w:r></w:p>").encode("utf-8"),
            "word/document.xml")
        with self.assertRaisesRegex(ValueError, "declares a DTD"):
            _assert_no_dtd(_entity_bomb_xml().encode("utf-8"),
                           "word/document.xml")

    def test_a_dtd_in_an_unparsed_member_is_inert(self):
        r"""Pins the NARROWED claim, and makes it mechanical.

        _assert_no_dtd has exactly ONE call site, on word/document.xml. A DTD
        in any other member is inert BECAUSE no other member is ever
        decompressed or parsed -- not because it was screened. Saying the
        archive "is refused if it declares a DTD" overstates that.

        The payload here is the SAME entity ladder that IS refused outright in
        word/document.xml, so this row cannot pass by using a weak bomb. It
        still passes the per-member metadata screens (468 declared bytes at
        2.07:1), so it reaches the point where scope is the only thing
        deciding the outcome.

        A future change that begins decompressing or parsing another member
        turns this green row red, rather than silently widening the gap
        between what the docstring claims and what the code does.

        The return value alone CANNOT express that claim.
        An implementation that opens word/styles.xml, parses it, expands its
        entity ladder and discards the result still returns "clean body" --
        having done precisely the work this row is supposed to prove never
        happens. "Never decompressed or parsed" is a statement about CALLS, so
        the calls are what must be asserted."""
        path = self._word_shaped_docx(
            "<w:p><w:r><w:t>clean body</w:t></w:r></w:p>",
            member_overrides={"word/styles.xml": _entity_bomb_xml()})
        opened, parsed = [], []
        real_open = zipfile.ZipFile.open
        real_fromstring = docx_text.ET.fromstring

        def recording_open(self_archive, name, *args, **kwargs):
            opened.append(getattr(name, "filename", name))
            return real_open(self_archive, name, *args, **kwargs)

        def recording_fromstring(text, *args, **kwargs):
            parsed.append(len(text))
            return real_fromstring(text, *args, **kwargs)

        with mock.patch.object(zipfile.ZipFile, "open", recording_open):
            with mock.patch.object(
                    docx_text.ET, "fromstring", recording_fromstring):
                self.assertEqual(extract_text(path), "clean body")
        # EXACTLY one member decompressed, and it is the document.
        self.assertEqual(
            opened, ["word/document.xml"],
            "a member other than word/document.xml was decompressed: %s"
            % opened)
        # EXACTLY one parse. Two would mean the bomb member reached the parser
        # even if its result was thrown away.
        self.assertEqual(
            len(parsed), 1,
            "more than one member was parsed: %d parses" % len(parsed))

    def test_malformed_xml_becomes_a_valueerror(self):
        # The prescreen converts ExpatError to ValueError, so malformed input
        # surfaces as ValueError rather than ElementTree's ParseError.
        #
        # All three source forms: the ValueError contract is advertised to
        # EVERY caller, and file_tools reaches this module through BytesIO in
        # production, so an implementation that converts the error only on
        # the path branch and leaks it for BytesIO or a handle must fail.
        path = self._raw_docx(b"<w:document><unclosed>")

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, make_source in (("path", lambda p: p),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            with self.assertRaisesRegex(ValueError, "not well-formed",
                                        msg=label):
                extract_text(make_source(path))

    # ----------------------------------------------------- ORDER OF OPERATIONS
    # Every screen in this module derives its value from WHERE it runs, and a
    # row that asserts only the eventual exception cannot see position at all.
    # These rows exist because moving a guard after the act it guards would
    # otherwise leave the entire suite green.

    def test_the_dtd_refusal_runs_BEFORE_elementtree_builds_the_tree(self):
        r"""The basic bomb rows assert only the eventual ValueError, so
        relocating _assert_no_dtd to AFTER ET.fromstring would leave all of
        them green -- while performing the entity expansion the guard exists
        to prevent. The refusal's whole value is its POSITION, and position
        is not observable from an exception.

        So make ET.fromstring detonate if it is reached at all. A correct
        implementation never reaches it for a DTD-declaring member; the
        impostor raises RuntimeError and this row fails on the exception
        TYPE, which is the signal."""
        path = self._raw_docx(_entity_bomb_xml().encode("utf-8"))

        def must_not_be_called(*args, **kwargs):
            raise RuntimeError(
                "ET.fromstring was reached for a member declaring a DTD: "
                "the DTD screen ran too late, or not at all")

        with mock.patch.object(docx_text.ET, "fromstring", must_not_be_called):
            with self.assertRaisesRegex(ValueError, "declares a DTD"):
                extract_text(path)

    def test_undeclared_namespace_prefix_is_a_ValueError_not_a_ParseError(self):
        r"""test_malformed_xml_becomes_a_valueerror covers only an error the
        EXPAT prescreen catches. The prescreen parser is created WITHOUT
        namespace processing, so an undeclared namespace
        prefix is well-formed to expat and reaches ET.fromstring, which
        rejects it with ParseError -- a SyntaxError subclass, NOT the
        ValueError this module's contract advertises to every caller.

        This is the seam BETWEEN the two parsers, and it is the only place the
        contract can leak: ET.fromstring must be wrapped, not just the
        prescreen.

        THIS is the row that reaches ET.fromstring, so it drives all three
        source forms. The all-route malformed row above does NOT exercise the
        wrapper at all: its unclosed XML is rejected by the expat prescreen
        before ET.fromstring is ever called. An implementation that wraps
        ParseError only on the path branch and leaks it for BytesIO or an
        open handle would otherwise pass -- including on the BytesIO route
        file_tools uses in production.
        """
        path = self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(
                body="<w:p><w:r><zz:t>x</zz:t></w:r></w:p>").encode("utf-8"))

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        # assertRaises(ValueError) FAILS on ParseError, which is the point.
        for label, make_source in (("path", lambda p: p),
                                   ("bytesio", self._bytesio_of),
                                   ("handle", handle_of)):
            with self.assertRaisesRegex(
                    ValueError, "not well-formed|could not be parsed",
                    msg=label):
                extract_text(make_source(path))




    # ------------------------------------------------------------------
    # FILE-LIKE ROUTES AND CORRUPT INPUT. Every row below drives a FILE-LIKE
    # source (BytesIO / open handle) or corrupt input. file_tools routes the
    # production read through exactly those forms, so a control verified
    # only on the path route is not verified where it matters.
    # ------------------------------------------------------------------

    def test_MID_READ_corruption_FAILS_CLOSED_on_all_three_routes(self):
        """Corruption that surfaces mid-READ, not at open time.

        The non-ZIP row (test_a_CORRUPT_archive_FAILS_CLOSED_and_returns_no_
        text) proves refusal only while *opening* an archive. Measured: with
        one byte flipped inside the deflate stream, archive.open() SUCCEEDS
        and the failure arrives at member.read(). An implementation catching
        that and returning partial or empty text passes the open-time row,
        then hands the model a short digest instead of a refusal, on the
        production read path.

        Measured -- all three source forms behave identically:
            zlib.error: Error -3 while decompressing data:
            invalid distance too far back

        NOTE THE TYPES -- there are TWO, and they are not interchangeable.
        Flipping a byte in a DEFLATE stream only yields zlib.error, which
        never exercises the CRC-32 check, so both compressions are needed:

            ZIP_STORED   -> zipfile.BadZipFile: Bad CRC-32 for file '...'
            ZIP_DEFLATED -> zlib.error: Error -3 ... invalid distance too far

        Neither is a ValueError, so a (ValueError, BadZipFile) tuple would
        not even catch the deflate case. An implementation catching only one
        of them and returning empty or partial text passes the other, which is
        why both compressions are driven here, on all three source forms.

        **MID-MEMBER TRUNCATION IS NOT CONSTRUCTIBLE and is deliberately not
        claimed.** Removing bytes from inside a member while leaving the
        central directory intact desynchronises every following offset;
        measured, it produces `ValueError: negative seek value -500` -- an
        artifact of the malformed container, not the failure mode being
        tested. Truncating the FILE removes the central directory and fails
        at *open* time, which the open-time corruption row already covers.
        So this row claims the two mid-READ modes that can be produced
        deterministically, and says plainly that it claims no third.

        What this row pins is that the call RAISES rather than RETURNS -- the
        fail-closed property. The un-normalised exception types are a known
        residual rather than silently widened away. **Note also that
        `Bad CRC-32` is attacker-controlled**, so this row is evidence about
        not-swallowing-errors, never a security guarantee.
        """
        document_bytes = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>hello world</w:t></w:r></w:p>" * 200
        ).encode("utf-8")

        def corrupt_copy(compression):
            """A valid archive with ONE byte flipped inside the member data.
            The central directory is untouched, so the archive opens cleanly
            and only the read fails."""
            good = self._raw_docx(document_bytes, compression=compression)
            raw = bytearray(open(good, "rb").read())
            raw[30 + len("word/document.xml") + 50] ^= 0xFF
            bad = self._zip_path()
            with open(bad, "wb") as out:
                out.write(bytes(raw))
            return bad

        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        # BOTH deterministic mid-read failure modes, each on ALL THREE source
        # forms. Measured, and they are genuinely different errors:
        #   ZIP_STORED   -> zipfile.BadZipFile: Bad CRC-32 for file '...'
        #   ZIP_DEFLATED -> zlib.error: Error -3 ... invalid distance too far back
        # An implementation catching only ONE of them and returning empty or
        # partial text passes the other case -- which is why both are here.
        for kind, compression in (("CRC-32", zipfile.ZIP_STORED),
                                  ("deflate stream", zipfile.ZIP_DEFLATED)):
            bad = corrupt_copy(compression)
            for label, source in (("path", bad),
                                  ("BytesIO", self._bytesio_of(bad)),
                                  ("handle", handle_of(bad))):
                with self.assertRaises(
                        (ValueError, zipfile.BadZipFile, zlib.error),
                        msg="%s/%s: mid-read corruption RETURNED instead of "
                            "raising -- the model receives a truncated digest"
                            % (kind, label)):
                    extract_text(source)

    def test_a_ZERO_COMPRESSED_member_is_screened_through_file_like_routes(self):
        # The infinite-ratio branch on the file-like routes. The
        # forged-infolist patch is source-independent, but that does not make
        # a file-like repetition redundant: source-specific screening logic
        # can skip the branch entirely no matter where the forged metadata
        # came from.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, wrap in (("BytesIO", self._bytesio_of),
                            ("handle", handle_of)):
            path, forged = self._docx_with_forged_ratio(200000, 0)
            with forged:
                def run():
                    with self.assertRaisesRegex(
                            ValueError, "compression ratio is infinite"):
                        extract_text(wrap(path))

                self.assertEqual(
                    self._opens_during(run), [],
                    "%s: a zero-compressed member was opened before the "
                    "infinite-ratio screen refused it" % label)

    def test_UPPERCASE_and_RELS_members_are_screened_through_file_like_routes(
            self):
        # The uppercase and .rels suffix rows above use paths, and every other
        # file-like row uses a lowercase .xml member. A file-like branch using
        # a case-SENSITIVE endswith(".xml") would therefore pass all of them
        # while skipping uppercase parts and every relationship part.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for member in ("word/HEADER1.XML", "word/_rels/DOCUMENT.XML.RELS"):
            for label, wrap in (("BytesIO", self._bytesio_of),
                                ("handle", handle_of)):
                path = self._word_shaped_docx(
                    "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
                    extra_members=((member, b"A" * 200000),))

                def run():
                    with self.assertRaisesRegex(ValueError, "compression ratio"):
                        extract_text(wrap(path))

                self.assertEqual(
                    self._opens_during(run), [],
                    "%s/%s escaped the file-like sweep -- a case-sensitive "
                    "or .xml-only predicate" % (label, member))

    def test_LATER_members_use_the_CALLERS_RATIO_through_file_like_routes(self):
        # The file-like DECLARED-SIZE rows above use a caller cap (200000
        # against 100000). This is the RATIO counterpart: the other file-like
        # ratio fixtures sit at 384000/100 -- the library DEFAULTS -- and so
        # cannot tell a default-hardcoded sweep apart. Forged 200000/4000 =
        # 50:1 is OVER a caller cap of 25 and UNDER the default of 100, so
        # only a sweep reading the caller's value refuses it. file_size stays
        # below both size caps, so the size branch cannot fire first.
        def handle_of(p):
            fh = open(p, "rb")
            self.addCleanup(fh.close)
            return fh

        for label, wrap in (("BytesIO", self._bytesio_of),
                            ("handle", handle_of)):
            path, forged = self._docx_with_forged_ratio(200000, 4000)
            with forged:
                def run():
                    with self.assertRaisesRegex(ValueError, "compression ratio"):
                        extract_text(wrap(path), max_bytes=384000, max_ratio=25)

                self.assertEqual(
                    self._opens_during(run), [],
                    "%s: the file-like sweep used DEFAULT_MAX_RATIO instead "
                    "of the caller's 25" % label)



    # ==================================================================
    # ROUTE PARITY -- the sweep that CLOSES the class.
    #
    # A recurring bug shape: "this control is verified on the PATH route
    # and not on the file-like routes." One more row per control fixes an
    # instance and leaves the next one open, so the SHAPE is closed once,
    # here, with a rule rather than an instance.
    #
    # The two rows below are complementary and neither is redundant --
    # measured against named impostor implementations:
    #   the census   catches a screen ADDED or DELETED
    #   the sweep    catches a screen that BEHAVES DIFFERENTLY per route
    #                or is skipped on one
    # ==================================================================

    # ==================================================================

    def test_EVERY_refusal_site_has_a_route_parity_case(self):
        """A CENSUS of docx_text.py's refusal sites, so that adding a
        screen without route coverage cannot pass silently.

        This is what makes the sweep below a SEAL rather than a habit. A
        hand-written list of screens is only as good as the next author's
        memory; an AST census of every `raise ValueError` in the shipped
        module fails the moment the module grows or loses one.

        Measured: ELEVEN sites, each matched by exactly one token below. RED
        at `12 != 11` against an added screen and `10 != 11` against a deleted
        one.

        This row does NOT claim any screen is CORRECT -- only that every one
        is accounted for by the sweep. An added screen is the case the sweep
        alone misses (no existing case can reach it), which is why both
        exist.
        """
        # LOCAL import: the module header does not import `ast`, and this is
        # the only row in this file that needs it.
        import ast

        source = io.open(docx_text.__file__, "r", encoding="utf-8").read()
        found = []
        for node in ast.walk(ast.parse(source)):
            if not isinstance(node, ast.Raise):
                continue
            exc = node.exc
            if not (isinstance(exc, ast.Call)
                    and isinstance(exc.func, ast.Name)
                    and exc.func.id == "ValueError"):
                continue
            literals = [n.value for n in ast.walk(exc)
                        if isinstance(n, ast.Constant)
                        and isinstance(n.value, str)]
            self.assertTrue(
                literals,
                "a ValueError at line %d carries no literal message, so no "
                "sweep case can be keyed to it" % node.lineno)
            # Adjacent string literals are folded by the parser into ONE
            # Constant, so the longest literal in the call IS the format
            # string. Verified on all eleven sites.
            found.append(max(literals, key=len))

        self.assertEqual(
            len(found), len(self._REFUSAL_SITES),
            "docx_text.py raises ValueError from %d sites but _REFUSAL_SITES "
            "enumerates %d -- a screen was added or removed without a "
            "route-parity case" % (len(found), len(self._REFUSAL_SITES)))

        for site_id, token in self._REFUSAL_SITES:
            matches = [m for m in found if token in m]
            self.assertEqual(
                len(matches), 1,
                "token %r for site %r matched %d refusal messages, expected "
                "exactly 1" % (token, site_id, len(matches)))

    def test_every_screen_behaves_IDENTICALLY_on_path_BytesIO_and_handle(self):
        """THE ROUTE-PARITY SWEEP. Every screen, every source form, one row.

        For each case the archive is built ONCE and handed to extract_text
        three ways -- path, BytesIO, open handle -- and the three OUTCOMES
        are required to be equal to each other AND to match the expected
        refusal. file_tools routes the production read through the file-like
        forms, so a control that holds only on the path route is not a
        control on the path that ships.

        **WHY OUTCOME EQUALITY AND NOT assertRaises.** Measured against an
        impostor in which file-like sources skip screening: all three routes
        RAISE, so an assertRaises row PASSES it. They raise DIFFERENT
        refusals, because the file-like routes fall through to the post-read
        guard instead of refusing on metadata first -- one route refuses
        before reading, the other reads first. Only comparing the outcomes
        to EACH OTHER sees that.

        **BOTH DIRECTIONS ARE LOAD-BEARING.** The refusal cases catch a
        screen that fails to fire; the acceptance cases catch one that fires
        when it should not. A cap tested only from above cannot tell `>`
        from `>=`, and a `>=` slip does not let bombs through -- it REJECTS
        VALID DOCUMENTS on the deployed route.
        """
        cases = self._parity_cases()
        # Only REFUSAL cases count as driving a site. A site carrying only a
        # boundary-acceptance case would otherwise look covered while
        # nothing ever proved it refuses.
        driven = {c["id"] for c in cases
                  if c["id"] and not c.get("accepts_boundary")}
        # PINNED so that exempting a screen is a deliberate edit to a literal
        # rather than a quiet dictionary entry. Without this, a future screen
        # could be "covered" by adding it to _NOT_ROUTE_REACHABLE and the
        # assertion below would still pass -- a seal by discipline, not a
        # seal. VERIFIED by adding a second entry: RED, "Items in the first
        # set but not the second".
        self.assertEqual(
            set(self._NOT_ROUTE_REACHABLE), {"post_read_cap"},
            "the route-unreachable exemption list changed; a screen may have "
            "been exempted instead of covered")
        self.assertEqual(
            driven | set(self._NOT_ROUTE_REACHABLE),
            {site_id for site_id, _token in self._REFUSAL_SITES},
            "a refusal site is neither driven across the three routes nor "
            "recorded in _NOT_ROUTE_REACHABLE with a reason")
        # EVERY CAP SCREEN OWES ITS BOUNDARY AN ACCEPTANCE, not just its
        # breach a refusal. A screen exercised only with over-limit input
        # cannot tell `>` from `>=` -- and the `>=` slip does not let bad
        # documents through, it REJECTS GOOD ONES, which no refusal case
        # can ever notice.
        self.assertEqual(
            {c["id"] for c in cases if c.get("accepts_boundary")},
            self._CAP_SITES,
            "a cap screen has no exactly-at-the-limit acceptance case, so "
            "nothing in this battery separates `>` from `>=` on any route")

        for case in cases:
            made = case["make"]()
            path, patch = made if isinstance(made, tuple) else (made, None)
            outcomes, opens = {}, {}
            for label in ("path", "BytesIO", "handle"):
                if patch is None:
                    outcome, opened, alive = self._route_outcome(
                        path, label, case["kwargs"])
                else:
                    with patch:
                        outcome, opened, alive = self._route_outcome(
                            path, label, case["kwargs"])
                outcomes[label] = outcome
                opens[label] = opened
                if label == "handle":
                    self.assertTrue(
                        alive,
                        "%s: extract_text CLOSED the caller's handle. It "
                        "borrows the handle, it does not own it -- and "
                        "file_tools hands down a handle it still needs"
                        % case["label"])

            self.assertEqual(
                outcomes["path"], outcomes["BytesIO"],
                "%s: the BytesIO route did not behave identically to the "
                "path route" % case["label"])
            self.assertEqual(
                outcomes["path"], outcomes["handle"],
                "%s: the open-handle route did not behave identically to "
                "the path route" % case["label"])

            kind, detail = outcomes["path"]
            # A case ACCEPTS if it names no refusal site at all, or if it is
            # a cap site's exactly-at-the-limit boundary.
            if case["id"] is None or case.get("accepts_boundary"):
                self.assertEqual((kind, detail), ("returned", case["expect"]),
                                 "%s was not accepted" % case["label"])
            else:
                self.assertEqual(kind, "raised",
                                 "%s RETURNED instead of raising: %r"
                                 % (case["label"], detail))
                self.assertIn(case["expect"], detail,
                              "%s raised the wrong refusal" % case["label"])

            if case["pre_open"]:
                for label in ("path", "BytesIO", "handle"):
                    self.assertEqual(
                        opens[label], [],
                        "%s/%s: a member was decompressed before the "
                        "metadata screen refused the archive"
                        % (case["label"], label))

    # ------------------------- helpers and data for the two rows above

    def _raw_docx_with_members(self, document_bytes, extra_members,
                               compression=zipfile.ZIP_STORED):
        """`_raw_docx` writes word/document.xml ALONE; this variant adds
        extra members after it. ZIP_STORED by default so declared ==
        compressed and the ratio branch cannot fire while a size or
        acceptance case is being measured."""
        path = self._zip_path()
        with zipfile.ZipFile(path, "w", compression) as archive:
            archive.writestr("word/document.xml", document_bytes)
            for name, payload in extra_members:
                archive.writestr(name, payload)
        return path

    def _not_a_zip(self):
        path = self._zip_path()
        with open(path, "wb") as handle:
            handle.write(b"not a zip archive at all")
        return path

    def _route_source(self, path, label):
        """One archive, three ways in. A FRESH BytesIO/handle per call --
        ZipFile seeks, so a reused file-like source is not a second
        independent read."""
        if label == "path":
            return path
        if label == "BytesIO":
            return self._bytesio_of(path)
        handle = open(path, "rb")
        # Registered AFTER _track(path) inside make(), so LIFO cleanup
        # closes the handle BEFORE the path is removed. A tearDown that
        # deletes these paths would run while the handle is still open and
        # fail on Windows with PermissionError [WinError 32].
        self.addCleanup(handle.close)
        return handle

    def _route_outcome(self, path, label, kwargs):
        """Run extract_text through ONE route and record what happened.

        Returns ((kind, detail), opened_member_names) where kind is
        "returned" or "raised". Catching bare Exception is deliberate: an
        unexpected exception TYPE differing between routes is itself a
        parity failure, and normalising it here would hide exactly what
        this row exists to find."""
        opened = []
        real_open = zipfile.ZipFile.open

        def recording_open(self_archive, name, *a, **kw):
            opened.append(getattr(name, "filename", name))
            return real_open(self_archive, name, *a, **kw)

        source = self._route_source(path, label)
        with mock.patch.object(zipfile.ZipFile, "open", recording_open):
            try:
                outcome = ("returned", extract_text(source, **kwargs))
            except Exception as exc:
                outcome = ("raised", "%s: %s" % (type(exc).__name__, exc))
        # A caller-owned handle must SURVIVE the call, on success and on
        # refusal alike -- extract_text BORROWS it, it does not consume it.
        # The surrounding cleanups tolerate a double close, so without this
        # check an implementation that closed the caller's handle would pass
        # every row. The addCleanup-based temp-file cleanup in this file also
        # depends on this contract: it exists BECAUSE a correct extract_text
        # leaves caller handles open.
        still_open = (not source.closed) if label == "handle" else None
        return outcome, opened, still_open

    def _parity_cases(self):
        """One entry per REACHABLE refusal site, plus one acceptance.

        Every `expect` string below is VERBATIM from the measured run, not
        predicted. Two of them are deliberately narrow: "declared size"
        alone also matches the infinite-ratio and post-read messages, and
        "compression ratio" alone also matches the infinite-ratio message,
        so both carry enough of the rendered text to name ONE site."""
        long_body = "<w:p><w:r><w:t>hello</w:t></w:r></w:p>" * 50
        return (
            dict(id=None, label="a clean Word-shaped document",
                 make=lambda: self._word_shaped_docx(
                     "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"),
                 kwargs={}, expect="hello", pre_open=False),

            # ---- BOUNDARY ACCEPTANCE. Every cap screen rejects on `>`, so
            # the value EXACTLY AT the cap must be ACCEPTED, on every route.
            # A `>=` slip confined to the file-like branch REJECTS
            # LEGITIMATE DOCUMENTS on the route file_tools actually uses,
            # while every over-limit refusal case still passes.
            # Each literal below is MEASURED, not chosen.
            dict(id="declared_size", accepts_boundary=True,
                 label="declared size EXACTLY at the cap",
                 make=lambda: self._raw_docx(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
                     ).encode("utf-8"), compression=zipfile.ZIP_STORED),
                 # MEASURED: the member is 206 bytes and ZIP_STORED, so
                 # declared == compressed and the ratio is exactly 1.00 --
                 # the declared-size branch is the ONLY one that can act.
                 # At max_bytes=205 it refuses; at 206 it must accept.
                 kwargs=dict(max_bytes=206), expect="hello", pre_open=False),
            dict(id="ratio_over", accepts_boundary=True,
                 label="compression ratio EXACTLY at the limit",
                 # 100000 / 1000 == 100.0 against max_ratio=100.
                 make=lambda: self._docx_with_forged_ratio(100000, 1000),
                 kwargs=dict(max_bytes=384000, max_ratio=100),
                 expect="x", pre_open=False),
            dict(id="char_cap", accepts_boundary=True,
                 label="visible text EXACTLY at the character cap",
                 make=lambda: self._raw_docx(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>" * 4
                     ).encode("utf-8")),
                 # MEASURED: 4 paragraphs x 5 chars + 3 separators = 23.
                 # MULTI-paragraph deliberately -- a single paragraph never
                 # charges a separator, so it could not catch a cap that
                 # ignores separators. At max_text_chars=22 it refuses; at
                 # 23 it accepts.
                 kwargs=dict(max_text_chars=23),
                 expect="\n".join(["hello"] * 4), pre_open=False),

            # ---- ACCEPTANCE THAT IS NOT A BOUNDARY. Both of these surfaces
            # are exercised ONLY by hostile fixtures elsewhere in the
            # battery, so an over-refusing implementation passes every
            # existing row while rejecting valid Word output.
            dict(id=None, label="BENIGN uppercase .XML and .RELS members",
                 # Every other uppercase fixture in this battery is
                 # hostile, and the clean Word-shaped fixture is entirely
                 # lowercase -- so a blanket refusal of uppercase OPC parts
                 # would pass everything else. Uppercase part names are
                 # VALID.
                 make=lambda: self._raw_docx_with_members(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
                     ).encode("utf-8"),
                     (("word/HEADER1.XML", self._CLEAN_PART),
                      ("word/_rels/DOCUMENT.XML.RELS", self._CLEAN_PART))),
                 kwargs={}, expect="hello", pre_open=False),
            dict(id=None, label="screening is PER MEMBER, not cumulative",
                 # The per-member acceptance row covers path and BytesIO
                 # only, and the generic clean case cannot separate
                 # per-member from cumulative screening because its total
                 # stays under the default cap.
                 make=lambda: self._raw_docx_with_members(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
                     ).encode("utf-8"),
                     tuple(("word/part%d.xml" % i,
                            b"<r>" + b"z" * 3000 + b"</r>")
                           for i in range(5))),
                 # MEASURED: 6 members, largest 3007, SUM 15241. The cap sits
                 # ABOVE the largest member and far BELOW the sum, so ONLY a
                 # running-total implementation refuses this archive.
                 kwargs=dict(max_bytes=3017), expect="hello", pre_open=False),

            dict(id="compression_method", label="a BZIP2 member",
                 make=lambda: self._raw_docx(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body="<w:p><w:r><w:t>%s</w:t></w:r></w:p>"
                         % " ".join(str(n * 7919) for n in range(2000))
                     ).encode("utf-8"), compression=zipfile.ZIP_BZIP2),
                 kwargs={}, expect="uses compression method 12",
                 pre_open=True),
            dict(id="unreadable_archive", label="bytes that are not a zip",
                 make=lambda: self._not_a_zip(),
                 kwargs={}, expect="could not be read as a zip archive",
                 pre_open=True),
            dict(id="declared_size", label="declared size over the cap",
                 make=lambda: self._raw_docx(b"A" * 5000,
                                             compression=zipfile.ZIP_STORED),
                 kwargs=dict(max_bytes=1000),
                 expect="declared size 5000 exceeds the 1000 byte cap",
                 pre_open=True),
            dict(id="ratio_over", label="compression ratio over the limit",
                 make=lambda: self._raw_docx(b"A" * 200000),
                 kwargs={}, expect=":1 exceeds the 100:1 limit",
                 pre_open=True),
            dict(id="infinite_ratio", label="zero-compressed member",
                 make=lambda: self._docx_with_forged_ratio(200000, 0),
                 kwargs={}, expect="compression ratio is infinite",
                 pre_open=True),
            dict(id="dtd", label="a DOCTYPE in word/document.xml",
                 make=lambda: self._raw_docx(
                     _entity_bomb_xml("UTF-8").encode("utf-8")),
                 kwargs={}, expect="declares a DTD", pre_open=False),
            dict(id="malformed", label="malformed XML",
                 make=lambda: self._raw_docx(b"<w:document><w:body>"),
                 kwargs={}, expect="not well-formed XML", pre_open=False),
            dict(id="char_cap", label="visible text over the character cap",
                 make=lambda: self._raw_docx(
                     _DOCUMENT_XML_TEMPLATE.format(
                         body=long_body).encode("utf-8")),
                 kwargs=dict(max_text_chars=100),
                 expect="visible text exceeded the 100 character cap",
                 pre_open=False),
            dict(id="unsupported_namespace",
                 label="a root in an unsupported namespace",
                 make=lambda: self._raw_docx(
                     b'<w:document xmlns:w="urn:example:not-word"><w:body>'
                     b"<w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>"
                     b"</w:document>"),
                 kwargs={}, expect="unsupported WordprocessingML namespace",
                 pre_open=False),
            dict(id="parse_error", label="an undeclared namespace prefix",
                 make=lambda: self._raw_docx(
                     b"<w:document><w:body><w:p><w:r><w:t>x</w:t></w:r>"
                     b"</w:p></w:body></w:document>"),
                 kwargs={}, expect="could not be parsed", pre_open=False),
        )

    # Every `raise ValueError` in docx_text.py, keyed to a parity case.
    # The token must appear in EXACTLY ONE refusal message -- the census
    # row asserts the bijection in both directions.
    _REFUSAL_SITES = (
        ("compression_method", "uses compression method %d"),
        ("unreadable_archive", "could not be read as a zip archive"),
        ("declared_size", "declared size %d exceeds"),
        ("infinite_ratio", "compression ratio is infinite"),
        ("ratio_over", "compression ratio %.1f:1 exceeds"),
        ("post_read_cap", "expanded past the %d byte cap"),
        ("dtd", "declares a DTD (DOCTYPE)"),
        ("malformed", "is not well-formed XML"),
        ("char_cap", "visible text exceeded the %d character cap"),
        ("parse_error", "could not be parsed"),
        ("unsupported_namespace", "unsupported WordprocessingML namespace"),
    )

    # A minimal well-formed part, used where a member must be BENIGN.
    _CLEAN_PART = b'<?xml version="1.0" encoding="UTF-8"?><root a="1"/>'

    # The screens that compare against a numeric limit. Each rejects on
    # `>`, so each owes an exactly-at-the-limit ACCEPTANCE case. The
    # infinite-ratio and DTD screens are absent deliberately: neither is a
    # numeric threshold, so neither has a "one under the line" value.
    _CAP_SITES = frozenset({"declared_size", "ratio_over", "char_cap"})

    # The ONE site with no end-to-end case, with the reason stated rather
    # than left as an omission. Pinned by the sweep, so growing this dict
    # is a deliberate act and not a quiet way to skip coverage.
    _NOT_ROUTE_REACHABLE = {
        "post_read_cap": (
            "unreachable end-to-end on EVERY route, not merely on two of them: "
            "ZipExtFile slices each read to the DECLARED size, and a member "
            "declaring more than max_bytes is refused by the metadata screen "
            "first. Reachable only by CALLING _assert_within_cap."),
    }

    # ENCODING x START-FORM x OFFSET x ROUTE.
    # label, xml-declaration encoding (None = NO declaration), codec, BOM?
    _ENCODING_FORMS = (
        ("UTF-8", "UTF-8", "utf-8", False),
        ("UTF-8 +BOM", "UTF-8", "utf-8", True),
        ("UTF-8 no-decl", None, "utf-8", False),
        ("UTF-16LE +BOM", "UTF-16", "utf-16-le", True),
        ("UTF-16BE +BOM", "UTF-16", "utf-16-be", True),
        ("UTF-16LE no BOM", "UTF-16", "utf-16-le", False),
        ("UTF-16BE no BOM", "UTF-16", "utf-16-be", False),
        ("UTF-16LE no-decl +BOM", None, "utf-16-le", True),
        ("UTF-8 no-decl +BOM", None, "utf-8", True),
        ("UTF-16BE no-decl +BOM", None, "utf-16-be", True),
    )
    _BOM = {"utf-8": bytes([0xEF, 0xBB, 0xBF]),
            "utf-16-le": bytes([0xFF, 0xFE]),
            "utf-16-be": bytes([0xFE, 0xFF])}

    # NON-ASCII clean text. An ASCII literal such as "clean" would prove only
    # that each start form was RECOGNISED, never that its text was DECODED
    # INTACT: a UTF-16-only lossy decode would pass the entire matrix while
    # corrupting ordinary document text. Measured against exactly that
    # impostor: the correct implementation returns this string byte-for-byte
    # on UTF-16LE+BOM and UTF-16BE+BOM; the lossy impostor returns
    # 'caf? ? ? ??'. With an ASCII literal, both would return that literal
    # intact on every form. (The same impostor is invisible on the UTF-8
    # forms by construction, which is why the discriminating cells are the
    # UTF-16 ones.) ESCAPES, not literal characters, keep this source file
    # pure ASCII.
    _CLEAN_TEXT = "caf\u00e9 \u2014 \u03a9 \u65e5\u672c"

    def _et_watch(self):
        """Returns `(reached, patches)` -- a recorder over EVERY ElementTree
        parse entry point, and the companion to `_expat_watch`.

        It is a shared helper so that every hostile cell installs the same
        instrumentation; a row carrying only the expat half would run with
        **half** of it. Because ElementTree's C accelerator creates **zero**
        watched `ParserCreate` calls (measured, and recorded in
        `_expat_watch`), an
        implementation could pre-parse a payload through `ET.fromstring`,
        expand the ladder, then run the correct expat screen and return the
        expected `ValueError` -- invisible to `_expat_watch`, and reachable
        only on the DTD forms the matrix does not vary. **Neither watch is
        redundant and neither can see the other; every hostile cell in BOTH
        sweeps installs BOTH.**

        `ET.XML` is an alias for `fromstring` and an explicit `XMLParser`
        bypasses both, so all three names are watched rather than just the
        obvious one.
        """
        reached = []
        originals = {}
        for name in ("fromstring", "XML", "XMLParser"):
            originals[name] = getattr(docx_text.ET, name, None)

        def make(name, real):
            def watch(*a, **k):
                reached.append(name)
                return real(*a, **k)
            return watch

        return reached, [
            mock.patch.object(docx_text.ET, n, make(n, originals[n]))
            for n in originals if originals[n] is not None]

    def _expat_watch(self):
        """Returns `(parses, patch)`. `parses` gains ONE entry per
        **Parse CALL**, recording whether a REAL
        `StartDoctypeDeclHandler` was installed **before that call began**,
        whether that call **RETURNED** (a correct screen's never does), and
        how many characters expat handed out. The assertions over those three
        live in `_assert_screened_without_expanding`, which documents why the
        second is load-bearing and the third only corroborating.

        A naive watcher -- one that records mere ASSIGNMENT of the handler
        attribute and reads that flag after the route returns -- is defeated
        two ways, BOTH MEASURED against variants of `_assert_no_dtd`
        (468-byte hostile UTF-8 ladder, characters counted with a
        `CharacterDataHandler` on the spied parser):

        | screen | outcome | naive watcher | this watcher | chars expanded |
        |---|---|---|---|---|
        | correct | `ValueError` DTD | PASS | **PASS** | **0** |
        | assigns the attribute to `None` | `ValueError` DTD | **PASS** | **FAIL** | **500,000** |
        | parses, THEN installs the handler | `ValueError` DTD | **PASS** | **FAIL** | **500,000** |
        | entity-expanding pre-pass, then the correct screen | `ValueError` DTD | FAIL | **FAIL** | **500,000** |

        Both middle rows are real defects wearing a correct refusal: the
        message is byte-identical to the correct one and half a million
        characters have already expanded. The last row is the pre-pass
        impostor, kept here to show this watcher LOSES NOTHING the naive one
        caught.

        Two changes, and each closes exactly one of the two:
        `callable(v)` rather than mere assignment (setting the attribute to
        `None` installs NOTHING -- expat calls nothing on the DOCTYPE), and a
        SNAPSHOT taken when `Parse` is entered rather than a flag read after
        the route returns (reading afterwards cannot see ORDER, which is what
        let parse-then-install look guarded).

        This watcher and the `ET` watch above it are COMPLEMENTARY, never
        redundant -- measured on CPython with the C accelerator:
        `ET.fromstring`, `ET.XML` and `ET.XMLParser()` produce **zero** calls
        to `xml.parsers.expat.ParserCreate` (the C accelerator builds its own
        parser, and `XMLParser` has no `.parser` attribute -- the same fact
        `_assert_no_dtd` records). So the expat watch sees ONLY the screen's
        own parse and the ET watch sees ONLY the tree parse; neither can see
        the other, and that is why `parses` is expected to hold EXACTLY ONE
        entry.
        """
        parses = []
        _real_create = xml.parsers.expat.ParserCreate

        def spy_create(*a, **k):
            _p = _real_create(*a, **k)
            state = {"armed": False, "chars": 0, "cd_replaced": False}

            def _count(data):
                state["chars"] += len(data)
            _p.CharacterDataHandler = _count

            class _Watched:
                def __setattr__(self, n, v):
                    if n == "StartDoctypeDeclHandler":
                        state["armed"] = callable(v)
                    if n == "CharacterDataHandler":
                        # The counter below is only trustworthy while it is
                        # still installed. Record replacement rather than
                        # silently reporting zero expansion.
                        state["cd_replaced"] = True
                    setattr(_p, n, v)

                def __getattr__(self, n):
                    if n == "Parse":
                        def _run(*pa, **pk):
                            rec = {"armed": state["armed"],
                                   "completed": False, "state": state}
                            parses.append(rec)
                            result = _p.Parse(*pa, **pk)
                            # Only reached if Parse RETURNED. A screen that
                            # aborts at the DOCTYPE never gets here.
                            rec["completed"] = True
                            return result
                        return _run
                    return getattr(_p, n)

            return _Watched()

        return parses, mock.patch.object(
            xml.parsers.expat, "ParserCreate", spy_create)

    def _assert_screened_without_expanding(self, parses, where):
        """Every assertion a hostile cell owes, in ONE place so the encoding
        matrix and the DTD-form sweep cannot drift apart.

        **Mechanism is not outcome.** Proving that a REAL handler was
        installed BEFORE `Parse` is still only the MECHANISM. It does not
        prove the handler STOPPED anything. MEASURED
        against a handler that records the DOCTYPE, **returns**, lets expat
        finish, and raises the expected `ValueError` afterwards:

        | screen | outcome | armed before parse | parse COMPLETED | chars |
        |---|---|---|---|---|
        | correct | `ValueError` DTD | yes | **no** | **0** |
        | handler returns, raises after `Parse` | `ValueError` DTD | **yes** | **yes** | **500,000** |

        The armed check passes it. **What separates them is that a correct
        screen's `Parse` never returns** -- the handler's exception propagates
        out of it -- so `completed` stays False. That is an OUTCOME (nothing
        was parsed past the DOCTYPE), not a mechanism, which is the point of
        this check.

        `chars` corroborates and bounds the harm, but it is the WEAKER of the
        two: an implementation installing its own `CharacterDataHandler`
        displaces the counter. That is why replacement is recorded and
        asserted rather than left to report a comfortable zero, and why
        `completed` -- which lives in the wrapper and cannot be displaced --
        is the load-bearing check.
        """
        unguarded = [r for r in parses if not r["armed"]]
        self.assertEqual(
            len(unguarded), 0,
            "%s: %d expat parse(s) began WITHOUT a StartDoctypeDeclHandler "
            "installed. The refusal can still carry the correct message while "
            "the entity ladder has already expanded -- MEASURED at 500,000 "
            "characters, from 468 bytes." % (where, len(unguarded)))
        self.assertEqual(
            len(parses), 1,
            "%s: expected EXACTLY ONE expat parse (the DTD screen itself); "
            "saw %d. An extra parse is an extra chance to expand."
            % (where, len(parses)))
        completed = [r for r in parses if r["completed"]]
        self.assertEqual(
            len(completed), 0,
            "%s: %d expat parse(s) RAN TO COMPLETION on a hostile payload. "
            "The handler was installed but did not ABORT -- it recorded the "
            "DOCTYPE, let the ladder expand, and the refusal was raised "
            "afterwards. MEASURED: 500,000 characters expanded behind a "
            "byte-identical refusal message." % (where, len(completed)))
        for r in parses:
            self.assertFalse(
                r["state"]["cd_replaced"],
                "%s: the implementation installed its own "
                "CharacterDataHandler, displacing the expansion counter -- "
                "its zero would be meaningless, so this is a refusal to "
                "report rather than a pass." % where)
            self.assertEqual(
                r["state"]["chars"], 0,
                "%s: %d characters were handed out by expat on a hostile "
                "payload. A correct screen aborts AT the DOCTYPE token, "
                "before any character data."
                % (where, r["state"]["chars"]))

    def _encoded_docx(self, codec, decl, bom, hostile, offset):
        def assemble(pad_chars):
            prolog = "" if decl is None else (
                '<?xml version="1.0" encoding="%s" standalone="yes"?>' % decl)
            if pad_chars:
                prolog += "<!-- %s -->" % ("p" * pad_chars)
            if hostile:
                prolog += "<!DOCTYPE w:document [%s]>" % _ENTITY_LADDER
            body = (
                '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>%s</w:t>'
                '</w:r></w:p></w:body></w:document>'
                % ("&e;" if hostile else self._CLEAN_TEXT))
            return ((self._BOM[codec] if bom else b"")
                    + (prolog + body).encode(codec))

        if offset == "early":
            payload = assemble(0)
        elif offset == "late":
            payload = assemble(4000)
        else:
            # NEAR-CAP: derive the pad from the ASSEMBLED size rather than a
            # constant. The DOCTYPE + entity ladder are themselves ~300 bytes
            # and are ENCODING-dependent, so a hand-computed pad overshoots
            # the byte cap and the archive is refused for SIZE -- which reads
            # as "the DTD screen missed it". Measured overshoot with a
            # constant pad: declared 384,077 against the 384,000 cap.
            per_char = 2 if "16" in codec else 1
            overhead = len(assemble(0))
            payload = assemble((DEFAULT_MAX_BYTES - overhead - 64) // per_char)
            # BOTH bounds. An upper bound alone lets the padding collapse
            # silently, and a collapsed pad no longer excludes the
            # 64 KB-prefix impostor while the row stays green.
            assert len(payload) <= DEFAULT_MAX_BYTES, len(payload)
            assert len(payload) > DEFAULT_MAX_BYTES - 4096, len(payload)
        return self._raw_docx(payload, compression=zipfile.ZIP_STORED)

    def test_the_DTD_screen_is_ENCODING_and_OFFSET_agnostic_on_all_routes(self):
        for label, decl, codec, bom in self._ENCODING_FORMS:
            for offset in ("early", "late", "near-cap"):
                for hostile in (True, False):
                    path = self._encoded_docx(codec, decl, bom, hostile, offset)
                    seen = []
                    for route in ("path", "BytesIO", "handle"):
                        if hostile:
                            # RECORD the call; do NOT raise. A tripwire that
                            # raises is swallowed by an implementation whose
                            # pre-parse sits in `try/except Exception: pass`
                            # -- MEASURED: the raising version left that
                            # impostor GREEN. A recorder cannot be caught.
                            # Watch EVERY parse entry point, not just
                            # fromstring: ET.XML is an alias and an explicit
                            # XMLParser bypasses both, so an implementation
                            # can pre-parse and expand through either while
                            # `fromstring` is never touched.
                            # The ET entry points are NOT the parser this
                            # module screens with -- `_assert_no_dtd` calls
                            # xml.parsers.expat.ParserCreate() DIRECTLY, and
                            # the two watches are provably disjoint. Both are
                            # helpers so that BOTH sweeps get BOTH; a row with
                            # only the expat half would silently run with
                            # half the instrumentation.
                            reached, patches = self._et_watch()
                            expat_parses, expat_patch = self._expat_watch()
                            patches.append(expat_patch)
                            for pt in patches:
                                pt.start()
                            try:
                                outcome, _o, alive = self._route_outcome(
                                    path, route, {})
                            finally:
                                for pt in patches:
                                    pt.stop()
                            self.assertEqual(
                                reached, [],
                                "%s/%s/%s: an ElementTree parse entry point "
                                "was REACHED for a hostile payload (via %s) "
                                "-- the bomb expanded before the refusal, "
                                "even though the refusal message looks correct"
                                % (label, offset, route,
                                   reached[0] if reached else "?"))
                            self._assert_screened_without_expanding(
                                expat_parses,
                                "%s/%s/%s" % (label, offset, route))
                        else:
                            outcome, _o, alive = self._route_outcome(
                                path, route, {})
                        # `_route_outcome` MEASURES whether a caller-owned
                        # handle survived. Every other row that checks it
                        # uses UTF-8 fixtures, so an implementation closing
                        # the caller's handle only on the UTF-16 / BOM paths
                        # would otherwise pass. Asserted on BOTH the hostile
                        # and the clean cells: a refusal must not consume the
                        # handle either.
                        if route == "handle":
                            self.assertTrue(
                                alive,
                                "%s/%s/%s: extract_text CLOSED the caller's "
                                "handle -- it BORROWS the handle, it does not "
                                "consume it, on refusal as much as on success"
                                % (label, offset, route))
                        seen.append(outcome)
                        kind, detail = outcome
                        if hostile:
                            self.assertEqual(
                                kind, "raised", "%s/%s/%s: not refused"
                                % (label, offset, route))
                            # The TYPE is part of the contract, not just the
                            # message: a RuntimeError carrying "declares a DTD"
                            # would otherwise pass every cell.
                            self.assertTrue(
                                detail.startswith("ValueError: "),
                                "%s/%s/%s: refused with the wrong EXCEPTION "
                                "TYPE -- %s" % (label, offset, route, detail))
                            self.assertIn(
                                "declares a DTD", detail,
                                "%s/%s/%s: wrong refusal -- %s"
                                % (label, offset, route, detail))
                        else:
                            self.assertEqual(
                                (kind, detail),
                                ("returned", self._CLEAN_TEXT),
                                "%s/%s/%s: a VALID document was refused, OR "
                                "its non-ASCII text was not decoded intact "
                                "-- this cell asserts the EXACT text, not "
                                "merely that the document was accepted"
                                % (label, offset, route))
                    self.assertEqual(
                        len(set(seen)), 1,
                        "%s/%s: the three routes disagreed" % (label, offset))

        # FALSE REJECTION, in every start form -- the other half of what a
        # STRUCTURAL screen buys. The other DTD rows carry a CDATA lookalike
        # (UTF-8 only), but their comments are PADDING in front of a real
        # hostile DTD; none is a clean comment or processing-instruction
        # lookalike in ANY encoding.
        #
        # MEASURED against a screen that decodes and substring-searches while
        # special-casing CDATA -- the natural shape of a textual screen that
        # has already been told about CDATA once:
        #
        #   CDATA lookalike   -> correct ACCEPTS, textual screen ACCEPTS
        #   comment lookalike -> correct ACCEPTS, textual screen REFUSES
        #   PI lookalike      -> correct ACCEPTS, textual screen REFUSES
        #
        # in BOTH UTF-8 and UTF-16LE+BOM. A false rejection is a real defect
        # on ordinary documents, and no hostile cell can ever surface it.
        # These cases live here because this row already owns start forms.
        for label, decl, codec, bom in self._ENCODING_FORMS:
            for kind, fragment in (("comment", "<!-- <!DOCTYPE x> -->"),
                                   ("processing instruction",
                                    "<?pi <!DOCTYPE x> ?>")):
                prolog = "" if decl is None else (
                    '<?xml version="1.0" encoding="%s" standalone="yes"?>'
                    % decl)
                xml_text = (
                    prolog + fragment
                    + '<w:document xmlns:w="http://schemas.openxmlformats.org'
                    '/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>%s'
                    '</w:t></w:r></w:p></w:body></w:document>'
                    % self._CLEAN_TEXT)
                path = self._raw_docx(
                    (self._BOM[codec] if bom else b"")
                    + xml_text.encode(codec),
                    compression=zipfile.ZIP_STORED)
                for route in ("path", "BytesIO", "handle"):
                    outcome, _o, alive = self._route_outcome(path, route, {})
                    self.assertEqual(
                        outcome, ("returned", self._CLEAN_TEXT),
                        "%s/%s/%s: a DOCTYPE-lookalike inside a %s is NOT a "
                        "declaration -- a VALID document was refused, or its "
                        "text was not decoded intact. Got %r"
                        % (label, kind, route, kind, outcome))
                    if route == "handle":
                        self.assertTrue(
                            alive,
                            "%s/%s/%s: extract_text CLOSED the caller's handle"
                            % (label, kind, route))

    # SEMANTIC DOCUMENTS -- each shaped to be observable ONLY under one of
    # the traversal/counting impostors the extraction-semantics rows exclude.
    _SEMANTIC_DOCS = (
        ("nested paragraph (text box)",
         "<w:p><w:r><w:t>before </w:t></w:r>"
         "<w:p><w:r><w:t>inside</w:t></w:r></w:p>"
         "<w:r><w:t> after</w:t></w:r></w:p>", {}),
        ("hyperlink and SDT wrappers",
         "<w:p><w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>"
         "<w:sdt><w:sdtContent><w:r><w:t> ctrl</w:t></w:r>"
         "</w:sdtContent></w:sdt></w:p>", {}),
        ("predefined and numeric references",
         "<w:p><w:r><w:t>a &amp; b &#65; c</w:t></w:r></w:p>", {}),
        # AT the character cap and OVER the utf-8 BYTE count, deliberately:
        # "cafe-acute space world" is 7 CHARACTERS and 12 UTF-8 BYTES, so at
        # max_text_chars=8 a character cap ACCEPTS (7 <= 8) and a byte cap
        # REFUSES (12 > 8). It separates the two counts; it is NOT the
        # inclusive non-ASCII boundary and does not claim to be.
        # Without the tightened cap this document is nowhere near 384,000 and
        # byte-vs-character counting is invisible -- measured: without this
        # kwarg the byte-cap impostor stays GREEN.
        ("multibyte text at the character cap",
         # ESCAPES, not literal characters: they keep this source file pure
         # ASCII, and a pasted non-ASCII literal is easily mangled.
         "<w:p><w:r><w:t>caf\u00e9 \u4e16\u754c</w:t></w:r></w:p>",
         dict(max_text_chars=8)),
        ("empty run then text",
         "<w:p><w:r><w:t/></w:r><w:r><w:t>x</w:t></w:r></w:p>", {}),
        # AT a tight cap on purpose: a file-like branch charging RAW
        # pre-strip() text counts 10 where a correct one counts 6, so the
        # cap is what makes the difference observable. Untightened, both
        # counts sit far under 384,000 and the impostor is invisible.
        ("edge whitespace at the cap",
         "<w:p><w:r><w:t>  padded  </w:t></w:r></w:p>",
         dict(max_text_chars=6)),
        # An EMPTY paragraph BETWEEN two visible ones, at the exact cap:
        # "one" + sep + "two" = 7. An implementation charging a separator
        # for the skipped empty paragraph counts 8 and refuses.
        ("empty paragraph between visible ones, at the cap",
         "<w:p><w:r><w:t>one</w:t></w:r></w:p>"
         "<w:p><w:r><w:t></w:t></w:r></w:p>"
         "<w:p><w:r><w:t>two</w:t></w:r></w:p>",
         dict(max_text_chars=7)),
        ("orphan w:t outside any paragraph",
         "<w:p><w:r><w:t>owned</w:t></w:r></w:p>"
         "<w:r><w:t>ORPHAN</w:t></w:r>", {}),
        ("two paragraphs (separator)",
         "<w:p><w:r><w:t>one</w:t></w:r></w:p>"
         "<w:p><w:r><w:t>two</w:t></w:r></w:p>", {}),
        # Every other cap boundary in this suite is ASCII, so a counter
        # charging only ASCII characters -- sum(ch.isascii() for ch in text)
        # -- would pass all of them while permitting arbitrarily long
        # non-ASCII output, defeating the max_text_chars bound outright.
        # MEASURED at max_text_chars=50 against exactly that impostor:
        #     ASCII     at cap    correct accepted   impostor accepted
        #     ASCII     over cap  correct RAISED     impostor RAISED   <- blind
        #     non-ASCII at cap    correct accepted   impostor accepted
        #     non-ASCII over cap  correct RAISED     impostor ACCEPTED <- sees
        # Only the last case discriminates, which is why it is added here and
        # the ASCII rows above are left alone.
        # ROUTE PARITY ALONE CANNOT SEE IT EITHER: extract_text converges on a
        # single `_visible_text` call regardless of source type, so this
        # impostor is uniformly wrong on all three routes and the parity
        # assertion stays GREEN. That is exactly why this entry carries an
        # EXPECTED OUTCOME and the nine above do not -- parity is the wrong
        # instrument for a defect that is not route-dependent.
        # chr(), not a literal: it keeps this source file pure ASCII, and an
        # escape sequence is easily mangled by tools that normalise them.
        ("non-ASCII text ONE OVER the character cap",
         "<w:p><w:r><w:t>" + chr(0x00E9) * 51 + "</w:t></w:r></w:p>",
         dict(max_text_chars=50),
         ("raised", "character cap")),
    )

    def test_the_SAME_TEXT_comes_back_on_path_BytesIO_and_handle(self):
        """**The extraction-semantics rows call `extract_text(path)` and
        nothing else**, and the other three-route clean fixtures each carry
        a single short run of text. So a file-like branch with a
        direct-child traversal, a descendant-joining traversal, or a
        UTF-8-byte character cap would pass every one of them while
        dropping, duplicating or rejecting real text on the route
        `file_tools` uses.

        This closes that as a CLASS rather than converting each row: each
        document below is shaped to be observable only under one of the
        named impostors, and all three routes must return the IDENTICAL
        string. It asserts EQUALITY ACROSS ROUTES, not a literal, so it
        stays correct if the traversal is ever legitimately changed -- the
        semantics rows own the literals; this row owns the parity.
        """
        for entry in self._SEMANTIC_DOCS:
            label, body, kwargs = entry[0], entry[1], entry[2]
            # An optional FOURTH element pins the expected outcome as
            # (kind, substring). Route parity is the wrong instrument for an
            # impostor that is uniformly wrong on every route, so a row that
            # needs its outcome asserted supplies it here. Three-element
            # entries have a parity-only contract.
            expected = entry[3] if len(entry) > 3 else None
            path = self._docx(body)
            seen = {}
            for route in ("path", "BytesIO", "handle"):
                outcome, _o, alive = self._route_outcome(
                    path, route, kwargs)
                seen[route] = outcome
                if route == "handle":
                    self.assertTrue(alive, "%s: handle was closed" % label)
                if expected is not None:
                    kind, detail = outcome
                    self.assertEqual(
                        kind, expected[0],
                        "%s/%s: expected the document to be %s, got %r -- a "
                        "cap counting only ASCII characters does exactly this"
                        % (label, route, expected[0], outcome))
                    self.assertIn(
                        expected[1], detail,
                        "%s/%s: %s for the WRONG reason -- %s"
                        % (label, route, expected[0], detail))
            self.assertEqual(
                seen["path"], seen["BytesIO"],
                "%s: BytesIO returned different TEXT from the path route "
                "-- %r vs %r" % (label, seen["path"], seen["BytesIO"]))
            self.assertEqual(
                seen["path"], seen["handle"],
                "%s: the open handle returned different TEXT from the path "
                "route -- %r vs %r" % (label, seen["path"], seen["handle"]))
            # Only cases WITHOUT a pinned outcome must be accepted here. The
            # one case that correctly expects a REFUSAL -- 51 non-ASCII
            # characters under a 50-character cap, there to catch an
            # ASCII-only counter -- has already been asserted per route at
            # the top of this loop; an unconditional "returned" check would
            # make this row impossible to pass against the CORRECT
            # implementation. Parity across routes is asserted for every case
            # either way, so nothing is weakened here.
            if expected is None:
                kind, detail = seen["path"]
                self.assertEqual(
                    kind, "returned",
                    "%s: a VALID document was refused -- %s" % (label, detail))


    # THE DTD AXIS IS GENERATED, NOT LISTED.
    #
    # A hand-kept tuple of DOCTYPE forms kept turning out to be missing ONE
    # more cell: the alternate DOCTYPE name, external-with-internal-subset,
    # post-DOCTYPE whitespace, PUBLIC-with-internal-subset. Each addition
    # was correct and each left the axis PARTIAL, because a hand-kept list
    # of a PRODUCT is only ever as complete as its author's imagination.
    #
    # The product is small and closed: a DOCTYPE carries an external
    # identifier that is absent, SYSTEM or PUBLIC, and an internal subset
    # that is absent, empty or entity-bearing. That is 3 x 3 = 9, and
    # generating it means no cell of it can be missing. Two further variants
    # are ORTHOGONAL to that product rather than members of it -- the DOCTYPE
    # NAME and the whitespace between the token and the name -- so they are
    # appended explicitly and each is justified by its own impostor.
    #
    # MEASURED over all 11 forms: the correct screen refuses 11/11 naming the
    # DTD, expanding 0 characters. Six impostors, each gated on a different
    # attribute of the declaration, are ALL caught:
    #
    #   name-gated (refuses only `w:document`)      -> 1 cell,  500,000 chars
    #   subset-gated (`sys is None or not internal`)-> 4 cells, 500,000 chars
    #   public-gated (`pub is None or not internal`)-> 2 cells, 500,000 chars
    #   internal-only (refuses only a bare DOCTYPE) -> 6 cells, 500,000 chars
    #   requires-internal-subset                    -> 3 cells
    #   byte prefilter on `<!DOCTYPE ` + space      -> 1 cell
    #
    # **The last three were never anticipated by the hand-kept list.** The
    # generated product caught them because it is a product, which is the
    # argument for generating it rather than the argument for any one cell.
    _EXTERNAL_IDS = (
        ("internal-only", ""),
        ("SYSTEM", ' SYSTEM "x.dtd"'),
        ("PUBLIC", ' PUBLIC "-//X//EN" "x.dtd"'),
    )
    _INTERNAL_SUBSETS = (
        ("no subset", "", "text"),
        ("empty subset", " []", "text"),
        ("entity subset", " [%s]" % _ENTITY_LADDER, "&e;"),
    )
    # The body is per-form because only an entity-DECLARING form may legally
    # REFERENCE an entity: "&e;" in a shared body makes every no-subset and
    # empty-subset fixture reference an UNDECLARED entity, so a correct screen
    # raises "is not well-formed XML" instead of the DTD refusal and the cell
    # proves nothing.
    #
    # `_build_dtd_forms` is a CLASS-BODY LOCAL, called during class creation
    # and then deleted, so it never becomes a method and adds no
    # module-level name. Its parameters are passed in rather than read from
    # the class, because a class body is not an enclosing scope for the
    # functions defined inside it. VERIFIED: 11 forms built, and
    # `hasattr(cls, "_build_dtd_forms")` is False.
    def _build_dtd_forms(externals, subsets, ladder):
        forms = []
        for ext_label, ext in externals:
            for sub_label, sub, body in subsets:
                forms.append(("%s / %s" % (ext_label, sub_label),
                              "<!DOCTYPE w:document%s%s>" % (ext, sub), body))
        # ORTHOGONAL to the product above, not members of it: the DOCTYPE
        # NAME and the whitespace between token and name.
        forms.append(("alternate DOCTYPE name",
                      "<!DOCTYPE x [%s]>" % ladder, "&e;"))
        forms.append(("whitespace after DOCTYPE",
                      "<!DOCTYPE\n\tw:document\n[%s]>" % ladder, "&e;"))
        return tuple(forms)

    _DTD_FORMS = _build_dtd_forms(_EXTERNAL_IDS, _INTERNAL_SUBSETS,
                                  _ENTITY_LADDER)
    del _build_dtd_forms

    def test_EVERY_DTD_FORM_refused_in_EVERY_ENCODING_OFFSET_and_ROUTE(self):
        """**The encoding matrix crosses start form, offset and route -- but
        every one of its hostile payloads carries the SAME entity-bearing
        internal subset.** The minimal, empty-subset, SYSTEM and PUBLIC
        fixtures elsewhere are UTF-8 only. So an implementation that refuses
        all UTF-8 DTDs, but refuses a non-UTF-8 DTD only when it has an
        internal subset or declares entities, would pass every other row --
        and then accept a minimal or external UTF-16 DOCTYPE.

        The DTD refusal on word/document.xml is UNCONDITIONAL on the
        DOCTYPE, so the form must be crossed with the encoding.

        **This row runs TWO crossings.** Crossing the forms with only some
        of the start forms at a fixed offset would leave **A x C partial and
        B x C crossed by nothing at all** -- a late or near-cap non-ladder
        DTD would pass every cell. Both crossings live here, in ONE row.

        MEASURED on the correct screen, every cell naming the DTD:

        | crossing | cells | result |
        |---|---|---|
        | **A x C** -- 11 forms x 10 start forms, offset `early` | 110 | **110/110 refused** |
        | **B x C** -- 11 forms x 3 offsets, start form UTF-8 | 33 | **33/33 refused** |

        The near-cap cells derive their pad from the assembled size, so every
        one lands at 383,945 bytes against the 384,000 cap -- inside it, and
        not so far inside that the padding has collapsed.

        **And every cell carries the no-parse recorder.** A sweep asserting
        OUTCOME only would let the same entity-expanding pre-pass the
        encoding matrix excludes walk straight through all of it. The
        recorder is `_expat_watch`, shared with the matrix so the two cannot
        drift.

        The alternate-name form is the DOCTYPE **NAME** branch. Every other
        hostile fixture hard-codes the name `w:document`, so a handler gated
        on that literal would pass all of them.
        MEASURED over the 110 A x C cells: the correct screen refused 110/110;
        the name-gated impostor accepted **exactly 10** -- the ten start forms
        of that one row -- and refused the other 100. So the cell reds on the
        case it targets and on nothing else, which is the check a red needs
        before it counts as evidence. **Six impostors, each gated on a
        different attribute of the declaration, are all caught by the
        generated axis -- and three of the six were never anticipated by a
        hand-kept list.** That is the argument for generating the axis
        rather than the argument for any one cell in it.

        This is the same two-axes-at-each-other's-default shape the route
        sweep and the encoding matrix each close, and the reason the axis is
        enumerated rather than sampled.
        """
        def build(codec, decl, bom, doctype, body, offset):
            def assemble(pad):
                prolog = "" if decl is None else (
                    '<?xml version="1.0" encoding="%s" standalone="yes"?>'
                    % decl)
                if pad:
                    prolog += "<!-- %s -->" % ("p" * pad)
                prolog += doctype
                inner = (
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                    'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>%s'
                    '</w:t></w:r></w:p></w:body></w:document>' % body)
                return ((self._BOM[codec] if bom else b"")
                        + (prolog + inner).encode(codec))

            if offset == "early":
                return assemble(0)
            if offset == "late":
                return assemble(4000)
            # Same derivation as _encoded_docx, and for the same reason: a
            # constant pad overshoots the byte cap and the archive is refused
            # for SIZE, which reads as "the DTD screen missed it".
            per_char = 2 if "16" in codec else 1
            payload = assemble(
                (DEFAULT_MAX_BYTES - len(assemble(0)) - 64) // per_char)
            assert len(payload) <= DEFAULT_MAX_BYTES, len(payload)
            assert len(payload) > DEFAULT_MAX_BYTES - 4096, len(payload)
            return payload

        def sweep(axis, cells):
            done = set()
            for enc_label, codec, decl, bom, form, doctype, body, offset in cells:
                path = self._raw_docx(
                    build(codec, decl, bom, doctype, body, offset),
                    compression=zipfile.ZIP_STORED)
                seen = []
                for route in ("path", "BytesIO", "handle"):
                    where = "%s/%s/%s/%s/%s" % (axis, enc_label, form,
                                                offset, route)
                    # BOTH watches, on EVERY hostile cell. The expat watch
                    # alone cannot see an
                    # ElementTree pre-parse, and ET is exactly where an
                    # implementation would expand the forms this row adds.
                    reached, patches = self._et_watch()
                    expat_parses, expat_patch = self._expat_watch()
                    patches.append(expat_patch)
                    for pt in patches:
                        pt.start()
                    try:
                        outcome, _o, alive = self._route_outcome(
                            path, route, {})
                    finally:
                        for pt in patches:
                            pt.stop()
                    self.assertEqual(
                        reached, [],
                        "%s: an ElementTree parse entry point was REACHED "
                        "for a hostile payload (via %s) -- the bomb expanded "
                        "before the refusal, even though the refusal message "
                        "looks correct"
                        % (where, reached[0] if reached else "?"))
                    self._assert_screened_without_expanding(
                        expat_parses, where)
                    done.add((enc_label, offset, form))
                    if route == "handle":
                        self.assertTrue(
                            alive,
                            "%s: extract_text CLOSED the caller's handle on a "
                            "refusal -- it BORROWS, it does not consume"
                            % where)
                    seen.append(outcome)
                    kind, detail = outcome
                    self.assertEqual(
                        kind, "raised",
                        "%s: a %s DTD was ACCEPTED" % (where, form))
                    self.assertTrue(
                        detail.startswith("ValueError: "),
                        "%s: wrong exception TYPE -- %s" % (where, detail))
                    self.assertIn(
                        "declares a DTD", detail,
                        "%s: wrong refusal -- %s" % (where, detail))
                self.assertEqual(
                    len(set(seen)), 1,
                    "%s/%s/%s/%s: the three routes disagreed"
                    % (axis, enc_label, form, offset))
            return done

        OFFSETS = ("early", "late", "near-cap")
        enc_labels = [e[0] for e in self._ENCODING_FORMS]
        form_labels = [f[0] for f in self._DTD_FORMS]

        # A x C -- every DTD form against every START FORM, offset fixed.
        axc = sweep("AxC",
                    [(enc_label, codec, decl, bom, form, doctype, body,
                      "early")
                     for enc_label, decl, codec, bom in self._ENCODING_FORMS
                     for form, doctype, body in self._DTD_FORMS])

        # B x C -- every DTD form against every OFFSET, start form fixed.
        bxc = sweep("BxC",
                    [("UTF-8", "utf-8", "UTF-8", False, form, doctype, body,
                      offset)
                     for offset in OFFSETS
                     for form, doctype, body in self._DTD_FORMS])

        # THE COVERAGE CLAIM IS ASSERTED BY THE ROW, NOT WRITTEN IN PROSE.
        # Cell arithmetic stated only in prose can be wrong while the test
        # count still looks right, because nothing executes the claim. A
        # shortened or vacuous loop is a FAILING row rather than a silently
        # smaller sweep.
        self.assertEqual(len(set(enc_labels)), len(enc_labels),
                         "duplicate start-form label: %s" % enc_labels)
        self.assertEqual(len(set(form_labels)), len(form_labels),
                         "duplicate DTD-form label: %s" % form_labels)
        self.assertEqual(
            axc, {(e, "early", f) for e in enc_labels for f in form_labels},
            "A x C is not the COMPLETE product of start form x DTD form")
        self.assertEqual(
            bxc, {("UTF-8", o, f) for o in OFFSETS for f in form_labels},
            "B x C is not the COMPLETE product of offset x DTD form")

        # The matrix's own product, derived from the SAME constants its loop
        # runs over; the matrix row asserts its own completeness separately.
        entity_form = "internal-only / entity subset"
        self.assertIn(entity_form, form_labels,
                      "the encoding matrix's fixed DTD form left the axis")
        axb = {(e, o, entity_form) for e in enc_labels for o in OFFSETS}
        distinct = axb | axc | bxc
        space = len(enc_labels) * len(OFFSETS) * len(form_labels)
        self.assertEqual(
            (len(axb), len(axc), len(bxc)), (30, 110, 33),
            "cells RUN per crossing changed -- update the counts")
        self.assertEqual(
            (len(distinct), space, space - len(distinct)), (150, 330, 180),
            "DISTINCT / SPACE / UNTESTED changed -- update the counts")

    def test_the_character_cap_counts_CHARACTERS_not_utf8_bytes(self):
        # Every other cap row is ASCII, where characters and bytes are the
        # same number, so an implementation measuring
        # len(text.encode("utf-8")) -- or reusing the byte count it already
        # has from the read -- passes all of them while falsely REJECTING
        # legitimate non-ASCII documents. Each of these 300 characters is
        # three bytes in UTF-8: a byte-counting impostor sees 900 and refuses
        # at a 500-character cap; a correct implementation sees 300 and
        # accepts. Written as an escape, never a literal -- this file is
        # ASCII-only.
        body = "\u4e16" * 300          # U+4E16, three bytes in UTF-8
        path = self._docx("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % body)
        self.assertEqual(extract_text(path, max_text_chars=500), body)

        # Every cap fixture above is BMP, where one character is one UTF-16
        # code unit, so a counter written as len(text.encode("utf-16-le")) //
        # 2 passes all of them and then FALSELY REJECTS ordinary emoji.
        #
        # MEASURED on U+1F600 -- 1 Python character, 2 UTF-16 code units,
        # 4 UTF-8 bytes. `max_text_chars=1` is the ONLY cap that separates
        # all three: correct ACCEPTS, the UTF-16 counter REFUSES, the UTF-8
        # byte counter REFUSES. At cap=2 the UTF-16 counter passes again, so
        # the 1 is load-bearing, not decoration. It lives in this row because
        # this row already owns character-vs-byte accounting.
        astral = "\U0001F600"
        self.assertEqual(len(astral), 1)          # pins WHY cap=1 works
        path = self._docx("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % astral)
        self.assertEqual(extract_text(path, max_text_chars=1), astral)

    def test_an_EMPTY_text_run_is_joined_as_an_empty_string(self):
        # _visible_text deliberately writes `run.text or ""` because
        # ElementTree gives text=None for a well-formed EMPTY element --
        # <w:t/> or <w:t></w:t>, both of which a real Word document produces.
        # Without a fixture containing one, replacing that expression with a
        # bare `run.text` would leave every other row green and then raise
        # TypeError ("sequence item N: expected str, NoneType found") the
        # first time a real document with an empty run was read.
        #
        # The empty run must sit BETWEEN two non-empty ones: a trailing empty
        # run can be dropped by an implementation that filters falsy values and
        # still produce the right string, which would make this row pass
        # against the very defect it exists to catch.
        path = self._docx(
            "<w:p><w:r><w:t>before</w:t></w:r>"
            "<w:r><w:t/></w:r>"
            "<w:r><w:t>after</w:t></w:r></w:p>")
        self.assertEqual(extract_text(path), "beforeafter")

    def test_runs_inside_HYPERLINK_and_SDT_wrappers_are_preserved(self):
        # Every other positive fixture puts <w:t> inside a DIRECT <w:r> child
        # of the paragraph, so a traversal written as ./w:r/w:t passes the
        # basic, nested-depth and text-box rows while silently DROPPING text
        # inside the wrappers Word actually emits: w:hyperlink for every link,
        # w:sdt for every content control. That is data loss on ordinary
        # documents, not an attack -- and nearest-paragraph ownership is
        # exactly the kind of change that invites it, since the bug it
        # corrects is over-collection.
        #
        # The invariant is "every w:t the paragraph OWNS, in document order".
        # Ownership is about the nearest enclosing w:p, NOT about depth.
        body = ("<w:p>"
                "<w:r><w:t>plain </w:t></w:r>"
                "<w:hyperlink><w:r><w:t>linked</w:t></w:r></w:hyperlink>"
                "<w:r><w:t> and </w:t></w:r>"
                "<w:sdt><w:sdtContent><w:r><w:t>boxed</w:t></w:r>"
                "</w:sdtContent></w:sdt>"
                "</w:p>")
        path = self._docx(body)
        # EXACT and ORDERED: a traversal collecting the right runs in the wrong
        # order is also wrong, and a set comparison would not see it.
        self.assertEqual(extract_text(path), "plain linked and boxed")

    def test_predefined_and_numeric_character_references_are_PRESERVED(self):
        # The DTD screen relies on predefined and numeric character references
        # being SAFE -- each expands to exactly one character and cannot bomb.
        # Without a clean fixture containing one, an implementation that
        # refuses any payload containing "&" would pass every bomb row, the
        # UTF-16 rows, the CDATA row and every clean control -- while
        # rejecting ordinary Word output, because Word writes "&" as "&amp;"
        # in every document that contains one.
        body = ("<w:p><w:r><w:t>Tom &amp; Jerry &lt;ok&gt; "
                "&#65;&#x42;</w:t></w:r></w:p>")
        path = self._docx(body)
        self.assertEqual(extract_text(path), "Tom & Jerry <ok> AB")

    def test_the_cap_boundary_is_correct_across_an_EMPTY_paragraph(self):
        # Every other cap-boundary row uses paragraphs that all emit text. An
        # implementation that charges a SEPARATOR for an empty paragraph
        # before skipping it still passes the basic
        # test_empty_paragraph_skipped output row and every cap row, while
        # falsely rejecting a legitimate document sitting just under the cap.
        # The empty-paragraph form matches that basic row:
        # <w:p><w:pPr/></w:p>.
        body = ("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("a" * 10)
                + "<w:p><w:pPr/></w:p>"
                + "<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("b" * 10))
        path = self._docx(body)
        joined = extract_text(path)
        # Two visible paragraphs and ONE separator: 10 + 1 + 10. The empty
        # paragraph contributes nothing -- not even a separator.
        self.assertEqual(len(joined), 21, repr(joined))
        # The discriminating assertion: an implementation charging the empty
        # paragraph a separator counts 22 and refuses this.
        self.assertEqual(extract_text(path, max_text_chars=21), joined)
        with self.assertRaisesRegex(ValueError, "visible text exceeded"):
            extract_text(path, max_text_chars=20)

    def test_the_cap_boundary_is_correct_across_WHITESPACE(self):
        # The empty-paragraph boundary row above uses a paragraph carrying no
        # <w:t> at all. Two shapes one step smaller than that behave
        # identically under the skip-empty rule: a paragraph whose <w:t>
        # holds ONLY whitespace (strips to "", contributes nothing, not even a
        # separator), and edge whitespace on a paragraph that IS visible
        # (stripped off before it is counted).
        #
        # An implementation that accumulates RAW text and charges it against
        # the cap before strip() passes every other row -- including the
        # empty-paragraph row, whose paragraph has no text node to charge --
        # while falsely refusing a legitimate document sitting just under the
        # cap. The cap must be accounted on the text actually RETURNED.
        body = ("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("a" * 10)
                + "<w:p><w:r><w:t>     </w:t></w:r></w:p>"
                + "<w:p><w:r><w:t>  %s  </w:t></w:r></w:p>" % ("b" * 10))
        path = self._docx(body)
        joined = extract_text(path)
        # Two visible paragraphs and ONE separator: 10 + 1 + 10.
        self.assertEqual(joined, "a" * 10 + "\n" + "b" * 10)
        self.assertEqual(len(joined), 21, repr(joined))
        # The discriminating pair. A raw-text impostor counts 10 + 5 + 14
        # plus two separators = 31 and refuses this document at a cap of 21.
        self.assertEqual(extract_text(path, max_text_chars=21), joined)
        with self.assertRaisesRegex(ValueError, "visible text exceeded"):
            extract_text(path, max_text_chars=20)

    def test_the_text_cap_STOPS_THE_WALK_rather_than_trimming_at_the_end(self):
        r"""Every other cap row asserts only that a `ValueError` is raised.
        An implementation that builds the whole joined string and checks
        `len()` once at the end satisfies all of them, while `_visible_text`
        promises the accumulated total is bounded "as it is built".

        What this row measures is the paragraph-JOINING loop -- the one the cap
        actually governs -- and it measures it through the `reached %d` value
        the refusal message already carries. That number IS the accumulated
        total at the moment the cap fired, so it reports directly how much was
        accumulated before the walk stopped. No test double is involved.

        **Why no test double.** A proxy root that counts `root.iter()` calls
        is written against one implementation shape. The single-pass
        traversal carries the owner DOWN and never calls `root.iter()` -- it
        walks with `list(node)`. Python looks up `__iter__` on the TYPE, so a
        proxy's `__getattr__` never fires and the proxy raises
        `TypeError: ... object is not iterable`. **A test double is written
        against a specific implementation shape, and rewriting that shape
        silently kills it.**

        Counting walks is avoided too: a node budget tight enough to matter
        (say, fewer than 60 nodes) is one a conforming multi-pass traversal
        (807 nodes handed out over three walks) can never meet. Reading
        `reached` avoids the whole question -- it is insensitive to how many
        passes the traversal makes, and only sensitive to WHERE the cap is
        applied, which is the one thing this row exists to pin.

        **WHAT THIS ROW CLAIMS, NARROWED -- a SCOPING statement, because no
        stronger row is reachable.** It does NOT prove "the cap REJECTS
        DURING ACCUMULATION rather than after it". An implementation that
        joins all 200 paragraphs and THEN scans prefixes to report the first
        crossing reports `reached 5004` too. **MEASURED against exactly that
        impostor: identical message, and `tracemalloc` peak 757,777 vs the
        correct 758,183 -- so
        allocation does not separate them either**, because the traversal
        already holds every paragraph's text in `owned` regardless, and the
        joined string is noise against it. The one-pass design precomputes
        ownership, so NOTHING in the joining loop is observable from outside.

        **What it DOES claim, and does prove: the reported figure is the
        ACCUMULATION POINT, not the document total** -- 5004, not 200199. That
        excludes the naive build-then-measure implementation, which is the one
        that actually misreports, and it is measured both ways.

        **Why the unexcluded variant is ACCEPTED rather than chased:** it
        refuses the same documents with the same message and the same memory
        profile. The character cap exists to bound what reaches the model, and
        both variants refuse identically -- so the residual is a micro-
        inefficiency, not a hole. **Do not add machinery to chase it**: the
        guarantee is not observable from outside."""
        paragraph = "<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("z" * 1000)
        root = docx_text.ET.fromstring(
            _DOCUMENT_XML_TEMPLATE.format(body=paragraph * 200))

        with self.assertRaises(ValueError) as caught:
            docx_text._visible_text(root, 5000, "word/document.xml")
        message = str(caught.exception)

        # 200 paragraphs of 1,000 characters against a 5,000-character cap.
        # Measured against the real implementation and against the
        # build-then-measure impostor:
        #
        #   correct (stops during accumulation)  -> "reached 5004"
        #   build-then-measure (joins all 200)   -> "reached 200199"
        #
        # 5004 because the cap is tested after each paragraph is added:
        # 1000, 2001, 3002, 4003, 5004 -- the fifth crosses 5000 and raises
        # BEFORE the sixth is ever joined. The impostor reports the FULL
        # 200*1000 + 199 separator characters. BOTH raise ValueError with the
        # identical message shape, which is exactly why no other cap row can
        # separate them and why this row reads the NUMBER rather than the
        # exception.
        self.assertIn("visible text exceeded", message)
        self.assertRegex(
            message, r"reached 5004\b",
            "the refusal reported the DOCUMENT TOTAL rather than the "
            "accumulation point -- a build-then-measure implementation: %s"
            % message)

    def test_visible_text_cap_boundary_is_inclusive(self):
        r"""MULTI-PARAGRAPH ON PURPOSE. Do not reduce this to one paragraph.

        _visible_text accumulates len(text) + (1 if lines else 0), so the cap
        bounds the JOINED string. A single-paragraph row never reaches the
        separator term at all.

        The impostor excluded here is total += len(text) -- the separator
        dropped. Measured: it leaves every other row green, and at a cap of
        max_text_chars it RETURNS max_text_chars + (N - 1) characters, which
        breaches the exact bound the cap promises.

        Four paragraphs of 250 characters join to 4*250 + 3 = 1003.
        """
        count, width = 4, 250
        joined = count * width + (count - 1)
        body = "".join("<w:p><w:r><w:t>%s</w:t></w:r></w:p>" % ("x" * width)
                       for _ in range(count))
        path = self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(body=body).encode("utf-8"))
        # Inclusive at the boundary, and the RETURNED length is the joined
        # length -- which is what the cap means. Both the real implementation
        # and the impostor satisfy this line; it pins the bound's meaning
        # rather than excluding the impostor.
        self.assertEqual(
            len(extract_text(path, max_text_chars=joined)), joined)
        # THIS is the discriminating assertion. The impostor counts only
        # 4*250 = 1000, never exceeds the 1002 cap, does not raise, and hands
        # back 1003 characters against it.
        with self.assertRaisesRegex(ValueError, "visible text exceeded"):
            extract_text(path, max_text_chars=joined - 1)

    def test_word_text_box_paragraph_is_not_duplicated(self):
        # The REAL shape: Word embeds a whole <w:p> inside a run via
        # <w:txbxContent>. A descendant-joining traversal returns
        # ['Paragraph before the box.TEXT INSIDE THE BOX', 'TEXT INSIDE THE
        # BOX', ...] -- the box text duplicated AND glued onto the preceding
        # paragraph. This is a correctness bug before it is a size bug.
        #
        # THE ORPHAN <w:t> IS LOAD-BEARING. It exercises `owner is None`,
        # the branch that DROPS text outside any paragraph. Measured: correct
        # ownership drops it and the expected list below is UNCHANGED, while
        # an impostor carrying the LAST-SEEN
        # paragraph down instead of the ENCLOSING one returns
        # 'TEXT INSIDE THE BOXORPHAN OUTSIDE EVERY PARAGRAPH' and reds here.
        path = self._raw_docx(_DOCUMENT_XML_TEMPLATE.format(body=(
            "<w:p><w:r><w:t>Paragraph before the box.</w:t></w:r>"
            "<w:r><w:pict><v:shape xmlns:v='urn:schemas-microsoft-com:vml'>"
            "<v:textbox><w:txbxContent>"
            "<w:p><w:r><w:t>TEXT INSIDE THE BOX</w:t></w:r></w:p>"
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>"
            # OUTER-PARAGRAPH TEXT AFTER THE NESTED ONE. Without a run here,
            # a traversal carrying ONE mutable "current owner" and CLEARING
            # it when the inner </w:p> closes would pass the nested-depth,
            # text-box, orphan, wrapper and cap rows while silently DROPPING
            # every run that follows a nested paragraph inside its own.
            # Measured: the correct traversal returns 'Paragraph before the
            # box. AND AFTER THE BOX'; that impostor returns 'Paragraph
            # before the box.'
            "<w:r><w:t> AND AFTER THE BOX</w:t></w:r></w:p>"
            "<w:t>ORPHAN OUTSIDE EVERY PARAGRAPH</w:t>"
            "<w:p><w:r><w:t>Paragraph after the box.</w:t></w:r></w:p>"
        )).encode("utf-8"))
        self.assertEqual(extract_text(path).splitlines(), [
            "Paragraph before the box. AND AFTER THE BOX",
            "TEXT INSIDE THE BOX",
            "Paragraph after the box.",
        ])

    def test_a_non_deflate_compression_method_is_refused_before_decompression(self):
        # BZIP2 and LZMA members are decompressed in whole blocks BEFORE
        # ZipExtFile slices the result to the declared size, so a member that
        # declares a few bytes can still make the decompressor allocate far
        # more than max_bytes (measured: about 2.4 MB for BZIP2 and 10.8 MB
        # for LZMA against a 384000-byte cap). The metadata screen therefore
        # accepts only the two methods real Word output uses, STORED and
        # DEFLATED, and refuses anything else before a member is opened.
        document_xml = _DOCUMENT_XML_TEMPLATE.format(
            body="<w:p><w:r><w:t>%s</w:t></w:r></w:p>"
            % " ".join(str(n * 7919) for n in range(20000))
        ).encode("utf-8")
        for label, method in (("bzip2", zipfile.ZIP_BZIP2),
                              ("lzma", zipfile.ZIP_LZMA)):
            path = self._raw_docx(document_xml, compression=method)

            def run():
                with self.assertRaisesRegex(ValueError, "compression method"):
                    extract_text(path)

            self.assertEqual(
                self._opens_during(run), [],
                "%s: a member was opened before the method was refused"
                % label)

    def test_stored_and_deflated_members_are_still_accepted(self):
        body = "<w:p><w:r><w:t>Plain text</w:t></w:r></w:p>"
        document_xml = _DOCUMENT_XML_TEMPLATE.format(body=body).encode("utf-8")
        for method in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
            path = self._raw_docx(document_xml, compression=method)
            self.assertEqual(extract_text(path), "Plain text")

    def test_archive_level_failures_surface_as_ValueError_on_every_route(self):
        # zipfile reports an unreadable archive as BadZipFile, an absent
        # word/document.xml as KeyError, and a damaged member as BadZipFile
        # (bad CRC) -- none of them the ValueError every other refusal in this
        # module uses. All three must come out as ValueError so a caller needs
        # to catch only one type, and none may return text.
        not_a_zip = self._zip_path()
        with open(not_a_zip, "wb") as handle:
            handle.write(b"not a zip archive at all")

        no_document = self._zip_path()
        with zipfile.ZipFile(no_document, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("word/styles.xml", "<w:styles/>")

        damaged = self._raw_docx(
            _DOCUMENT_XML_TEMPLATE.format(
                body="<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
            ).encode("utf-8"), compression=zipfile.ZIP_STORED)
        with open(damaged, "rb") as handle:
            blob = bytearray(handle.read())
        blob[blob.index(b"hello")] ^= 0xFF   # flip one stored byte: CRC no longer matches
        with open(damaged, "wb") as handle:
            handle.write(bytes(blob))

        for name, path in (("not a zip", not_a_zip),
                           ("no word/document.xml", no_document),
                           ("damaged member", damaged)):
            for label in ("path", "BytesIO", "handle"):
                with self.assertRaises(
                        ValueError,
                        msg="%s via %s did not raise ValueError" % (name, label)):
                    extract_text(self._route_source(path, label))

    def _body_text(self, body_xml, **kwargs):
        document_xml = _DOCUMENT_XML_TEMPLATE.format(
            body=body_xml).encode("utf-8")
        return extract_text(self._raw_docx(document_xml), **kwargs)

    def test_a_tab_run_separates_the_text_around_it(self):
        # Without this, "12", a tab, "34" extracted as "1234" and the file
        # tools handed the model a number that is not in the document.
        text = self._body_text(
            "<w:p><w:r><w:t>12</w:t></w:r><w:r><w:tab/></w:r>"
            "<w:r><w:t>34</w:t></w:r></w:p>")
        self.assertEqual(text, "12\t34")

    def test_a_tab_between_text_in_ONE_run_is_kept_in_order(self):
        text = self._body_text(
            "<w:p><w:r><w:t>12</w:t><w:tab/><w:t>34</w:t><w:tab/>"
            "<w:t>56</w:t></w:r></w:p>")
        self.assertEqual(text, "12\t34\t56")

    def test_soft_breaks_and_carriage_returns_become_newlines(self):
        for element in ("<w:br/>", '<w:br w:type="page"/>', "<w:cr/>"):
            with self.subTest(element=element):
                text = self._body_text(
                    "<w:p><w:r><w:t>12</w:t>%s<w:t>34</w:t></w:r></w:p>"
                    % element)
                self.assertEqual(text, "12\n34")

    def test_a_positional_tab_is_a_tab_and_a_no_break_hyphen_is_a_hyphen(self):
        self.assertEqual(self._body_text(
            '<w:p><w:r><w:t>12</w:t><w:ptab w:relativeTo="margin" '
            'w:alignment="left" w:leader="none"/><w:t>34</w:t></w:r></w:p>'),
            "12\t34")
        self.assertEqual(self._body_text(
            "<w:p><w:r><w:t>well</w:t><w:noBreakHyphen/><w:t>known</w:t>"
            "</w:r></w:p>"), "well-known")

    def test_a_tab_STOP_definition_is_not_a_tab_character(self):
        # <w:tabs><w:tab .../></w:tabs> inside paragraph properties defines a
        # stop; it is not a tab in the text. Only a tab that is a child of a
        # run counts. (The properties element is placed between two runs so a
        # leading-whitespace strip cannot hide a wrongly emitted tab.)
        text = self._body_text(
            "<w:p><w:r><w:t>a</w:t></w:r><w:pPr><w:tabs>"
            '<w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>'
            "<w:r><w:t>b</w:t></w:r></w:p>")
        self.assertEqual(text, "ab")

    def test_a_paragraph_holding_only_a_tab_or_break_is_still_skipped(self):
        text = self._body_text(
            "<w:p><w:r><w:tab/></w:r></w:p>"
            "<w:p><w:r><w:br/></w:r></w:p>"
            "<w:p><w:r><w:t>kept</w:t></w:r></w:p>")
        self.assertEqual(text, "kept")

    def test_separators_in_a_text_box_stay_with_their_own_paragraph(self):
        text = self._body_text(
            "<w:p><w:r><w:t>outer</w:t></w:r><w:r><w:pict>"
            "<w:txbxContent><w:p><w:r><w:t>12</w:t><w:tab/><w:t>34</w:t>"
            "</w:r></w:p></w:txbxContent></w:pict></w:r></w:p>")
        self.assertEqual(text.splitlines(), ["outer", "12\t34"])

    def test_separator_characters_count_toward_the_text_cap(self):
        # The cap bounds what is returned, so emitted separators must count.
        body = ("<w:p><w:r><w:t>a</w:t>%s<w:t>b</w:t></w:r></w:p>"
                % ("<w:tab/>" * 50))
        self.assertEqual(len(self._body_text(body, max_text_chars=52)), 52)
        with self.assertRaisesRegex(ValueError, "character cap"):
            self._body_text(body, max_text_chars=51)

    def test_an_UNKNOWN_xml_encoding_is_a_ValueError_on_every_route(self):
        # expat reports an encoding it does not know as LookupError, not
        # ExpatError, so the prescreen's except clause let it escape the
        # module's ValueError contract.
        document_xml = (
            '<?xml version="1.0" encoding="NOPE-9"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t>'
            "</w:r></w:p></w:body></w:document>").encode("ascii")
        path = self._raw_docx(document_xml)
        for label in ("path", "BytesIO", "handle"):
            with self.assertRaisesRegex(
                    ValueError, "not well-formed XML",
                    msg="%s did not raise ValueError" % label):
                extract_text(self._route_source(path, label))


    _STRICT_NS = "http://purl.oclc.org/ooxml/wordprocessingml/main"
    _TRANSITIONAL_NS = ("http://schemas.openxmlformats.org/wordprocessingml/"
                        "2006/main")
    _MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

    def _ns_docx(self, namespace, body_xml, extra_declarations=""):
        document_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="%s" xmlns:mc="%s"%s><w:body>%s</w:body>'
            "</w:document>" % (namespace, self._MC_NS, extra_declarations,
                               body_xml)).encode("utf-8")
        return self._raw_docx(document_xml)

    def test_a_STRICT_ooxml_document_extracts_the_same_text(self):
        # A valid Strict OOXML .docx used to extract as EMPTY text with no
        # error, because only the Transitional namespace was recognized.
        body = ("<w:p><w:r><w:t>12</w:t><w:tab/><w:t>34</w:t><w:br/>"
                "<w:t>56</w:t></w:r></w:p>"
                "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p>"
                "</w:tc></w:tr></w:tbl>")
        strict = extract_text(self._ns_docx(self._STRICT_NS, body))
        transitional = extract_text(self._ns_docx(self._TRANSITIONAL_NS, body))
        self.assertEqual(strict, "12\t34\n56\ncell")
        self.assertEqual(strict, transitional)

    def test_an_UNSUPPORTED_namespace_is_refused_not_returned_empty(self):
        body = "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"
        for label, namespace in (("unknown", "urn:example:not-word"),
                                 ("empty", "")):
            path = self._ns_docx(namespace, body) if namespace else \
                self._raw_docx(
                    b"<document><body><p><r><t>hello</t></r></p></body>"
                    b"</document>")
            for route in ("path", "BytesIO", "handle"):
                with self.assertRaisesRegex(
                        ValueError, "unsupported WordprocessingML namespace",
                        msg="%s via %s" % (label, route)):
                    extract_text(self._route_source(path, route))

    def test_a_recognized_namespace_with_a_different_root_is_refused(self):
        document_xml = (
            '<w:wrong xmlns:w="%s"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p>'
            "</w:body></w:wrong>" % self._TRANSITIONAL_NS).encode("utf-8")
        with self.assertRaisesRegex(
                ValueError, "unsupported WordprocessingML namespace"):
            extract_text(self._raw_docx(document_xml))

    def test_a_recognized_document_with_no_text_returns_empty_text(self):
        # extract_text's own contract stays "the visible text"; the file tools
        # are what refuse an empty result (see the public-seam rows).
        for namespace in (self._STRICT_NS, self._TRANSITIONAL_NS):
            self.assertEqual(extract_text(self._ns_docx(namespace, "<w:p/>")), "")

    _TEXTBOX = ("<w:txbxContent><w:p><w:r><w:t>BOX</w:t></w:r></w:p>"
                "</w:txbxContent>")

    def _alt_content(self, choice, fallback):
        parts = ["<mc:AlternateContent>"]
        if choice is not None:
            parts.append('<mc:Choice Requires="wps">%s</mc:Choice>' % choice)
        if fallback is not None:
            parts.append("<mc:Fallback>%s</mc:Fallback>" % fallback)
        parts.append("</mc:AlternateContent>")
        return "<w:p><w:r><w:t>before</w:t>%s</w:r></w:p>" % "".join(parts)

    def test_alternate_content_is_read_ONCE_not_per_branch(self):
        # Choice and Fallback are two renderings of the SAME content; reading
        # both extracted a text box twice.
        for namespace in (self._TRANSITIONAL_NS, self._STRICT_NS):
            body = self._alt_content(self._TEXTBOX, "<w:pict>%s</w:pict>"
                                     % self._TEXTBOX)
            text = extract_text(self._ns_docx(namespace, body))
            self.assertEqual(text.count("BOX"), 1, text)
            self.assertEqual(text.splitlines(), ["before", "BOX"])

    def test_alternate_content_prefers_the_choice_branch(self):
        body = self._alt_content(
            "<w:txbxContent><w:p><w:r><w:t>FROM CHOICE</w:t></w:r></w:p>"
            "</w:txbxContent>",
            "<w:txbxContent><w:p><w:r><w:t>FROM FALLBACK</w:t></w:r></w:p>"
            "</w:txbxContent>")
        text = extract_text(self._ns_docx(self._TRANSITIONAL_NS, body))
        self.assertEqual(text.splitlines(), ["before", "FROM CHOICE"])

    def test_alternate_content_uses_the_fallback_only_when_there_is_no_choice(
            self):
        body = self._alt_content(None, self._TEXTBOX)
        text = extract_text(self._ns_docx(self._TRANSITIONAL_NS, body))
        self.assertEqual(text.splitlines(), ["before", "BOX"])

    def test_text_outside_alternate_content_is_unaffected(self):
        body = (self._alt_content(self._TEXTBOX, self._TEXTBOX)
                + "<w:p><w:r><w:t>after</w:t></w:r></w:p>")
        text = extract_text(self._ns_docx(self._TRANSITIONAL_NS, body))
        self.assertEqual(text.splitlines(), ["before", "BOX", "after"])


if __name__ == "__main__":
    unittest.main()
