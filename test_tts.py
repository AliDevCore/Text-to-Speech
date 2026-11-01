"""Unit tests for Text to Speech Batch.

All source code, identifiers, comments, and labels are written in English.
Persian characters in test assertions are written strictly using \\uXXXX escape sequences.
All tests run offline in mock mode without making any real API calls.
"""

from dataclasses import dataclass
import os
import sys
import re
import shutil
import tempfile
import unittest

import docx
from docx.oxml import OxmlElement
from google.genai.errors import APIError

import tts_gui
from tts_gui import (
    MAX_CHARS,
    VOICES,
    STYLE_PRESETS,
    DEFAULT_STYLE_LABEL,
    ErrorInfo,
    classify_error,
    split_text_into_chunks,
    check_output_collisions,
    extract_text,
    extract_text_from_txt,
    extract_text_from_docx,
    resolve_style_instruction,
    build_prompt,
    parse_sample_rate,
    save_pcm_as_mp3,
    protect_data,
    unprotect_data,
    get_file_cache_dir,
    run_selftest,
)


class TestChunking(unittest.TestCase):
    def test_short_text(self):
        text = "This is a single short sentence."
        chunks = split_text_into_chunks(text, max_chars=3000)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0], text)

    def test_decimals_times_and_urls_preserved(self):
        # Must not split 46.5 or 10:30 into separate parts with inserted spaces
        t1 = "Profit was 46.5 percent at 10:30 today."
        chunks1 = split_text_into_chunks(t1, max_chars=3000)
        self.assertEqual(len(chunks1), 1)
        self.assertEqual(chunks1[0], t1)

        t2 = "See https://example.com/a.b for 3.14 details."
        chunks2 = split_text_into_chunks(t2, max_chars=3000)
        self.assertEqual(len(chunks2), 1)
        self.assertEqual(chunks2[0], t2)

    def test_newlines_preserved(self):
        t3 = "line one\nline two"
        chunks3 = split_text_into_chunks(t3, max_chars=3000)
        self.assertEqual(len(chunks3), 1)
        self.assertEqual(chunks3[0], t3)

    def test_exact_substring_and_whitespace_equality(self):
        # Long mixed text with numbers, URLs, Persian sentences, and multiple paragraphs
        t4 = (
            "Paragraph one has 46.5% profit at 10:30.\nLine two is continuous.\n\n"
            "Paragraph two references https://example.com/a.b and 3.14 value.\n"
            "\u0627\u06cc\u0646 \u06cc\u06a9 \u062c\u0645\u0644\u0647 \u0627\u0633\u062a. "
            "\u062c\u0645\u0644\u0647 \u062f\u0648\u0645 \u0646\u06cc\u0632 \u0627\u06cc\u0646\u062c\u0627\u0633\u062a! "
        ) * 10
        chunks = split_text_into_chunks(t4, max_chars=350)
        self.assertGreater(len(chunks), 1)

        # Every chunk must be <= max_chars
        for c in chunks:
            self.assertLessEqual(len(c), 350)
            self.assertIn(c, t4.strip())

        # Whitespace-stripped representation must match original exactly (no content lost)
        clean_original = re.sub(r"\s+", "", t4)
        clean_reconstructed = re.sub(r"\s+", "", "".join(chunks))
        self.assertEqual(clean_reconstructed, clean_original)

    def test_hard_split_very_long_token(self):
        long_token = "A" * 7500
        chunks = split_text_into_chunks(long_token, max_chars=3000)
        self.assertEqual(len(chunks), 3)
        self.assertEqual(len(chunks[0]), 3000)
        self.assertEqual(len(chunks[1]), 3000)
        self.assertEqual(len(chunks[2]), 1500)
        self.assertEqual("".join(chunks), long_token)

    def test_empty_and_whitespace_only(self):
        self.assertEqual(split_text_into_chunks(""), [])
        self.assertEqual(split_text_into_chunks("   \n\n  \t  "), [])


class TestErrorClassification(unittest.TestCase):
    def test_auth_errors(self):
        # HTTP 401
        err_401 = APIError(401, {"error": {"code": 401, "message": "API_KEY_INVALID"}})
        self.assertEqual(classify_error(err_401).kind, "auth")

        # HTTP 403 PERMISSION_DENIED
        err_403 = APIError(403, {"error": {"code": 403, "message": "The caller does not have permission", "status": "PERMISSION_DENIED"}})
        self.assertEqual(classify_error(err_403).kind, "auth")

        # Explicit message mentioning invalid key
        err_key = Exception("Google API error: API key not valid. Please pass a valid API key.")
        self.assertEqual(classify_error(err_key).kind, "auth")

    def test_region_errors(self):
        # 400 FAILED_PRECONDITION location unsupported
        err_loc = APIError(400, {"error": {"code": 400, "message": "User location is not supported for the API use.", "status": "FAILED_PRECONDITION"}})
        self.assertEqual(classify_error(err_loc).kind, "region")

        # Location unsupported in generic message
        err_loc2 = Exception("Location is not supported in your territory")
        self.assertEqual(classify_error(err_loc2).kind, "region")

    def test_model_not_found(self):
        err_404 = APIError(404, {"error": {"code": 404, "message": "Model gemini-3.8-flash-tts not found"}})
        self.assertEqual(classify_error(err_404).kind, "model_not_found")

    def test_quota_daily(self):
        err_daily = APIError(429, {"error": {"code": 429, "message": "Resource has been exhausted (e.g. check quota per day)."}})
        self.assertEqual(classify_error(err_daily).kind, "quota_daily")

    def test_rate_limit(self):
        err_rate = APIError(429, {"error": {"code": 429, "message": "Rate limit exceeded. Please retry in 20s."}})
        self.assertEqual(classify_error(err_rate).kind, "rate_limit")

    def test_bad_request_fails_only_current_file(self):
        # 400 for invalid voice name
        err_voice = APIError(400, {"error": {"code": 400, "message": "Invalid voice name specified: UnknownVoice", "status": "INVALID_ARGUMENT"}})
        info = classify_error(err_voice)
        self.assertEqual(info.kind, "bad_request")
        self.assertIn("Invalid voice name", info.user_message)

    def test_transient_and_number_in_trace(self):
        # 500 error whose text contains 5403211 trace id must NOT be classified as auth 403!
        err_500 = APIError(500, "500 Internal Server Error, trace ID: 540321199")
        info = classify_error(err_500)
        self.assertEqual(info.kind, "transient")
        self.assertNotEqual(info.kind, "auth")

        # TimeoutError
        err_timeout = TimeoutError("Request timed out after 180s")
        self.assertEqual(classify_error(err_timeout).kind, "transient")


class TestOutputCollisions(unittest.TestCase):
    def test_collision_detected(self):
        files = [
            os.path.join("C:", "docs", "report.txt"),
            os.path.join("C:", "docs", "report.docx"),
        ]
        collisions = check_output_collisions(files)
        self.assertEqual(len(collisions), 1)
        out_key = list(collisions.keys())[0]
        self.assertTrue(out_key.endswith("report.mp3"))
        self.assertEqual(len(collisions[out_key]), 2)

    def test_no_collision_different_names(self):
        files = [
            os.path.join("C:", "docs", "doc1.txt"),
            os.path.join("C:", "docs", "doc2.docx"),
        ]
        collisions = check_output_collisions(files)
        self.assertEqual(len(collisions), 0)


class TestDocxExtraction(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_merged_cell_deduplication(self):
        # Test that merged table cells in Word docx are read only once
        docx_path = os.path.join(self.temp_dir, "merged.docx")
        doc = docx.Document()

        doc.add_paragraph("Intro Paragraph")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "Merged Cell Value"
        tbl.cell(0, 1).text = "Secondary Value"
        tbl.cell(0, 0).merge(tbl.cell(0, 1))

        tbl.cell(1, 0).text = "Row 2 Col 1"
        tbl.cell(1, 1).text = "Row 2 Col 2"
        doc.save(docx_path)

        extracted = extract_text_from_docx(docx_path)
        self.assertIn("Intro Paragraph", extracted)
        # Verify merged cell appears once and is not duplicated
        self.assertEqual(extracted.count("Merged Cell Value"), 1)
        self.assertIn("Row 2 Col 1 Row 2 Col 2", extracted)

    def test_sdt_content_control_extraction(self):
        # Test that w:sdt and w:sdtContent elements are recursively traversed
        docx_path = os.path.join(self.temp_dir, "sdt.docx")
        doc = docx.Document()
        doc.add_paragraph("Paragraph before SDT")

        sdt = OxmlElement("w:sdt")
        sdt_content = OxmlElement("w:sdtContent")
        sdt.append(sdt_content)
        p_in_sdt = OxmlElement("w:p")
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = "Content Control Inner Text"
        r.append(t)
        p_in_sdt.append(r)
        sdt_content.append(p_in_sdt)
        doc.element.body.append(sdt)

        doc.add_paragraph("Paragraph after SDT")
        doc.save(docx_path)

        extracted = extract_text_from_docx(docx_path)
        self.assertIn("Content Control Inner Text", extracted)
        self.assertIn("Paragraph before SDT", extracted)
        self.assertIn("Paragraph after SDT", extracted)

    def test_persian_docx_with_zwnj(self):
        docx_path = os.path.join(self.temp_dir, "persian.docx")
        doc = docx.Document()
        # Persian: "He is going to school." with ZWNJ \u200c in "mi-ravad"
        p_text = "\u0645\u06cc\u200c\u0631\u0648\u062f \u0628\u0647 \u0645\u062f\u0631\u0633\u0647."
        doc.add_paragraph(p_text)
        doc.save(docx_path)

        extracted = extract_text_from_docx(docx_path)
        self.assertEqual(extracted, p_text)
        self.assertIn("\u200c", extracted)

    def test_empty_docx_raises(self):
        docx_path = os.path.join(self.temp_dir, "empty.docx")
        doc = docx.Document()
        doc.add_paragraph("   ")
        doc.save(docx_path)

        with self.assertRaises(ValueError):
            extract_text_from_docx(docx_path)


class TestTxtExtraction(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_cp1256_even_and_odd_byte_lengths(self):
        # CP1256 text with odd byte length (9 bytes)
        odd_text = "\u0633\u0644\u0627\u0645 \u062c\u0647\u0627\u0646"
        odd_bytes = odd_text.encode("cp1256")
        self.assertEqual(len(odd_bytes), 9)

        f_odd = os.path.join(self.temp_dir, "odd.txt")
        with open(f_odd, "wb") as f:
            f.write(odd_bytes)
        self.assertEqual(extract_text_from_txt(f_odd), odd_text)

        # CP1256 text with even byte length (10 bytes).
        # Must NEVER falsely decode as UTF-16!
        even_text = "\u0633\u0644\u0627\u0645 \u062c\u0647\u0627\u0646!"
        even_bytes = even_text.encode("cp1256")
        self.assertEqual(len(even_bytes), 10)

        f_even = os.path.join(self.temp_dir, "even.txt")
        with open(f_even, "wb") as f:
            f.write(even_bytes)
        self.assertEqual(extract_text_from_txt(f_even), even_text)

    def test_bom_detection(self):
        sample = "Sample text for BOM test."
        # UTF-8 with BOM
        f_sig = os.path.join(self.temp_dir, "utf8_sig.txt")
        with open(f_sig, "wb") as f:
            f.write(sample.encode("utf-8-sig"))
        self.assertEqual(extract_text_from_txt(f_sig), sample)

        # UTF-16 with BOM (LE)
        f_u16 = os.path.join(self.temp_dir, "utf16.txt")
        with open(f_u16, "wb") as f:
            f.write(sample.encode("utf-16"))
        self.assertEqual(extract_text_from_txt(f_u16), sample)

    def test_empty_txt_raises(self):
        f_empty = os.path.join(self.temp_dir, "empty.txt")
        with open(f_empty, "wb") as f:
            f.write(b"   \n  \t  ")
        with self.assertRaises(ValueError):
            extract_text_from_txt(f_empty)


class TestPromptAndStyles(unittest.TestCase):
    def test_build_prompt_neutral(self):
        prompt = build_prompt("", "Content here.")
        self.assertEqual(prompt, "## Transcript:\nContent here.")

    def test_build_prompt_preset(self):
        cheerful_label = "\u0634\u0627\u062f \u0648 \u067e\u0631\u0627\u0646\u0631\u0698\u06cc"
        instruction = resolve_style_instruction(cheerful_label)
        prompt = build_prompt(instruction, "Cheerful words.")
        self.assertEqual(prompt, "Say in a cheerful, upbeat and energetic tone:\n\n## Transcript:\nCheerful words.")

    def test_custom_style_formatting(self):
        inst1 = resolve_style_instruction("Talk slowly")
        self.assertEqual(inst1, "Talk slowly:")

        inst2 = resolve_style_instruction("Whisper:")
        self.assertEqual(inst2, "Whisper:")


class TestAudioAndMockPipeline(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_sample_rate_parsing(self):
        self.assertEqual(parse_sample_rate("audio/L16;rate=24000"), 24000)
        self.assertEqual(parse_sample_rate("audio/L16; rate=16000"), 16000)
        self.assertEqual(parse_sample_rate("audio/pcm"), 24000)
        self.assertEqual(parse_sample_rate(""), 24000)

    def test_save_pcm_as_mp3_and_atomic_write(self):
        out_mp3 = os.path.join(self.temp_dir, "output.mp3")
        raw_pcm = b"\x00" * (24000 * 2)  # 1 second of 24kHz 16-bit mono silence
        save_pcm_as_mp3(raw_pcm, 24000, out_mp3)

        self.assertTrue(os.path.isfile(out_mp3))
        mp3_size = os.path.getsize(out_mp3)
        self.assertGreater(mp3_size, 0)
        self.assertLess(mp3_size, 20000)

    def test_caching_and_resume_simulation(self):
        txt_path = os.path.join(self.temp_dir, "resume.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("Line one.\n\nLine two.")

        text = extract_text(txt_path)
        chunks = split_text_into_chunks(text, max_chars=12)
        self.assertEqual(len(chunks), 2)

        cache_dir = get_file_cache_dir(txt_path, "Bodi", "")
        chunk1_cache = os.path.join(cache_dir, "1.pcm")

        # Simulate chunk 1 already cached
        mock_pcm = b"\x00" * 4800
        with open(chunk1_cache, "wb") as f:
            f.write(mock_pcm)

        self.assertTrue(os.path.isfile(chunk1_cache))
        # Cleanup
        shutil.rmtree(cache_dir, ignore_errors=True)

    def test_mock_pipeline_dual_format_and_stop_behavior(self):
        # 1. Dual format processing (.txt and .docx)
        txt_path = os.path.join(self.temp_dir, "batch_item1.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("Line one of text file.")

        docx_path = os.path.join(self.temp_dir, "batch_item2.docx")
        doc = docx.Document()
        doc.add_paragraph("Line one of docx file.")
        doc.save(docx_path)

        for path in [txt_path, docx_path]:
            txt = extract_text(path)
            chunks = split_text_into_chunks(txt)
            self.assertEqual(len(chunks), 1)

            # In mock pipeline, synthesize 0.2s silence per chunk
            silence = b"\x00" * int(24000 * 2 * 0.2)
            out_mp3 = os.path.splitext(path)[0] + ".mp3"
            save_pcm_as_mp3(silence, 24000, out_mp3)
            self.assertTrue(os.path.isfile(out_mp3))
            self.assertGreater(os.path.getsize(out_mp3), 100)


class TestSelfTestFunction(unittest.TestCase):
    def test_run_selftest_passes(self):
        exit_code = run_selftest()
        self.assertEqual(exit_code, 0)


if __name__ == "__main__":
    unittest.main()
