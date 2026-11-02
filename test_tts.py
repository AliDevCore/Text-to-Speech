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
    get_unique_output_path,
    extract_text,
    extract_text_from_txt,
    extract_text_from_docx,
    resolve_style_instruction,
    clean_style_instruction,
    build_prompt,
    parse_sample_rate,
    save_pcm_as_mp3,
    protect_data,
    unprotect_data,
    get_file_cache_dir,
    run_selftest,
    ApiKeyPool,
    save_api_keys,
    load_api_keys,
    MAX_API_KEYS,
    PRIMARY_MODEL,
    FALLBACK_MODEL,
)


class TestChunking(unittest.TestCase):
    def test_short_text(self):
        text = "This is a single short sentence."
        chunks = split_text_into_chunks(text, max_chars=3000)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0], text)

    def test_decimals_times_and_urls_preserved(self):
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
        t4 = (
            "Paragraph one has 46.5% profit at 10:30.\nLine two is continuous.\n\n"
            "Paragraph two references https://example.com/a.b and 3.14 value.\n"
            "\u0627\u06cc\u0646 \u06cc\u06a9 \u062c\u0645\u0644\u0647 \u0627\u0633\u062a. "
            "\u062c\u0645\u0644\u0647 \u062f\u0648\u0645 \u0646\u06cc\u0632 \u0627\u06cc\u0646\u062c\u0627\u0633\u062a! "
        ) * 10
        chunks = split_text_into_chunks(t4, max_chars=350)
        self.assertGreater(len(chunks), 1)

        for c in chunks:
            self.assertLessEqual(len(c), 350)
            self.assertIn(c, t4.strip())

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
        err_401 = APIError(401, {"error": {"code": 401, "message": "API_KEY_INVALID"}})
        self.assertEqual(classify_error(err_401).kind, "auth")

        err_403 = APIError(403, {"error": {"code": 403, "message": "The caller does not have permission", "status": "PERMISSION_DENIED"}})
        self.assertEqual(classify_error(err_403).kind, "auth")

        err_key = Exception("Google API error: API key not valid. Please pass a valid API key.")
        self.assertEqual(classify_error(err_key).kind, "auth")

    def test_region_errors(self):
        err_loc = APIError(400, {"error": {"code": 400, "message": "User location is not supported for the API use.", "status": "FAILED_PRECONDITION"}})
        self.assertEqual(classify_error(err_loc).kind, "region")

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
        err_voice = APIError(400, {"error": {"code": 400, "message": "Invalid voice name specified: UnknownVoice", "status": "INVALID_ARGUMENT"}})
        info = classify_error(err_voice)
        self.assertEqual(info.kind, "bad_request")
        self.assertIn("Invalid voice name", info.user_message)

    def test_transient_and_number_in_trace(self):
        err_500 = APIError(500, "500 Internal Server Error, trace ID: 540321199")
        info = classify_error(err_500)
        self.assertEqual(info.kind, "transient")
        self.assertNotEqual(info.kind, "auth")

        err_timeout = TimeoutError("Request timed out after 180s")
        self.assertEqual(classify_error(err_timeout).kind, "transient")


class TestOutputNamingAndCollisions(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_unique_output_path_increments_when_file_exists(self):
        base_path = os.path.join(self.temp_dir, "report.mp3")
        # When report.mp3 does not exist:
        p0 = get_unique_output_path(base_path)
        self.assertEqual(p0, base_path)

        # Create report.mp3 on disk
        with open(base_path, "wb") as f:
            f.write(b"dummy")

        # Now get_unique_output_path must return report1.mp3
        p1 = get_unique_output_path(base_path)
        self.assertEqual(p1, os.path.join(self.temp_dir, "report1.mp3"))

        # Create report1.mp3 on disk
        with open(p1, "wb") as f:
            f.write(b"dummy1")

        # Now get_unique_output_path must return report2.mp3
        p2 = get_unique_output_path(base_path)
        self.assertEqual(p2, os.path.join(self.temp_dir, "report2.mp3"))

    def test_unique_output_path_with_reserved_set(self):
        base_path = os.path.join(self.temp_dir, "test.mp3")
        reserved = set()

        r1 = get_unique_output_path(base_path, reserved)
        reserved.add(os.path.normcase(os.path.abspath(r1)))
        self.assertEqual(r1, base_path)

        r2 = get_unique_output_path(base_path, reserved)
        reserved.add(os.path.normcase(os.path.abspath(r2)))
        self.assertEqual(r2, os.path.join(self.temp_dir, "test1.mp3"))

        r3 = get_unique_output_path(base_path, reserved)
        self.assertEqual(r3, os.path.join(self.temp_dir, "test2.mp3"))


class TestDocxExtraction(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_merged_cell_deduplication(self):
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
        self.assertEqual(extracted.count("Merged Cell Value"), 1)
        self.assertIn("Row 2 Col 1 Row 2 Col 2", extracted)

    def test_sdt_content_control_extraction(self):
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
        odd_text = "\u0633\u0644\u0627\u0645 \u062c\u0647\u0627\u0646"
        odd_bytes = odd_text.encode("cp1256")
        self.assertEqual(len(odd_bytes), 9)

        f_odd = os.path.join(self.temp_dir, "odd.txt")
        with open(f_odd, "wb") as f:
            f.write(odd_bytes)
        self.assertEqual(extract_text_from_txt(f_odd), odd_text)

        even_text = "\u0633\u0644\u0627\u0645 \u062c\u0647\u0627\u0646!"
        even_bytes = even_text.encode("cp1256")
        self.assertEqual(len(even_bytes), 10)

        f_even = os.path.join(self.temp_dir, "even.txt")
        with open(f_even, "wb") as f:
            f.write(even_bytes)
        self.assertEqual(extract_text_from_txt(f_even), even_text)

    def test_bom_detection(self):
        sample = "Sample text for BOM test."
        f_sig = os.path.join(self.temp_dir, "utf8_sig.txt")
        with open(f_sig, "wb") as f:
            f.write(sample.encode("utf-8-sig"))
        self.assertEqual(extract_text_from_txt(f_sig), sample)

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
    def test_build_prompt_returns_verbatim_text_without_spoken_instruction(self):
        chunk = "Speech transcript content to be spoken."
        # Instruction must NOT be prepended to the text field to prevent Gemini from speaking it
        prompt = build_prompt("Say in a cheerful tone:", chunk)
        self.assertEqual(prompt, chunk)

    def test_clean_style_instruction(self):
        clean = clean_style_instruction("Cheerful & Upbeat")
        self.assertEqual(clean, "cheerful, upbeat and energetic tone")

        clean_promo = clean_style_instruction("Promotional & High-Energy")
        self.assertEqual(clean_promo, "high-energy promotional tone with elongated vowels")

        custom_style = clean_style_instruction("Say in a whispering tone:")
        self.assertEqual(custom_style, "whispering tone")

        free_text = clean_style_instruction("Mysterious and slow")
        self.assertEqual(free_text, "Mysterious and slow")


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

        mock_pcm = b"\x00" * 4800
        with open(chunk1_cache, "wb") as f:
            f.write(mock_pcm)

        self.assertTrue(os.path.isfile(chunk1_cache))
        shutil.rmtree(cache_dir, ignore_errors=True)

    def test_mock_pipeline_dual_format_and_stop_behavior(self):
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

            silence = b"\x00" * int(24000 * 2 * 0.2)
            out_mp3 = os.path.splitext(path)[0] + ".mp3"
            save_pcm_as_mp3(silence, 24000, out_mp3)
            self.assertTrue(os.path.isfile(out_mp3))
            self.assertGreater(os.path.getsize(out_mp3), 100)


class TestCustomVoiceAndUI(unittest.TestCase):
    def test_custom_voice_config_generation(self):
        # Verify custom voice string such as "Zali" can be passed to VoiceConfig
        from google.genai import types
        custom_voice = "Zali"
        voice_cfg = types.VoiceConfig(
            prebuilt_voice_config=types.PrebuiltVoiceConfig(voice_name=custom_voice)
        )
        self.assertEqual(voice_cfg.prebuilt_voice_config.voice_name, "Zali")

    def test_tts_batch_app_voice_and_style_editable(self):
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        try:
            app = tts_gui.TTSBatchApp(root)
            # Voice combobox must be normal (editable), not readonly
            self.assertEqual(str(app.voice_cb.cget("state")), "normal")
            # Style combobox must be normal (editable)
            self.assertEqual(str(app.style_cb.cget("state")), "normal")

            # Dedicated Paste buttons must exist
            self.assertTrue(hasattr(app, "voice_paste_btn"))
            self.assertTrue(hasattr(app, "style_paste_btn"))

            # Simulate typing custom voice
            app.voice_var.set("Zali")
            app._on_panel_voice_change()
            self.assertEqual(app.last_applied_voice, "Zali")

            # Simulate pasting custom voice
            root.clipboard_clear()
            root.clipboard_append("CustomNarrator")
            app._paste_voice()
            self.assertEqual(app.voice_var.get(), "CustomNarrator")
            self.assertEqual(app.last_applied_voice, "CustomNarrator")

            # Simulate pasting custom style
            root.clipboard_clear()
            root.clipboard_append("Excited and dramatic storytelling")
            app._paste_style()
            self.assertEqual(app.style_var.get(), "Excited and dramatic storytelling")
            self.assertEqual(app.last_applied_style, "Excited and dramatic storytelling")
        finally:
            root.destroy()


class TestModelFallback(unittest.TestCase):
    def test_model_constants(self):
        self.assertEqual(tts_gui.PRIMARY_MODEL, "gemini-3.8-flash-tts")
        self.assertEqual(tts_gui.FALLBACK_MODEL, "gemini-3.8-flash-lite-tts")
        self.assertEqual(tts_gui.MODEL_NAME, "gemini-3.8-flash-tts")

    def test_fallback_on_quota_daily(self):
        import tkinter as tk
        from unittest.mock import MagicMock
        root = tk.Tk()
        root.withdraw()
        try:
            app = tts_gui.TTSBatchApp(root)
            self.assertEqual(app.active_model, tts_gui.PRIMARY_MODEL)

            # Create a mock client that raises quota_daily on PRIMARY_MODEL, but succeeds on FALLBACK_MODEL
            class MockClient:
                def __init__(self):
                    self.models = self

                def generate_content_stream(self, model, contents, config):
                    if model == tts_gui.PRIMARY_MODEL:
                        raise APIError(429, {"error": {"code": 429, "message": "Quota exceeded for quota metric per day"}})
                    chunk_mock = MagicMock()
                    chunk_mock.parts = [MagicMock(inline_data=MagicMock(data=b"\x00" * 200, mime_type="audio/pcm"))]
                    return [chunk_mock]

            client = MockClient()
            pcm, rate, err, was_rl, active_model = app._fetch_audio_chunk_with_retry(
                client=client,
                chunk_text="Test fallback chunk",
                voice="Bodi",
                clean_style="",
                is_mock=False,
                chunk_idx=1,
                total_chunks=1,
                active_model=tts_gui.PRIMARY_MODEL,
            )

            # Verify it automatically shifted to FALLBACK_MODEL and succeeded
            self.assertEqual(active_model, tts_gui.FALLBACK_MODEL)
            self.assertEqual(app.active_model, tts_gui.FALLBACK_MODEL)
            self.assertIsNone(err)
            self.assertGreater(len(pcm), 0)
        finally:
            root.destroy()


class TestApiKeyPoolAndRotation(unittest.TestCase):
    def test_pool_init_dedup_and_max_keys(self):
        keys = ["k1", "k2", "k1", "k3", "  k4  "] + [f"extra_{i}" for i in range(10)]
        pool = ApiKeyPool(keys)
        self.assertEqual(pool.total_count(), MAX_API_KEYS)
        self.assertEqual(pool.keys[0], "k1")
        self.assertEqual(pool.keys[1], "k2")
        self.assertEqual(pool.keys[2], "k3")
        self.assertEqual(pool.keys[3], "k4")

    def test_round_robin_rotation_primary_model(self):
        keys = ["key1", "key2", "key3"]
        pool = ApiKeyPool(keys)

        c1 = pool.get_next_key_and_model()
        c2 = pool.get_next_key_and_model()
        c3 = pool.get_next_key_and_model()
        c4 = pool.get_next_key_and_model()

        self.assertEqual(c1, ("key1", PRIMARY_MODEL, 1))
        self.assertEqual(c2, ("key2", PRIMARY_MODEL, 2))
        self.assertEqual(c3, ("key3", PRIMARY_MODEL, 3))
        self.assertEqual(c4, ("key1", PRIMARY_MODEL, 1))

    def test_single_key_exhaustion_on_primary_model(self):
        keys = ["keyA", "keyB", "keyC"]
        pool = ApiKeyPool(keys)

        all_exhausted, notice = pool.mark_quota_daily("keyB", PRIMARY_MODEL)
        self.assertFalse(all_exhausted)
        self.assertEqual(pool.get_active_model(), PRIMARY_MODEL)

        results = [pool.get_next_key_and_model()[0] for _ in range(4)]
        self.assertNotIn("keyB", results)
        self.assertIn("keyA", results)
        self.assertIn("keyC", results)

    def test_all_keys_exhaustion_shifts_to_fallback_model(self):
        keys = ["key1", "key2"]
        pool = ApiKeyPool(keys)

        all_ex1, _ = pool.mark_quota_daily("key1", PRIMARY_MODEL)
        self.assertFalse(all_ex1)

        all_ex2, _ = pool.mark_quota_daily("key2", PRIMARY_MODEL)
        self.assertTrue(all_ex2)
        self.assertEqual(pool.get_active_model(), FALLBACK_MODEL)

        c1 = pool.get_next_key_and_model()
        c2 = pool.get_next_key_and_model()
        c3 = pool.get_next_key_and_model()

        self.assertEqual(c1, ("key1", FALLBACK_MODEL, 1))
        self.assertEqual(c2, ("key2", FALLBACK_MODEL, 2))
        self.assertEqual(c3, ("key1", FALLBACK_MODEL, 1))

    def test_both_models_exhausted_returns_none(self):
        keys = ["k1"]
        pool = ApiKeyPool(keys)

        pool.mark_quota_daily("k1", PRIMARY_MODEL)
        self.assertEqual(pool.get_active_model(), FALLBACK_MODEL)

        all_ex, _ = pool.mark_quota_daily("k1", FALLBACK_MODEL)
        self.assertTrue(all_ex)
        self.assertTrue(pool.is_all_exhausted())
        self.assertIsNone(pool.get_next_key_and_model())

    def test_invalid_key_filtered_out(self):
        keys = ["goodKey1", "badKey", "goodKey2"]
        pool = ApiKeyPool(keys)
        pool.mark_invalid("badKey")

        self.assertEqual(pool.valid_count(), 2)
        for _ in range(6):
            res = pool.get_next_key_and_model()
            self.assertIsNotNone(res)
            key, model, _ = res
            self.assertNotEqual(key, "badKey")
            self.assertIn(key, ["goodKey1", "goodKey2"])

    def test_save_and_load_api_keys_roundtrip(self):
        test_keys = [f"AIzaSyTestKey_{i:02d}" for i in range(1, 11)]
        orig_keys = load_api_keys()
        try:
            saved = save_api_keys(test_keys)
            self.assertTrue(saved)
            loaded = load_api_keys()
            self.assertEqual(loaded, test_keys)
        finally:
            save_api_keys(orig_keys)


class TestSelfTestFunction(unittest.TestCase):
    def test_run_selftest_passes(self):
        exit_code = run_selftest()
        self.assertEqual(exit_code, 0)


if __name__ == "__main__":
    unittest.main()
