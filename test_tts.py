"""Unit tests for Text to Speech Batch.

All source code, identifiers, comments, and labels are written in English.
Persian characters in test assertions are written using \\uXXXX escape sequences.
"""

import os
import sys
import tempfile
import unittest
import shutil

import docx

import tts_gui
from tts_gui import (
    MAX_CHARS,
    VOICES,
    STYLE_PRESETS,
    DEFAULT_STYLE_LABEL,
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
)


class TestChunking(unittest.TestCase):
    def test_short_text(self):
        text = "This is a single short sentence."
        chunks = split_text_into_chunks(text, max_chars=3000)
        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0], text)

    def test_paragraph_splitting(self):
        # Paragraphs separated by blank lines
        p1 = "Paragraph one content."
        p2 = "Paragraph two content."
        combined = f"{p1}\n\n{p2}"
        
        # When max_chars is large enough, they merge with \n\n
        merged = split_text_into_chunks(combined, max_chars=3000)
        self.assertEqual(len(merged), 1)
        self.assertEqual(merged[0], f"{p1}\n\n{p2}")

        # When max_chars is small, they split into 2 chunks
        split_chunks = split_text_into_chunks(combined, max_chars=len(p1) + 5)
        self.assertEqual(len(split_chunks), 2)
        self.assertEqual(split_chunks[0], p1)
        self.assertEqual(split_chunks[1], p2)

    def test_sentence_delimiters(self):
        # Split after . ! ? : newline and Arabic question mark (\u061f) and Arabic semicolon (\u061b)
        text = (
            "Sentence one. Sentence two! Sentence three? Sentence four: "
            "Line five\nSentence six\u061f Sentence seven\u061b Final."
        )
        chunks = split_text_into_chunks(text, max_chars=25)
        self.assertTrue(len(chunks) >= 4)
        for c in chunks:
            self.assertLessEqual(len(c), 25)

    def test_hard_split_long_sentence(self):
        # A single sentence longer than max_chars without any punctuation
        long_sent = "A" * 7500
        chunks = split_text_into_chunks(long_sent, max_chars=3000)
        self.assertEqual(len(chunks), 3)
        self.assertEqual(len(chunks[0]), 3000)
        self.assertEqual(len(chunks[1]), 3000)
        self.assertEqual(len(chunks[2]), 1500)
        self.assertEqual("".join(chunks), long_sent)

    def test_empty_and_whitespace(self):
        self.assertEqual(split_text_into_chunks(""), [])
        self.assertEqual(split_text_into_chunks("   \n\n  \t  "), [])


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

    def test_no_collision_different_folders(self):
        files = [
            os.path.join("C:", "folder1", "report.txt"),
            os.path.join("C:", "folder2", "report.docx"),
        ]
        collisions = check_output_collisions(files)
        self.assertEqual(len(collisions), 0)


class TestDocxExtraction(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_persian_docx_extraction(self):
        # Programmatically construct a Persian .docx file
        docx_path = os.path.join(self.temp_dir, "sample_persian.docx")
        doc = docx.Document()

        # Paragraph 1 (Persian: "This is a test paragraph.")
        p1_text = "\u0627\u06cc\u0646 \u06cc\u06a9 \u067e\u0627\u0631\u0627\u06af\u0631\u0627\u0641 \u062a\u0633\u062a \u0627\u0633\u062a."
        doc.add_paragraph(p1_text)

        # Paragraph 2: Empty paragraph (should be skipped)
        doc.add_paragraph("   ")

        # Table: 2 rows, 2 cols (Persian cell labels)
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "\u0633\u0644\u0648\u0644 \u06f1"  # Cell 1
        tbl.cell(0, 1).text = "\u0633\u0644\u0648\u0644 \u06f2"  # Cell 2
        tbl.cell(1, 0).text = "\u0633\u0644\u0648\u0644 \u06f3"  # Cell 3
        tbl.cell(1, 1).text = "\u0633\u0644\u0648\u0644 \u06f4"  # Cell 4

        # Paragraph 3: Persian text containing ZWNJ (\u200c)
        # Persian: "He is going home." with ZWNJ in "mi-ravad"
        p3_text = "\u0645\u06cc\u200c\u0631\u0648\u062f \u0628\u0647 \u062e\u0627\u0646\u0647."
        doc.add_paragraph(p3_text)

        doc.save(docx_path)

        # Extract text using extract_text_from_docx
        extracted = extract_text_from_docx(docx_path)

        # Verify body order and formatting
        expected_table = (
            "\u0633\u0644\u0648\u0644 \u06f1 \u0633\u0644\u0648\u0644 \u06f2\n"
            "\u0633\u0644\u0648\u0644 \u06f3 \u0633\u0644\u0648\u0644 \u06f4"
        )
        expected_full = f"{p1_text}\n\n{expected_table}\n\n{p3_text}"

        self.assertEqual(extracted, expected_full)
        # Check ZWNJ was preserved
        self.assertIn("\u200c", extracted)

    def test_empty_docx_rejected(self):
        docx_path = os.path.join(self.temp_dir, "empty.docx")
        doc = docx.Document()
        doc.add_paragraph("   ")
        doc.save(docx_path)

        with self.assertRaises(ValueError) as ctx:
            extract_text_from_docx(docx_path)
        self.assertIn("no text found", str(ctx.exception))


class TestTxtExtraction(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_encodings_support(self):
        test_string = "Persian test: \u0633\u0644\u0627\u0645 \u062c\u0647\u0627\u0646"
        for enc in ["utf-8-sig", "utf-8", "utf-16", "cp1256"]:
            fpath = os.path.join(self.temp_dir, f"test_{enc}.txt")
            with open(fpath, "wb") as f:
                f.write(test_string.encode(enc))
            result = extract_text_from_txt(fpath)
            self.assertEqual(result, test_string)

    def test_empty_txt_rejected(self):
        fpath = os.path.join(self.temp_dir, "empty.txt")
        with open(fpath, "wb") as f:
            f.write(b"   \n  \t  ")
        with self.assertRaises(ValueError) as ctx:
            extract_text_from_txt(fpath)
        self.assertIn("no text found", str(ctx.exception))


class TestPromptAndStyles(unittest.TestCase):
    def test_build_prompt_without_style(self):
        chunk = "Sample chunk content."
        prompt = build_prompt("", chunk)
        self.assertEqual(prompt, f"## Transcript:\n{chunk}")

    def test_build_prompt_with_preset_style(self):
        # Cheerful preset
        label = "\u0634\u0627\u062f \u0648 \u067e\u0631\u0627\u0646\u0631\u0698\u06cc"
        instruction = resolve_style_instruction(label)
        self.assertEqual(instruction, "Say in a cheerful, upbeat and energetic tone:")
        
        chunk = "Sample chunk content."
        prompt = build_prompt(instruction, chunk)
        self.assertEqual(prompt, f"{instruction}\n\n## Transcript:\n{chunk}")

    def test_custom_free_text_style(self):
        # Without colon: append colon
        custom_raw = "Speak like a wise teacher"
        resolved = resolve_style_instruction(custom_raw)
        self.assertEqual(resolved, "Speak like a wise teacher:")
        prompt = build_prompt(resolved, "Text here.")
        self.assertEqual(prompt, "Speak like a wise teacher:\n\n## Transcript:\nText here.")

        # With colon: unchanged
        custom_with_colon = "Whisper with a mysterious tone:"
        resolved2 = resolve_style_instruction(custom_with_colon)
        self.assertEqual(resolved2, custom_with_colon)


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
        # 1 second of 24000Hz 16-bit mono silence (48,000 bytes)
        raw_pcm = b"\x00" * (24000 * 2)
        save_pcm_as_mp3(raw_pcm, 24000, out_mp3)

        self.assertTrue(os.path.isfile(out_mp3))
        mp3_size = os.path.getsize(out_mp3)
        self.assertGreater(mp3_size, 0)
        # MP3 encoded at 64kbps must be significantly smaller than 48,000 bytes PCM
        self.assertLess(mp3_size, 20000)

    def test_mock_pipeline_end_to_end(self):
        # Create text file
        txt_path = os.path.join(self.temp_dir, "sample.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("Line one of sample.\nLine two of sample.")

        # Create docx file
        docx_path = os.path.join(self.temp_dir, "sample.docx")
        doc = docx.Document()
        doc.add_paragraph("Word document line.")
        doc.save(docx_path)

        # Run extraction, chunking, and mock audio generation for both files
        for test_file in [txt_path, docx_path]:
            text = extract_text(test_file)
            chunks = split_text_into_chunks(text, max_chars=3000)
            self.assertGreaterEqual(len(chunks), 1)

            # In mock mode, generate silence per chunk
            total_pcm = bytearray()
            for chunk in chunks:
                pcm_silence = b"\x00" * 48000
                total_pcm.extend(pcm_silence)

            out_mp3 = os.path.splitext(test_file)[0] + ".mp3"
            save_pcm_as_mp3(bytes(total_pcm), 24000, out_mp3)

            self.assertTrue(os.path.isfile(out_mp3))
            self.assertGreater(os.path.getsize(out_mp3), 1000)


if __name__ == "__main__":
    unittest.main()

