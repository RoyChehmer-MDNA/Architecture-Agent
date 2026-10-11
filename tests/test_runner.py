import base64
import json
from pathlib import Path
import tempfile
import unittest
import os
import sys
import types
import zipfile
from unittest.mock import patch

import agent_runner as runner


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / "agents.json").write_text(json.dumps({"architecture": {
            "skill": "skill.md", "input": "input", "output": "output",
            "task": "Document this", "model": "test-model"}}))
        (self.root / "skill.md").write_text("# Required sections\nExplain components")
        self.image = self.root / "diagram.png"
        self.image.write_bytes(base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII="))
        self.calls = []

    def provider(self, payload):
        self.calls.append(payload)
        return "# Architecture\nTest document", {"response_id": "test"}

    def run_image(self, **kwargs):
        return runner.process(self.root, "architecture", self.image, provider=self.provider, **kwargs)

    def test_output_snapshot_and_deduplication(self):
        run = self.run_image()
        self.assertTrue((run / "diagram.docx").is_file())
        self.assertEqual((run / "skill.md").read_text(), (self.root / "skill.md").read_text())
        self.assertEqual(json.loads((run / "run.json").read_text())["artifact"], "diagram.docx")
        self.assertEqual(json.loads((run / "run.json").read_text())["source_image"], "diagram.png")
        self.assertEqual((run / "diagram.png").read_bytes(), self.image.read_bytes())
        self.assertIsNone(self.run_image())
        self.assertEqual(len(self.calls), 1)
        from docx import Document

        document = Document(run / "diagram.docx")
        self.assertEqual(len(document.inline_shapes), 1)
        with zipfile.ZipFile(run / "diagram.docx") as package:
            self.assertEqual(package.read("word/media/image1.png"), self.image.read_bytes())

    def test_saved_chart_remains_after_inbox_source_is_archived(self):
        inbox = self.root / "input"
        inbox.mkdir()
        source = inbox / "My chart.PNG"
        original = self.image.read_bytes()
        source.write_bytes(original)
        run = runner.process(self.root, "architecture", source, provider=self.provider)
        archived = runner.archive_processed_image(
            self.root, runner.read_config(self.root, "architecture"), source)
        self.assertFalse(source.exists())
        self.assertEqual((run / source.name).read_bytes(), original)
        self.assertEqual(archived.read_bytes(), original)

    def test_chart_copy_failure_does_not_mark_success(self):
        with patch.object(runner.shutil, "copy2", side_effect=OSError("Copy failed")):
            with self.assertRaisesRegex(OSError, "Copy failed"):
                self.run_image()
        self.assertTrue(self.image.is_file())
        self.assertEqual(list((self.root / "output").glob("*/run.json")), [])

    def test_chart_follows_description_before_detailed_sections(self):
        from docx import Document
        from docx.oxml.ns import qn

        for rtl in (False, True):
            with self.subTest(rtl=rtl):
                output = self.root / "opening.docx"
                runner.write_word_document(
                    "# Architecture\n\n## Executive summary\n\n"
                    "The client calls the API.\n\n## Components\n\nAPI details.",
                    output, rtl=rtl, image_path=self.image,
                )
                document = Document(output)
                paragraphs = document.paragraphs
                chart_index = next(index for index, paragraph in enumerate(paragraphs)
                                   if paragraph._p.find('.//' + qn('w:drawing')) is not None)
                self.assertEqual(paragraphs[chart_index - 1].text, "The client calls the API.")
                self.assertEqual(paragraphs[chart_index + 1].text, "Components")
                self.assertEqual(len(document.inline_shapes), 1)

    def test_jpeg_and_webp_charts_fit_page_and_preserve_proportions(self):
        from PIL import Image
        from docx import Document
        from docx.shared import Inches

        for extension in ("jpg", "webp"):
            with self.subTest(extension=extension):
                image_path = self.root / f"portrait.{extension}"
                Image.new("RGB", (100, 400), "white").save(image_path)
                output = self.root / f"portrait-{extension}.docx"
                runner.write_word_document("# Architecture\n\nDescription.", output,
                                           image_path=image_path)
                document = Document(output)
                chart = document.inline_shapes[0]
                section = document.sections[0]
                self.assertAlmostEqual(chart.width / chart.height, 0.25, places=5)
                self.assertLessEqual(chart.width,
                                     section.page_width - section.left_margin - section.right_margin)
                self.assertLessEqual(chart.height, Inches(6))

    def test_skill_and_task_changes_regenerate(self):
        first = self.run_image()
        (self.root / "skill.md").write_text("New instructions")
        second = self.run_image()
        third = self.run_image(task="Focus on security")
        self.assertEqual(len({first, second, third}), 3)
        self.assertIn("New instructions", self.calls[-1]["instructions"])

    def test_dry_run_has_no_side_effects(self):
        self.run_image(dry_run=True)
        self.assertEqual(self.calls, [])
        self.assertFalse((self.root / "output").exists())

    def test_default_watch_detects_new_image_after_stable_poll(self):
        folder = self.root / "input"
        folder.mkdir()
        polls = 0

        def next_poll(seconds):
            nonlocal polls
            polls += 1
            if polls == 1:
                (folder / "new.png").write_bytes(self.image.read_bytes())
            if polls == 4:
                raise KeyboardInterrupt

        with patch.object(runner, "ROOT", self.root), \
             patch.object(runner.time, "sleep", side_effect=next_poll), \
             patch.object(runner, "process") as process:
            self.assertEqual(runner.main([]), 0)
        process.assert_called_once()
        self.assertEqual(process.call_args.args[2], folder / "new.png")

    def test_failed_call_does_not_mark_success(self):
        def fail(payload):
            raise RuntimeError("Service unavailable")
        with self.assertRaises(RuntimeError):
            runner.process(self.root, "architecture", self.image, provider=fail)
        self.assertFalse((self.root / "output").exists())
        self.assertIsNotNone(self.run_image())

    def test_reject_empty_skill_and_unsupported_file(self):
        (self.root / "skill.md").write_text("")
        with self.assertRaisesRegex(ValueError, "empty"):
            self.run_image()

    def test_api_extracts_text_and_rejects_incomplete_response(self):
        from io import BytesIO
        payload = {"status": "completed", "output": [{"type": "message", "content": [
            {"type": "output_text", "text": "# Document"}]}]}
        with patch.dict("os.environ", {"OPENAI_API_KEY": "test"}):
            with patch("urllib.request.urlopen", return_value=BytesIO(json.dumps(payload).encode())):
                self.assertEqual(runner.call_api({})[0], "# Document")
            payload["status"] = "incomplete"
            with patch("urllib.request.urlopen", return_value=BytesIO(json.dumps(payload).encode())):
                with self.assertRaises(RuntimeError):
                    runner.call_api({})

    def test_env_file_loads_values_without_overriding_environment(self):
        env_file = self.root / ".env"
        env_file.write_text(
            "# Local configuration\nOPENAI_API_KEY='file-key'\nexport OPENAI_MODEL=gpt-file\n",
            encoding="utf-8",
        )
        with patch.dict(os.environ, {"OPENAI_API_KEY": "shell-key"}, clear=True):
            runner.load_env_file(env_file)
            self.assertEqual(os.environ["OPENAI_API_KEY"], "shell-key")
            self.assertEqual(os.environ["OPENAI_MODEL"], "gpt-file")

    def test_bedrock_provider_and_model_are_selected_from_environment(self):
        with patch.dict(os.environ, {
            "AI_PROVIDER": "bedrock", "BEDROCK_MODEL_ID": "us.example-model-v1:0"
        }, clear=True):
            _, identity, _, payload = runner.prepare(
                self.root, "architecture", self.image
            )
        self.assertEqual(identity["provider"], "bedrock")
        self.assertEqual(identity["model"], "us.example-model-v1:0")
        self.assertEqual(payload["model"], "us.example-model-v1:0")

    def test_bedrock_converse_uses_local_profile_and_image_bytes(self):
        calls = {}

        class FakeClient:
            def converse(self, **request):
                calls["request"] = request
                return {
                    "output": {"message": {"content": [{"text": "# Bedrock document"}]}},
                    "usage": {"inputTokens": 10, "outputTokens": 3},
                    "stopReason": "end_turn",
                    "ResponseMetadata": {"RequestId": "aws-request"},
                }

        class FakeSession:
            def __init__(self, **kwargs):
                calls["session"] = kwargs

            def client(self, service):
                calls["service"] = service
                return FakeClient()

        boto3 = types.ModuleType("boto3")
        boto3.Session = FakeSession
        botocore = types.ModuleType("botocore")
        exceptions = types.ModuleType("botocore.exceptions")
        exceptions.BotoCoreError = type("BotoCoreError", (Exception,), {})
        exceptions.ClientError = type("ClientError", (Exception,), {})
        with patch.dict(sys.modules, {
            "boto3": boto3, "botocore": botocore, "botocore.exceptions": exceptions
        }), patch.dict(os.environ, {
            "AI_PROVIDER": "bedrock", "BEDROCK_MODEL_ID": "us.example-model-v1:0",
            "AWS_PROFILE": "local-sso", "AWS_REGION": "us-east-1",
            "BEDROCK_MAX_TOKENS": "4096",
        }, clear=True):
            _, _, _, payload = runner.prepare(self.root, "architecture", self.image)
            document, metadata = runner.call_bedrock_api(payload)

        self.assertEqual(document, "# Bedrock document")
        self.assertEqual(metadata["response_id"], "aws-request")
        self.assertEqual(calls["session"], {
            "profile_name": "local-sso", "region_name": "us-east-1"
        })
        self.assertEqual(calls["service"], "bedrock-runtime")
        request = calls["request"]
        self.assertEqual(request["modelId"], "us.example-model-v1:0")
        self.assertEqual(request["inferenceConfig"], {"maxTokens": 4096})
        image = request["messages"][0]["content"][1]["image"]
        self.assertEqual(image["format"], "png")
        self.assertEqual(image["source"]["bytes"], self.image.read_bytes())

    def test_invalid_provider_is_rejected(self):
        with patch.dict(os.environ, {"AI_PROVIDER": "unknown"}, clear=True):
            with self.assertRaisesRegex(ValueError, "AI_PROVIDER"):
                runner.prepare(self.root, "architecture", self.image)

    def test_env_language_overrides_prompt_and_enables_full_rtl_word_layout(self):
        with patch.dict(os.environ, {
            "AI_PROVIDER": "openai", "DOCUMENT_LANGUAGE": "IL"
        }, clear=True):
            _, identity, _, payload = runner.prepare(self.root, "architecture", self.image)
        self.assertEqual(identity["language"], "IL")
        self.assertEqual(identity["document_direction"], "rtl")
        self.assertIn("Write the entire document in Hebrew", payload["instructions"])

        output = self.root / "rtl.docx"
        runner.write_word_document("# English identifier\n\n| Name | Value |\n| --- | --- |\n| API | test |", output, rtl=True)
        with zipfile.ZipFile(output) as package:
            xml = package.read("word/document.xml")
        self.assertIn(b"<w:bidi", xml)
        self.assertIn(b"<w:bidiVisual", xml)

    def test_headings_are_centered_and_rtl_body_starts_on_right(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn

        markdown = "\n\n".join("#" * level + " Heading" for level in range(1, 7))
        markdown += ("\n\nBody text.\n\n- Bullet item\n\n1. Numbered item\n\n"
                     "```\nAPI identifier\n```\n\n"
                     "| Name | Value |\n| --- | --- |\n| API | test |")
        for rtl in (False, True):
            with self.subTest(rtl=rtl):
                output = self.root / "alignment.docx"
                runner.write_word_document(markdown, output, rtl=rtl)
                document = Document(output)
                headings = document.paragraphs[:6]
                self.assertTrue(all(p.alignment == WD_ALIGN_PARAGRAPH.CENTER for p in headings))
                if rtl:
                    body = document.paragraphs[6:] + [
                        cell.paragraphs[0] for row in document.tables[0].rows for cell in row.cells]
                    self.assertTrue(all(p.alignment == WD_ALIGN_PARAGRAPH.RIGHT for p in body))
                    for paragraph in headings + body:
                        self.assertIsNotNone(paragraph._p.pPr.find(qn("w:bidi")))

        # Auto-detected RTL text also aligns right without a runtime language override.
        output = self.root / "detected-rtl.docx"
        runner.write_word_document("# כותרת\n\nתיאור בעברית", output)
        document = Document(output)
        self.assertEqual(document.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(document.paragraphs[1].alignment, WD_ALIGN_PARAGRAPH.RIGHT)

    def test_heading_settings_apply_to_all_levels_and_rtl_fonts(self):
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.shared import Pt

        registry_path = self.root / "agents.json"
        registry = json.loads(registry_path.read_text())
        registry["architecture"]["heading_style"] = {
            "color": "#123abc", "font": "Calibri", "size": 19.5}
        registry_path.write_text(json.dumps(registry))
        with patch.dict(os.environ, {"DOCUMENT_LANGUAGE": "HE"}, clear=True):
            def provider(payload):
                return "\n\n".join("#" * level + " Heading" for level in range(1, 7)), {}
            run = runner.process(self.root, "architecture", self.image, provider=provider)
        document = Document(run / "diagram.docx")
        for paragraph in document.paragraphs[:6]:
            self.assertEqual(paragraph.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertIsNotNone(paragraph._p.pPr.find(qn("w:bidi")))
            for word_run in paragraph.runs:
                self.assertEqual(word_run.font.name, "Calibri")
                self.assertEqual(word_run.font.size, Pt(19.5))
                self.assertEqual(str(word_run.font.color.rgb), "123ABC")
                self.assertEqual(word_run._r.rPr.rFonts.get(qn("w:cs")), "Calibri")
                self.assertEqual(word_run._r.rPr.find(qn("w:szCs")).get(qn("w:val")), "39")
        record = json.loads((run / "run.json").read_text())
        self.assertEqual(record["heading_style"]["color"], "123ABC")

    def test_heading_style_change_regenerates_document(self):
        first = self.run_image()
        registry_path = self.root / "agents.json"
        registry = json.loads(registry_path.read_text())
        registry["architecture"]["heading_style"] = {"size": 20}
        registry_path.write_text(json.dumps(registry))
        self.assertNotEqual(first, self.run_image())
        self.assertEqual(len(self.calls), 2)

    def test_invalid_heading_settings_are_rejected_before_provider_call(self):
        for settings in ([], {"color": "red"}, {"font": " "}, {"size": True},
                         {"size": 0}, {"size": 97}, {"size": float("nan")},
                         {"size": "16"}, {"colour": "000000"}, {"bold": "true"},
                         {"italic": 1}, {"underline": None}):
            with self.subTest(settings=settings):
                registry_path = self.root / "agents.json"
                registry = json.loads(registry_path.read_text())
                registry["architecture"]["heading_style"] = settings
                registry_path.write_text(json.dumps(registry))
                with self.assertRaisesRegex(ValueError, "heading_style"):
                    self.run_image()
        self.assertEqual(self.calls, [])
        self.assertFalse((self.root / "output").exists())

    def test_heading_emphasis_options_override_markdown_including_rtl(self):
        from docx import Document

        for enabled in (False, True):
            with self.subTest(enabled=enabled):
                output = self.root / "heading-emphasis.docx"
                runner.write_word_document(
                    "# **Bold** and *italic*\n\n## Second heading", output, rtl=True,
                    heading_style={"bold": enabled, "italic": enabled, "underline": enabled})
                for paragraph in Document(output).paragraphs:
                    for word_run in paragraph.runs:
                        self.assertEqual(word_run.bold, enabled)
                        self.assertEqual(word_run.italic, enabled)
                        self.assertEqual(word_run.underline, enabled)
                        self.assertEqual(word_run.font.cs_bold, enabled)
                        self.assertEqual(word_run.font.cs_italic, enabled)

    def test_enter_listener_is_disabled_without_interactive_stdin(self):
        with patch.object(runner.sys.stdin, "isatty", return_value=False):
            self.assertIsNone(runner.start_enter_listener())

    def test_successful_inbox_image_is_archived_without_overwriting(self):
        inbox = self.root / "input"
        inbox.mkdir()
        first = inbox / "diagram.png"
        first.write_bytes(self.image.read_bytes())
        config = runner.read_config(self.root, "architecture")

        archived = runner.archive_processed_image(self.root, config, first)
        self.assertEqual(archived, self.root / "archive" / "diagram.png")
        self.assertFalse(first.exists())
        self.assertTrue(archived.is_file())

        first.write_bytes(self.image.read_bytes())
        second = runner.archive_processed_image(self.root, config, first)
        self.assertNotEqual(second, archived)
        self.assertTrue(second.is_file())
        self.assertTrue(archived.is_file())


if __name__ == "__main__":
    unittest.main()
