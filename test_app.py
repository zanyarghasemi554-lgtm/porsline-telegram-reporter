import importlib
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch


os.environ["DISABLE_TELEGRAM_POLLING"] = "true"
PROJECT_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_DIR))
app = importlib.import_module("app")


class DynamicFormTests(unittest.TestCase):
    def test_parse_survey_identifier(self):
        self.assertEqual(app.parse_survey_identifier("AbC123"), "AbC123")
        self.assertEqual(app.parse_survey_identifier("12345"), "12345")
        self.assertEqual(
            app.parse_survey_identifier("https://survey.porsline.ir/s/AbC123"),
            "AbC123",
        )
        with self.assertRaises(ValueError):
            app.parse_survey_identifier("not a valid identifier")

    def test_custom_forms_are_added_to_main_menu(self):
        with patch.object(app, "get_custom_forms", return_value=[("/form_x", "AbC123", "فرم آزمایشی")]):
            labels = [button["text"] for row in app.main_menu()["keyboard"] for button in row]
        self.assertIn("➕ ثبت فرم جدید", labels)
        self.assertTrue(any(label.startswith("📋 فرم آزمایشی") for label in labels))

    def test_all_reports_include_custom_forms(self):
        with patch.object(app, "get_custom_forms", return_value=[("/form_x", "AbC123", "فرم آزمایشی")]):
            reports = app.all_report_definitions()
        self.assertIn(("AbC123", "فرم آزمایشی"), reports)
        self.assertEqual(len(reports), len(app.FIVE_REPORTS) + 1)

    def test_registration_name_step_requests_identifier(self):
        with patch.object(app, "set_state") as set_state, patch.object(app, "send_message") as send_message:
            app.process_form_registration("فرم جدید", "7", "8", {"step": "name"})
        self.assertIn('"step": "identifier"', set_state.call_args.args[1])
        self.assertIn("شناسه فرم", send_message.call_args.args[0])

    def test_registration_identifier_saves_form(self):
        with (
            patch.object(app, "all_report_definitions", return_value=[]),
            patch.object(app, "validate_survey_access") as validate,
            patch.object(app, "save_custom_form") as save,
            patch.object(app, "set_state"),
            patch.object(app, "main_menu", return_value={"keyboard": []}),
            patch.object(app, "send_message"),
        ):
            app.process_form_registration(
                "https://survey.porsline.ir/s/AbC123",
                "7",
                "8",
                {"step": "identifier", "report_name": "فرم جدید"},
            )
        validate.assert_called_once_with("AbC123")
        self.assertEqual(save.call_args.args[1:3], ("AbC123", "فرم جدید"))

    def test_custom_full_export_includes_processed_rows(self):
        pending = {"action": "custom_full", "survey_code": "AbC123", "report_name": "فرم جدید"}
        with patch.object(app, "send_message"), patch.object(app, "run_report", return_value={}) as run:
            app.execute_custom_export(pending)
        self.assertTrue(run.call_args.kwargs["include_processed"])
        self.assertEqual(run.call_args.kwargs["selected_reports"], [("AbC123", "فرم جدید")])

    def test_excel_report_is_created_from_template(self):
        filename, stream = app.build_report(
            [{"persian_name": "نام آزمایشی", "english_name": "Test Name", "national_id": "0123456789"}],
            1,
            "فرم آزمایشی",
        )
        self.assertEqual(filename, "فرم آزمایشی 1.xlsx")
        self.assertTrue(stream.read(2).startswith(b"PK"))


if __name__ == "__main__":
    unittest.main()
