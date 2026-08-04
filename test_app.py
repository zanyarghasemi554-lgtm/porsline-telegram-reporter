import importlib
import os
import sys
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch


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

    def test_teachers_are_added_to_main_menu(self):
        with patch.object(app, "get_teachers", return_value=["اسکالپل", "مدرس آزمایشی"]):
            labels = [button["text"] for row in app.main_menu()["keyboard"] for button in row]
        self.assertIn("⚙️ مدیریت ربات", labels)
        self.assertIn("🔍 جست‌وجوی ثبت‌نام‌کنندگان", labels)
        self.assertIn("🩺 بررسی سلامت", labels)
        self.assertIn("👨‍🏫 اسکالپل", labels)
        self.assertIn("👨‍🏫 مدرس آزمایشی", labels)

    def test_nursing_menu_uses_clear_numbered_steps(self):
        labels = [button["text"] for row in app.NURSING_MENU["keyboard"] for button in row]
        self.assertIn("1️⃣ مدارک از نظام رسید", labels)
        self.assertNotIn("2️⃣ مدارک را به پست دادم", labels)
        self.assertNotIn("📥 دریافت دوباره فایل گزارش", labels)
        self.assertIn("📘 راهنمای این بخش", labels)
        help_text = app.nursing_help_text()
        self.assertIn("مرحله", help_text)
        self.assertIn("اصلاح انتخاب اشتباه", help_text)

    def test_custom_form_is_only_in_its_teacher_menu(self):
        def forms(_active_only=True, teacher_name=None):
            if teacher_name == "مدرس الف":
                return [("custom", "AbC123", "فرم آزمایشی", "مدرس الف")]
            return []

        with patch.object(app, "get_form_records", side_effect=forms):
            first_labels = [button["text"] for row in app.teacher_menu("مدرس الف")["keyboard"] for button in row]
            second_labels = [button["text"] for row in app.teacher_menu("مدرس ب")["keyboard"] for button in row]
        self.assertTrue(any(label.startswith("📋 فرم آزمایشی") for label in first_labels))
        self.assertFalse(any(label.startswith("📋 فرم آزمایشی") for label in second_labels))

    def test_all_reports_include_custom_forms(self):
        records = [
            ("builtin", app.INJECTION_SURVEY_CODE, "تزریقات", "اسکالپل"),
            ("custom", "AbC123", "فرم آزمایشی", "مدرس الف"),
        ]
        with patch.object(app, "get_form_records", return_value=records):
            reports = app.all_report_definitions()
        self.assertIn(("AbC123", "فرم آزمایشی"), reports)
        self.assertEqual(len(reports), 2)

    def test_registration_name_step_requests_identifier(self):
        with patch.object(app, "set_state") as set_state, patch.object(app, "send_message") as send_message:
            app.process_form_registration("فرم جدید", "7", "8", {"step": "name"})
        self.assertIn('"step": "identifier"', set_state.call_args.args[1])
        self.assertIn("شناسه فرم", send_message.call_args.args[0])

    def test_registration_identifier_requests_teacher(self):
        with (
            patch.object(app, "all_report_definitions", return_value=[]),
            patch.object(app, "validate_survey_access") as validate,
            patch.object(app, "get_teachers", return_value=["اسکالپل"]),
            patch.object(app, "set_state") as set_state,
            patch.object(app, "send_message") as send_message,
        ):
            app.process_form_registration(
                "https://survey.porsline.ir/s/AbC123",
                "7",
                "8",
                {"step": "identifier", "report_name": "فرم جدید"},
            )
        validate.assert_called_once_with("AbC123")
        self.assertIn('"step": "teacher"', set_state.call_args.args[1])
        self.assertIn("کدام مدرس", send_message.call_args.args[0])

    def test_teacher_selection_saves_form_under_teacher(self):
        registration = {"step": "teacher", "report_name": "فرم جدید", "survey_code": "AbC123"}
        with (
            patch.object(app, "get_teachers", return_value=["اسکالپل"]),
            patch.object(app, "set_state"),
            patch.object(app, "send_message") as send,
        ):
            app.process_form_registration("👨‍🏫 اسکالپل", "7", "8", registration)
        self.assertEqual(registration["step"], "certificate_type")
        self.assertEqual(registration["teacher_name"], "اسکالپل")
        self.assertIs(send.call_args.args[2], app.CERTIFICATE_TYPE_MENU)

    def test_new_teacher_is_created_and_form_is_assigned(self):
        registration = {"step": "new_teacher", "report_name": "فرم جدید", "survey_code": "AbC123"}
        with (
            patch.object(app, "get_teachers", return_value=[]),
            patch.object(app, "available_teacher_colors", return_value=[("blue", "آبی", "BDD7EE")]),
            patch.object(app, "set_state") as set_state,
            patch.object(app, "send_message"),
        ):
            app.process_form_registration("مدرس جدید", "7", "8", registration)
        self.assertIn('"step": "new_teacher_color"', set_state.call_args.args[1])

    def test_new_teacher_color_is_saved_before_form_assignment(self):
        registration = {
            "step": "new_teacher_color", "report_name": "فرم جدید",
            "survey_code": "AbC123", "teacher_name": "مدرس جدید",
        }
        with (
            patch.object(app, "available_teacher_colors", return_value=[("blue", "آبی", "BDD7EE")]),
            patch.object(app, "save_teacher", return_value="مدرس جدید") as save_teacher,
            patch.object(app, "set_state"),
            patch.object(app, "send_message") as send,
        ):
            app.process_form_registration("🎨 آبی", "7", "8", registration)
        save_teacher.assert_called_once_with("مدرس جدید", "7", "blue")
        self.assertEqual(registration["step"], "certificate_type")
        self.assertEqual(registration["teacher_name"], "مدرس جدید")
        self.assertIs(send.call_args.args[2], app.CERTIFICATE_TYPE_MENU)

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

    def test_colored_report_colors_the_whole_student_row(self):
        _filename, stream = app.build_report(
            [{
                "persian_name": "نام آزمایشی", "english_name": "Test Name",
                "national_id": "0123456789", "color_key": "green",
            }],
            1, "فرم آزمایشی", colored=True,
        )
        workbook = app.load_workbook(stream)
        fills = [workbook.active.cell(2, col).fill.fgColor.rgb for col in range(1, 6)]
        self.assertTrue(all(str(value).endswith("C6EFCE") for value in fills))

    def test_caption_contains_first_and_last_person(self):
        rows = [
            {"persian_name": "نفر اول"},
            {"persian_name": "نفر آخر"},
        ]
        caption = app.report_caption(rows, ["فرم تست"], batch_short_id="R-1")
        self.assertIn("از نفر اول تا نفر آخر", caption)
        self.assertIn("R-1", caption)

    def test_combined_new_report_sends_white_and_colored_files(self):
        fake_lock = MagicMock()
        fake_lock.acquire.return_value = True
        first_person = {"persian_name": "اول", "english_name": "First", "national_id": "001"}
        second_person = {"persian_name": "آخر", "english_name": "Last", "national_id": "002"}
        fake_stream = MagicMock()
        with (
            patch.object(app, "RUN_LOCK", fake_lock),
            patch.object(app, "require_settings"),
            patch.object(app, "init_db"),
            patch.object(app, "get_form_records", return_value=[
                ("custom", "A1", "فرم اول", "اسکالپل"),
                ("custom", "A2", "فرم دوم", "خانم ظهیری"),
            ]),
            patch.object(app, "get_teacher_colors", return_value={"اسکالپل": "white", "خانم ظهیری": "green"}),
            patch.object(app, "resolve_surveys", return_value={"A1": 1, "A2": 2}),
            patch.object(app, "collect_new_rows", side_effect=[
                ([first_person], [("A1", "R1")], 1),
                ([second_person], [("A2", "R2")], 1),
            ]),
            patch.object(app, "create_report_batch", return_value={
                "id": 1, "short_id": "R-TEST", "created_at": app.datetime.now(app.timezone.utc),
            }),
            patch.object(app, "build_report", side_effect=[
                ("white.xlsx", fake_stream), ("color.xlsx", fake_stream),
            ]) as build,
            patch.object(app, "send_document") as send,
            patch.object(app, "mark_processed") as mark,
        ):
            result = app.run_report(selected_reports=[("A1", "فرم اول"), ("A2", "فرم دوم")], combine=True)
        self.assertEqual(result["files"], ["white.xlsx", "color.xlsx"])
        self.assertEqual(send.call_count, 2)
        self.assertFalse(build.call_args_list[0].kwargs.get("colored", False))
        self.assertTrue(build.call_args_list[1].kwargs["colored"])
        mark.assert_called_once()

    def test_first_delete_confirmation_only_requests_second_warning(self):
        pending = {
            "kind": "form", "source": "builtin", "survey_code": "A1",
            "report_name": "فرم", "stage": 1, "created_at": app.time.time(),
        }
        with (
            patch.object(app, "get_state", return_value=app.json.dumps(pending)),
            patch.object(app, "set_state") as set_state,
            patch.object(app, "send_message") as send_message,
            patch.object(app, "deactivate_form") as deactivate,
        ):
            app.process_deactivation_callback("deactivate:first", "7", "8")
        deactivate.assert_not_called()
        self.assertIn('"stage": 2', set_state.call_args.args[1])
        self.assertIn("هشدار نهایی", send_message.call_args.args[0])

    def test_second_delete_confirmation_deactivates_form(self):
        pending = {
            "kind": "form", "source": "builtin", "survey_code": "A1",
            "report_name": "فرم", "stage": 2, "created_at": app.time.time(),
        }
        with (
            patch.object(app, "get_state", return_value=app.json.dumps(pending)),
            patch.object(app, "set_state"),
            patch.object(app, "deactivate_form", return_value=True) as deactivate,
            patch.object(app, "main_menu", return_value={"keyboard": []}),
            patch.object(app, "send_message"),
        ):
            app.process_deactivation_callback("deactivate:second", "7", "8")
        deactivate.assert_called_once_with("builtin", "A1")

    def test_form_name_can_be_edited(self):
        editing = {
            "kind": "form", "source": "custom", "survey_code": "A1",
            "report_name": "قدیمی", "teacher_name": "مدرس الف",
        }
        with (
            patch.object(app, "rename_form", return_value="نام جدید") as rename,
            patch.object(app, "set_state"),
            patch.object(app, "teacher_menu", return_value={"keyboard": []}),
            patch.object(app, "send_message"),
        ):
            app.process_edit("نام جدید", "7", "8", editing)
        rename.assert_called_once_with("custom", "A1", "نام جدید")

    def test_teacher_name_can_be_edited(self):
        editing = {"kind": "teacher", "teacher_name": "مدرس قدیمی"}
        with (
            patch.object(app, "rename_teacher", return_value="مدرس جدید") as rename,
            patch.object(app, "set_state"),
            patch.object(app, "teacher_menu", return_value={"keyboard": []}),
            patch.object(app, "send_message"),
        ):
            app.process_edit("مدرس جدید", "7", "8", editing)
        rename.assert_called_once_with("مدرس قدیمی", "مدرس جدید")

    def test_second_confirmation_deactivates_teacher_and_forms(self):
        pending = {
            "kind": "teacher", "teacher_name": "اسکالپل",
            "stage": 2, "created_at": app.time.time(),
        }
        with (
            patch.object(app, "get_state", return_value=app.json.dumps(pending)),
            patch.object(app, "set_state"),
            patch.object(app, "deactivate_teacher", return_value=True) as deactivate,
            patch.object(app, "main_menu", return_value={"keyboard": []}),
            patch.object(app, "send_message"),
        ):
            app.process_deactivation_callback("deactivate:second", "7", "8")
        deactivate.assert_called_once_with("اسکالپل")

    def test_search_matches_name_and_national_id(self):
        person = {"persian_name": "علی رضایی", "english_name": "Ali Rezaei", "national_id": "0012345678"}
        self.assertTrue(app.registration_matches(person, "علی"))
        self.assertTrue(app.registration_matches(person, "ali"))
        self.assertTrue(app.registration_matches(person, "۰۰۱۲۳۴۵۶۷۸"))
        self.assertFalse(app.registration_matches(person, "محمد"))

    def test_search_returns_form_teacher_and_latest_sent_date(self):
        records = [("custom", "A1", "فرم تست", "مدرس تست")]
        headers = ["نام فارسی", "نام انگلیسی", "کد ملی", "response id", "تاریخ ثبت پاسخ"]
        rows = [["علی رضایی", "Ali Rezaei", "0012345678", "R1", "2026-08-01T12:30:00Z"]]
        sent_at = app.datetime(2026, 8, 2, 10, 0, tzinfo=app.timezone.utc)
        fake_lock = MagicMock()
        fake_lock.acquire.return_value = True
        with (
            patch.object(app, "RUN_LOCK", fake_lock),
            patch.object(app, "require_settings"),
            patch.object(app, "init_db"),
            patch.object(app, "get_form_records", return_value=records),
            patch.object(app, "resolve_surveys", return_value={"A1": 1}),
            patch.object(app, "fetch_results", return_value=(headers, rows, 1)),
            patch.object(app, "get_processed_times", return_value={"R1": sent_at}),
            patch.object(app, "get_certificate_history", return_value={}),
            patch.object(app, "get_self_certificate_history", return_value={}),
        ):
            matches, total = app.search_sent_registrations("علی")
        self.assertEqual(total, 1)
        self.assertEqual(matches[0]["form"], "فرم تست")
        self.assertEqual(matches[0]["teacher"], "مدرس تست")
        self.assertEqual(matches[0]["sent_at"], sent_at)
        self.assertEqual(matches[0]["submitted_at"], "2026-08-01T12:30:00Z")

    def test_dates_are_tehran_time_and_jalali(self):
        sent_at = app.datetime(2026, 8, 2, 10, 0, tzinfo=app.timezone.utc)
        self.assertEqual(app.format_sent_at(sent_at), "11-05-1405، ساعت 13:30")
        self.assertEqual(
            app.format_tehran_jalali("2026-08-01T12:30:00Z"),
            "10-05-1405، ساعت 16:00",
        )
        self.assertEqual(
            app.format_tehran_jalali("1405/05/11 17:45"),
            "11-05-1405، ساعت 17:45",
        )

    def test_porsline_separate_date_and_time_are_combined(self):
        mapping = {"تاریخ ثبت پاسخ": "1405/05/11", "زمان ثبت پاسخ": "18:20"}
        self.assertEqual(app.extract_submission_value(mapping), "1405/05/11 18:20")

    def test_restore_teacher_reactivates_all_its_forms(self):
        with (
            patch.object(app, "get_inactive_teachers", return_value=["مدرس تست"]),
            patch.object(app, "restore_teacher", return_value=True) as restore,
            patch.object(app, "send_message"),
        ):
            handled = app.process_restore_selection("♻️ مدرس • مدرس تست", "8")
        self.assertTrue(handled)
        restore.assert_called_once_with("مدرس تست")

    def test_transfer_form_to_selected_teacher(self):
        transfer = {
            "source": "custom", "survey_code": "A1", "report_name": "فرم تست",
            "teacher_name": "مدرس اول",
        }
        with (
            patch.object(app, "get_teachers", return_value=["مدرس اول", "مدرس دوم"]),
            patch.object(app, "move_form", return_value=True) as move,
            patch.object(app, "set_state"),
            patch.object(app, "send_message"),
        ):
            app.process_transfer_target("👨‍🏫 مدرس دوم", "7", "8", transfer)
        move.assert_called_once_with("custom", "A1", "مدرس دوم")

    def test_health_report_checks_three_services(self):
        fake_response = MagicMock()
        fake_response.json.return_value = {"ok": True}
        with (
            patch.object(app, "db") as database,
            patch.object(app, "porsline_get", return_value=[]),
            patch.object(app.requests, "get", return_value=fake_response),
        ):
            database.return_value.__enter__.return_value.cursor.return_value.__enter__.return_value.fetchone.return_value = (1,)
            checks = app.run_health_checks()
        self.assertEqual([name for name, _ok, _elapsed, _detail in checks], ["پایگاه‌داده", "پرس‌لاین", "تلگرام"])
        self.assertTrue(all(ok for _name, ok, _elapsed, _detail in checks))


if __name__ == "__main__":
    unittest.main()
