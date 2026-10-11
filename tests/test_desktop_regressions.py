"""Run with: python3 -m unittest discover -s tests -v."""
import importlib.util
import inspect
import pathlib
import subprocess
import tempfile
import unittest
from unittest.mock import patch

ROOT = pathlib.Path(__file__).resolve().parents[1]


class DesktopRegressionTests(unittest.TestCase):
    def setUp(self):
        spec = importlib.util.spec_from_file_location("schoolbot_test", ROOT / "public/schoolbot-mail.py")
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        self.app = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.app)

    def test_embedded_javascript_parses_after_python_string_decoding(self):
        script = self.app.PAGE.split("<script>", 1)[1].split("</script>", 1)[0]
        with tempfile.TemporaryDirectory() as folder:
            path = pathlib.Path(folder) / "desktop.js"
            path.write_text(script, encoding="utf-8")
            result = subprocess.run(["node", "--check", str(path)], capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_saved_account_survives_reload(self):
        with tempfile.TemporaryDirectory() as folder:
            self.app.CONF_PATH = str(pathlib.Path(folder) / ".mailpilot/config.json")
            self.app.CFG = {}
            saved = {"email": "test@example.invalid", "imap_host": "imap.example.invalid",
                     "pass": "fake-test-password", "days": "2", "ollama_model": "auto"}
            self.app.save_cfg(saved)
            self.app.CFG = {}
            self.app.load_cfg()
            self.assertEqual(self.app.CFG, saved)

    def support_storage(self, folder):
        self.app.SUPPORT_PATH = str(pathlib.Path(folder) / "support.json")

    def test_support_on_first_launch(self):
        with tempfile.TemporaryDirectory() as folder:
            self.support_storage(folder)
            self.app.register_support_launch()
            self.assertTrue(self.app.SUPPORT["show"])

    def test_support_every_ten_further_launches(self):
        with tempfile.TemporaryDirectory() as folder:
            self.support_storage(folder)
            self.app.register_support_launch()
            for cycle in range(2):
                for opening in range(9):
                    self.app.register_support_launch()
                    self.assertFalse(self.app.SUPPORT["show"], (cycle, opening))
                self.app.register_support_launch()
                self.assertTrue(self.app.SUPPORT["show"])

    def test_installed_update_shows_support_and_restarts_interval(self):
        with tempfile.TemporaryDirectory() as folder:
            self.support_storage(folder)
            self.app.register_support_launch()
            self.app.register_support_launch()
            self.assertFalse(self.app.SUPPORT["show"])
            self.app.VERSION = "test-next-version"
            self.app.register_support_launch()
            self.assertTrue(self.app.SUPPORT["show"])
            self.app.register_support_launch()
            self.assertFalse(self.app.SUPPORT["show"])

    def test_support_does_not_change_mail_settings(self):
        with tempfile.TemporaryDirectory() as folder:
            self.support_storage(folder)
            self.app.CFG = {"signature": "Test", "days": "2"}
            self.app.register_support_launch()
            self.assertEqual(self.app.CFG, {"signature": "Test", "days": "2"})

    def test_support_storage_failure_does_not_block_launch(self):
        with tempfile.TemporaryDirectory() as folder:
            self.support_storage(folder)
            with patch.object(self.app, "save_support", side_effect=OSError("read only")):
                self.app.register_support_launch()
            self.assertTrue(self.app.SUPPORT["show"])

    def test_schoolbot_button_opens_the_fixed_site_in_the_browser(self):
        self.assertEqual(self.app.SCHOOLBOT_URL, "https://schoolbot.be")
        script = self.app.PAGE.split("<script>", 1)[1].split("</script>", 1)[0]
        self.assertIn('onclick="schoolbot()"', self.app.PAGE)
        self.assertIn("fetch('/api/schoolbot'", script)
        handler = inspect.getsource(self.app.H.do_POST)
        self.assertRegex(handler, r'"/api/schoolbot"[\s\S]{0,160}webbrowser\.open\(SCHOOLBOT_URL\)')

    def test_microsoft_login_requests_mail_permissions(self):
        scopes = self.app.MS_SCOPE.split()
        for needed in ("offline_access", "Mail.ReadWrite", "Mail.Send"):
            self.assertIn(needed, scopes)
        self.assertNotIn("https://graph.microsoft.com/.default", scopes)

    def test_old_microsoft_session_is_refreshed_with_mail_permissions(self):
        with tempfile.TemporaryDirectory() as folder:
            self.app.CONF_PATH = str(pathlib.Path(folder) / "config.json")
            self.app.CFG = {"ms_access": "old", "ms_refresh": "rt", "ms_expiry": 9e12}
            with patch.object(self.app, "_post", return_value={"access_token": "new", "expires_in": 3600}) as post:
                self.assertEqual(self.app.ms_token(), "new")
            self.assertEqual(post.call_args[0][1]["scope"], self.app.MS_SCOPE)


if __name__ == "__main__":
    unittest.main()