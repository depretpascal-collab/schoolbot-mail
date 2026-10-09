"""Run with: python3 -m unittest discover -s tests -v."""
import importlib.util
import pathlib
import subprocess
import tempfile
import unittest

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


if __name__ == "__main__":
    unittest.main()