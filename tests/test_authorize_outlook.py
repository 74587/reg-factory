import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import tools.authorize_outlook as authorize_outlook


class AuthorizeOutlookTests(unittest.TestCase):
    def test_load_accounts_accepts_email_password_and_deduplicates(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "accounts.txt"
            path.write_text(
                "User@Outlook.com----secret\n"
                "user@outlook.com----duplicate\n"
                "# comment\n",
                encoding="utf-8",
            )
            self.assertEqual(
                authorize_outlook.load_accounts(str(path)),
                [("user@outlook.com", "secret")],
            )

    def test_load_accounts_rejects_missing_password(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "accounts.txt"
            path.write_text("user@outlook.com\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "email----password"):
                authorize_outlook.load_accounts(str(path))

    def test_load_accounts_accepts_five_dash_separator(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "accounts.txt"
            path.write_text("user@outlook.com-----secret\n", encoding="utf-8")
            self.assertEqual(authorize_outlook.load_accounts(str(path)), [("user@outlook.com", "secret")])

    @patch("tools.extract_graph_tokens.get_graph_token")
    def test_authorize_one_returns_only_safe_result_fields(self, get_graph_token):
        get_graph_token.return_value = {
            "email": "user@outlook.com",
            "password": "secret",
            "refresh_token": "refresh",
            "client_id": "client",
            "access_token": "access",
        }
        result = authorize_outlook.authorize_one(("user@outlook.com", "secret"), 1)
        self.assertEqual(result, {
            "email": "user@outlook.com",
            "password": "secret",
            "refresh_token": "refresh",
            "client_id": "client",
        })
        get_graph_token.assert_called_once_with("user@outlook.com", "secret", 1)


if __name__ == "__main__":
    unittest.main()
