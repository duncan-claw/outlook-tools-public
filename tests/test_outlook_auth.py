import test_support
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from outlook import outlook_personal_graph as auth

class AuthTests(unittest.TestCase):
    def test_missing_app_id_fails_closed_before_login(self):
        with patch.object(auth.config, 'GRAPH_CLIENT_ID', ''):
            with self.assertRaisesRegex(RuntimeError, 'OUTLOOK_PERSONAL_GRAPH_CLIENT_ID_REQUIRED'):
                auth.device_login()

    def test_atomic_private_save(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            with patch.object(auth, 'STATE_DIR', root), patch.object(auth, 'TOKEN_PATH', root/'token.json'):
                auth.save_token({'access_token':'synthetic', 'expires_in':3600})
                self.assertEqual(json.loads((root/'token.json').read_text())['access_token'], 'synthetic')
                self.assertEqual((root/'token.json').stat().st_mode & 0o777, 0o600)
                self.assertEqual(list(root.glob('.token-*')), [])
    def test_refresh_preserves_scopes(self):
        with patch.object(auth, 'post_form', return_value={'access_token':'synthetic'}) as post, patch.object(auth, 'save_token'):
            auth._refresh_token_locked({'refresh_token':'synthetic', 'scope':'Custom.Scope Contacts.Read'})
            scopes = post.call_args.args[1]['scope'].split()
            self.assertIn('Custom.Scope', scopes)
            self.assertIn('Contacts.Read', scopes)
            self.assertIn('Mail.Send', scopes)
    def test_fresh_token_rechecked_under_lock(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(auth, 'STATE_DIR', Path(directory)), patch.object(auth, 'load_token', return_value={'access_token':'new','expires_at':9999999999}), patch.object(auth, '_refresh_token_locked') as refresh:
            self.assertEqual(auth.refresh_token({'refresh_token':'old'})['access_token'], 'new')
            refresh.assert_not_called()
