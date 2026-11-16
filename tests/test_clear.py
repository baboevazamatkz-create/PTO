import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
import app
from learn import Knowledge


class ClearTests(unittest.TestCase):
    def test_same_file_updates_changed_text(self):
        records = [dict(act_no='377', works=text, file='C:/Acts/377.docx')
                   for text in ('Установка печи.', 'Установка  печи.')]
        k = Knowledge(records)
        self.assertEqual(len(k.recs), 1)
        self.assertFalse(k.upsert({**records[0], 'works': 'Исправленная работа'}))
        self.assertEqual(len(k.recs), 1)
        self.assertEqual(k.recs[0]['works'], 'Исправленная работа')

    def test_clear_scopes_and_cancel_preserve_documents(self):
        for scope in ('acts', 'certs', 'all'):
            for confirm in (False, True):
                with self.subTest(scope=scope, confirm=confirm), tempfile.TemporaryDirectory() as tmp:
                    root = Path(tmp)
                    document = root / 'акт.docx'
                    document.write_bytes(b'original document')
                    settings = dict(base_dir=tmp, cert_dir=tmp, cert_registry=str(root/'registry.xlsx'))
                    fake = SimpleNamespace(settings=settings, know=Knowledge([dict(act_no='1', works='Работа')]),
                        certs=[dict(file='cert.pdf')], _acts_generation=0, _certs_generation=0,
                        room_refresh=lambda: None, update_cert_status=lambda: None,
                        winfo_children=lambda: [], flash_saved=lambda text: None)
                    with patch.multiple(app, SETTINGS_FILE=root/'settings.json',
                                        KNOW_FILE=root/'knowledge.json', CERT_FILE=root/'certs.json'), \
                         patch.object(app.messagebox, 'askyesno', return_value=confirm):
                        app.App.clear_database(fake, scope)
                    self.assertEqual(bool(fake.know.recs), not(confirm and scope in ('acts', 'all')))
                    self.assertEqual(bool(fake.certs), not(confirm and scope in ('certs', 'all')))
                    self.assertEqual(document.read_bytes(), b'original document')
                    if confirm and scope in ('acts', 'all'):
                        self.assertNotIn('base_dir', settings)
                        self.assertEqual(app.load_json(root/'knowledge.json', None), [])
                    if confirm and scope in ('certs', 'all'):
                        self.assertNotIn('cert_dir', settings)
                        self.assertNotIn('cert_registry', settings)
                        self.assertEqual(app.load_json(root/'certs.json', None), [])

    def test_background_results_do_not_restore_cleared_records(self):
        class InlineThread:
            def __init__(self, target, **kwargs):
                self.target = target
            def start(self):
                self.target()

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pending = []
            fake = SimpleNamespace(settings=dict(base_dir=tmp, cert_dir=tmp),
                know=Knowledge([]), certs=[], _acts_generation=0, _certs_generation=0,
                base_info=SimpleNamespace(configure=lambda **kwargs: None),
                room_refresh=lambda: None, update_cert_status=lambda: None,
                winfo_children=lambda: [], flash_saved=lambda text: None,
                after=lambda delay, callback: pending.append(callback))
            with patch.multiple(app, SETTINGS_FILE=root/'settings.json',
                                KNOW_FILE=root/'knowledge.json', CERT_FILE=root/'certs.json'), \
                 patch.object(app.threading, 'Thread', InlineThread), \
                 patch.object(app, 'import_folder', return_value=[dict(act_no='1', works='Работа')]), \
                 patch.object(app, 'scan_folder', return_value=[dict(file=str(root/'cert.pdf'))]), \
                 patch.object(app.messagebox, 'askyesno', return_value=True):
                app.App.scan_base(fake)
                app.App.scan_certs(fake)
                self.assertEqual(len(pending), 2)
                app.App.clear_database(fake, 'all')
                for callback in pending:
                    callback()
                self.assertEqual(fake.know.recs, [])
                self.assertEqual(fake.certs, [])
                self.assertEqual(app.load_json(root/'knowledge.json', None), [])
                self.assertEqual(app.load_json(root/'certs.json', None), [])
