import tempfile
import unittest
from pathlib import Path

from docx import Document
from openpyxl import Workbook
from app import DEFAULT, table_sort_key
from certs import load_registry, make_entry, merge_certificates, scan_folder
from generator import generate
from learn import Knowledge, parse_act


class Regressions(unittest.TestCase):
    def test_numeric_and_date_sort(self):
        self.assertEqual(sorted(['Акт 10', 'Акт 2', 'Акт 1'], key=table_sort_key),
                         ['Акт 1', 'Акт 2', 'Акт 10'])
        self.assertEqual(sorted(['01.02.2026', '31.01.2026'], key=table_sort_key),
                         ['31.01.2026', '01.02.2026'])

    def test_excel_and_folder_merge(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / 'docs').mkdir()
            cert = root / 'docs' / 'certificate.jpg'
            cert.touch()
            reg = root / 'registry.xlsx'
            wb = Workbook()
            ws = wb.active
            ws.append(['Номер документа', 'Наименование материала', 'Файл'])
            for _ in range(2):
                ws.append(['KZ.1234567', 'Материал из Excel', 'docs\\certificate.jpg'])
            wb.save(reg)
            rows = merge_certificates(load_registry(reg), scan_folder(root), reg)
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]['material'], 'Материал из Excel')
            self.assertTrue(Path(rows[0]['file']).exists())

    def test_both_document_types_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            knowledge = Knowledge([])
            for kind in ('АОСР', 'АППР'):
                data = {**DEFAULT, 'act_no': '12', 'act_type': kind,
                        'works': 'Устройство стяжки. Помещение № 38, 2 этаж.',
                        'materials': ['Смесь'], 'start': '01.11.2026', 'end': '02.11.2026'}
                path = knowledge.file_path(data, tmp)
                path.parent.mkdir(parents=True, exist_ok=True)
                generate(data, path)
                text = '\n'.join(p.text for p in Document(path).paragraphs)
                self.assertIn('АКТ № 12', text)
                self.assertIn('промежуточной приемки работ' if kind == 'АППР'
                              else 'освидетельствования скрытых работ', text)
                if kind == 'АППР':
                    self.assertTrue(path.name.startswith('АППР №12 п.38'))
                    self.assertIn('стяжки', path.name)
                record = parse_act(path)
                self.assertIsNotNone(record)
                self.assertEqual(record['act_no'], '12')
                knowledge.upsert(record)
            self.assertEqual(len(knowledge.recs), 2)
            for r in list(knowledge.recs):
                knowledge.upsert(r)
            self.assertEqual(len(knowledge.recs), 2)

    def test_import_adds_room_to_choices(self):
        knowledge = Knowledge([])
        before = knowledge.rooms()
        knowledge.upsert({'act_no': '1', 'works': 'Устройство стяжки. Помещение № 99, 2 этаж.',
                          'docs': '', 'file': '/base/2 этаж/99/акт.docx'})
        self.assertFalse(before)
        self.assertIn('99.2', knowledge.rooms())


if __name__ == '__main__':
    unittest.main()
