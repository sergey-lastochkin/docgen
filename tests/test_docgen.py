import json
import pathlib
import unittest

import docgen

ROOT = pathlib.Path(__file__).resolve().parent.parent


class Words(unittest.TestCase):
    def test_rubles_in_words(self):
        self.assertEqual(docgen.rubles_in_words(40000), 'Сорок тысяч')
        self.assertEqual(docgen.rubles_in_words(21001), 'Двадцать одна тысяча один')
        self.assertEqual(docgen.rubles_in_words(1252000), 'Один миллион двести пятьдесят две тысячи')

    def test_money_full(self):
        self.assertEqual(docgen.money_full(40000), '40 000 (Сорок тысяч) рублей 00 копеек')


class Spec(unittest.TestCase):
    def test_example_is_valid_and_sums_up(self):
        spec = json.loads((ROOT / 'examples' / 'romashka-ozon.json').read_text(encoding='utf-8'))
        self.assertEqual(docgen.check(spec), [])
        self.assertEqual(docgen.total(spec), 40000)

    def test_check_catches_bad_dates(self):
        spec = json.loads((ROOT / 'examples' / 'romashka-ozon.json').read_text(encoding='utf-8'))
        spec['start'], spec['end'] = '2026-10-20', '2026-10-12'
        self.assertIn('start позже end', docgen.check(spec))


if __name__ == '__main__':
    unittest.main()
