import unittest
import datetime
from music_library import parse_rating, parse_genres, parse_woke, parse_type, MusicLibrary, Record

class TestParsers(unittest.TestCase):
    def test_parse_rating_numeric(self):
        self.assertEqual(parse_rating(8.0), 8.0)
        self.assertEqual(parse_rating(7), 7.0)
        self.assertEqual(parse_rating(9.75), 9.75)
        self.assertIsNone(parse_rating(None))
        self.assertIsNone(parse_rating(""))

    def test_parse_rating_datetime_modern(self):
        # 2025-05-09 -> Day 9, Month 5 -> 9.5
        val = datetime.datetime(2025, 5, 9, 0, 0)
        self.assertEqual(parse_rating(val), 9.5)

        # 2026-05-08 -> Day 8, Month 5 -> 8.5
        val2 = datetime.datetime(2026, 5, 8, 0, 0)
        self.assertEqual(parse_rating(val2), 8.5)

        # 2025-07-09 -> Day 9, Month 7 -> 9.7
        val3 = datetime.datetime(2025, 7, 9, 0, 0)
        self.assertEqual(parse_rating(val3), 9.7)

    def test_parse_rating_datetime_1900(self):
        # 1900-01-07 -> Day 7 -> 7.0
        val = datetime.datetime(1900, 1, 7, 0, 0)
        self.assertEqual(parse_rating(val), 7.0)

        # 1900-01-08 06:00:00 -> Day 8 + 6 hours (6/24 = 0.25) -> 8.25
        val2 = datetime.datetime(1900, 1, 8, 6, 0)
        self.assertEqual(parse_rating(val2), 8.25)

    def test_parse_woke(self):
        self.assertTrue(parse_woke("Yeh"))
        self.assertTrue(parse_woke("Yeh "))
        self.assertTrue(parse_woke("Yah"))
        self.assertTrue(parse_woke("yes"))
        self.assertFalse(parse_woke("Nah"))
        self.assertFalse(parse_woke(""))
        self.assertFalse(parse_woke(None))

    def test_parse_genres(self):
        self.assertEqual(parse_genres("Alternative/Indie"), ["Alternative", "Indie"])
        self.assertEqual(parse_genres("Lounge, Barroque Pop"), ["Lounge", "Barroque Pop"])
        self.assertEqual(parse_genres("Indie Rock, Indie Hiphop"), ["Indie Rock", "Indie Hiphop"])
        self.assertEqual(parse_genres(None), [])
        self.assertEqual(parse_genres(""), [])


class TestMusicLibrary(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.library = MusicLibrary("Albums List.xlsx")

    def test_library_loading(self):
        self.assertGreater(len(self.library.records), 0)
        # Verify first row matches what we saw in the print
        # Row 2: Serotonin Dreams | BoyWithUke | rating: 8.0
        rec = self.library.records[0]
        self.assertEqual(rec.title, "Serotonin Dreams")
        self.assertEqual(rec.creator, "BoyWithUke")
        self.assertEqual(rec.rating, 8.0)

    def test_search_by_title(self):
        results = self.library.search_by_title("Discovery")
        self.assertGreaterEqual(len(results), 1)
        self.assertEqual(results[0].creator, "Daft Punk")
        self.assertEqual(results[0].rating, 9.5)

    def test_search_by_artist(self):
        results = self.library.search_by_artist("Arctic Monkeys")
        self.assertGreaterEqual(len(results), 5)
        for r in results:
            self.assertIn("arctic monkeys", r.creator.lower())

    def test_filter_by_year(self):
        # Single year
        results_2018 = self.library.filter_by_year(2018)
        self.assertTrue(all(r.year == 2018 for r in results_2018))

        # Year range
        results_range = self.library.filter_by_year((2010, 2015))
        self.assertTrue(all(2010 <= r.year <= 2015 for r in results_range))

    def test_filter_by_genre(self):
        results = self.library.filter_by_genre("Indie Rock")
        self.assertGreater(len(results), 0)
        for r in results:
            self.assertTrue(any("indie rock" in g.lower() for g in r.genres))

    def test_filter_by_rating_greater_equal(self):
        results = self.library.filter_by_rating(9.5, op=">=")
        self.assertTrue(all(r.rating >= 9.5 for r in results))

    def test_filter_by_rating_less_equal(self):
        results = self.library.filter_by_rating(5.0, op="<=")
        self.assertTrue(all(r.rating <= 5.0 for r in results))
        self.assertGreater(len(results), 0)

    def test_filter_by_rating_truncate_equal(self):
        results = self.library.filter_by_rating(9.0, op="==", truncate=True)
        self.assertTrue(all(int(r.rating) == 9 for r in results))
        titles = [r.title for r in results]
        self.assertIn("Discovery", titles)

    def test_filter_by_rating_truncate_less_equal(self):
        results = self.library.filter_by_rating(5.0, op="<=", truncate=True)
        self.assertTrue(all(int(r.rating) <= 5 for r in results))
        self.assertGreater(len(results), 0)

    def test_to_dataframe(self):
        df = self.library.to_dataframe()
        self.assertEqual(len(df), len(self.library.records))
        self.assertIn("RECORDS", df.columns)
        self.assertIn("MARK (/10)", df.columns)

if __name__ == "__main__":
    unittest.main()
