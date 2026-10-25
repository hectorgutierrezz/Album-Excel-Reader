import unittest
from music_library import MusicLibrary, Record, parse_rating, parse_artists, parse_genres, parse_rating_criteria

class TestMusicLibrary(unittest.TestCase):
    def test_parse_artists_single(self):
        self.assertEqual(parse_artists("Mac Miller"), ["Mac Miller"])
        self.assertEqual(parse_artists("  Daft Punk  "), ["Daft Punk"])
        # Commas in artist name should not be split when using semicolon separator
        self.assertEqual(parse_artists("Tyler, The Creator"), ["Tyler, The Creator"])
        self.assertEqual(parse_artists("Black Country, New Road"), ["Black Country, New Road"])

    def test_parse_artists_multiple_semicolon(self):
        self.assertEqual(parse_artists("Bruno Mars; Anderson .Paak"), ["Bruno Mars", "Anderson .Paak"])
        self.assertEqual(parse_artists("Clipse; Pusha T; Malice"), ["Clipse", "Pusha T", "Malice"])
        self.assertEqual(parse_artists("Artist 1 ;  Artist 2  ; Artist 3"), ["Artist 1", "Artist 2", "Artist 3"])

    def test_parse_artists_empty(self):
        self.assertEqual(parse_artists(""), [])
        self.assertEqual(parse_artists(None), [])

    def test_parse_rating_numbers_and_commas(self):
        self.assertEqual(parse_rating(9), 9.0)
        self.assertEqual(parse_rating(8.5), 8.5)
        self.assertEqual(parse_rating("8.5"), 8.5)
        self.assertEqual(parse_rating("8,5"), 8.5)
        self.assertEqual(parse_rating("9,75"), 9.75)
        self.assertEqual(parse_rating(" 10 "), 10.0)

    def test_parse_rating_criteria(self):
        crit1 = parse_rating_criteria(">=8,5")
        self.assertEqual(crit1["val"], 8.5)
        self.assertEqual(crit1["op"], ">=")

        crit2 = parse_rating_criteria("<=7.5")
        self.assertEqual(crit2["val"], 7.5)
        self.assertEqual(crit2["op"], "<=")

        crit3 = parse_rating_criteria("8")
        self.assertEqual(crit3["val"], 8.0)
        self.assertEqual(crit3["op"], "==")
        self.assertTrue(crit3["truncate"])

    def test_multi_artist_search(self):
        rec1 = Record(title="An Evening With Silk Sonic", creator="Bruno Mars; Anderson .Paak", record_type="Album",
                      genres=["R&B", "Soul"], language="English", year=2021, rating=9.0, woke=False)
        rec2 = Record(title="24K Magic", creator="Bruno Mars", record_type="Album",
                      genres=["Pop", "Funk"], language="English", year=2016, rating=8.5, woke=False)
        rec3 = Record(title="Swimming", creator="Mac Miller", record_type="Album",
                      genres=["Hip Hop"], language="English", year=2018, rating=8.0, woke=False)
        rec4 = Record(title="IGOR", creator="Tyler, The Creator", record_type="Album",
                      genres=["Hip Hop"], language="English", year=2019, rating=9.5, woke=False)

        records = [rec1, rec2, rec3, rec4]
        lib = MusicLibrary.__new__(MusicLibrary)
        lib.records = records

        # Search Bruno Mars (should find rec1 and rec2)
        results = lib.search_by_artist("Bruno Mars")
        self.assertEqual(len(results), 2)
        self.assertIn(rec1, results)
        self.assertIn(rec2, results)

        # Search Anderson .Paak (should find rec1)
        results_paak = lib.search_by_artist("Anderson .Paak")
        self.assertEqual(len(results_paak), 1)
        self.assertIn(rec1, results_paak)

        # Exact search for Anderson .Paak
        results_exact = lib.search_by_artist("Anderson .Paak", exact=True)
        self.assertEqual(len(results_exact), 1)
        self.assertIn(rec1, results_exact)

        # Substring search for paak
        results_sub = lib.search_by_artist("paak")
        self.assertEqual(len(results_sub), 1)
        self.assertIn(rec1, results_sub)

        # Search Mac Miller (solo artist)
        results_mac = lib.search_by_artist("Mac Miller")
        self.assertEqual(len(results_mac), 1)
        self.assertIn(rec3, results_mac)

        # Search Tyler, The Creator
        results_tyler = lib.search_by_artist("Tyler, The Creator", exact=True)
        self.assertEqual(len(results_tyler), 1)
        self.assertIn(rec4, results_tyler)
        self.assertEqual(rec4.artists, ["Tyler, The Creator"])

    def test_load_real_file(self):
        lib = MusicLibrary("Albums List.xlsx")
        self.assertGreater(len(lib.records), 500)
        
        # Verify Silk Sonic has both artists
        silk = [r for r in lib.records if "Silk Sonic" in r.title]
        self.assertTrue(len(silk) > 0)
        self.assertIn("Bruno Mars", silk[0].artists)
        self.assertIn("Anderson .Paak", silk[0].artists)

        # Verify Tyler, The Creator is single artist
        igor = [r for r in lib.records if r.title == "IGOR"]
        self.assertTrue(len(igor) > 0)
        self.assertEqual(igor[0].artists, ["Tyler, The Creator"])

        # Verify year and rating are correct on first record
        first = lib.records[0]
        self.assertGreater(first.year, 1950)
        self.assertGreater(first.rating, 0.0)

if __name__ == '__main__':
    unittest.main()
