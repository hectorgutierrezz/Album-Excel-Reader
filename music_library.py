import datetime
import os
import re
import sys
from dataclasses import dataclass
from typing import List, Optional, Tuple, Union
import openpyxl
import pandas as pd

@dataclass
class Record:
    """Represents a clean, normalized album/EP/record entry from the music library."""
    title: str
    creator: str
    record_type: str
    genres: List[str]
    language: str
    year: int
    rating: float
    woke: bool
    original_rating: Optional[str] = None
    artists: Optional[List[str]] = None

    def __post_init__(self):
        if self.artists is None:
            self.artists = parse_artists(self.creator)

def parse_artists(val) -> List[str]:
    """Splits creator string by semicolons (;) and normalizes them into a list of individual artist strings."""
    if not val:
        return []
    if isinstance(val, (list, tuple)):
        return [str(p).strip() for p in val if str(p).strip()]
    val_str = str(val)
    parts = val_str.split(';')
    return [p.strip() for p in parts if p.strip()]


def parse_rating(val) -> Optional[float]:
    """
    Parses various ratings formats out of 10, correcting Excel auto-conversion bugs.
    
    Excel auto-converts some fractional values (like 9.5 represented as 9/5 or 9-5) 
    into dates. 
    - For current/recent year datetimes (e.g. 2025-05-09):
        The day is the integer rating (9), and the month is the decimal part (5 tenths -> .5).
        Result: 9.5
    - For 1900 year datetimes (Excel raw serial dates):
        The day is the integer rating, and time (hour/minute) is the decimal.
        Result: day + hour/24.0 + minute/1440.0
    - For strings: supports both dot (9.5) and comma (9,5) decimal notations.
    """
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, datetime.datetime) or hasattr(val, 'year'):
        if val.year == 1900:
            rating = val.day + (val.hour / 24.0) + (val.minute / 1440.0)
            return round(rating, 3)
        else:
            rating = val.day + (val.month / 10.0)
            return round(rating, 3)
    if isinstance(val, str):
        val_clean = val.strip().replace(',', '.')
        try:
            return float(val_clean)
        except ValueError:
            pass
    return None

def parse_woke(val) -> bool:
    """Normalizes the WOKE? column values to booleans."""
    if not val:
        return False
    val_str = str(val).strip().lower()
    if val_str in ('yeh', 'yah', 'yes', 'true', '1', 'y'):
        return True
    return False

def parse_genres(val) -> List[str]:
    """Splits genre string by commas or slashes and normalizes them into a list of strings."""
    if not val:
        return []
    val_str = str(val)
    parts = re.split(r'[,/]', val_str)
    return [p.strip() for p in parts if p.strip()]

def parse_type(val) -> str:
    """Strips whitespace and normalizes record type."""
    if not val:
        return "Unknown"
    return str(val).strip()

def parse_rating_criteria(criteria: Union[str, float, int, dict]) -> dict:
    """
    Parses a rating criteria (e.g., '>=8.5', '<=6', '8', 9.5) into a standard dictionary:
    {'val': float, 'op': str, 'truncate': bool}
    Supports comma decimal input (e.g., '>=8,5').
    """
    if isinstance(criteria, dict):
        return {
            'val': float(str(criteria.get('val', 0.0)).replace(',', '.')),
            'op': str(criteria.get('op', '>=')),
            'truncate': bool(criteria.get('truncate', False))
        }
    if isinstance(criteria, (int, float)):
        val = float(criteria)
        if val.is_integer():
            return {'val': val, 'op': '==', 'truncate': True}
        else:
            return {'val': val, 'op': '>=', 'truncate': False}

    criteria_str = str(criteria).strip().replace(',', '.')
    if not criteria_str:
        return {'val': 0.0, 'op': '>=', 'truncate': False}

    if criteria_str.startswith(">="):
        try:
            val = float(criteria_str[2:].strip())
            return {'val': val, 'op': '>=', 'truncate': False}
        except ValueError:
            pass
    elif criteria_str.startswith("<="):
        try:
            val = float(criteria_str[2:].strip())
            return {'val': val, 'op': '<=', 'truncate': False}
        except ValueError:
            pass
    elif criteria_str.startswith("=="):
        try:
            val = float(criteria_str[2:].strip())
            return {'val': val, 'op': '==', 'truncate': False}
        except ValueError:
            pass
    else:
        try:
            val = float(criteria_str)
            if val.is_integer():
                return {'val': val, 'op': '==', 'truncate': True}
            else:
                return {'val': val, 'op': '>=', 'truncate': False}
        except ValueError:
            pass

    return {'val': 0.0, 'op': '>=', 'truncate': False}

class MusicLibrary:
    """Backend engine for loading, parsing, and filtering the Excel music sheet."""
    def __init__(self, excel_path: str = "Albums List.xlsx"):
        # Resolve absolute path relative to executable/script directory if path is relative
        if not os.path.isabs(excel_path):
            if getattr(sys, 'frozen', False):
                # Running as a compiled PyInstaller executable
                base_dir = os.path.dirname(sys.executable)
            else:
                # Running as normal script
                base_dir = os.path.dirname(os.path.abspath(__file__))
            self.excel_path = os.path.join(base_dir, excel_path)
        else:
            self.excel_path = excel_path

        self.records: List[Record] = []
        self.load()

    def load(self):
        """Loads and cleans records from the Excel spreadsheet using openpyxl."""
        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        sheet = wb.active
        
        # Read header and extract column mappings
        header = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]
        
        col_map = {
            "RECORDS": -1,
            "CREATOR": -1,
            "TYPE": -1,
            "GENRES": -1,
            "LANGUAGE": -1,
            "YEAR": -1,
            "MARK (/10)": -1,
            "WOKE?": -1
        }
        for idx, col_name in enumerate(header):
            cleaned = col_name.upper().strip()
            if not cleaned:
                continue
            if cleaned == "RECORDS" and col_map["RECORDS"] == -1:
                col_map["RECORDS"] = idx
            elif cleaned in ("CREATOR", "ARTIST", "ARTISTS", "ARTISTA", "ARTISTAS") and col_map["CREATOR"] == -1:
                col_map["CREATOR"] = idx
            elif cleaned == "TYPE" and col_map["TYPE"] == -1:
                col_map["TYPE"] = idx
            elif cleaned in ("GENRES", "GENRE") and col_map["GENRES"] == -1:
                col_map["GENRES"] = idx
            elif cleaned in ("LANGUAGE", "LANG") and col_map["LANGUAGE"] == -1:
                col_map["LANGUAGE"] = idx
            elif cleaned == "YEAR" and col_map["YEAR"] == -1:
                col_map["YEAR"] = idx
            elif any(k in cleaned for k in ("MARK", "RATING", "SCORE", "NOTA")) and col_map["MARK (/10)"] == -1:
                col_map["MARK (/10)"] = idx
            elif "WOKE" in cleaned and col_map["WOKE?"] == -1:
                col_map["WOKE?"] = idx
            elif col_name in col_map and col_map[col_name] == -1:
                col_map[col_name] = idx

        # Fallback to standard index expectations if columns aren't found by name
        idx_records = col_map["RECORDS"] if col_map["RECORDS"] != -1 else 0
        idx_creator = col_map["CREATOR"] if col_map["CREATOR"] != -1 else 1
        idx_type = col_map["TYPE"] if col_map["TYPE"] != -1 else 2
        idx_genres = col_map["GENRES"] if col_map["GENRES"] != -1 else 3
        idx_lang = col_map["LANGUAGE"] if col_map["LANGUAGE"] != -1 else 4
        idx_year = col_map["YEAR"] if col_map["YEAR"] != -1 else 5
        idx_mark = col_map["MARK (/10)"] if col_map["MARK (/10)"] != -1 else 6
        idx_woke = col_map["WOKE?"] if col_map["WOKE?"] != -1 else 7

        records = []
        for r_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
            if r_idx == 1:
                # Skip header row
                continue
            
            title_val = row[idx_records] if idx_records < len(row) else None
            creator_val = row[idx_creator] if idx_creator < len(row) else None
            
            # Skip empty rows
            if title_val is None and creator_val is None:
                continue
                
            title = str(title_val).strip() if title_val is not None else ""
            creator = str(creator_val).strip() if creator_val is not None else ""
            
            rec_type = parse_type(row[idx_type] if idx_type < len(row) else "")
            genres = parse_genres(row[idx_genres] if idx_genres < len(row) else "")
            lang = str(row[idx_lang]).strip() if idx_lang < len(row) and row[idx_lang] is not None else ""
            
            # Parse year safely
            raw_year = row[idx_year] if idx_year < len(row) else None
            try:
                year = int(float(raw_year)) if raw_year is not None else 0
            except (ValueError, TypeError):
                year = 0
                
            # Parse rating cleanly
            raw_rating = row[idx_mark] if idx_mark < len(row) else None
            rating = parse_rating(raw_rating)
            if rating is None:
                rating = 0.0
                
            woke = parse_woke(row[idx_woke] if idx_woke < len(row) else "")
            
            record = Record(
                title=title,
                creator=creator,
                record_type=rec_type,
                genres=genres,
                language=lang,
                year=year,
                rating=rating,
                woke=woke,
                original_rating=str(raw_rating) if raw_rating is not None else None
            )
            records.append(record)
            
        self.records = records

    # --- Search and Filtering APIs ---

    def search_by_title(self, title: str, exact: bool = False, records_list: Optional[List[Record]] = None) -> List[Record]:
        """Searches records by title (case-insensitive, substring by default)."""
        target = records_list if records_list is not None else self.records
        query = title.strip().lower()
        if not query:
            return target
        if exact:
            return [r for r in target if r.title.lower() == query]
        return [r for r in target if query in r.title.lower()]

    def search_by_artist(self, artist: str, exact: bool = False, records_list: Optional[List[Record]] = None) -> List[Record]:
        """Searches records by creator/artist/band (case-insensitive, substring by default).
        Supports multi-artist fields separated by commas."""
        target = records_list if records_list is not None else self.records
        query = artist.strip().lower()
        if not query:
            return target
        if exact:
            return [r for r in target if any(a.lower() == query for a in r.artists)]
        return [r for r in target if any(query in a.lower() for a in r.artists)]

    def filter_by_year(self, year_range: Union[int, Tuple[int, int]], records_list: Optional[List[Record]] = None) -> List[Record]:
        """Filters records by a single year or a range of years (inclusive)."""
        target = records_list if records_list is not None else self.records
        if isinstance(year_range, int):
            return [r for r in target if r.year == year_range]
        elif isinstance(year_range, tuple) and len(year_range) == 2:
            start, end = year_range
            return [r for r in target if start <= r.year <= end]
        return target

    def filter_by_genre(self, genre: str, records_list: Optional[List[Record]] = None) -> List[Record]:
        """Filters records where any genre matches the query (case-insensitive)."""
        target = records_list if records_list is not None else self.records
        query = genre.strip().lower()
        if not query:
            return target
        return [r for r in target if any(query in g.lower() for g in r.genres)]

    def filter_by_rating(self, rating: float, op: str = ">=", truncate: bool = False, records_list: Optional[List[Record]] = None) -> List[Record]:
        """
        Filters records by rating. 
        op can be ">=", "<=", or "==".
        If truncate is True, truncates the database rating to match the integer rating.
        """
        target = records_list if records_list is not None else self.records
        if truncate:
            val_limit = int(rating)
            if op == "<=":
                return [r for r in target if int(r.rating) <= val_limit]
            elif op == "==":
                return [r for r in target if int(r.rating) == val_limit]
            else:
                return [r for r in target if int(r.rating) >= val_limit]

        if op == "<=":
            return [r for r in target if r.rating <= rating]
        elif op == "==":
            return [r for r in target if r.rating == rating]
        else:
            return [r for r in target if r.rating >= rating]

    def filter_by_language(self, language: str, records_list: Optional[List[Record]] = None) -> List[Record]:
        """Filters records by main language (case-insensitive)."""
        target = records_list if records_list is not None else self.records
        query = language.strip().lower()
        if not query:
            return target
        return [r for r in target if r.language.lower() == query]

    def filter_by_type(self, type_name: str, records_list: Optional[List[Record]] = None) -> List[Record]:
        """Filters records by type (case-insensitive, e.g., 'Album', 'EP')."""
        target = records_list if records_list is not None else self.records
        query = type_name.strip().lower()
        if not query:
            return target
        return [r for r in target if r.record_type.lower() == query]

    def apply_filters(self, filters: dict, records_list: Optional[List[Record]] = None) -> List[Record]:
        """
        Applies multiple filters at once.
        `filters` is a dictionary where keys map to filter parameters:
        - 'artist': str (artist/band name)
        - 'artist_exact': bool (default False)
        - 'title': str (album title)
        - 'title_exact': bool (default False)
        - 'year': int or Tuple[int, int]
        - 'genre': str
        - 'rating': float, int, str, or dict/tuple criteria.
        - 'language': str
        - 'type': str
        """
        current_records = records_list if records_list is not None else self.records
        
        if not filters:
            return current_records

        # 1. Artist
        if 'artist' in filters and filters['artist']:
            exact = filters.get('artist_exact', False)
            current_records = self.search_by_artist(filters['artist'], exact=exact, records_list=current_records)
            
        # 2. Title
        if 'title' in filters and filters['title']:
            exact = filters.get('title_exact', False)
            current_records = self.search_by_title(filters['title'], exact=exact, records_list=current_records)
            
        # 3. Year
        if 'year' in filters and filters['year'] is not None and filters['year'] != '':
            current_records = self.filter_by_year(filters['year'], records_list=current_records)
            
        # 4. Genre
        if 'genre' in filters and filters['genre']:
            current_records = self.filter_by_genre(filters['genre'], records_list=current_records)
            
        # 5. Rating
        if 'rating' in filters and filters['rating'] is not None and filters['rating'] != '':
            rating_crit = parse_rating_criteria(filters['rating'])
            current_records = self.filter_by_rating(
                rating=rating_crit['val'],
                op=rating_crit['op'],
                truncate=rating_crit['truncate'],
                records_list=current_records
            )
            
        # 6. Language
        if 'language' in filters and filters['language']:
            current_records = self.filter_by_language(filters['language'], records_list=current_records)
            
        # 7. Type
        if 'type' in filters and filters['type']:
            current_records = self.filter_by_type(filters['type'], records_list=current_records)
            
        return current_records

    def to_dataframe(self, records_list: Optional[List[Record]] = None) -> pd.DataFrame:
        """Helper to convert a list of Records back to a clean pandas DataFrame for presentation."""
        target = records_list if records_list is not None else self.records
        data = []
        for r in target:
            data.append({
                "RECORDS": r.title,
                "CREATOR": r.creator,
                "TYPE": r.record_type,
                "GENRES": ", ".join(r.genres),
                "LANGUAGE": r.language,
                "YEAR": r.year,
                "MARK (/10)": r.rating,
                "WOKE?": "Yeh" if r.woke else "Nah"
            })
        # Specify index as range to mimic the original dataframe
        return pd.DataFrame(data)

    def get_unique_artists(self) -> List[str]:
        """Returns sorted list of all unique individual artists."""
        artists = set()
        for r in self.records:
            for a in r.artists:
                if a:
                    artists.add(a)
        return sorted(list(artists), key=lambda s: s.lower())

    def get_unique_genres(self) -> List[str]:
        """Returns sorted list of all unique genres."""
        genres = set()
        for r in self.records:
            for g in r.genres:
                if g:
                    genres.add(g)
        return sorted(list(genres), key=lambda s: s.lower())

    def get_unique_languages(self) -> List[str]:
        """Returns sorted list of all unique languages."""
        languages = {r.language for r in self.records if r.language}
        return sorted(list(languages), key=lambda s: s.lower())

    def get_unique_types(self) -> List[str]:
        """Returns sorted list of all unique record types."""
        types = {r.record_type for r in self.records if r.record_type}
        return sorted(list(types), key=lambda s: s.lower())

