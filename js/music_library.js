/**
 * MusicLibraryJS - Client-side engine for reading, parsing, and filtering
 * album excel spreadsheets in browser environments (GitHub Pages).
 * Replicates the exact logic of music_library.py with multi-artist support.
 */

class Record {
  constructor({
    title = '',
    creator = '',
    recordType = 'Unknown',
    genres = [],
    language = '',
    year = 0,
    rating = 0.0,
    originalRating = null,
    artists = null
  }) {
    this.title = title;
    this.creator = creator;
    this.recordType = recordType;
    this.genres = genres;
    this.language = language;
    this.year = year;
    this.rating = rating;
    this.originalRating = originalRating;
    this.artists = artists || MusicLibraryJS.parseArtists(creator);
  }
}

class MusicLibraryJS {
  constructor() {
    this.records = [];
    this.fileName = '';
    this.minYear = 1900;
    this.maxYear = 2026;
  }

  static parseArtists(val) {
    if (!val) return [];
    if (Array.isArray(val)) return val.map(s => String(s).trim()).filter(Boolean);
    const parts = String(val).split(';');
    return parts.map(p => p.trim()).filter(Boolean);
  }

  static parseGenres(val) {
    if (!val) return [];
    const parts = String(val).split(/[,/]/);
    return parts.map(p => p.trim()).filter(Boolean);
  }

  static parseType(val) {
    if (!val) return 'Unknown';
    const trimmed = String(val).trim();
    return trimmed || 'Unknown';
  }

  static parseRating(val) {
    if (val === null || val === undefined || val === '') return 0.0;
    if (typeof val === 'number') return isNaN(val) ? 0.0 : Math.round(val * 1000) / 1000;
    
    // Handle JavaScript Date objects (from Excel date conversions)
    if (val instanceof Date) {
      const year = val.getFullYear();
      if (year === 1900) {
        const rating = val.getDate() + (val.getHours() / 24.0) + (val.getMinutes() / 1440.0);
        return Math.round(rating * 1000) / 1000;
      } else {
        const rating = val.getDate() + ((val.getMonth() + 1) / 10.0);
        return Math.round(rating * 1000) / 1000;
      }
    }

    if (typeof val === 'string') {
      const valClean = val.trim().replace(',', '.');
      const parsed = parseFloat(valClean);
      return isNaN(parsed) ? 0.0 : Math.round(parsed * 1000) / 1000;
    }

    return 0.0;
  }

  static parseRatingCriteria(criteria) {
    if (typeof criteria === 'object' && criteria !== null) {
      return {
        val: parseFloat(String(criteria.val || 0).replace(',', '.')),
        op: String(criteria.op || '>='),
        truncate: Boolean(criteria.truncate)
      };
    }
    if (typeof criteria === 'number') {
      return {
        val: criteria,
        op: Number.isInteger(criteria) ? '==' : '>=',
        truncate: Number.isInteger(criteria)
      };
    }

    const str = String(criteria || '').trim().replace(',', '.');
    if (!str) return { val: 0.0, op: '>=', truncate: false };

    if (str.startsWith('>=')) {
      const val = parseFloat(str.slice(2).trim());
      return { val: isNaN(val) ? 0.0 : val, op: '>=', truncate: false };
    }
    if (str.startsWith('<=')) {
      const val = parseFloat(str.slice(2).trim());
      return { val: isNaN(val) ? 0.0 : val, op: '<=', truncate: false };
    }
    if (str.startsWith('==')) {
      const val = parseFloat(str.slice(2).trim());
      return { val: isNaN(val) ? 0.0 : val, op: '==', truncate: false };
    }

    const val = parseFloat(str);
    if (!isNaN(val)) {
      const isInt = Number.isInteger(val);
      return { val, op: isInt ? '==' : '>=', truncate: isInt };
    }

    return { val: 0.0, op: '>=', truncate: false };
  }

  loadFromArrayBuffer(arrayBuffer, fileName = 'Albums List.xlsx') {
    if (!window.XLSX) {
      throw new Error('SheetJS (XLSX) library is not loaded');
    }

    const workbook = window.XLSX.read(arrayBuffer, {
      type: 'array',
      cellDates: true,
      raw: true
    });

    if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
      throw new Error('No sheets found in Excel file');
    }

    const firstSheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[firstSheetName];
    const rows = window.XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      defval: null
    });

    if (!rows || rows.length < 2) {
      this.records = [];
      this.fileName = fileName;
      return;
    }

    const header = (rows[0] || []).map(cell => (cell !== null && cell !== undefined ? String(cell).trim() : ''));

    // Robust column mapping with duplicate header protection
    const colMap = {
      records: -1,
      creator: -1,
      type: -1,
      genres: -1,
      language: -1,
      year: -1,
      mark: -1
    };

    header.forEach((colName, idx) => {
      const upper = colName.toUpperCase().trim();
      if (!upper) return;

      if (upper === 'RECORDS' && colMap.records === -1) {
        colMap.records = idx;
      } else if (['CREATOR', 'ARTIST', 'ARTISTS', 'ARTISTA', 'ARTISTAS'].includes(upper) && colMap.creator === -1) {
        colMap.creator = idx;
      } else if (upper === 'TYPE' && colMap.type === -1) {
        colMap.type = idx;
      } else if (['GENRES', 'GENRE'].includes(upper) && colMap.genres === -1) {
        colMap.genres = idx;
      } else if (['LANGUAGE', 'LANG'].includes(upper) && colMap.language === -1) {
        colMap.language = idx;
      } else if (upper === 'YEAR' && colMap.year === -1) {
        colMap.year = idx;
      } else if ((upper.includes('MARK') || upper.includes('RATING') || upper.includes('SCORE') || upper.includes('NOTA')) && colMap.mark === -1) {
        colMap.mark = idx;
      }
    });

    // Fallbacks
    const idxRecords = colMap.records !== -1 ? colMap.records : 0;
    const idxCreator = colMap.creator !== -1 ? colMap.creator : 1;
    const idxType = colMap.type !== -1 ? colMap.type : 2;
    const idxGenres = colMap.genres !== -1 ? colMap.genres : 3;
    const idxLang = colMap.language !== -1 ? colMap.language : 4;
    const idxYear = colMap.year !== -1 ? colMap.year : 5;
    const idxMark = colMap.mark !== -1 ? colMap.mark : 6;

    const records = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || row.length === 0) continue;

      const titleVal = idxRecords < row.length ? row[idxRecords] : null;
      const creatorVal = idxCreator < row.length ? row[idxCreator] : null;

      if ((titleVal === null || titleVal === '') && (creatorVal === null || creatorVal === '')) {
        continue;
      }

      const title = titleVal !== null && titleVal !== undefined ? String(titleVal).trim() : '';
      const creator = creatorVal !== null && creatorVal !== undefined ? String(creatorVal).trim() : '';
      const recordType = MusicLibraryJS.parseType(idxType < row.length ? row[idxType] : '');
      const genres = MusicLibraryJS.parseGenres(idxGenres < row.length ? row[idxGenres] : '');
      const language = idxLang < row.length && row[idxLang] !== null ? String(row[idxLang]).trim() : '';

      // Year parsing
      const rawYear = idxYear < row.length ? row[idxYear] : null;
      let year = 0;
      if (rawYear !== null && rawYear !== undefined) {
        const parsedYear = parseInt(String(rawYear).trim(), 10);
        if (!isNaN(parsedYear)) year = parsedYear;
      }

      // Rating parsing
      const rawRating = idxMark < row.length ? row[idxMark] : null;
      const rating = MusicLibraryJS.parseRating(rawRating);

      records.push(new Record({
        title,
        creator,
        recordType,
        genres,
        language,
        year,
        rating,
        originalRating: rawRating !== null ? String(rawRating) : null
      }));
    }

    this.records = records;
    this.fileName = fileName;

    const validYears = records.map(r => r.year).filter(y => y > 0);
    this.minYear = validYears.length > 0 ? Math.min(...validYears) : 1900;
    this.maxYear = validYears.length > 0 ? Math.max(...validYears) : 2026;
  }

  // --- Search and Filtering Methods ---

  searchByTitle(title, exact = false, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    const query = String(title || '').trim().toLowerCase();
    if (!query) return target;
    if (exact) {
      return target.filter(r => r.title.toLowerCase() === query);
    }
    return target.filter(r => r.title.toLowerCase().includes(query));
  }

  searchByArtist(artist, exact = false, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    const query = String(artist || '').trim().toLowerCase();
    if (!query) return target;

    // Checks query against every artist in the record's parsed artist list
    if (exact) {
      return target.filter(r => r.artists.some(a => a.toLowerCase() === query));
    }
    return target.filter(r => r.artists.some(a => a.toLowerCase().includes(query)));
  }

  filterByYear(yearRange, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    if (typeof yearRange === 'number') {
      return target.filter(r => r.year === yearRange);
    }
    if (Array.isArray(yearRange) && yearRange.length === 2) {
      const [start, end] = yearRange;
      return target.filter(r => r.year >= start && r.year <= end);
    }
    return target;
  }

  filterByGenre(genre, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    const query = String(genre || '').trim().toLowerCase();
    if (!query) return target;
    return target.filter(r => r.genres.some(g => g.toLowerCase().includes(query)));
  }

  filterByRating(rating, op = '>=', truncate = false, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    if (truncate) {
      const valLimit = Math.floor(rating);
      if (op === '<=') return target.filter(r => Math.floor(r.rating) <= valLimit);
      if (op === '==') return target.filter(r => Math.floor(r.rating) === valLimit);
      return target.filter(r => Math.floor(r.rating) >= valLimit);
    }

    if (op === '<=') return target.filter(r => r.rating <= rating);
    if (op === '==') return target.filter(r => Math.abs(r.rating - rating) < 0.001);
    return target.filter(r => r.rating >= rating);
  }

  filterByLanguage(language, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    const query = String(language || '').trim().toLowerCase();
    if (!query) return target;
    return target.filter(r => r.language.toLowerCase() === query);
  }

  filterByType(typeName, recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    const query = String(typeName || '').trim().toLowerCase();
    if (!query) return target;
    return target.filter(r => r.recordType.toLowerCase() === query);
  }

  applyFilters(filters, recordsList = null) {
    let current = recordsList !== null ? recordsList : this.records;
    if (!filters) return current;

    // 1. Artist
    if (filters.artist) {
      current = this.searchByArtist(filters.artist, Boolean(filters.artistExact), current);
    }

    // 2. Title
    if (filters.title) {
      current = this.searchByTitle(filters.title, Boolean(filters.titleExact), current);
    }

    // 3. Year
    if (filters.year !== undefined && filters.year !== null && filters.year !== '') {
      current = this.filterByYear(filters.year, current);
    }

    // 4. Genre
    if (filters.genre && filters.genre !== 'Any') {
      current = this.filterByGenre(filters.genre, current);
    }

    // 5. Rating
    if (filters.rating !== undefined && filters.rating !== null && filters.rating !== '') {
      const crit = MusicLibraryJS.parseRatingCriteria(filters.rating);
      current = this.filterByRating(crit.val, crit.op, crit.truncate, current);
    }

    // 6. Language
    if (filters.language && filters.language !== 'Any') {
      current = this.filterByLanguage(filters.language, current);
    }

    // 7. Type
    if (filters.type && filters.type !== 'Any') {
      current = this.filterByType(filters.type, current);
    }

    return current;
  }

  // --- Unique Value Helpers ---

  getUniqueArtists() {
    const set = new Set();
    this.records.forEach(r => {
      r.artists.forEach(a => {
        if (a) set.add(a);
      });
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }

  getUniqueGenres() {
    const set = new Set();
    this.records.forEach(r => {
      r.genres.forEach(g => {
        if (g) set.add(g);
      });
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }

  getUniqueLanguages() {
    const set = new Set();
    this.records.forEach(r => {
      if (r.language) set.add(r.language);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }

  getUniqueTypes() {
    const set = new Set();
    this.records.forEach(r => {
      if (r.recordType) set.add(r.recordType);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base' }));
  }

  /**
   * Generates a monospaced DataFrame string replicating pandas df.to_string(index=True)
   * exactly as displayed in the readerUI python application.
   */
  toDataFrameString(recordsList = null) {
    const target = recordsList !== null ? recordsList : this.records;
    if (target.length === 0) {
      return '\n\n   No records found matching current active filters.';
    }

    // Max column width in pandas readerUI is set to 22
    const maxColWidth = 22;
    const truncate = (val, maxLen = maxColWidth) => {
      const s = String(val === null || val === undefined ? '' : val);
      if (s.length > maxLen) {
        return s.slice(0, maxLen - 3) + '...';
      }
      return s;
    };

    const rows = target.map((r, idx) => ({
      index: String(idx),
      records: truncate(r.title),
      creator: truncate(r.creator),
      type: truncate(r.recordType),
      genres: truncate(r.genres.join(', ')),
      language: truncate(r.language),
      year: String(r.year || 0),
      mark: (r.rating || 0).toFixed(3)
    }));

    // Calculate column widths
    const cols = [
      { key: 'index', header: '', align: 'right' },
      { key: 'records', header: 'RECORDS', align: 'left' },
      { key: 'creator', header: 'CREATOR', align: 'left' },
      { key: 'type', header: 'TYPE', align: 'left' },
      { key: 'genres', header: 'GENRES', align: 'left' },
      { key: 'language', header: 'LANGUAGE', align: 'left' },
      { key: 'year', header: 'YEAR', align: 'right' },
      { key: 'mark', header: 'MARK (/10)', align: 'right' }
    ];

    cols.forEach(col => {
      let maxW = col.header.length;
      rows.forEach(row => {
        const len = row[col.key].length;
        if (len > maxW) maxW = len;
      });
      col.width = maxW;
    });

    const pad = (text, width, align) => {
      const s = String(text);
      if (align === 'right') {
        return s.padStart(width, ' ');
      }
      return s.padEnd(width, ' ');
    };

    const headerLine = cols.map(c => pad(c.header, c.width, c.align)).join('  ');
    const lines = [headerLine];

    rows.forEach((row, idx) => {
      const line = cols.map(c => pad(row[c.key], c.width, c.align)).join('  ');
      lines.push(line);
    });

    return lines;
  }
}

// Export for ES modules and browser global
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { MusicLibraryJS, Record };
}
if (typeof window !== 'undefined') {
  window.MusicLibraryJS = MusicLibraryJS;
  window.Record = Record;
}
