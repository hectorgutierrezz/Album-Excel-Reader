/**
 * Music Library Browser - Main Web Application Script
 * Provides UI interaction, Excel file management, filters, and rendering.
 */

document.addEventListener('DOMContentLoaded', () => {
  // App State
  const state = {
    library: new MusicLibraryJS(),
    currentRecords: [],
    selectedGenre: 'Any',
    selectedLanguage: 'Any',
    viewMode: 'mono', // 'mono' or 'table'
    fontSize: 13,
    theme: localStorage.getItem('music_theme') || 'Dark',
    selectedRowIndex: null,
    sortCol: null,
    sortAsc: true
  };

  // DOM Elements
  const elements = {
    fileInput: document.getElementById('file-input'),
    dropzone: document.getElementById('file-dropzone'),
    dropzoneText: document.getElementById('dropzone-text'),
    fileStatusBadge: document.getElementById('file-status-badge'),
    loadedFileName: document.getElementById('loaded-file-name'),
    loadedFileCount: document.getElementById('loaded-file-count'),
    btnChangeFile: document.getElementById('btn-change-file'),

    // Filter Inputs
    artistInput: document.getElementById('artist-input'),
    titleInput: document.getElementById('title-input'),
    btnOpenGenreModal: document.getElementById('btn-open-genre-modal'),
    selectedGenreLabel: document.getElementById('selected-genre-label'),
    btnOpenLangModal: document.getElementById('btn-open-lang-modal'),
    selectedLangLabel: document.getElementById('selected-lang-label'),
    typeSelect: document.getElementById('type-select'),
    yearStartInput: document.getElementById('year-start-input'),
    yearEndInput: document.getElementById('year-end-input'),
    ratingOpSelect: document.getElementById('rating-op-select'),
    ratingSlider: document.getElementById('rating-slider'),
    ratingFeedbackLabel: document.getElementById('rating-feedback-label'),

    // Action buttons
    btnApplyFilters: document.getElementById('btn-apply-filters'),
    btnClearFilters: document.getElementById('btn-clear-filters'),
    themeSelect: document.getElementById('theme-select'),

    // Header Controls
    statsText: document.getElementById('stats-text'),
    activeFilterTags: document.getElementById('active-filter-tags'),
    btnViewMono: document.getElementById('btn-view-mono'),
    btnViewTable: document.getElementById('btn-view-table'),
    btnZoomOut: document.getElementById('btn-zoom-out'),
    btnZoomIn: document.getElementById('btn-zoom-in'),

    // Display areas
    resultsContainer: document.getElementById('results-container'),
    monoDisplay: document.getElementById('mono-display'),
    tableDisplay: document.getElementById('table-display'),
    tableBody: document.getElementById('table-body'),

    // Modals
    genreModal: document.getElementById('genre-modal'),
    btnCloseGenreModal: document.getElementById('btn-close-genre-modal'),
    genreSearchInput: document.getElementById('genre-search-input'),
    genreOptionsList: document.getElementById('genre-options-list'),

    langModal: document.getElementById('lang-modal'),
    btnCloseLangModal: document.getElementById('btn-close-lang-modal'),
    langSearchInput: document.getElementById('lang-search-input'),
    langOptionsList: document.getElementById('lang-options-list'),

    artistModal: document.getElementById('artist-modal'),
    btnOpenArtistsModal: document.getElementById('btn-open-artists-modal'),
    btnCloseArtistModal: document.getElementById('btn-close-artist-modal'),
    artistSearchInput: document.getElementById('artist-search-input'),
    artistOptionsList: document.getElementById('artist-options-list')
  };

  // --- Theme Management ---
  function applyTheme(theme) {
    state.theme = theme;
    localStorage.setItem('music_theme', theme);
    elements.themeSelect.value = theme;

    if (theme === 'System') {
      const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      document.documentElement.setAttribute('data-theme', prefersDark ? 'dark' : 'light');
    } else {
      document.documentElement.setAttribute('data-theme', theme.toLowerCase());
    }
  }

  elements.themeSelect.addEventListener('change', (e) => applyTheme(e.target.value));
  window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
    if (state.theme === 'System') applyTheme('System');
  });
  applyTheme(state.theme);

  // --- Zoom Font Size Management ---
  function setFontSize(newSize) {
    state.fontSize = Math.max(9, Math.min(28, newSize));
    document.documentElement.style.setProperty('--font-size-table', `${state.fontSize}px`);
  }

  elements.btnZoomOut.addEventListener('click', () => setFontSize(state.fontSize - 1));
  elements.btnZoomIn.addEventListener('click', () => setFontSize(state.fontSize + 1));

  // --- View Mode Toggle ---
  function setViewMode(mode) {
    state.viewMode = mode;
    if (mode === 'mono') {
      elements.btnViewMono.classList.add('active');
      elements.btnViewTable.classList.remove('active');
      elements.monoDisplay.style.display = 'block';
      elements.tableDisplay.style.display = 'none';
    } else {
      elements.btnViewMono.classList.remove('active');
      elements.btnViewTable.classList.add('active');
      elements.monoDisplay.style.display = 'none';
      elements.tableDisplay.style.display = 'table';
    }
    renderResults();
  }

  elements.btnViewMono.addEventListener('click', () => setViewMode('mono'));
  elements.btnViewTable.addEventListener('click', () => setViewMode('table'));

  // --- IndexedDB Cache for Local File Persistence ---
  const DB_NAME = 'AlbumExcelReaderDB';
  const STORE_NAME = 'libraryCache';

  function openDatabase() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, 1);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME);
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  async function cacheFileBuffer(fileName, buffer) {
    try {
      const db = await openDatabase();
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({ fileName, buffer, date: new Date().toISOString() }, 'currentFile');
    } catch (err) {
      console.warn('Could not cache file to IndexedDB:', err);
    }
  }

  async function loadCachedFileBuffer() {
    try {
      const db = await openDatabase();
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      return new Promise((resolve) => {
        const req = store.get('currentFile');
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      });
    } catch (err) {
      return null;
    }
  }

  // --- File Processing ---
  async function processArrayBuffer(buffer, fileName) {
    try {
      state.library.loadFromArrayBuffer(buffer, fileName);
      await cacheFileBuffer(fileName, buffer);

      // Update File Status Card
      elements.fileStatusBadge.style.display = 'flex';
      elements.loadedFileName.textContent = fileName;
      elements.loadedFileCount.textContent = `${state.library.records.length} records`;

      // Populate filter dropdowns based on loaded library
      populateFilterOptions();

      // Reset & Apply Filters
      clearAllFilters();
    } catch (err) {
      console.error('Error processing Excel spreadsheet:', err);
      alert(`Could not load Excel file: ${err.message}`);
    }
  }

  function handleFileSelected(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      processArrayBuffer(e.target.result, file.name);
    };
    reader.readAsArrayBuffer(file);
  }

  // File Picker & Dropzone listeners
  elements.dropzone.addEventListener('click', () => elements.fileInput.click());
  elements.btnChangeFile.addEventListener('click', () => elements.fileInput.click());
  elements.fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFileSelected(e.target.files[0]);
    }
  });

  elements.dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.dropzone.classList.add('dragover');
  });

  elements.dropzone.addEventListener('dragleave', () => {
    elements.dropzone.classList.remove('dragover');
  });

  elements.dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.dropzone.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFileSelected(e.dataTransfer.files[0]);
    }
  });

  // --- Initialize Options from Library ---
  function populateFilterOptions() {
    const lib = state.library;

    // 1. Record Types
    const types = ['Any', ...lib.getUniqueTypes()];
    elements.typeSelect.innerHTML = types.map(t => `<option value="${t}">${t}</option>`).join('');
    elements.typeSelect.value = 'Any';

    // 2. Year limits
    elements.yearStartInput.placeholder = String(lib.minYear);
    elements.yearEndInput.placeholder = String(lib.maxYear);
    elements.yearStartInput.value = '';
    elements.yearEndInput.value = '';
  }

  // --- Rating Slider Feedback ---
  function updateRatingLabel() {
    const val = parseFloat(elements.ratingSlider.value);
    const op = elements.ratingOpSelect.value;
    if (val === 0.0) {
      elements.ratingFeedbackLabel.textContent = 'Min Rating: 0.0 (Disabled)';
    } else {
      elements.ratingFeedbackLabel.textContent = `Rating Criteria: ${op} ${val.toFixed(1)}`;
    }
  }

  elements.ratingSlider.addEventListener('input', updateRatingLabel);
  elements.ratingOpSelect.addEventListener('change', updateRatingLabel);

  // --- Modals Management (Genre, Language & Quick Artist) ---
  function openModal(modalEl) {
    modalEl.classList.add('open');
    modalEl.setAttribute('aria-hidden', 'false');
    const searchInput = modalEl.querySelector('.modal-search input');
    if (searchInput) {
      searchInput.value = '';
      searchInput.focus();
    }
  }

  function closeModal(modalEl) {
    modalEl.classList.remove('open');
    modalEl.setAttribute('aria-hidden', 'true');
  }

  // Backdrop click and ESC listener
  document.querySelectorAll('.modal-backdrop').forEach(modal => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) closeModal(modal);
    });
  });

  window.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      document.querySelectorAll('.modal-backdrop.open').forEach(closeModal);
    }
  });

  // --- Genre Modal ---
  function renderGenreOptions(filterText = '') {
    const genres = ['Any', ...state.library.getUniqueGenres()];
    const query = filterText.trim().toLowerCase();
    const filtered = query ? genres.filter(g => g.toLowerCase().includes(query)) : genres;

    elements.genreOptionsList.innerHTML = filtered.map(genre => `
      <button type="button" class="btn-option-item ${state.selectedGenre === genre ? 'selected' : ''}" data-val="${genre}">
        ${genre}
      </button>
    `).join('');

    elements.genreOptionsList.querySelectorAll('.btn-option-item').forEach(btn => {
      btn.addEventListener('click', () => {
        selectGenre(btn.dataset.val);
        closeModal(elements.genreModal);
      });
    });
  }

  function selectGenre(genre) {
    state.selectedGenre = genre;
    elements.selectedGenreLabel.textContent = genre;
  }

  elements.btnOpenGenreModal.addEventListener('click', () => {
    renderGenreOptions();
    openModal(elements.genreModal);
  });
  elements.btnCloseGenreModal.addEventListener('click', () => closeModal(elements.genreModal));
  elements.genreSearchInput.addEventListener('input', (e) => renderGenreOptions(e.target.value));

  // --- Language Modal ---
  function renderLangOptions(filterText = '') {
    const langs = ['Any', ...state.library.getUniqueLanguages()];
    const query = filterText.trim().toLowerCase();
    const filtered = query ? langs.filter(l => l.toLowerCase().includes(query)) : langs;

    elements.langOptionsList.innerHTML = filtered.map(lang => `
      <button type="button" class="btn-option-item ${state.selectedLanguage === lang ? 'selected' : ''}" data-val="${lang}">
        ${lang}
      </button>
    `).join('');

    elements.langOptionsList.querySelectorAll('.btn-option-item').forEach(btn => {
      btn.addEventListener('click', () => {
        selectLanguage(btn.dataset.val);
        closeModal(elements.langModal);
      });
    });
  }

  function selectLanguage(lang) {
    state.selectedLanguage = lang;
    elements.selectedLangLabel.textContent = lang;
  }

  elements.btnOpenLangModal.addEventListener('click', () => {
    renderLangOptions();
    openModal(elements.langModal);
  });
  elements.btnCloseLangModal.addEventListener('click', () => closeModal(elements.langModal));
  elements.langSearchInput.addEventListener('input', (e) => renderLangOptions(e.target.value));

  // --- Quick Artist Modal ---
  function renderArtistOptions(filterText = '') {
    const artists = state.library.getUniqueArtists();
    const query = filterText.trim().toLowerCase();
    const filtered = query ? artists.filter(a => a.toLowerCase().includes(query)) : artists;

    if (filtered.length === 0) {
      elements.artistOptionsList.innerHTML = '<div style="padding: 12px; color: var(--text-muted); text-align: center;">No matching artists found</div>';
      return;
    }

    elements.artistOptionsList.innerHTML = filtered.map(artist => `
      <button type="button" class="btn-option-item ${elements.artistInput.value.trim().toLowerCase() === artist.toLowerCase() ? 'selected' : ''}" data-val="${artist}">
        ${artist}
      </button>
    `).join('');

    elements.artistOptionsList.querySelectorAll('.btn-option-item').forEach(btn => {
      btn.addEventListener('click', () => {
        elements.artistInput.value = btn.dataset.val;
        closeModal(elements.artistModal);
        applyFilters();
      });
    });
  }

  elements.btnOpenArtistsModal.addEventListener('click', () => {
    if (state.library.records.length === 0) {
      alert('Please load an Excel file first.');
      return;
    }
    renderArtistOptions();
    openModal(elements.artistModal);
  });
  elements.btnCloseArtistModal.addEventListener('click', () => closeModal(elements.artistModal));
  elements.artistSearchInput.addEventListener('input', (e) => renderArtistOptions(e.target.value));

  // --- Filter Application ---
  function getActiveFilterCriteria() {
    const filters = {};

    const artist = elements.artistInput.value.trim();
    if (artist) filters.artist = artist;

    const title = elements.titleInput.value.trim();
    if (title) filters.title = title;

    if (state.selectedGenre && state.selectedGenre !== 'Any') {
      filters.genre = state.selectedGenre;
    }

    if (state.selectedLanguage && state.selectedLanguage !== 'Any') {
      filters.language = state.selectedLanguage;
    }

    const typeVal = elements.typeSelect.value;
    if (typeVal && typeVal !== 'Any') {
      filters.type = typeVal;
    }

    // Year Range
    const startStr = elements.yearStartInput.value.trim();
    const endStr = elements.yearEndInput.value.trim();
    const startYear = startStr ? parseInt(startStr, 10) : state.library.minYear;
    const endYear = endStr ? parseInt(endStr, 10) : state.library.maxYear;

    if (!isNaN(startYear) && !isNaN(endYear)) {
      filters.year = [startYear, endYear];
      filters.hasYearCustom = Boolean(startStr || endStr);
    }

    // Rating
    const ratingVal = parseFloat(elements.ratingSlider.value);
    if (ratingVal > 0.0) {
      filters.rating = {
        val: ratingVal,
        op: elements.ratingOpSelect.value,
        truncate: false
      };
    }

    return filters;
  }

  function applyFilters() {
    const filters = getActiveFilterCriteria();
    state.currentRecords = state.library.applyFilters(filters);
    state.selectedRowIndex = null;
    renderResults();
    renderActiveFilterBadges(filters);
  }

  function clearAllFilters() {
    elements.artistInput.value = '';
    elements.titleInput.value = '';
    selectGenre('Any');
    selectLanguage('Any');
    elements.typeSelect.value = 'Any';
    elements.yearStartInput.value = '';
    elements.yearEndInput.value = '';
    elements.ratingSlider.value = 0;
    elements.ratingOpSelect.value = '>=';
    updateRatingLabel();

    state.currentRecords = [...state.library.records];
    state.selectedRowIndex = null;
    renderResults();
    renderActiveFilterBadges({});
  }

  elements.btnApplyFilters.addEventListener('click', applyFilters);
  elements.btnClearFilters.addEventListener('click', clearAllFilters);

  // Trigger apply on Enter key in text inputs
  [elements.artistInput, elements.titleInput, elements.yearStartInput, elements.yearEndInput].forEach(inp => {
    inp.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') applyFilters();
    });
  });

  // --- Active Filter Tags ---
  function renderActiveFilterBadges(filters) {
    const tags = [];

    if (filters.artist) {
      tags.push({ label: `Artist: ${filters.artist}`, clear: () => { elements.artistInput.value = ''; applyFilters(); } });
    }
    if (filters.title) {
      tags.push({ label: `Title: ${filters.title}`, clear: () => { elements.titleInput.value = ''; applyFilters(); } });
    }
    if (filters.genre) {
      tags.push({ label: `Genre: ${filters.genre}`, clear: () => { selectGenre('Any'); applyFilters(); } });
    }
    if (filters.language) {
      tags.push({ label: `Lang: ${filters.language}`, clear: () => { selectLanguage('Any'); applyFilters(); } });
    }
    if (filters.type) {
      tags.push({ label: `Type: ${filters.type}`, clear: () => { elements.typeSelect.value = 'Any'; applyFilters(); } });
    }
    if (filters.hasYearCustom && filters.year) {
      tags.push({ label: `Year: ${filters.year[0]}-${filters.year[1]}`, clear: () => { elements.yearStartInput.value = ''; elements.yearEndInput.value = ''; applyFilters(); } });
    }
    if (filters.rating) {
      tags.push({ label: `Rating ${filters.rating.op} ${filters.rating.val}`, clear: () => { elements.ratingSlider.value = 0; updateRatingLabel(); applyFilters(); } });
    }

    elements.activeFilterTags.innerHTML = tags.map((t, idx) => `
      <span class="filter-tag">
        ${t.label}
        <span class="tag-remove" data-idx="${idx}">&times;</span>
      </span>
    `).join('');

    elements.activeFilterTags.querySelectorAll('.tag-remove').forEach(rm => {
      rm.addEventListener('click', () => {
        const idx = parseInt(rm.dataset.idx, 10);
        if (tags[idx]) tags[idx].clear();
      });
    });
  }

  // --- Results Rendering ---
  function renderResults() {
    const total = state.library.records.length;
    const count = state.currentRecords.length;

    // Update Header Bar
    if (total === 0) {
      elements.statsText.textContent = 'No library loaded';
    } else {
      elements.statsText.textContent = `Showing ${count} of ${total} records`;
    }

    if (total === 0) {
      elements.monoDisplay.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">💿</div>
          <h3>No Records Loaded</h3>
          <p>Select your <strong>Albums List.xlsx</strong> file on the sidebar to get started.</p>
        </div>
      `;
      elements.tableBody.innerHTML = '';
      return;
    }

    if (count === 0) {
      const emptyHtml = `
        <div class="empty-state">
          <div class="empty-state-icon">🔍</div>
          <h3>No records found</h3>
          <p>No albums match your current active filters. Try loosening your search criteria.</p>
        </div>
      `;
      elements.monoDisplay.innerHTML = emptyHtml;
      elements.tableBody.innerHTML = `<tr><td colspan="9" style="text-align: center; padding: 40px; color: var(--text-muted);">No records found matching filters.</td></tr>`;
      return;
    }

    if (state.viewMode === 'mono') {
      renderMonospaceView();
    } else {
      renderTableView();
    }
  }

  // 1. Monospace DataFrame View (Replicating readerUI.py CTkTextbox)
  function renderMonospaceView() {
    const lines = state.library.toDataFrameString(state.currentRecords);
    if (!Array.isArray(lines) || lines.length === 0) return;

    let html = `<div class="monospace-row header-row">${lines[0]}</div>`;
    for (let i = 1; i < lines.length; i++) {
      const isSelected = state.selectedRowIndex === (i - 1);
      html += `<div class="monospace-row ${isSelected ? 'selected' : ''}" data-idx="${i - 1}">${lines[i]}</div>`;
    }

    elements.monoDisplay.innerHTML = html;

    // Attach row selection click handlers
    elements.monoDisplay.querySelectorAll('.monospace-row:not(.header-row)').forEach(rowEl => {
      rowEl.addEventListener('click', () => {
        const idx = parseInt(rowEl.dataset.idx, 10);
        if (state.selectedRowIndex === idx) {
          state.selectedRowIndex = null;
        } else {
          state.selectedRowIndex = idx;
        }
        renderMonospaceView();
      });
    });
  }

  // 2. Interactive Data Table View
  function renderTableView() {
    elements.tableBody.innerHTML = state.currentRecords.map((r, idx) => {
      const isSelected = state.selectedRowIndex === idx;
      let ratingClass = 'rating-low';
      if (r.rating >= 8.5) ratingClass = 'rating-high';
      else if (r.rating >= 7.0) ratingClass = 'rating-mid';

      const genresHtml = r.genres.map(g => `<span class="genre-tag-pill">${g}</span>`).join('');

      return `
        <tr class="${isSelected ? 'selected' : ''}" data-idx="${idx}">
          <td style="color: var(--text-muted);">${idx}</td>
          <td style="font-weight: 600; color: var(--text-primary);">${r.title}</td>
          <td>${r.creator}</td>
          <td>${r.recordType}</td>
          <td>${genresHtml}</td>
          <td>${r.language}</td>
          <td style="text-align: right;">${r.year || '-'}</td>
          <td style="text-align: right;"><span class="rating-badge ${ratingClass}">${(r.rating || 0).toFixed(2)}</span></td>
          <td>${r.woke ? 'Yeh' : 'Nah'}</td>
        </tr>
      `;
    }).join('');

    elements.tableBody.querySelectorAll('tr').forEach(tr => {
      tr.addEventListener('click', () => {
        const idx = parseInt(tr.dataset.idx, 10);
        state.selectedRowIndex = state.selectedRowIndex === idx ? null : idx;
        renderTableView();
      });
    });
  }

  // Table Column Sorting
  document.querySelectorAll('.table-view th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.sort;
      if (state.sortCol === col) {
        state.sortAsc = !state.sortAsc;
      } else {
        state.sortCol = col;
        state.sortAsc = true;
      }

      state.currentRecords.sort((a, b) => {
        let valA = a[col];
        let valB = b[col];
        if (col === 'index') return 0;
        if (typeof valA === 'string') {
          return state.sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        valA = valA || 0;
        valB = valB || 0;
        return state.sortAsc ? valA - valB : valB - valA;
      });

      renderTableView();
    });
  });

  // --- Auto-Load / Session Recovery ---
  async function initAutoLoad() {
    // 1. Check IndexedDB cache
    const cached = await loadCachedFileBuffer();
    if (cached && cached.buffer) {
      console.log('Restoring library from local browser cache:', cached.fileName);
      await processArrayBuffer(cached.buffer, cached.fileName);
      return;
    }

    // 2. Fallback: If running on local server where Albums List.xlsx might be served
    if (window.location.protocol.startsWith('http')) {
      try {
        const resp = await fetch('Albums List.xlsx');
        if (resp.ok) {
          const buf = await resp.arrayBuffer();
          console.log('Auto-loaded local Albums List.xlsx');
          await processArrayBuffer(buf, 'Albums List.xlsx');
        }
      } catch (err) {
        // Normal behavior when running on GitHub Pages (static, file not served)
      }
    }
  }

  initAutoLoad();
});
