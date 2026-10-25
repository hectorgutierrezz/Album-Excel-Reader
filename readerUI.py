import sys
import subprocess

# Ensure customtkinter is installed
try:
    import customtkinter as ctk
except ImportError:
    print("customtkinter not found. Installing now...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "customtkinter"])
    import customtkinter as ctk

from music_library import MusicLibrary
import pandas as pd

# Configure pandas formatting for display in the textbox
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)
pd.set_option('display.max_colwidth', 22)  # Limits and truncates column width to 22 chars
pd.set_option('display.colheader_justify', 'left')

ctk.set_appearance_mode("System")  # Modes: "System", "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue", "green", "dark-blue"

class MusicLibraryApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Load the library backend
        try:
            self.library = MusicLibrary('Albums List.xlsx')
        except Exception:
            try:
                import tkinter.filedialog as fd
                chosen = fd.askopenfilename(
                    title="Select Albums Excel File",
                    filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
                )
                if chosen:
                    self.library = MusicLibrary(chosen)
                else:
                    self.library = MusicLibrary()
            except Exception:
                self.library = MusicLibrary()


        # Window settings
        self.title("🎵 Music Library Browser")
        self.geometry("1700x1000")
        self.text_font_size = 12
        self.filter_font_size = 24
        self.genre_popup = None
        self.language_popup = None

        # Main Layout: 2 Columns (Sidebar & Main Content)
        self.grid_columnconfigure(0, weight=0)  # Sidebar fixed width
        self.grid_columnconfigure(1, weight=1)  # Content expands
        self.grid_rowconfigure(0, weight=1)     # Both expand vertically

        # --- Sidebar Frame ---
        self.sidebar_frame = ctk.CTkFrame(self, width=480, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(0, weight=1)  # Content frame expands
        self.sidebar_frame.grid_rowconfigure(1, weight=0)  # Theme selection stays bottom

        # Scrollable Frame inside sidebar for filters
        self.scroll_filters = ctk.CTkScrollableFrame(self.sidebar_frame, label_text="Filters")
        self.scroll_filters.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.scroll_filters.grid_columnconfigure(0, weight=1)

        # 1. Artist/Band
        self.artist_label = ctk.CTkLabel(self.scroll_filters, text="Artist/Band:", anchor="w")
        self.artist_label.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="w")
        self.artist_entry = ctk.CTkEntry(self.scroll_filters, placeholder_text="e.g. Daft Punk")
        self.artist_entry.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="ew")

        # 2. Record Title
        self.title_label = ctk.CTkLabel(self.scroll_filters, text="Record Name:", anchor="w")
        self.title_label.grid(row=2, column=0, padx=10, pady=(10, 0), sticky="w")
        self.title_entry = ctk.CTkEntry(self.scroll_filters, placeholder_text="e.g. Discovery")
        self.title_entry.grid(row=3, column=0, padx=10, pady=(0, 10), sticky="ew")

        # 3. Genre
        self.genre_label = ctk.CTkLabel(self.scroll_filters, text="Genre:", anchor="w")
        self.genre_label.grid(row=4, column=0, padx=10, pady=(10, 0), sticky="w")
        unique_genres = sorted({genre for record in self.library.records for genre in record.genres if genre})
        self.genre_options = ["Any"] + unique_genres
        self.genre_button = ctk.CTkButton(self.scroll_filters, text="Any", anchor="w", command=self.open_genre_selector)
        self.genre_button.grid(row=5, column=0, padx=10, pady=(0, 10), sticky="ew")

        # 4. Language
        self.lang_label = ctk.CTkLabel(self.scroll_filters, text="Language:", anchor="w")
        self.lang_label.grid(row=6, column=0, padx=10, pady=(10, 0), sticky="w")
        unique_languages = sorted({record.language for record in self.library.records if record.language})
        self.lang_options = ["Any"] + unique_languages
        self.lang_button = ctk.CTkButton(self.scroll_filters, text="Any", anchor="w", command=self.open_language_selector)
        self.lang_button.grid(row=7, column=0, padx=10, pady=(0, 10), sticky="ew")

        # 5. Type (Dynamic OptionMenu)
        self.type_label = ctk.CTkLabel(self.scroll_filters, text="Record Type:", anchor="w")
        self.type_label.grid(row=8, column=0, padx=10, pady=(10, 0), sticky="w")
        
        # Get unique types from library
        unique_types = sorted(list(set(r.record_type for r in self.library.records if r.record_type)))
        self.type_options = ["Any"] + unique_types
        self.type_optionmenu = ctk.CTkOptionMenu(self.scroll_filters, values=self.type_options)
        self.type_optionmenu.grid(row=9, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.type_optionmenu.set("Any")

        # 6. Year Range
        self.year_label = ctk.CTkLabel(self.scroll_filters, text="Year Range (Start - End):", anchor="w")
        self.year_label.grid(row=10, column=0, padx=10, pady=(10, 0), sticky="w")
        
        self.year_frame = ctk.CTkFrame(self.scroll_filters, fg_color="transparent")
        self.year_frame.grid(row=11, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.year_frame.grid_columnconfigure(0, weight=1)
        self.year_frame.grid_columnconfigure(1, weight=1)

        # Pre-fill year limits
        years = [r.year for r in self.library.records if r.year > 0]
        self.min_year = min(years) if years else 1900
        self.max_year = max(years) if years else 2026

        self.start_year_entry = ctk.CTkEntry(self.year_frame, placeholder_text=str(self.min_year))
        self.start_year_entry.grid(row=0, column=0, padx=(0, 5), sticky="ew")
        self.end_year_entry = ctk.CTkEntry(self.year_frame, placeholder_text=str(self.max_year))
        self.end_year_entry.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # 7. Rating
        self.rating_label = ctk.CTkLabel(self.scroll_filters, text="Rating:", anchor="w")
        self.rating_label.grid(row=12, column=0, padx=10, pady=(10, 0), sticky="w")
        
        self.rating_frame = ctk.CTkFrame(self.scroll_filters, fg_color="transparent")
        self.rating_frame.grid(row=13, column=0, padx=10, pady=(0, 10), sticky="ew")
        self.rating_frame.grid_columnconfigure(0, weight=0)
        self.rating_frame.grid_columnconfigure(1, weight=1)

        self.rating_op_menu = ctk.CTkOptionMenu(self.rating_frame, values=[">=", "<=", "=="], width=70)
        self.rating_op_menu.grid(row=0, column=0, padx=(0, 5), sticky="w")
        self.rating_op_menu.set(">=")

        self.rating_slider = ctk.CTkSlider(self.rating_frame, from_=0.0, to=10.0, number_of_steps=20, command=self.on_slider_move)
        self.rating_slider.grid(row=0, column=1, padx=(5, 0), sticky="ew")
        self.rating_slider.set(0.0)

        self.rating_val_label = ctk.CTkLabel(self.scroll_filters, text="Min Rating: 0.0 (Disabled)", anchor="w")
        self.rating_val_label.grid(row=14, column=0, padx=10, pady=(0, 10), sticky="w")

        # --- Sidebar Bottom buttons ---
        self.btn_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.btn_frame.grid(row=1, column=0, padx=15, pady=(10, 20), sticky="ew")
        self.btn_frame.grid_columnconfigure(0, weight=1)
        self.btn_frame.grid_columnconfigure(1, weight=1)

        self.apply_btn = ctk.CTkButton(self.btn_frame, text="Apply", command=self.update_results, fg_color="#1f77b4", hover_color="#155d8f")
        self.apply_btn.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.clear_btn = ctk.CTkButton(self.btn_frame, text="Clear", command=self.clear_all_filters, fg_color="#555555", hover_color="#333333")
        self.clear_btn.grid(row=0, column=1, padx=(5, 0), sticky="ew")

        # Theme OptionMenu
        self.theme_label = ctk.CTkLabel(self.sidebar_frame, text="Theme Mode:", anchor="w")
        self.theme_label.grid(row=2, column=0, padx=20, pady=(0, 5), sticky="w")
        self.theme_menu = ctk.CTkOptionMenu(self.sidebar_frame, values=["System", "Dark", "Light"], command=self.change_theme_mode)
        self.theme_menu.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")

        self.filter_widgets = [
            self.artist_label, self.artist_entry, self.title_label, self.title_entry,
            self.genre_label, self.genre_button, self.lang_label, self.lang_button,
            self.type_label, self.type_optionmenu, self.year_label, self.start_year_entry,
            self.end_year_entry, self.rating_label, self.rating_op_menu, self.rating_val_label,
            self.apply_btn, self.clear_btn, self.theme_label, self.theme_menu
        ]
        for widget in self.filter_widgets:
            widget.configure(font=ctk.CTkFont(size=self.filter_font_size))

        # --- Main Content Frame ---
        self.main_content = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_content.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
        self.main_content.grid_columnconfigure(0, weight=1)
        self.main_content.grid_rowconfigure(1, weight=1)

        # Header Info Bar
        self.info_frame = ctk.CTkFrame(self.main_content)
        self.info_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self.info_frame.grid_columnconfigure(0, weight=1)

        self.result_stats_label = ctk.CTkLabel(self.info_frame, text="Loading library records...", font=ctk.CTkFont(size=14, weight="bold"))
        self.result_stats_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

        self.zoom_out_btn = ctk.CTkButton(self.info_frame, text="-", width=32, command=lambda: self.change_text_size(-1))
        self.zoom_out_btn.grid(row=0, column=1, padx=(5, 0), pady=8)
        self.zoom_in_btn = ctk.CTkButton(self.info_frame, text="+", width=32, command=lambda: self.change_text_size(1))
        self.zoom_in_btn.grid(row=0, column=2, padx=(0, 10), pady=8)

        # Results monospaced text display
        # We use a CTkTextbox and configure a clean monospaced font
        self.results_display = ctk.CTkTextbox(self.main_content, wrap="none")
        self.results_display.grid(row=1, column=0, sticky="nsew")
        
        # Configure monospaced font (fallback hierarchy)
        self.text_font_family = "Courier New" if sys.platform == "win32" else "DejaVu Sans Mono"
        self.results_display.configure(font=(self.text_font_family, self.text_font_size))
        self.results_textbox = self.results_display._textbox
        self.results_textbox.tag_configure("selected_row", background="#2f6f9f", foreground="white")
        self.results_textbox.bind("<Button-1>", self.select_result_row)

        # Initial data update
        self.update_results()

    def on_slider_move(self, value):
        if value == 0.0:
            self.rating_val_label.configure(text="Min Rating: 0.0 (Disabled)")
        else:
            self.rating_val_label.configure(text=f"Rating Criteria: {self.rating_op_menu.get()} {value:.1f}")

    def change_theme_mode(self, mode):
        ctk.set_appearance_mode(mode)

    def change_text_size(self, change):
        self.text_font_size = max(8, min(32, self.text_font_size + change))
        self.results_display.configure(font=(self.text_font_family, self.text_font_size))

    def open_genre_selector(self):
        self.open_selector("Genre", self.genre_options, "genre_popup", self.select_genre)

    def open_language_selector(self):
        self.open_selector("Language", self.lang_options, "language_popup", self.select_language)

    def open_selector(self, title, options, popup_attribute, select_command):
        popup = getattr(self, popup_attribute)
        if popup is not None and popup.winfo_exists():
            popup.focus()
            return

        popup = ctk.CTkToplevel(self)
        setattr(self, popup_attribute, popup)
        popup.title(f"Select {title}")
        popup.geometry("380x520")
        popup.resizable(False, False)
        popup.transient(self)
        popup.grid_columnconfigure(0, weight=1)
        popup.grid_rowconfigure(0, weight=1)

        option_list = ctk.CTkScrollableFrame(popup, label_text=f"{title}s")
        option_list.grid(row=0, column=0, padx=12, pady=12, sticky="nsew")
        option_list.grid_columnconfigure(0, weight=1)

        for option in options:
            option_button = ctk.CTkButton(
                option_list,
                text=option,
                anchor="w",
                font=ctk.CTkFont(size=self.filter_font_size),
                command=lambda value=option: select_command(value)
            )
            option_button.grid(row=option_list.grid_size()[1], column=0, padx=6, pady=3, sticky="ew")

    def select_genre(self, genre):
        self.genre_button.configure(text=genre)
        if self.genre_popup is not None and self.genre_popup.winfo_exists():
            self.genre_popup.destroy()
        self.genre_popup = None

    def select_language(self, language):
        self.lang_button.configure(text=language)
        if self.language_popup is not None and self.language_popup.winfo_exists():
            self.language_popup.destroy()
        self.language_popup = None

    def select_result_row(self, event):
        line_number = int(self.results_textbox.index(f"@{event.x},{event.y}").split(".")[0])
        if line_number < 3:
            return
        self.results_textbox.tag_remove("selected_row", "1.0", "end")
        self.results_textbox.tag_add("selected_row", f"{line_number}.0", f"{line_number}.end+1c")

    def clear_all_filters(self):
        # Reset text entries
        self.artist_entry.delete(0, "end")
        self.title_entry.delete(0, "end")
        self.genre_button.configure(text="Any")
        self.lang_button.configure(text="Any")
        
        # Reset options
        self.type_optionmenu.set("Any")
        self.rating_op_menu.set(">=")
        self.rating_slider.set(0.0)
        self.rating_val_label.configure(text="Min Rating: 0.0 (Disabled)")
        
        # Reset year ranges
        self.start_year_entry.delete(0, "end")
        self.end_year_entry.delete(0, "end")
        
        # Refresh display
        self.update_results()

    def update_results(self):
        filters = {}

        # 1. Text Filters
        artist = self.artist_entry.get().strip()
        title = self.title_entry.get().strip()
        genre = self.genre_button.cget("text")
        lang = self.lang_button.cget("text")
        rec_type = self.type_optionmenu.get()

        if artist: filters['artist'] = artist
        if title: filters['title'] = title
        if genre != "Any": filters['genre'] = genre
        if lang != "Any": filters['language'] = lang
        if rec_type != "Any": filters['type'] = rec_type

        # 2. Year Range
        start_year_str = self.start_year_entry.get().strip()
        end_year_str = self.end_year_entry.get().strip()
        
        # Parse inputs, fall back to actual min/max in the database if empty
        try:
            start_year = int(start_year_str) if start_year_str else self.min_year
        except ValueError:
            start_year = self.min_year

        try:
            end_year = int(end_year_str) if end_year_str else self.max_year
        except ValueError:
            end_year = self.max_year

        filters['year'] = (start_year, end_year)

        # 3. Rating filter
        rating_value = self.rating_slider.get()
        if rating_value > 0.0:
            filters['rating'] = {
                'val': rating_value,
                'op': self.rating_op_menu.get(),
                'truncate': False
            }

        # Apply using backend API
        results = self.library.apply_filters(filters)
        
        # Render clean pandas dataframe string
        df = self.library.to_dataframe(results)
        df_str = df.to_string(index=True)

        # Update statistics header
        self.result_stats_label.configure(
            text=f"📊 Showing {len(results)} of {len(self.library.records)} records"
        )

        # Update main text area
        self.results_display.configure(state="normal")
        self.results_display.delete("1.0", "end")
        self.results_textbox.tag_remove("selected_row", "1.0", "end")
        
        if len(results) == 0:
            self.results_display.insert("1.0", "\n\n   No records found matching current active filters.")
        else:
            self.results_display.insert("1.0", df_str)
            
        self.results_display.configure(state="disabled")

if __name__ == "__main__":
    app = MusicLibraryApp()
    app.mainloop()
