import pandas as pd
from music_library import MusicLibrary

# Initial settings for the pandas data frame display
pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', 30)
pd.set_option('display.colheader_justify', 'right')
pd.set_option('display.precision', 3)
pd.set_option('display.float_format', '{:.3f}'.format)

# Load the music library backend
try:
    library = MusicLibrary('Albums List.xlsx')
except Exception as e:
    print(f"Error loading Excel sheet: {e}")
    exit(1)

print("Hello! This is your music library. What would you like to do?")

# Option loop for the user
while True:
    print("\n" + "="*40)
    print("1. Filter albums")
    print("2. Exit")
    choice = input("What would you like to do? ")
    
    if choice == "1":
        active_filters = {}
        while True:
            # Display active filters nicely
            if not active_filters:
                print("\nActive filters: None")
            else:
                filter_desc = []
                for k, v in active_filters.items():
                    if k == 'year':
                        if isinstance(v, tuple):
                            filter_desc.append(f"Year: {v[0]}-{v[1]}")
                        else:
                            filter_desc.append(f"Year: {v}")
                    else:
                        filter_desc.append(f"{k.capitalize()}: {v}")
                print("\nActive filters: " + ", ".join(filter_desc))

            # Get filter option from the user
            print("\nFILTERING OPTIONS:")
            print("1.- ARTIST/BAND")
            print("2.- RECORD NAME")
            print("3.- YEAR OF RELEASE")
            print("4.- RECORD GENRES")
            print("5.- MY RATING")
            print("6.- RECORD LANGUAGES")
            print("7.- RECORD TYPE")
            print("8.- CLEAR ALL FILTERS")
            print("9.- BACK TO MAIN MENU")
            filter_option = input("What would you like to do? (Please type the number): ").strip()

            if filter_option == "9":
                break

            if filter_option == "8":
                active_filters.clear()
                print("All filters cleared.")
                continue

            results = None

            # Filter by artist/band
            if filter_option == "1":
                filter_value = input("Enter the artist name: ").strip()
                if filter_value:
                    active_filters['artist'] = filter_value

            # Filter by album title
            elif filter_option == "2":
                filter_value = input("Enter the Record name: ").strip()
                if filter_value:
                    active_filters['title'] = filter_value

            # Filter by year or range
            elif filter_option == "3":
                filter_value = input("Enter the year (e.g., 2018) or a range (e.g., 2018-2022): ").strip()
                if "-" in filter_value:
                    parts = filter_value.split("-")
                    try:
                        start = int(parts[0].strip())
                        end = int(parts[1].strip())
                        active_filters['year'] = (start, end)
                    except ValueError:
                        print("Invalid range format. Please use 'YYYY-YYYY'.")
                        continue
                else:
                    try:
                        year = int(filter_value)
                        active_filters['year'] = year
                    except ValueError:
                        print("Invalid year format.")
                        continue

            # Filter by genre
            elif filter_option == "4":
                filter_value = input("Enter the genre: ").strip()
                if filter_value:
                    active_filters['genre'] = filter_value

            # Filter by rating
            elif filter_option == "5":
                print("Rating filtering formats:")
                print("- A whole number (e.g., 9) to search for exactly that rating (matches 9.0 to 9.99)")
                print("- A decimal number (e.g., 9.5) to search for ratings >= 9.5")
                print("- An expression with >= (e.g., >=8.5) to search for ratings >= 8.5")
                print("- An expression with <= (e.g., <=6.0) to search for ratings <= 6.0")

                filter_value = input("Enter the rating criteria: ").strip()
                if filter_value:
                    active_filters['rating'] = filter_value

            # Filter by language
            elif filter_option == "6":
                filter_value = input("Enter language (in English): ").strip()
                if filter_value:
                    active_filters['language'] = filter_value

            # Filter by type
            elif filter_option == "7":
                filter_value = input("Enter record type (e.g., Album, EP): ").strip()
                if filter_value:
                    active_filters['type'] = filter_value

            # Invalid option
            else:
                print("Invalid filter option.")
                continue

            # Apply accumulated filters
            results = library.apply_filters(active_filters)

            # Print result dataframe
            if results is not None:
                if len(results) == 0:
                    print("\nNo records found matching your active filters.")
                else:
                    print(f"\nFound {len(results)} records:")
                    print(library.to_dataframe(results))

            print("\n" + "-"*40)
            print("Would you like to perform another filter (accumulating with current ones)?")
            print("1. Yes")
            print("2. No (clears filters and returns to main menu)")
            sub_choice = input().strip()
            if sub_choice == "1":
                continue
            else:
                break

    # Exit the system
    elif choice == "2":
        break

    # Invalid option
    else:
        print("Invalid option.")
