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
        # Get filter option from the user
        print("\nFILTERING OPTIONS:")
        print("1.- ARTIST/BAND")
        print("2.- RECORD NAME")
        print("3.- YEAR OF RELEASE")
        print("4.- RECORD GENRES")
        print("5.- MY RATING")
        print("6.- RECORD LANGUAGES")
        print("7.- NO FILTERS")
        filter_option = input("What would you like to filter by? (Please type the number of the filter you want): ").strip()

        results = None

        # Filter by artist/band
        if filter_option == "1":
            filter_value = input("Enter the artist name: ")
            results = library.search_by_artist(filter_value)
            
        # Filter by album title
        elif filter_option == "2":
            filter_value = input("Enter the Record name: ")
            results = library.search_by_title(filter_value)
            
        # Filter by year or range
        elif filter_option == "3":
            filter_value = input("Enter the year (e.g., 2018) or a range (e.g., 2018-2022): ").strip()
            if "-" in filter_value:
                parts = filter_value.split("-")
                try:
                    start = int(parts[0].strip())
                    end = int(parts[1].strip())
                    results = library.filter_by_year((start, end))
                except ValueError:
                    print("Invalid range format. Please use 'YYYY-YYYY'.")
                    continue
            else:
                try:
                    year = int(filter_value)
                    results = library.filter_by_year(year)
                except ValueError:
                    print("Invalid year format.")
                    continue

        # Filter by genre
        elif filter_option == "4":
            filter_value = input("Enter the genre: ")
            results = library.filter_by_genre(filter_value)

        # Filter by rating
        elif filter_option == "5":
            print("Rating filtering formats:")
            print("- A whole number (e.g., 9) to search for exactly that rating (matches 9.0 to 9.99)")
            print("- A decimal number (e.g., 9.5) to search for ratings >= 9.5")
            print("- An expression with >= (e.g., >=8.5) to search for ratings >= 8.5")
            print("- An expression with <= (e.g., <=6.0) to search for ratings <= 6.0")
            
            filter_value = input("Enter the rating criteria: ").strip()
            
            if filter_value.startswith(">="):
                try:
                    val = float(filter_value[2:].strip())
                    results = library.filter_by_rating(val, truncate=False)
                except ValueError:
                    print("Invalid rating format.")
                    continue
            
            else:
                try:
                    val = float(filter_value)
                    # If it's an integer rating, truncate and match exactly that integer range (e.g., 9.0 to 9.9)
                    if val.is_integer():
                        results = library.filter_by_rating(val, truncate=True)
                    else:
                        results = library.filter_by_rating(val, truncate=False)
                except ValueError:
                    print("Invalid rating format.")
                    continue

        # Filter by language
        elif filter_option == "6":
            filter_value = input("Enter language (in English): ")
            results = library.filter_by_language(filter_value)
        
        # No filter option selected
        elif filter_option == "7":
            results = library.records
        
        # Invalid filter option
        else:
            print("Invalid filter option.")
            continue

        # Print result dataframe
        if results is not None:
            if len(results) == 0:
                print("\nNo records found matching your filters.")
            else:
                print(f"\nFound {len(results)} records:")
                print(library.to_dataframe(results))

        print("\n" + "-"*40)
        print("Would you like to perform another filter?")
        print("1. Yes")
        print("2. No")
        sub_choice = input()
        if sub_choice == "1":
            continue
        elif sub_choice == "2":
            break
        else:
            print("Invalid option. Returning to main menu.")

    # Exit the system
    elif choice == "2":
        break

    # Invalid option
    else:
        print("Invalid option.")
