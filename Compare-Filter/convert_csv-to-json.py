Entimport csv
import json

def convert_csv_to_json(csv_filename, json_filename):
    try:
        with open(csv_filename, 'r', encoding='utf-8', errors='ignore') as csvfile: # Added errors='ignore'
            csv_reader = csv.DictReader(csvfile)
            data_list = list(csv_reader)

        with open(json_filename, 'w', encoding='utf-8') as jsonfile:
            json.dump(data_list, jsonfile, indent=4)

        print(f"Successfully converted '{csv_filename}' to '{json_filename}'")

    except FileNotFoundError:
        print(f"Error: Could not find the file '{csv_filename}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

convert_csv_to_json('all_inventory_list.csv', 'all_inventory_list.json')er file contents here
