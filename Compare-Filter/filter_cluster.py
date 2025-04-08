import pandas as pd
import json

def filter_json_to_csv(json_filename, csv_filename, cluster_filter):
    """
    Filters a JSON file based on a cluster value and saves the result to a CSV file.

    Args:
        json_filename (str): The path to the input JSON file.
        csv_filename (str): The path to the output CSV file.
        cluster_filter (str): The cluster value to filter by.
    """
    try:
        with open(json_filename, 'r') as f:
            json_data = json.load(f)

        df = pd.DataFrame(json_data)
        filtered_df = df[df['cluster'] == cluster_filter]
        filtered_df.to_csv(csv_filename, index=False)

        print(f"Successfully filtered '{json_filename}' and saved to '{csv_filename}'.")

    except FileNotFoundError:
        print(f"Error: Could not find the file '{json_filename}'.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage:
filter_json_to_csv('all_inventory_list.json', 'HQ-VMWARE-Cluster.csv', 'HQ-VMWARE')
