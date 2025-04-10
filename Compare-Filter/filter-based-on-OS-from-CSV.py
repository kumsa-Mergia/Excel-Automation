import pandas as pd

def filter_csv_to_excel_sheets(csv_filename, excel_filename):
    try:
        df = pd.read_csv(csv_filename, encoding='latin-1')

        # Define OS groups where multiple names should be in the same sheet
        os_groups = {
            "Windows Server": ["Window Server", "Windows Server"],
            "Linux Appliance": ["linux_Appliance", "Linux-appliance"],
            "Alpine Linux": ["Alpine Linux"],
            "AIX": ["AIX"],
            "Debian": ["Debian"],
            "CentOS": ["Centos"],
            "KEMP": ["KEMP"],
            "Gaia OS": ["Gaia OS"],
            "Other Linux": ["Other Linux"],
            "Oracle Linux": ["Oracle Linux"],
            "RockyLinux": ["RockyLinux"],
            "RedHat": ["RedHat"],
            "SuSe Linux": ["SuSe Linux"],
            "Ubuntu": ["Ubuntu", "ubunu"],
            "_": ["_"],
            "nan": ["nan"]  # If "nan" is a string in your data
        }

        with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
            for sheet_name, group in os_groups.items():
                filtered_df = df[df['OSType'].isin(group)]
                if not filtered_df.empty:
                    filtered_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)

        print(f"Successfully filtered '{csv_filename}' and saved to '{excel_filename}'.")

    except FileNotFoundError:
        print(f"Error: Could not find the file '{csv_filename}'.")
    except KeyError:
        print("Error: The column 'cluster' does not exist in the CSV file.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage:
filter_csv_to_excel_sheets('all_inventory_list.csv', 'osFilter.xlsx')
er file contents here
