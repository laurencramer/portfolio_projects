import pandas as pd 
import re
from datetime import datetime
import os
import csv



### Helper function to parse and sort data from an Excel file; returns a sorted list
def get_sorted_column_values(file_path, sheet_name='[REDACTED]', column_name='AD'):
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    if column_name not in df.columns:
        raise ValueError(f"Column '{column_name}' not found in sheet '{sheet_name}'.")

    return sorted(df[column_name].dropna())


### Helper function to extract table data from Excel; returns a 2D list of values
def extract_table_data(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')

    if df.empty:
        raise ValueError(f"No data found in sheet '{sheet_name}'.")

    return df.values.tolist()


### Helper function to extract substring
def extract_substring(text):
    text = text.upper()
    first_index = text.find('DPC')
    second_index = text.find('LPT')
    
    # Return substring starting from index, if found
    if first_index != -1:
        return text[first_index:]
    elif second_index != -1:
        return text[second_index:]
    
    # Return fallback message if neither substring is found
    return "No Substring Found"


### Helper function to check if extracted substring is in the data and return the data from the second column
def find_value_in_data(extracted_value, data_array, return_column=2, compare_column=0):
    for row in data_array:
        if extract_substring.lower() == str(row[compare_column]).lower():  # Compare with the specified column
            return row[return_column]  # Return value from the specified column
    return "No Match Found"


### Helper function to get predefined barcode ranges for different categories
def get_category_ranges():
    """Returns predefined barcode ranges for different categories."""
    return [
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"},
        {"start": "[REDACTED]", "end": "[REDACTED]", "name": "[REDACTED]"}
    ]   


### Helper function to check if the provided data has a valid structure
def valid_data_structure(items):
    """Checks if the provided data has a valid structure."""
    return isinstance(items, list) and all(len(item) >= 3 for item in items)


### Helper function to format an item for output
def format_item(item):
    """Formats a single item for output."""
    return f"{item[0]} ({item[1]}) [{item[2]}]"


### Helper function to map each item key to its associated identifier
def map_item_to_identifier(item_key, identifier, item_to_identifiers_map):
    """Maps each item key to its associated identifier."""
    if item_key in item_to_identifiers_map:
        item_to_identifiers_map[item_key].add(identifier)
    else:
        item_to_identifiers_map[item_key] = {identifier}


### Helper function to categorize an item based on its model and add it to the appropriate category
def categorize_item(item, formatted_entry, category_dict, category_ranges):
    """Categorizes an item based on its model and adds it to the appropriate category."""
    model = item[2].lower()

    # Category 1: Match certain keywords or barcode ranges
    if any(keyword in model for keyword in ["[REDACTED]", "[REDACTED]", "[REDACTED]"]) or in_category_ranges(item[0], category_ranges):
        category_dict["category1"].append(f"{formatted_entry} *Manual Add*")

    # Category 2: Conditions based on the model
    if "[REDACTED]" in model and "[REDACTED]" not in item[0].lower():
        category_dict["category2"].append(f"{formatted_entry} *Manual Add*")

    # Category 3: Laptops
    if "[REDACTED]" in item[0].lower():
        category_dict["category3"].append(formatted_entry)


### Helper function to check if an item falls within any of the predefined barcode ranges
def in_category_ranges(item_key, category_ranges):
    """Checks if an item falls within any of the predefined barcode ranges."""
    last_six_str = item_key[-6:]
    if last_six_str.isdigit():
        last_six = int(last_six_str)
        for category in category_ranges:
            if category["start"] <= last_six <= category["end"]:
                return True
    return False


### Helper function to filter items by the most recent year found in timestamps
def filter_by_recent_year(category_dict, categories):
    """Filters items in categories based on the most recent year found in timestamps."""
    for category in categories:
        entries = category_dict[category]
        years = extract_years_from_entries(entries)
        if years:
            most_recent_year = max(years)
            category_dict[category] = [entry for entry in entries if str(most_recent_year) in entry]


### Helper function to extract years from timestamps
def extract_years_from_entries(entries):
    """Extracts years from timestamps found in a list of entries."""
    years = []
    for entry in entries:
        match = re.search(r'\((.*?)\)', entry)
        if match:
            try:
                date = datetime.strptime(match.group(1), '%Y-%m-%d %H:%M:%S')
                years.append(date.year)
            except ValueError:
                continue
    return years

### Helper function to mark items as duplicates
def mark_duplicates(item_dict, item_to_identifiers_map):
    """Marks items that are associated with more than one identifier as duplicates."""
    for identifier, categories in item_dict.items():
        for category in categories:
            for i, entry in enumerate(categories[category]):
                item_key = entry.split(' ')[0]
                if item_key in item_to_identifiers_map and len(item_to_identifiers_map[item_key]) > 1:
                    categories[category][i] = f"{entry} DUPLICATE"


### Helper function to prepare the final Excel output
def prepare_output(item_dict):
    """Prepares the final output by formatting each identifier and its categorized items."""
    return [[identifier, 
             ";\n ".join(entries["original"]), 
             ";\n ".join(entries["category1"]), 
             ";\n ".join(entries["category2"]), 
             ";\n ".join(entries["category3"])]
            for identifier, entries in sorted(item_dict.items())]


def find_matching_items_with_attributes(sorted_items, primary_data, secondary_data):
    """
    This function finds matching items in the primary data for each sorted item and checks them against secondary data.

    Parameters:
    sorted_items (list): A list of items to search for.
    primary_data (list): A 2D list containing primary data. Each row represents an entity with attributes like item name, key, and date.
    secondary_data (list): A 2D list containing secondary data. Each row represents an entity with attributes like key and additional info.

    The function searches the primary data for each item, extracts relevant attributes, and matches them against the secondary data.
    It keeps track of unique keys, updating attributes if newer data is found. If no matches are found, it appends a default entry.

    Returns:
    all_results (list): A list of lists, where each inner list contains an item and its matching attributes.
    """
    all_results = []

    for item in sorted_items:
        item_results = [item]
        attribute_dict = {}
        search_key = f"\\{item}".lower()

        for row in primary_data:
            row_key1 = str(row[1]).lower()
            row_key2 = str(row[2]).lower()

            # Search for the key in the primary data
            if search_key in row_key1 or search_key in row_key2:
                unique_key = row[0]
                date_value = row[4]

                # Extract substring from the key
                extracted_substring = extract_substring(unique_key)

                # Find corresponding value in secondary data
                additional_info = find_value_in_data(extracted_substring, secondary_data)


                # Update dictionary with latest date for each unique key
                if unique_key in attribute_dict:
                    existing_date = attribute_dict[unique_key][1]
                    if date_value > existing_date:
                        attribute_dict[unique_key] = [unique_key, date_value, additional_info]
                else:
                    attribute_dict[unique_key] = [unique_key, date_value, additional_info]


        # Remove duplicates by keeping the latest date for each key
        final_attributes = {}
        for key, value in attribute_dict.items():
            unique_key, date_value, info_value = value
            if key in final_attributes:
                if date_value > final_attributes[key][1]:
                    final_attributes[key] = [unique_key, date_value, info_value]
            else:
                final_attributes[key] = [unique_key, date_value, info_value]


        # Keep only the latest entries based on the last 6 characters of the unique key
        filtered_attributes = {}
        for key, value in final_attributes.items():
            last_six = key[-6:]
            if last_six in filtered_attributes:
                if value[1] > filtered_attributes[last_six][1]:
                    filtered_attributes[last_six] = value
            else:
                filtered_attributes[last_six] = value

        # Convert final attributes to a list of lists
        final_list = list(filtered_attributes.values())

        # Append to item results or add a default entry if no matches are found
        if final_list:
            item_results.append(final_list)
        else:
            item_results.append([["No Device Found", "01/01/1900", "No Info Found"]])

        all_results.append(item_results)

    return all_results


def format_for_excel(data_with_items):
    """
    The 'format_for_excel' function processes a list of tuples where each tuple contains an identifier and a list of items. 
    It categorizes these items into different groups based on specific criteria and returns them in a formatted structure.

    Parameters:
    data_with_items (list): A list of tuples. Each tuple contains an identifier (str) and a list of items. 
    Each item is a list with a name (str), a timestamp (str), and a category (str).

    Returns:
    list: A sorted list of lists where each inner list contains the identifier and semicolon-separated categorized items.
    """

    item_dict = {}
    item_to_identifiers_map = {}
    category_ranges = get_category_ranges()

    for entry in data_with_items:
        identifier, items = entry[0], entry[1]

        if not valid_data_structure(items):
            print(f"Unexpected format in items: {items}")
            continue

        # Initialize dictionary for each identifier
        item_dict.setdefault(identifier, {"original": [], "category1": [], "category2": [], "category3": []})

        for item in items:
            item_key = item[0]
            formatted_entry = format_item(item)

            # Track which identifiers each item is associated with
            map_item_to_identifier(item_key, identifier, item_to_identifiers_map)

            # Append to original list
            item_dict[identifier]["original"].append(formatted_entry)

            # Categorize and append items
            categorize_item(item, formatted_entry, item_dict[identifier], category_ranges)

        # Filter categories by the most recent year
        filter_by_recent_year(item_dict[identifier], ["category1", "category2", "category3"])

    mark_duplicates(item_dict, item_to_identifiers_map)

    # Prepare and return final output
    return prepare_output(item_dict)


def write_to_csv(data, file_path):
    """
    Writes formatted data to a CSV file with specific columns.

    Parameters:
    data (list of lists): Formatted data where each inner list contains values for each column.
    file_path (str): Path to the output CSV file.

    Returns:
    None
    """
    headers = ['Identifier', 'All Items', 'Category1', 'Category2', 'Category3']
    
    with open(file_path, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # Write the header row
        writer.writerow(headers)
        # Write the data rows
        writer.writerows(data)


def get_file_paths(location):
    """Returns a dictionary of file paths based on the location."""
    if location == 0:
        return {
            'location1': r"[REDACTED]",
            'location2': r"[REDACTED]",
            'location3': r"[REDACTED]"
        }
    elif location == 1:
        return {
            'location1': r"[REDACTED]",
            'location2': r"[REDACTED]",
            'location3': r"[REDACTED]"
        }
    else:
        print("Invalid location")
        return None


def main():
    """
    Main function to orchestrate the process of extracting data from Excel files,
    processing it, and writing the results to a CSV file. 

    The function performs the following steps:
    1. Prompts the user for their name and location to determine the appropriate file paths.
    2. Extracts data from the specified Excel files using helper functions.
    3. Finds matching vaalues and models from the extracted data.
    4. Formats the data for output.
    5. Writes the formatted data to a CSV file.
    
    It prints status updates to the console and handles invalid user input or location selections.
    """
    user_name = input("Enter your name: ").strip().lower()
    if user_name == "lauren":
        location = int(input("Enter 0 for work, 1 for home: "))
        file_paths = get_file_paths(location)
    else:
        print("Invalid user")
        return

    print(f"\nRunning!\nTime Start: {datetime.now().strftime('%H:%M:%S')}")

    # Extract data from the Excel files
    print("- Extracting data from Excel files...")
    sorted_values_location1 = get_sorted_column_values(file_paths['location1'])
    extracted_data_location2 = extract_table_data(file_paths['location2'], '[REDACTED]')
    extracted_data_location3 = extract_table_data(file_paths['location3'], '[REDACTED]')

    # Find matching value and models
    print("- Finding matching values and models...")
    matching_items = find_matching_items_with_attributes(sorted_values_location1, extracted_data_location2, extracted_data_location3)

    print("- Formatting data...")
    formatted_data = format_for_excel(matching_items)

    # Write the results to a CSV file
    print("- Writing data to CSV file...\n")
    current_directory = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(current_directory, "output.csv")
    write_to_csv(formatted_data, output_path)
    print(f"*** Data has been written to: {output_path} ***\n")
    print(f"Time End: {datetime.now().strftime('%H:%M:%S')}")


if __name__ == '__main__':
    main()