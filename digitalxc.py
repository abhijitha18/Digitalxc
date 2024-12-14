import pandas as pd
import re
from collections import Counter


def extract_groups(file_path, sheet_name=None, column_name="Additional comments", keyword="Groups"):
    """
    Extracts group names and their occurrence counts from the specified column in an Excel sheet.

    :param file_path: Path to the Excel file.
    :param sheet_name: Optional. Name of the sheet to read. If None, reads the first sheet.
    :param column_name: The column to search for group strings.
    :param keyword: The keyword to identify the group information.
    :return: Dictionary with group names and their counts.
    """
    # Load the Excel file
    if sheet_name:
        data = pd.read_excel(file_path, sheet_name=sheet_name)  # Load specific sheet
    else:
        data = pd.read_excel(file_path)  # Load the first sheet by default
    
    if not isinstance(data, pd.DataFrame):
        raise ValueError(f"Expected a DataFrame but got {type(data)}. Check the 'sheet_name' parameter.")

    if column_name not in data.columns:
        raise ValueError(f"Column '{column_name}' not found in the sheet.")
    
    # Extract relevant column
    comments = data[column_name].dropna()
    
    group_pattern = rf"{keyword} : \[code\]<I>(.*?)</I>\[/code\]"
    all_groups = []
    
    # Process each comment to extract groups
    for comment in comments:
        matches = re.findall(group_pattern, comment, re.IGNORECASE)
        for match in matches:
            # Split groups by commas and strip whitespace
            groups = [group.strip() for group in match.split(",")]
            all_groups.extend(groups)
    
    # Count occurrences of each group
    group_counts = Counter(all_groups)
    return group_counts

def save_to_file(group_counts, output_file="group_counts.txt"):
    """
    Saves the group counts to a text file in the required format.

    :param group_counts: Dictionary with group names and their counts.
    :param output_file: File path to save the output.
    """
    with open(output_file, "w") as file:
        file.write("Group_name\t\t\tNumber of occurrences\n")
        for group, count in group_counts.items():
            file.write(f"{group}\t\t\t{count}\n")
    print(f"Output saved to {output_file}")

# Example usage
if __name__ == "__main__":
    file_path = r"C:\Users\aksha\Downloads\coding challenge test.xlsx"  # Replace with the actual file path
    sheet_name = "Input Data sheet"
    column_to_search = "Additional comments"  # Change to "Comments and Worknotes" if needed
    keyword_to_search = "Groups"
    
    group_data = extract_groups(file_path,sheet_name=sheet_name, column_name=column_to_search, keyword=keyword_to_search)
    save_to_file(group_data)
