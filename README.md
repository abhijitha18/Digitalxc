This repository contains a Python script that extracts group names and their counts from a specified column in an Excel sheet. The script uses Pandas for reading Excel files and Regular Expressions (Regex) to identify and extract group names from the data.

Requirements
- Python 3.x
- pandas library for data manipulation and analysis
- openpyxl library for reading .xlsx files

Installation
1. Clone the repository (if applicable) or download the digitalxc.py script.
2. Create a virtual environment (optional but recommended):
       python -m venv venv
3. Activate the virtual environment:
     .\venv\Scripts\activate
For macOS/Linux:
     source venv/bin/activate

4. Install dependencies:

pip install pandas openpyxl


Usage
1. Ensure the Excel file you want to analyze is accessible and has the necessary column with group data.

2. In the script, adjust the following variables to match your use case:
    - file_path: Path to your Excel file.
    - sheet_name: Name of the sheet containing the relevant data (default is None, which loads the first sheet).
    - column_to_search: The column to extract group data from (e.g., "Additional comments").
    - keyword_to_search: The keyword to look for in the comments (e.g., "Groups").


Running the Script
Once the variables are set, run the script with:

python digitalxc.py


Output
The script will generate a file named group_counts.txt, containing the extracted group names and their occurrences.

Example Output:

Group_name               Number of occurrences
------------------------------------------------
Group A                  5
Group B                  3
Group C                  2



Code Explanation
extract_groups function:
- Reads the specified Excel file and sheet.
- Searches for comments containing the specified keyword (e.g., "Groups").
- Uses regular expressions to extract group names enclosed within specific markers (e.g., [code]<I>XXX</I>[/code]).
- Counts the occurrences of each group name and returns a dictionary with group names as keys and their counts as values.

save_to_file function:
- Saves the extracted group counts to a text file (group_counts.txt).

Troubleshooting
If you encounter the error ModuleNotFoundError: No module named 'pandas' or a similar error:
1. Ensure you've activated the virtual environment.
2. Install the missing libraries by running:
     pip install pandas openpyxl
