import pandas as pd

def extract_keywords_to_excel(input_file, output_file, sheet_name='Planilha1', column_name='Keywords'):
    """
    Extract unique keywords from a specified column in an Excel sheet and save them in a new Excel file.

    :param input_file: Path to the input Excel file
    :param output_file: Path to the output Excel file
    :param sheet_name: Name of the sheet in the Excel file
    :param column_name: Name of the column containing keywords
    """
    # Load the data from the specified sheet
    sheet_data = pd.read_excel(input_file, sheet_name=sheet_name)
    
    # Check if the column exists
    if column_name not in sheet_data.columns:
        raise ValueError(f"Column '{column_name}' not found in the sheet '{sheet_name}'.")
    
    # Extract keywords, split by delimiters (commas or semicolons), and normalize
    keywords_series = sheet_data[column_name].str.split(r'[;,]', expand=True).stack().str.strip()
    
    # Remove duplicates and sort the keywords
    unique_keywords = sorted(set(keywords_series))
    
    # Convert the list to a DataFrame
    keywords_df = pd.DataFrame(unique_keywords, columns=['Keyword'])
    
    # Save the DataFrame to a new Excel file
    keywords_df.to_excel(output_file, index=False)
    print(f"Keywords saved to '{output_file}'.")

# Example usage
input_file = 'keywords.xlsx'  # Replace with your input file path
output_file = 'output_keywords.xlsx'  # Replace with your desired output file path

try:
    extract_keywords_to_excel(input_file, output_file)
except Exception as e:
    print(f"Error: {e}")

