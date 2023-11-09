import pandas as pd

def fill_values_from_dictionary(input_excel, output_excel, key_column_name, new_column_name, dictionary, not_found_text='No Match Found'):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_excel)

    # Function to look up the value in the dictionary based on case-insensitive partial match
    def lookup_value(cell_value):
        cell_value_lower = str(cell_value).lower()
        for key, value in dictionary.items():
            if key.lower() in cell_value_lower:
                return value
        return not_found_text

    # Create a new column based on the values in the 'KeyColumn'
    df[new_column_name] = df[key_column_name].apply(lookup_value)

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_excel, index=False)

if __name__ == '__main__':
    input_excel = 'input.xlsx'     # Replace with your input Excel file
    output_excel = 'output.xlsx'   # Replace with the desired output Excel file
    key_column_name = 'color'   # Replace with the name of the column to check against the dictionary
    new_column_name = 'NewColumn'   # Replace with the desired name for the new column
    my_dictionary = {
        'Red (': 'Redcolor',
        'red flame': 'flamecolor',
        # Add more entries to the dictionary as needed
    }
    not_found_text = 'No Match Found'  # Replace with the desired text for no match found

    fill_values_from_dictionary(input_excel, output_excel, key_column_name, new_column_name, my_dictionary, not_found_text)
