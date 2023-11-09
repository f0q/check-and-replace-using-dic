import pandas as pd
from fuzzywuzzy import process

def fill_values_from_dictionary(input_excel, output_excel, key_column_name, new_column_name, dictionary, not_found_text='No Match Found'):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_excel)

    # Function to look up the value in the dictionary based on fuzzy matching
    def lookup_value(cell_value):
        cell_value_lower = str(cell_value).lower()
        matching_keys = [key for key in dictionary.keys() if key.lower() in cell_value_lower]
        if matching_keys:
            best_match = max(matching_keys, key=len)
            return dictionary[best_match]
        else:
            return not_found_text

    # Create a new column based on the values in the 'KeyColumn'
    df[new_column_name] = df[key_column_name].apply(lookup_value)

    # Save the updated DataFrame to a new Excel file
    df.to_excel(output_excel, index=False)

if __name__ == '__main__':
    input_excel = 'input.xlsx'     # Replace with your input Excel file
    output_excel = 'output.xlsx'   # Replace with the desired output Excel file
    key_column_name = 'id_key'   # Replace with the name of the column to check against the dictionary
    new_column_name = 'cmlid_values1'   # Replace with the desired name for the new column
    my_dictionary = {
        '702': '631',
        '732': '635',
        '730': '666',
        '750': '695',
        '752': '696',
        '728': '697',
        '726': '699',
        '724': '701',
        '722': '703',
        '720': '704',
        '710': '716',
        '706': '720',
        '704': '722',
        '702M': '809',
        '704M': '810',
        '706M': '811',
        '792': '814',
        '794': '816',
        '796': '818',
        '798': '819',
        '800': '820',
        '780': '824',
        '782': '825',
        '784': '826',
        '786': '827',
        '788': '828',
        '730M': '1360',
        '732M': '1361',
        '750M': '1362',
        '752M': '1363',
        '764M': '1668',
        '762M': '1682',
        '768M': '1717',
        '760M': '1731',
        '736M': '1741',
        '736': '1751',
        '734': '1757',
        '732-s': '1762',
        '730-s': '1767',
        '7322': '1772',
        '7302': '1777',
        '7324': '1782',
        '7304': '1787',
        '810': '1801',
        '812': '1806',
        '814': '1811',
        '818': '1821',
        '822': '1826',
        '824': '1831',
        '826': '1836',
        '828': '1841',
        '802': '1850',
        '804': '1855',
        '808': '1865',
        '820': '1875',
        '852': '1880',
        '854': '1885',
        '856': '1890',
        '858': '1895',
        '772': '1906',
        '774': '1911',
        '776': '1916',
        '778': '1921',
        '770': '1929',
        '714': '1936',
        '716': '1943',
        '768': '1957',
        '760': '1965',
        '762': '1970',
        '764': '1975',
        '661': '2044',
        '663': '2047',
        '665': '2050',
        '830': '2057',
        '832': '2062',
        '714M': '2154',
        '716M': '2159',
        '718': '2187',
        '740': '2578',
        '742': '2587',
        '744': '2596',
        '714-s': '3049',
        '714-ss': '3063',
        '716-s': '3074',
        '716-ss': '3107',
        '7162': '3117',
        '7164': '3125',
        '7142': '3133',
        '7144': '3141',
        '710M': '3196',
        # Add more entries to the dictionary as needed
    }
    not_found_text = 'No Match Found'  # Replace with the desired text for no match found

    fill_values_from_dictionary(input_excel, output_excel, key_column_name, new_column_name, my_dictionary, not_found_text)
