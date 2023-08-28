import pandas as pd
import re
from openpyxl.utils import get_column_letter
from openpyxl import Workbook
from tqdm import tqdm

def separate_words_numbers_emails(text):
    """
    Function to separate words, numbers, and emails within the text and ensure they are on the same line.
    Args:
        text (str): The input text.
    Returns:
        str: The modified text with words, numbers, and emails separated by spaces.
    """
    words = []
    numbers = []
    emails = []
    # Separate words, numbers, and emails using regex patterns
    words_pattern = r'\b[A-Za-z]+\b'
    numbers_pattern = r'\b\d+\b'
    emails_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b'
    # Find words, numbers, and emails in the text
    words = re.findall(words_pattern, text)
    numbers = re.findall(numbers_pattern, text)
    emails = re.findall(emails_pattern, text)
    # Join the separated words, numbers, and emails with spaces
    separated_text = ' '.join(words + numbers + emails)
    return separated_text

def highlight_matched_rows(row):
    """
    Function to apply background color to the entire row based on the 'Presence' values.
    Args:
        row (pandas.Series): A row in the DataFrame.
    Returns:
        list: A list of CSS background color styles for each cell in the row.
    """
    if row['Presence'] == 'Yes':
        return ['background-color: #90EE90'] * len(row)  # Light green
    elif row['Presence'] == 'No':
        return ['background-color: #FF9999'] * len(row)  # Light red
    return [''] * len(row)

def main():
    """
    Main function to detect presence based on the 'MRS Long Text' column.
    """
    print('Script initiating...')
    print('Reading data from Excel file.')
    # Read data from Excel file
    df_sheet1 = pd.read_excel('C:\\Users\\brduarte\\Documents\\Python Scripts\\Check of MRS long text.xlsx', sheet_name='Sheet1')
    df_sheet2 = pd.read_excel('C:\\Users\\brduarte\\Documents\\Python Scripts\\Check of MRS long text.xlsx', sheet_name='Sheet2')
    print('Extracting relevant columns.')
    # Extract relevant columns from Sheet2
    names_ids_emails = df_sheet2[['NOKIA ID', 'Email Address', 'NameFirstLast', 'NameLastFirst']]
    # Initialize a new column 'Presence' in Sheet1 with default values as 'No'
    df_sheet1['Presence'] = 'No'
    print('Iterating through DataFrame Index.')
    print('Applying filters to DataFrame Index.')
    print('Updating Presence column for matched rows... This may take a while...')
    # Iterate through each row in Sheet2 with a loading bar
    for index, row in tqdm(names_ids_emails.iterrows(), total=names_ids_emails.shape[0], desc='Processing rows'):
        nokiaid = row['NOKIA ID']
        email = row['Email Address']
        namefirstlast = row['NameFirstLast']
        namelastfirst = row['NameLastFirst']
        # Check if any of the elements is present in Sheet1 column 'MRS Long Text'
        presence_filter = (
            df_sheet1['MRS Long Text'].astype(str).apply(separate_words_numbers_emails)
            .apply(lambda x: any([
                re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', x),  # Matches emails
                str(nokiaid) in x,  # Matches Nokia ID
                any(re.search(r'\b\d{%d}\b' % len(str(nokiaid)), num) for num in re.findall(r'\b\d+\b', x)),  # Matches numbers with the same number of digits as Nokia ID
                re.search(r'\b%s\b' % re.escape(namefirstlast), x, flags=re.IGNORECASE),  # Matches First Name (case-insensitive)
                re.search(r'\b%s\b' % re.escape(namelastfirst), x, flags=re.IGNORECASE)  # Matches Last Name (case-insensitive)
            ]))
        )

        # Update 'Presence' column in Sheet1 for matched rows
        df_sheet1.loc[presence_filter, 'Presence'] = 'Yes'
    print('Styling DataFrame and saving to Excel file.')
    # Create a DataFrame with styled cells
    df_sheet1_styled = df_sheet1.style.apply(highlight_matched_rows, axis=1).set_properties(
        subset=pd.IndexSlice[:, 'Presence'], **{'font-weight': 'bold'}
    )
    # Save the styled DataFrame to a new Excel file with auto-fit column width
    with pd.ExcelWriter('CheckMRS_UpdatedFile.xlsx', engine='openpyxl') as writer:
        df_sheet1_styled.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        # Auto-fit column width for each column
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except TypeError:
                    pass
            adjusted_width = (max_length + 2) * 1.2  # Adding extra padding for better readability
            worksheet.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width
        writer._save()
    print('CheckMRS_UpdatedFile.xlsx file created successfully.')
    print('Script completed.')


if __name__ == '__main__':
    main()
