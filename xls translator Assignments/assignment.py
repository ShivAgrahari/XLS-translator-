import pandas as pd
from deep_translator import GoogleTranslator
import time
from concurrent.futures import ThreadPoolExecutor

# Function to translate a given text using Google Translator
def translate_text(text):
    if pd.isna(text) or not isinstance(text, str):
        return text
    try:
        translation = GoogleTranslator(source='auto', target='en').translate(text)
        return translation
    except Exception as e:
        print(f"Error translating text: {e}")
        return None

# Function to translate a single row in parallel
def translate_row_parallel(row):
    with ThreadPoolExecutor() as executor:
        translated_row = list(executor.map(translate_text, row))
    return translated_row

# Function to translate the entire dataframe, including headers
def translate_dataframe(dataframe):
    # Translate the headers
    translated_headers = translate_row_parallel(dataframe.columns)
    
    # Translate the entire dataframe in parallel
    with ThreadPoolExecutor() as executor:
        translated_dataframe = pd.DataFrame(list(executor.map(translate_row_parallel, dataframe.values)),
                                            columns=translated_headers)
    
    return translated_dataframe

# Load the Excel file into a Pandas DataFrame
file_path = 'Order Export.xls'
df = pd.read_excel(file_path)

# Record the start time
start_time = time.time()

# Translate the content of the DataFrame, including headers, using parallel processing
translated_df = translate_dataframe(df)

# Save the translated DataFrame to a new Excel file in XLS format using openpyxl
output_file_path = 'translated_order_export.xls'
translated_df.to_excel(output_file_path, index=False, engine='openpyxl')

# Record the end time
end_time = time.time()

# Calculate and print the elapsed time
elapsed_time = end_time - start_time
print(f"Translation and saving completed. Elapsed time: {elapsed_time:.2f} seconds. Translated data saved to: {output_file_path}")
