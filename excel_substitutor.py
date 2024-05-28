import pandas as pd
import os

# Dictionary for substituting values in specific columns of different sheets
substitution_dicts = {
    'Стълб': {'Подтип': {0: 'Стълб - недиференцирано', 1: 'Дървен', 2: 'Композитен', 3: 'Стоманорешетъчен',
                         4: 'Стоманотръбен', 5: 'Стоманобетонен'}},
    'Оборудване на стълб': {'Вид оборудване': {1: 'Изолатор', 2: 'Конзола', 3: 'Обтяжка/подпора', 4: 'Заземителен контур'}},
    'РОМ - РОС': {'Експлоатационно състояние': {0: 'Изключено', 1: 'Включено', 2: 'НЯМА ИНФОРМАЦИЯ'},
                  'Вид': {0: 'Секциониращ разединител', 1: 'Товаров разединител с ДУ', 2: 'Реклоузер', 3: 'РОМ',
                          4: 'РОС'}},
    'Вентилен отвод': {'Фази': {0: 'НЯМА ИНФОРМАЦИЯ', 1: 'L3', 2: 'L2', 3: 'L2-L3', 4: 'L1', 5: 'L1-L3', 6: 'L1-L2',
                                7: 'L1-L2-L3'}}
}


# Function to add additional headers to the Excel file
def add_additional_headers(intermediate_result_path, result_path):
    print(f"Adding additional headers to: {intermediate_result_path}")
    xlsx = pd.ExcelFile(intermediate_result_path)
    copied_header_length = {
        'Стълб': {'start': 1, 'end': -3},
        'Оборудване на стълб': {'start': 1, 'end': None},
        'РОМ - РОС': {'start': 1, 'end': None},
        'Вентилен отвод': {'start': 1, 'end': None},
        'ИЗКС': {'start': 1, 'end': None},
        'Проводник': {'start': 1, 'end': None},
    }
    modified_df = {}   # Dictionary to hold modified DataFrames for each sheet
    headers_dict = {}  # Dictionary to hold headers for each sheet

    # Iterate through each sheet in the Excel file
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None) # Load each sheet into a DataFrame

        # Get start and end indices for the current sheet
        indices = copied_header_length.get(sheet_name, {'start': 0, 'end': None})    # defaults
        
        # Extract headers based on start and end indices
        headers = df.iloc[0, indices['start']:indices['end']].dropna().tolist() 
        
         # Store the headers in the dictionary
        headers_dict[sheet_name] = headers   

        # Add additional headers to the existing one
        additional_headers = headers + ['Забележка ЕРМЗ', 'Pontech'] * 3

        # Create a new DataFrame to store the additional headers
        new_headers = pd.Series(additional_headers)   

        # Concatenate additional headers to DataFrame
        all_headers_df = pd.concat([df, pd.DataFrame([new_headers])], axis=1)   

        # Add the modified DataFrame to the dictionary
        modified_df[sheet_name] = all_headers_df  

    # Write modified DataFrames to a new Excel file
    with pd.ExcelWriter(result_path, engine='openpyxl') as result:    # openpyxl is used for writing the .xlsx file; default: xlsxwriter
        for sheet, modified_df in modified_df.items():
            modified_df.to_excel(result, sheet_name=sheet, index=False, header=False)
    print(f"Additional headers added and saved to: {result_path}")
    return result_path


# Function to apply substitutions and create additional columns
def apply_substitutions(source_file_path, result_path, substitution_dicts):
    print(f"Applying substitutions to: {source_file_path}")
    xlsx = pd.ExcelFile(source_file_path)
    with pd.ExcelWriter(result_path, engine='openpyxl') as result:  # openpyxl is used for writing the .xlsx file; default: xlsxwriter
        for sheet in xlsx.sheet_names:
            df = pd.read_excel(source_file_path, sheet_name=sheet)
            if sheet == "Стълб":
                df = df.drop(columns=['Линк'])  # # Drop 'Линк' column if it exists, errors='ignore' --> not raise an error if a column ('Линк') does not exist in the DataFrame.

            # Apply substitutions based on the dictionary
            if sheet in substitution_dicts:
                for column, mapping_value in substitution_dicts[sheet].items():
                    df[column] = df[column].map(mapping_value).fillna(df[column])
            df.to_excel(result, index=False, sheet_name=sheet)

        # Create a 'Коментари в ГИС' sheet with specific columns
        comments_column = pd.DataFrame(columns=['OBJECTID *', 'Текст', 'Забележка ЕРМЗ',
                                                'Pontech', 'Забележка ЕРМЗ', 'Pontech',
                                                'Забележка ЕРМЗ', 'Pontech'])
        comments_column.to_excel(result, index=False, sheet_name='Коментари в ГИС')
    print(f"Substitutions applied and saved to: {result_path}")
    return result_path


# Main function to call the above functions in sequence
def main(source_file_path, result_path):
    intermediate_result_path = apply_substitutions(source_file_path, result_path, substitution_dicts)
    final_result_path = add_additional_headers(intermediate_result_path, result_path)
    print(f"Updated file saved to: {final_result_path}")


# If this script is run directly, execute the main function
if __name__ == "__main__":
    env = input("Enter the directory path where your files are located: ").strip('\"')    #env = r"D:\scripts\excel проверки"

    source_file_name = input("Enter the source Excel file name (e.g., 'source.xlsx'): ").strip('\"')    #source_file_path = os.path.join(env, "ВП1057-ВЕЦ Брусен-Искър-VR6401-20 kV-2.1.xlsx")
    result_file_name = input("Enter the result Excel file name (e.g., 'result.xlsx'): ").strip('\"')    #result_path = os.path.join(env, "ВП1057 Искър - проверка.xlsx")

    if not source_file_name.endswith('.xlsx'):
        source_file_name += '.xlsx'
    if not result_file_name.endswith('.xlsx'):
        result_file_name += '.xlsx'

    source_file_path = os.path.join(env, source_file_name)
    result_path = os.path.join(env, result_file_name)
    
    main(source_file_path, result_path)
