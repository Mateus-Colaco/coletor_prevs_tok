import pandas as pd, os
from datetime import datetime


def df2file(df_in):
    
    # Get the current date and time
    now = datetime.now()
    # Format the date and time as yyyymmdd-hh_mm
    date_str = now.strftime("%Y%m%d-%H_%M")
    # Create the string with the desired format
    filename = f'ENAs_TOK_{date_str}.xlsx'
    arquivo_1 = os.path.join(os.getcwd(), filename)
    # Exportar para csv
    df_in.to_excel(arquivo_1, index=False)

    return f'Dataframe {df_in} salvo com sucesso.'


def ena_sin(df_in):
    
    #ENA SIN

    # Columns to aggregate (all columns after 'Subsistema')
    columns_to_sum = df_in.columns.difference(['Modelo', 'Grandeza', 'Subsistema'])

    # Create the aggregation dictionary
    agg_dict = {col: 'sum' for col in columns_to_sum}

    # Using .agg function for aggregation
    df_sin = df_in.groupby('Modelo').agg(agg_dict).reset_index()

    # Insert the 'Grandeza' column at the second position (index 1)
    df_sin.insert(1, 'Grandeza', 'ENA MWm')
    df_sin.insert(2, 'Subsistema', 'SIN')

    # Concatenate dfs
    df_in = pd.concat([df_in, df_sin], axis=0, sort=False)


    return df_in


def ler_enas():

    # Define the columns
    colunas = ['Subsistema', 'ENA %', 'ENA MWm', 'Montador', 'Modelo']
    # Create an empty DataFrame
    df = pd.DataFrame(columns=colunas)

    for root, dirs, files in os.walk(os.getcwd()):       
        for file_name in files:
            if file_name.endswith("Montador.xlsx"):
                file_path = os.path.join(root, file_name)
                # Load the Excel file
                excel_data = pd.read_excel(file_path, sheet_name='REEs_MWm', engine='openpyxl', header=None)
                # Extract the specific range K6:M9 (zero-indexed, adjust accordingly)
                # K6 is [5, 10], M9 is [8, 12], note that pandas uses 0-based indexing
                cell_range = excel_data.iloc[5:9, 10:13]
                # Convert to a DataFrame
                df_aux = pd.DataFrame(cell_range)
                # Split the file name into name and extension
                nome_montador, extensao = os.path.splitext(file_name)
                df_aux['Montador'] = nome_montador
                df_aux['Modelo'] = file_path.split('\\')[-2]
                df_aux.reset_index(drop=True, inplace=True)
                df_aux.columns = colunas
                df_aux['ENA %'] = df_aux['ENA %'].apply(lambda x: f"{x*100:.2f}%")
                df_aux['ENA MWm'] = df_aux['ENA MWm'].round(0).astype(int)

                # Concatenate dfs
                df = pd.concat([df, df_aux], axis=0, sort=False)
                df = df[['Modelo', 'Montador', 'Subsistema', 'ENA %', 'ENA MWm']]
    
    df = reshape_df(df)
    df2file(df)


def ordena_subs(df_in):

    # Order list
    order = ['SE/CO', 'SUL', 'NORDESTE', 'NORTE']

    # Create a categorical type with the specified order and add 'SIN' at the end
    df_in['Subsistema'] = pd.Categorical(df_in['Subsistema'], categories=order + ['SIN'], ordered=True)

    # Sort the DataFrame based on the categorical column
    df_out = df_in.sort_values(['Modelo', 'Grandeza', 'Subsistema'])
    df_out.drop(columns=['index'], axis=1, inplace=True)
    
    return df_out


def transform_list(input_list):
    output_list = []
    for item in input_list:
        # Check if the item starts with 'ENA MWm_'
        if item.startswith('ENA'):
            # Split the string by '__' and extract the relevant parts
            parts = item.split('_')
            year = parts[1]
            month = parts[3]
            revision = parts[5]
            # Concatenate the month and revision parts
            transformed_item = f"{year}_{month}_{revision}"
            output_list.append(transformed_item)
        else:
            # If the item doesn't match the pattern, keep it unchanged
            output_list.append(item)
    return output_list


def reshape_df(df_in):
    # Pivot the DataFrame
    df_ENA_perc = df_in.pivot_table(index=['Modelo', 'Subsistema'], columns='Montador', values=['ENA %'], aggfunc='first')
    df_ENA_MWm = df_in.pivot_table(index=['Modelo', 'Subsistema'], columns='Montador', values=['ENA MWm'], aggfunc='first')

    # Reset the index and flatten the MultiIndex columns
    df_ENA_perc.columns = ['_'.join(col).strip() for col in df_ENA_perc.columns.values]
    df_ENA_perc.reset_index(inplace=True)
    df_ENA_MWm.columns = ['_'.join(col).strip() for col in df_ENA_MWm.columns.values]
    df_ENA_MWm.reset_index(inplace=True)


    # Insert the 'Grandeza' column at the second position (index 1)
    df_ENA_perc.insert(1, 'Grandeza', 'ENA %')
    df_ENA_MWm.insert(1, 'Grandeza', 'ENA MWm')

    #Change columns names before concatenation
    df_ENA_perc.columns = transform_list(df_ENA_perc.columns.values)
    df_ENA_MWm.columns = transform_list(df_ENA_MWm.columns.values)

    # Create a ENA SIN line
    df_ENA_MWm = ena_sin(df_ENA_MWm)
    
    # Concatenate dfs
    df_out = pd.concat([df_ENA_perc, df_ENA_MWm], axis=0)
    df_out.reset_index(inplace=True)

    # Ordenar linhas por submercado
    df_out = ordena_subs(df_out)

    return df_out

