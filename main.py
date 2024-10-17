import os
import sys
from dotenv import load_dotenv
from processtable import process_client_table, backup_database, get_db_keys_from_env, get_database_name_from_path, save_and_format_dataframe_to_excel, find_code_matches_in_db, update_old_codes_in_db


# Load variables from the .env file
load_dotenv()

# Get the database path from the environment variable
access_db_0 = os.getenv('ACCESS_DB_PATH_0')
access_db_1 = os.getenv('ACCESS_DB_PATH_1')
client_table = os.getenv('CLIENT_TABLE')
client_key_field = os.getenv('CLIENT_KEY_FIELD')
db_dir_0 = os.path.dirname(access_db_0)
db_dir_1 = os.path.dirname(access_db_1)

# Backup database before process
backup_file_0 = backup_database(access_db_0)
backup_file_1 = backup_database(access_db_1)

# Process client table and obtain mapping codes
df_mappings = process_client_table(access_db_0, client_table, client_key_field, update_clients=True)
df_mappings_dict = {
    client_table: df_mappings
}
# Define the path to save the Excel file in the same directory as the database
excel_path = os.path.join(db_dir_0, f'{client_table}_mappings.xlsx')
# Call the function to save and format the mappings to Excel
save_and_format_dataframe_to_excel(df_mappings_dict, excel_path)

# Buscar las coincidencias en todas las tablas de la base de datos
database_name = get_database_name_from_path(access_db_0)
db_keys = get_db_keys_from_env(database_name)
df_matches_0 = find_code_matches_in_db(df_mappings, access_db_0, db_keys)

# # Define the path to save the Excel file in the same directory as the database
excel_path = os.path.join(db_dir_0, f'{database_name}_mappings.xlsx')
save_and_format_dataframe_to_excel(df_matches_0, excel_path)

# Actualizar valores en las tablas con matching
update_counts_0 = update_old_codes_in_db(df_matches_0, access_db_0, db_keys)
print(update_counts_0)

# # Buscar las coincidencias en todas las tablas de la base de datos
database_name = get_database_name_from_path(access_db_1)
db_keys = get_db_keys_from_env(database_name)
df_matches_1 = find_code_matches_in_db(df_mappings, access_db_1, db_keys)

# # Define the path to save the Excel file in the same directory as the database
excel_path = os.path.join(db_dir_1, f'{database_name}_mappings.xlsx')
save_and_format_dataframe_to_excel(df_matches_1, excel_path)

# Actualizar valores en las tablas con matching
update_counts_1 = update_old_codes_in_db(df_matches_1, access_db_1, db_keys)
print(update_counts_1)
