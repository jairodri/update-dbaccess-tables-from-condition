import os
from dotenv import load_dotenv
from processtable import process_client_table, backup_database


# Load variables from the .env file
load_dotenv()

# Get the database path from the environment variable
access_db_0 = os.getenv('ACCESS_DB_PATH_0')
access_db_1 = os.getenv('ACCESS_DB_PATH_1')
client_table = os.getenv('CLIENT_TABLE')
client_key_field = os.getenv('CLIENT_KEY_FIELD')

backup_file = backup_database(access_db_0)
process_client_table(access_db_0, client_table, client_key_field)

