import shutil
import os
from datetime import datetime
import sqlalchemy as sa
import sqlalchemy_access as sa_a
import pandas as pd
import re


def backup_database(db_path: str, backup_dir: str = None) -> str:
    """
    Creates a backup copy of the Access database file before any modifications are made.

    Args:
        db_path (str): The full path to the Access database file (.mdb or .accdb).
        backup_dir (str, optional): The directory where the backup will be stored. 
                                    If None, the backup will be stored in the same directory as the original file.
    
    Returns:
        str: The full path to the backup file.
    
    Raises:
        FileNotFoundError: If the original database file does not exist.
    
    Example:
        backup_file = backup_database('C:/path/to/database.mdb')
    """
    
    # Check if the database file exists
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"The database file {db_path} does not exist.")
    
    # Get the current timestamp to append to the backup file name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Get the directory and file name of the database
    db_dir, db_name = os.path.split(db_path)
    db_base, db_ext = os.path.splitext(db_name)
    
    # If no backup directory is provided, use the same directory as the database
    if backup_dir is None:
        backup_dir = db_dir
    
    # Create the backup file name with the timestamp
    backup_filename = f"{db_base}_backup_{timestamp}{db_ext}"
    backup_path = os.path.join(backup_dir, backup_filename)
    
    # Perform the file copy to create the backup
    shutil.copy2(db_path, backup_path)
    
    return backup_path


def get_dbaccess_connection(db_path: str):
    """
    Establishes a connection to a Microsoft Access database using SQLAlchemy 
    and returns a connection engine.

    Args:
        db_path (str): The full path to the Access database file (.mdb or .accdb).

    Returns:
        sqlalchemy.engine.Engine: An SQLAlchemy engine object representing the connection 
        to the Access database.

    Details:
        - Extracts the database name from the provided `db_path`.
        - Uses the ODBC driver "{Microsoft Access Driver (*.mdb, *.accdb)}" to establish the connection.
        - Configures the connection with the "ExtendedAnsiSQL=1" parameter for better ANSI compatibility.
        - Leverages `sqlalchemy_access` and `pyodbc` for SQLAlchemy integration.

    Example:
        engine = get_dbaccess_connection('C:/path/to/your_database.mdb')
    
    """
    # Extract the database name from the db_path
    db_name = os.path.splitext(os.path.basename(db_path))[0]

    #
    # Create the SQLAlchemy engine with an ODBC connection string
    # Ordinary unprotected Access database
    #  
    driver = "{Microsoft Access Driver (*.mdb, *.accdb)}"
    connection_string = (
        f"DRIVER={driver};"
        f"DBQ={db_path};"
        f"ExtendedAnsiSQL=1;"
        )
    connection_url = sa.engine.URL.create(
        "access+pyodbc",
        query={"odbc_connect": connection_string}
        )
    engine = sa.create_engine(connection_url)

    return engine


def has_trailing_letter(code) -> bool:
    """
    Checks if a given string has a letter (A-Z or a-z) at the end.

    Args:
        code (str): The string to be checked.

    Returns:
        bool: True if the string ends with a letter, False otherwise.

    Example:
        has_trailing_letter("12345A")  # Returns True
        has_trailing_letter("123456")  # Returns False
    """
    return bool(re.match(r'.*[A-Za-z]$', code))


def generate_unique_code(old_code, existing_codes) -> str:
    """
    Generates a new unique code by removing the last letter from the old code 
    and prefixing it with a number (starting with '9'). If the generated code 
    already exists, it tries successive prefixes ('8', '7', etc.) until a unique 
    code is found.

    Args:
        old_code (str): The original code containing a trailing letter.
        existing_codes (set): A set of codes that are already in use.

    Returns:
        str: A new unique code that is not present in `existing_codes`.

    Raises:
        ValueError: If a unique code could not be generated after trying all prefixes.

    Example:
        generate_unique_code("12345A", {"912345", "812345"})  # Returns a unique code like "712345"
    """
    base_code = old_code[:-1]  # Remove the trailing letter
    for prefix in ['9', '8', '7', '6', '5', '4', '3', '2', '1', '0']:
        new_code = prefix + base_code
        if new_code not in existing_codes:
            return new_code
    raise ValueError("Could not generate a unique code")


def process_client_table(db_path: str, client_table: str, client_key_field: str):

    engine = get_dbaccess_connection(db_path)
    
    query = f"SELECT * FROM {client_table}"
    df_clients = pd.read_sql(query, engine)

    # Filtrar los códigos erróneos
    df_clients['Has_Letter'] = df_clients[client_key_field].apply(has_trailing_letter)
    df_erroneous = df_clients[df_clients['Has_Letter'] == True]

    # Crear la tabla de mapeo
    df_mappings = pd.DataFrame(columns=['OLD_CODE', 'NEW_CODE'])

    existing_codes = set(df_clients[client_key_field])  # Códigos existentes

    # Generar códigos nuevos
    for idx, row in df_erroneous.iterrows():
        old_code = row[client_key_field]
        new_code = generate_unique_code(old_code, existing_codes)
        existing_codes.add(new_code)  # Add new code to the existing set

        # Create a new DataFrame from the row to append
        new_row = pd.DataFrame([{'OLD_CODE': old_code, 'NEW_CODE': new_code}])
        
        # Use pd.concat instead of append
        df_mappings = pd.concat([df_mappings, new_row], ignore_index=True)
        
    # Mostrar los mapeos
    print(df_mappings)

