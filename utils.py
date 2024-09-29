import os
import sqlalchemy as sa
import sqlalchemy_access as sa_a


def get_dbaccess_connection(db_path:str):
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
