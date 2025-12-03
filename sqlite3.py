import sqlite3  # no installation needed

# Connect to a database (creates a file if not exists)
conn = sqlite3.connect("sa_data.db")  # sa_data.db is your database file

# Create a cursor (to execute SQL commands)
cursor = conn.cursor()

# Create a table if it doesn’t exist
cursor.execute("""
CREATE TABLE IF NOT EXISTS Measurements (
    Timestamp TEXT,
    Machine TEXT,
    PartType TEXT,
    Hole TEXT,
    Feature TEXT,
    Value REAL,
    Status TEXT
)
""")

# Commit changes and close
conn.commit()
conn.close()
