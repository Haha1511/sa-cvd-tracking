import sqlite3

# Connect to database (creates file if it doesn't exist)
conn = sqlite3.connect("sa_data.db")
cursor = conn.cursor()

# Create a table
cursor.execute("""
CREATE TABLE IF NOT EXISTS Measurements (
    Timestamp TEXT,
    MeasuredDate TEXT,
    Machine TEXT,
    PartType TEXT,
    Chamber TEXT,
    PieceID TEXT,
    PartFlow TEXT,
    BatchCleaning TEXT,
    Hole TEXT,
    Feature TEXT,
    Value REAL,
    Nominal REAL,
    LSL REAL,
    USL REAL,
    Status TEXT,
    Notes TEXT,
    ImagePath TEXT
)
""")

conn.commit()
conn.close()
print("Database and table created successfully!")
