import pyodbc
from datetime import datetime
import configparser
import os

def check_tasks():
    try:
        config = configparser.ConfigParser()
        config.read('config.ini')
        
        server = config['Database']['Server']
        database = config['Database']['Database']
        username = config['Database']['Username']
        password = config['Database']['Password']
        
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        print(f"Connected to {server}/{database}")
        
        # 1. Check all active tasks with due dates
        print("\n--- ALL ACTIVE TASKS WITH DUE DATES ---")
        cursor.execute("""
            SELECT TOP 20 id, title, due_date, completed, deleted
            FROM Tasks 
            WHERE completed = 0 AND deleted = 0 AND due_date IS NOT NULL
            ORDER BY due_date DESC
        """)
        for row in cursor.fetchall():
            print(f"ID: {row.id}, Title: {row.title}, Due: {row.due_date}, Compl: {row.completed}")

        # 2. Run the specific query used in the app
        print("\n--- RUNNING APP QUERY ---")
        limit = 15
        query = f"""
                SELECT TOP {limit} id, title, due_date, priority, assigned_to
                FROM Tasks 
                WHERE completed = 0 AND deleted = 0 
                  AND due_date IS NOT NULL
                  AND due_date >= CAST(GETDATE() AS DATE)
                ORDER BY due_date ASC, 
                         CASE priority 
                             WHEN 'High' THEN 1 
                             WHEN 'Medium' THEN 2 
                             WHEN 'Low' THEN 3 
                         END
            """
        cursor.execute(query)
        rows = cursor.fetchall()
        print(f"Query returned {len(rows)} rows")
        for row in rows:
            print(f"ID: {row.id}, Title: {row.title}, Due: {row.due_date}")

        # 3. Check Date
        cursor.execute("SELECT GETDATE()")
        print(f"\nServer Date: {cursor.fetchone()[0]}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    check_tasks()
