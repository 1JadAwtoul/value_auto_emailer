import sqlite3

conn = sqlite3.connect('sent_invoices.db')

cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE sent_invoices (
               invoice_id TEXT PRIMARY KEY
        )
    '''
)

conn.commit()
conn.close()


