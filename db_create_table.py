from flask_setting import *
from sqlalchemy import create_engine
import os

def create_table():
    engine = create_engine(os.getenv('DB_URI'))
    print('Existing tables:',engine.table_names())

    db.create_all()

    print('Current tables:',engine.table_names())

if __name__=='__main__':
    create_table()