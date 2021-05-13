from flask_settings import *
from sqlalchemy import create_engine
from setting import db_uri
import os

def create_table():
    engine = create_engine(db_uri)
    print('Existing tables:',engine.table_names())

    db.create_all()

    print('Current tables:',engine.table_names())

if __name__=='__main__':
    create_table()