from flask_settings import *
from sqlalchemy import create_engine
import os
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import MetaData
from setting import db_uri

def drop_table():
   #engine = create_engine(db_uri)
   engine = create_engine("mysql+pymysql://kw:jackfruits4585@10.123.219.20:3306/kw_db")

   print('Existing tables:', engine.table_names())

   table_name = input('Input table name to drop:')

   base = declarative_base()
   metadata = MetaData(engine, reflect=True)
   table = metadata.tables.get(table_name)
   if table is not None:
       base.metadata.drop_all(engine, [table], checkfirst=True)

   engine = create_engine(db_uri)
   print('Current tables:', engine.table_names())


if __name__=='__main__':
    drop_table()
