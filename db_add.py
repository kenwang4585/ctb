# import cryptography

import pandas as pd
from flask_settings import *


def add_user_log(user='', location='', user_action='',summary=''):
    '''
    Add the user log to db
    '''

    log = UserLog(USER_NAME=user,
                    DATE=pd.Timestamp.now().date(),
                    TIME=pd.Timestamp.now().strftime('%H:%M:%S'),
                    LOCATION=location,
                    USER_ACTION=user_action,
                    SUMMARY=summary)


    db.session.add(log)  # can also use add_all() for multiple adding at one time
    db.session.commit()
    #print('User log added')

def add_data_from_file_initial():
    '''
    Add data from file: initial update
    '''
    fname='/users/wangken/downloads/program_lo (2020-01-17).xlsx'


    df = pd.read_excel(fname,parse_dates=['Date'])

    print(df.columns)
    print(df.shape)

    #print(df_summary.LINE_CREATION_DATE)
    df.loc[:, 'Date'] = df.Date.map(lambda x: x.date())
    df.fillna('',inplace=True)

    #df=df[['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PO_NUMBER','LINE_CREATION_DATE', 'OPTION_NUMBER',
    #   'PRODUCT_ID', 'ORDERED_QUANTITY','LABEL','COMMENTS','REPORT_DATE', 'UPLOAD_BY','ML_COLLECTED']]

    df_data = df.values

    db.session.bulk_insert_mappings(
                                    UserLog,
                                    [dict(
                                        USER_NAME=row[0],
                                        DATE=row[1],
                                        START_TIME=row[2],
                                        FINISH_TIME=row[3],
                                        PROCESSING_TIME=row[4],
                                        SELECTED_PROGRAMS=row[5],
                                        EMAIL_TO_ONLY=row[6],
                                        PROGRAM_LOG=row[7]
                                        )
                                     for row in df_data]
                                    )
    db.session.commit()
    print('data added')


def roll_back():
    try:
        db.session.commit()
    except:
        db.session.rollback()


if __name__ == '__main__':
    add_user_log(user='kwang2', location='Admin', user_action='Visit',
                 summary='Warning')