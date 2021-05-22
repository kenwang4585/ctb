import pandas as pd
from sqlalchemy import MetaData,create_engine
import os

def read_table(table_name,columns='*',show_last=False,criteria_string=None,records_limit=None):
    '''从表中按条件读取数据'''
    engine = create_engine(os.getenv('ENGINE'))
    if criteria_string==None and records_limit==None:
        sql = 'SELECT '+ columns +' FROM ' + table_name
    elif criteria_string!=None and records_limit==None:
        sql = 'SELECT '+ columns +' FROM ' + table_name + ' WHERE ' + criteria_string
    elif criteria_string==None and records_limit!=None:
        if show_last:
            sql='SELECT '+ columns +' FROM ' + table_name + ' ORDER BY id DESC LIMIT ' + records_limit
        else:
            sql = 'SELECT '+ columns +' FROM ' + table_name + ' LIMIT ' + records_limit
    else:
        sql = 'SELECT '+ columns +' FROM ' + table_name + ' WHERE ' + criteria_string + ' ORDER BY id DESC LIMIT ' + records_limit

    df=pd.read_sql(sql,engine)

    return df




if __name__=='__main__':

    #df = read_table('predicted_error_summary',"VALIDATION_RESULT='NOT_READY'")
    #df=read_table('collected_error_detail')
    #df.to_excel('collected error detail.xlsx', index=False)
    #df = read_table('module_slot_occupation')

    #df=read_table('predicted_error_detail')
    #df.to_excel('pred error detail.xlsx', index=False)
    #col=['PO_NUMBER','VALIDATION_RESULT']
    #print(df.shape)
    #print(df)
    '''
    start=pd.Timestamp.now()
    sql="select * from predicted_error_summary"
    df = pd.read_sql(sql, engine)

    df.to_excel('summary.xlsx')
    '''
    df = read_table('user_log')

    print(df)