# !/usr/bin/python
# coding=utf-8
# Created by shilwu at 2019-06-21
# from: https://gist.github.com/tbuckl/3c1eeb56f46904dbb143ac398ea79b40
import smartsheet
import pandas as pd


class SmartSheetClient:

    def __init__(self, access_token=None, proxies=None):
        self.smartsheet_client = smartsheet.Smartsheet(access_token=access_token, proxies=proxies)

    def get_sheet_as_df(self, sheet_id=None, add_row_id=False, add_att_id=False):
        """
        download sheet and return as df
        :param sheet_id:
        :param add_row_id: add row id in DataFrame for further processing
        :param add_att_id: add attachment id in DataFrame for further processing
        :return:  pd.DataFrame
        """
        sheet = self.smartsheet_client.Sheets.get_sheet(sheet_id)
        df = self.get_values_as_df(sheet_id)
        s2 = self.get_columns(sheet_id)
        df.columns = s2
        if add_row_id:
            row_info = dict()
            _ = list(map(lambda x: row_info.update({x['rowNumber']-1: x['id']}), sheet.to_dict()['rows']))
            df['row_id'] = df.index.map(row_info)
        if add_att_id:
            att_id_list = list(map(lambda x: self.get_attachment_id(x, sheet_id=sheet_id), df['row_id']))
            df['attachment_id'] = att_id_list
        return df

    def get_attachment_id(self, row_id=None, sheet_id=None):
        response = self.smartsheet_client.Attachments.list_row_attachments(sheet_id, row_id)
        #print(sheet_id, row_id)
        attachment_id = response.data[0].id  # only one attachment allowed in such case
        return attachment_id
    
    def get_columns(self, sheet_id=None):
        """
        get column as df
        :param sheet_id:
        :return:
        """
        sheet = self.smartsheet_client.Sheets.get_sheet(sheet_id)
        cl = sheet.get_columns()
        d3 = cl.to_dict()
        df = pd.DataFrame(d3['data'])
        df = df.set_index('id')
        return df.title
        
    def get_values_as_df(self, sheet_id=None):
        """
        get all values as df from a sheet
        :param sheet_id:
        :return:
        """
        sheet = self.smartsheet_client.Sheets.get_sheet(sheet_id)
        drows = sheet.to_dict()['rows']
        rownumber = [x['rowNumber'] for x in drows]
        rows = [x['cells'] for x in drows]
        # print(rows)
        values = [[x.get('displayValue') for x in y] for y in rows]
        return pd.DataFrame(values)

    def delete_row(self,sheet_id=None,row_id=None):
        '''
        Delete rows from Smartsheet
        Added by Ken
        '''
        self.smartsheet_client.Sheets.delete_rows(sheet_id, row_id)
        
        
        
    def update_row_with_dict(self, process_type='update', ss=None,sheet_id=None, row_id=None, update_dict=None):
        """
        update one row with update_dict, key in update_dict much match with sheet
        :param ss: smartsheet.Smartsheet(token) - Ken modified
        :param process_type: 'update'/'add': for attachment sheet it's 'update', for summary sheet it's 'add' - Ken added
        :param sheet_id:
        :param row_id:
        :param update_dict: change it to a list with dictionary(s) inside (one dic refers to one row of info to update ) - Ken modified
        
        """
        
        #sheet = self.smartsheet_client.Sheets.get_sheet(sheet_id, page_size=0) #Ken: remove this line
        new_row = ss.models.Row() #ken: change smartsheet_client to ss(ss=smartsheet.Smartsheet(access_token)
        col_dict = self.get_columns(sheet_id)   #Ken: change parameter(sheet) to sheet_id
        col_id_value = {v: k for k, v in col_dict.iteritems()}
        
        for dic in update_dict:
            # Build the row to update
            #print(dic)
            new_row = smartsheet.models.Row()
            
            for key, value in dic.items():   
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = col_id_value.get(key)
                new_cell.value = value
                # new_cell.strict = False
    
                
                #new_row.id = 2226873662498692
                if process_type=='update':   #Ken add: only update row id when it's update (of the attachment sheet)
                    new_row.id = row_id
                new_row.cells.append(new_cell)
    
        
            # Update/add rows
            if process_type=='update':
                updated_row = ss.Sheets.update_rows(sheet_id, [new_row]) #Ken: change smartsheet_client to ss(ss=smartsheet.Smartsheet(access_token)
            elif process_type=='add':
                updated_row = ss.Sheets.add_rows(sheet_id, [new_row])
    
    
    def get_attachment_per_row_as_df(self, sheet_id=None, attachment_id=None, row_id=None, file_format='xls'):
        """
        get attachment for a row in sheet
        :param sheet_id:
        :param attachment_id:
        :param row_id:
        :param file_format:
        :return:
        """
        if attachment_id is None:
            if row_id is None:
                raise ValueError("None of attachment_id, row_id has been specified")
            response = self.smartsheet_client.Attachments.list_row_attachments(sheet_id, row_id)
            attachment_id = response.data[0].id  # only one attachment allowed in such case
        attachment = self.smartsheet_client.Attachments.get_attachment(sheet_id, attachment_id)  #
        if 'xls' in file_format or file_format == 'Excel':
            attach_df = pd.read_excel(attachment.url)
        elif file_format == 'csv':
            attach_df = pd.read_csv(attachment.url)
        else:
            raise ValueError("format {} is not supported yet !".format(file_format))
        return attach_df
    

if __name__ == '__main__':
    #input token for smartsheet
    #token = 'TEST'
    proxies = None  # for proxy server
    smartsheet_client = SmartSheetClient(token, proxies)
    summary_sheet_id = 1955251633842052
    details_sheet_id = 503861925439364
    sample_row_id = 368999927703428
    # att_df = get_attachment_per_row(smartsheet_client=smartsheet_client,
    #                                 sheet_id=details_sheet_id, row_id=sample_row_id)
    # update_row_with_dict(smartsheet_client, details_sheet_id, row_id, update_dict={'processed': 5})
    sheet_df = smartsheet_client.get_sheet_as_df(details_sheet_id, add_row_id=True, add_att_id=True)
    print(sheet_df)



