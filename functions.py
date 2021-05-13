import matplotlib.pyplot as plt
from setting import *

#from send_sms import send_me_sms
import gc
import numpy as np
import re
import traceback
from smartsheet_handler import SmartSheetClient
from sending_email import *
import time

plt.rcParams.update({'figure.max_open_warning': 0})

def pick_out_zero_qty_order(df_3a4):
    """
    Pick out the error orders with 0 ordered qty
    :param df_3a4:
    :return:
    """

    df_3a4_zero_qty=df_3a4[df_3a4.ORDERED_QUANTITY==0].copy()
    df_3a4=df_3a4[df_3a4.ORDERED_QUANTITY>0].copy()

    return df_3a4,df_3a4_zero_qty

# define func for creating supply dic by tan&date
def created_supply_dict_per_df_supply(df_supply):
    """
    create supply dict based on df_supply
    supply_dic_tan={'800-42373-01':[{'2/10':25},{'2/12':4},{'2/15':10},{'2/22':20},{'3/1':10},{'3/5':15}],
               '800-42925-01':[{'2/12':4},{'2/13':3},{'2/15':12},{'2/23':25},{'3/1':8},{'3/6':10}]}
    """
    #print(df_supply.loc['10-2366',:])
    supply_dic_tan = {}
    col=df_supply.columns
    for tan in df_supply.index:
        date_qty_list = []
        for date in col:
            qty = df_supply.loc[tan, date]
            if type(qty)==str:
                print(tan,date,qty)
            if qty > 0:

                date_qty = {date: df_supply.loc[tan, date]}
                date_qty_list.append(date_qty)

        supply_dic_tan[tan] = date_qty_list

    return supply_dic_tan


# 根据tan list和df_3a4生成blg_dic_tan
def create_blg_dict_per_sorted_3a4_and_selected_tan(df_3a4, supply_dic_tan):
    """
    create backlog dict for selected tan list from the sorted 3a4 df which have packed PO excluded (considered order prioity and
    rank - a "production sequence")
    tan_list: from a predefined tan list, or from tans included in the supply df - depends on how we define it
    """

    tan = list(supply_dic_tan.keys())
    dfx=df_3a4[(df_3a4.PACKOUT_QUANTITY!='Packout Completed')&(df_3a4.ADDRESSABLE_FLAG!='PO_CANCELLED')]
    blg_dic_tan = {}
    for pn in tan:
        dfm = dfx[dfx.BOM_PN == pn]
        org_qty_po = []
        for org, qty, po, min_date in zip(dfm.ORGANIZATION_CODE, dfm.C_UNSTAGED_QTY, dfm.PO_NUMBER,
                                                      dfm.min_date):
            if qty > 0:
                org_qty_po.append({org: (qty, (po, min_date))})

        blg_dic_tan[pn] = org_qty_po
        #print(blg_dic_tan)
    return blg_dic_tan


def allocate_supply_to_backlog_and_calculate_shortage(supply_dic_tan, blg_dic_tan):
    """
    allocate supply to each PO by TAN based on supply dict and backlog dict
    supply dict is arranged in date order; backlog dict is arranged based on priority to fulfill

    examples:
        blg_dic_tan={'800-42373-01': [{'FJZ': (5, ('110077267-1','2020-4-1'))},{'FJZ': (23, ('110011089-4','2020-4-4'))},...]}
        supply_dic_tan={'800-42373-01':[{'2/10':25},{'2/12':4},{'2/15':10},{'2/22':20},{'3/1':10},{'3/5':15}],
                             '800-42925-01':[{'2/12':4},{'2/13':3},{'2/15':12},{'2/23':25},{'3/1':8},{'3/6':10}]}
    return: blg_with_allocation contains {po_tan:([(date,qty),(date,qty)], last_supply_date, po_uncovered_qty, lt_shortage_qty}
    """
    blg_with_allocation = {}

    for tan in blg_dic_tan.keys():
        blg_list_tan = blg_dic_tan[tan]  # 每一个tan对应的blg list

        if tan in supply_dic_tan.keys():  # 这一步总是成立，因为blg_dic_tan是根据supply 中的tan生成的 --- 可以考虑去除此判断
            supply_list_tan = supply_dic_tan[tan]
            # print(supply_list_tan)

            # 按顺序对每一个po进行数量分配
            for po in blg_list_tan:
                po_supply_allocation = []
                last_supply_date = None
                # allocated_supply=[] #已经分配完的supply

                po_qty = list(po.values())[0][0]
                po_number = list(po.values())[0][1][0]
                min_date = list(po.values())[0][1][1]
                #option_number = list(po.values())[0][1][1]
                #target_fcd = list(po_option.values())[0][1][2]
                #current_fcd = list(po_option.values())[0][1][3]

                shortage_qty = po_qty

                # 按顺序将supply分给此po
                for date_qty in supply_list_tan:
                    supply_date = list(date_qty.keys())[0]
                    supply_qty = list(date_qty.values())[0]

                    if po_qty < supply_qty:  # po数量小于supply数量：po被全额满足；supply数量被减掉；已分配的supply被记录 （后面跳转到下一个po）
                        po_supply_allocation.append((supply_date.strftime('%Y-%m-%d'), po_qty))

                        # 更新supply_list_tan
                        new_supply_qty = supply_qty - po_qty
                        ind = supply_list_tan.index(date_qty)
                        supply_list_tan[ind] = {supply_date: new_supply_qty}

                        po_qty = 0
                        last_supply_date = supply_date

                        # 计算shortage_to_target_fcd和shortage_to_current_fcd
                        if supply_date <= min_date:
                            shortage_qty = 0

                        break
                    elif po_qty == supply_qty:  # po数量等于supply数量：po_optio被全额满足；已分配的supply被记录；跳出本次po循环(进到下一个supply循环)
                        po_supply_allocation.append((supply_date.strftime('%Y-%m-%d'), po_qty))

                        # 更新supply_list_tan
                        new_supply_qty = 0
                        ind = supply_list_tan.index(date_qty)
                        supply_list_tan[ind] = {supply_date: new_supply_qty}

                        po_qty = 0
                        last_supply_date = supply_date

                        # 计算shortage_to_target_fcd和shortage_to_current_fcd
                        if supply_date <= min_date:
                            shortage_qty = 0

                        break
                    else:  # po数量大于supply数量：po被部分满足（=supply qty）；po数量被改小；进到下一个supply循环(同一个po)
                        if supply_qty > 0:
                            po_supply_allocation.append((supply_date.strftime('%Y-%m-%d'), supply_qty))

                            # 更新supply_list_tan
                            new_supply_qty = 0
                            ind = supply_list_tan.index(date_qty)
                            supply_list_tan[ind] = {supply_date: new_supply_qty}

                            # 计算shortage_to_target_fcd和shortage_to_current_fcd
                            if supply_date <= min_date:
                                shortage_qty = shortage_qty - supply_qty

                            # 更新订单数量
                            po_qty = po_qty - supply_qty

                # 完成一个po的分配，结果加入blg_with_allocation中
                uncovered_qty = po_qty
                blg_with_allocation[po_number + '_'  + tan] = (po_supply_allocation, last_supply_date, uncovered_qty, shortage_qty)


    return blg_with_allocation

def identify_top_gating_pn(df_3a4):
    """
    Identify the top gating items for missing lt_target_fcd or fcd.
    :param df_3a4:
    :return:
    """
    df_3a4.loc[:,'top_gating_target_fcd']=np.where(df_3a4.shortage_to_target_fcd>0,
                                                   np.where(df_3a4.tan_supply_ready_date==df_3a4.po_supply_ready_date,
                                                            'YES',
                                                            'NO'),
                                                   None)

    df_3a4.loc[:,'top_gating_fcd']=np.where(df_3a4.shortage_to_current_fcd>0,
                                                   np.where(df_3a4.tan_supply_ready_date==df_3a4.po_supply_ready_date,
                                                            'YES',
                                                            'NO'),
                                                   None)

    return df_3a4




# calculate the PO supply ready date
def calculate_po_supply_ready_date_and_add_to_3a4(df_3a4, pn_to_consider):
    dfx = df_3a4[df_3a4.BOM_PN.isin(pn_to_consider)].copy()

    dfx.sort_values(by=['PO_NUMBER', 'tan_supply_ready_date'], ascending=True, inplace=True)
    po_supply_date = [(x, y) for x, y in zip(dfx.PO_NUMBER, dfx.tan_supply_ready_date)]

    po_supply_ready_date = {}
    po = po_supply_date[0][0]
    supply_date = po_supply_date[0][1]

    po_supply_ready_date[po] = supply_date  # 第一个值

    for po_date in po_supply_date[1:]:
        if po_date[0] != po or po_date[1] != supply_date:
            po_supply_ready_date[po_date[0]] = po_date[1]
        po = po_date[0]
        supply_date = po_date[1]

    # 将po supply ready date加入3a4中
    df_3a4.loc[:, 'po_supply_ready_date'] = df_3a4.PO_NUMBER.map(
            lambda x: po_supply_ready_date[x] if x in po_supply_ready_date.keys() else None)

    return df_3a4


def calculate_po_ctb_in_3a4(df_3a4):
    """
    Add 2 days FLT to get PO CTB based on PO supply available date;  PO without
    supply coverage will have ctb as 180days out (ITF); CTB in the past will be changed to today; also
    conider the earliest allowed packout date(based on target date, hold, and whether is scheduled).
    Simplified approach on FLT.. need to be updated
    """
    today=pd.Timestamp.today().date()
    date_180=pd.Timestamp.today().date() + pd.Timedelta(180, 'd')
    # using a date instead of ITF for resample purpose
    df_3a4.loc[:, 'po_ctb'] = np.where((df_3a4.po_supply_ready_date.isnull()) | (df_3a4.ADDRESSABLE_FLAG=='MFG_HOLD'),
                                       date_180,
                                       np.where(df_3a4.earliest_packable_date>df_3a4.po_supply_ready_date + pd.Timedelta(3, 'd'),
                                                df_3a4.earliest_packable_date,
                                                df_3a4.po_supply_ready_date + pd.Timedelta(FLT, 'd'))
                                        )

    # change po ctb in the past to today
    df_3a4.loc[:, 'po_ctb'] = np.where(df_3a4.po_ctb<today,
                                       today,
                                       df_3a4.po_ctb)

    # add ctb comments
    df_3a4.loc[:,'ctb_comment']=np.where(df_3a4.po_ctb==today,
                                         'Ready to go',
                                         np.where(df_3a4.po_ctb>df_3a4.earliest_packable_date,
                                                  'Following supply',
                                                  np.where((df_3a4.CURRENT_FCD_NBD_DATE.isnull()),
                                                            'Unscheduled - CTB 2wk out',
                                                            'Following Target FCD or Exception')))

    df_3a4.loc[:, 'ctb_comment'] = np.where(df_3a4.po_supply_ready_date.isnull(),
                                            'ITF - no supply recovery',
                                            np.where(df_3a4.ADDRESSABLE_FLAG=='MFG_HOLD',
                                                     'ITF - MFG Hold',
                                                     df_3a4.ctb_comment))

    return df_3a4


def calculate_ss_ctb_and_add_to_3a4(df_3a4):
    """
    Calculate SS CTB based on PO CTB date - the latest one----- having bug to be fixed!!!!!!
    """

    ss = df_3a4[df_3a4.po_ctb.notnull()].SO_SS.unique()
    dfx = df_3a4[df_3a4.SO_SS.isin(ss)].copy()

    dfx.sort_values(by=['SO_SS', 'po_ctb'], ascending=True, inplace=True)
    po_ctb_date = [(x, y) for x, y in zip(dfx.SO_SS, dfx.po_ctb)]

    ss_ctb_date = {}
    ss = po_ctb_date[0][0]
    ctb_date = po_ctb_date[0][1]

    ss_ctb_date[ss] = ctb_date  # 第一个值

    for ss_date in po_ctb_date[1:]:
        if ss_date[0] != ss or ss_date[1] != ctb_date:
            ss_ctb_date[ss_date[0]] = ss_date[1]
        ss = ss_date[0]
        ctb_date = ss_date[1]

    # add to 3a4
    df_3a4.loc[:, 'ss_ctb'] = df_3a4.SO_SS.map(lambda x: ss_ctb_date[x] if x in ss_ctb_date.keys() else None)

    return df_3a4

def calculate_riso_status(df_3a4):
    """
    Identify current RISo status and to be RISO status based on SS_CTB
    :param df_3a4:
    :return:
    """

    df_3a4.loc[:,'RISO (as is)']=np.where((df_3a4.LT_TARGET_FCD.notnull() & df_3a4.CURRENT_FCD_NBD_DATE.notnull()),
                                          np.where(df_3a4.LT_TARGET_FCD<df_3a4.CURRENT_FCD_NBD_DATE,
                                                   'Yes',
                                                   'No'),
                                          None)

    df_3a4.loc[:, 'RISO (to be)'] = np.where((df_3a4.LT_TARGET_FCD.notnull() & df_3a4.ss_ctb.notnull()),
                                             np.where(df_3a4.LT_TARGET_FCD < df_3a4.ss_ctb,
                                                      'Yes',
                                                      'No'),
                                             None)

    return df_3a4

def write_excel_file(fname, data_to_write):
    '''
    Write the df into excel files as different sheets
    :param fname: fname of the output excel
    :param data_to_write: a dict that contains {sheet_name:df}
    :return: None
    '''
    # engine='xlsxwriter' is used to avoid illegal character which lead to failure of saving the file
    writer = pd.ExcelWriter(fname, engine='xlsxwriter')

    for sheet_name, df in data_to_write.items():
        df.to_excel(writer, sheet_name=sheet_name)

    writer.save()

def read_3a4_and_limit_org_bu(file_path_3a4, bu_list, org_list):
    """
    Read 3a4 data and limit to selected org and bu;
    """
    # formally read 3a4
    df_3a4 = pd.read_csv(file_path_3a4, encoding='iso-8859-1',
                             # parse_dates=['CURRENT_FCD_NBD_DATE', 'ORIGINAL_FCD_NBD_DATE', 'LT_TARGET_FCD'],
                             low_memory=False)
    """
    # 重新根据org存储文件
    orgs = df_3a4.ORGANIZATION_CODE.unique()
    fname = file_path_3a4[:-4] + ' (' + '_'.join(orgs) + ').csv'
    df_3a4.to_csv(fname,index=False)
    """

    # Limit 3a4 to selected BU and Org
    if bu_list==['']:
        bu_list=df_3a4.BUSINESS_UNIT.unique()

    df_3a4 = df_3a4[(df_3a4.BUSINESS_UNIT.isin(bu_list)) & (df_3a4.ORGANIZATION_CODE.isin(org_list))].copy()

    return df_3a4


def read_data_cm(fname_supply,fname_ct2r):
    """
    Read supply data collected through CM/SCR
    :param fname_supply:
    :param fname_ct2r:
    :return:
    """
    df_supply_pcba = pd.read_excel(fname_supply, sheet_name='PCBA')
    df_supply_df = pd.read_excel(fname_supply, sheet_name='DF')
    df_ct2r = pd.read_excel(fname_ct2r)

    return df_supply_pcba,df_supply_df,df_ct2r

def get_file_info_on_drive(base_path,keep_hours=100):
    """
    Collect the file info on a drive and make that into a df. Remove files if older than keep_hours.
    """
    now=time.time()
    file_list = os.listdir(base_path)
    files = []
    creation_time = []
    file_size = []
    file_path = []
    for file in file_list:
        c_time = os.stat(os.path.join(base_path, file)).st_ctime

        if (now - c_time) / 3600 > keep_hours: #hours
            os.remove(os.path.join(base_path, file))
        else:
            c_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(c_time))
            file_s = os.path.getsize(os.path.join(base_path, file))
            if file_s > 1024 * 1024:
                file_s = str(round(file_s / (1024 * 1024), 1)) + 'M'
            else:
                file_s = str(int(file_s / 1024)) + 'K'

            files.append(file)
            creation_time.append(c_time)
            file_size.append(file_s)
            file_path.append(os.path.join(base_path,file))

    df_file_info=pd.DataFrame({'File_name':files,'Creation_time':creation_time, 'File_size':file_size, 'File_path':file_path})
    df_file_info.sort_values(by='Creation_time',ascending=False,inplace=True)

    return df_file_info


def read_kinaxis_supply_and_initial_process(fname_supply_kinaxis):
    """
    从kinaxis文件中读取supply数据，去除不要的列; 并按Org存储新文件
    :param fname_supply_kinaxis:
    :return:
    """
    df_kinaxis=pd.read_excel(fname_supply_kinaxis,header=1)

    """
    # 重新根据org存储文件
    orgs = df_kinaxis.ORG.unique()
    fname = fname_supply_kinaxis[:-5] + ' (' + '_'.join(orgs) + ').xlsx'
    df_kinaxis.to_excel(fname, index=False)
    """

    #判断位置并取出需要的信息
    col=df_kinaxis.columns
    ind = col.tolist().index('Past') - 1
    df_supply=df_kinaxis[df_kinaxis[col[ind]]=='Total Supply'].copy()
    df_supply.set_index(['TAN'], inplace=True)
    df_supply = df_supply.iloc[:, ind:]
    #Past改为本周前一天的日期；列名改为日期格式
    current_wk_start=col[ind+2]
    oh_date=pd.to_datetime(current_wk_start)-pd.Timedelta(1,'d')
    df_supply.rename(columns={'Past':oh_date},inplace=True)
    df_supply.columns=[x.date() for x in pd.to_datetime(df_supply.columns)]
    #处理更正数据格式
    df_supply = df_supply.applymap(lambda x: str(x).replace(',', ''))
    df_supply = df_supply.applymap(lambda x: str(x).replace('nan', '0'))
    df_supply = df_supply.applymap(float)
    df_supply = df_supply.applymap(int)

    return df_supply


def basic_data_processing_3a4(df_3a4):
    """
    Do some basic 3a4 processing
    :param df_3a4:
    :return:
    """

    df_3a4.CURRENT_FCD_NBD_DATE = pd.to_datetime(df_3a4.CURRENT_FCD_NBD_DATE)
    df_3a4.ORIGINAL_FCD_NBD_DATE=pd.to_datetime(df_3a4.ORIGINAL_FCD_NBD_DATE)
    df_3a4.TARGET_SSD=pd.to_datetime(df_3a4.TARGET_SSD)
    df_3a4.LT_TARGET_FCD=pd.to_datetime(df_3a4.LT_TARGET_FCD)
    df_3a4.CUSTOMER_REQUESTED_SHIP_DATE=pd.to_datetime(df_3a4.CUSTOMER_REQUESTED_SHIP_DATE)
    df_3a4.CURRENT_FCD_NBD_DATE = df_3a4.CURRENT_FCD_NBD_DATE.map(lambda x: x.date() if x is not np.nan else np.nan)
    df_3a4.ORIGINAL_FCD_NBD_DATE=df_3a4.ORIGINAL_FCD_NBD_DATE.map(lambda x: x.date() if x is not np.nan else np.nan)
    df_3a4.TARGET_SSD=df_3a4.TARGET_SSD.map(lambda x: x.date() if x is not np.nan else np.nan)
    df_3a4.LT_TARGET_FCD=df_3a4.LT_TARGET_FCD.map(lambda x: x.date() if x is not np.nan else np.nan)
    df_3a4.CUSTOMER_REQUESTED_SHIP_DATE=df_3a4.CUSTOMER_REQUESTED_SHIP_DATE.map(lambda x: x.date() if x is not np.nan else np.nan)

    # 取最小日期作为排序依据
    df_3a4.loc[:,'min_date']=np.where(df_3a4.REVENUE_NON_REVENUE=='NO',#!!!!non rev orders consider CURRENT FCD instead and only
                                        df_3a4.CURRENT_FCD_NBD_DATE,
                                        np.where(df_3a4.ORIGINAL_FCD_NBD_DATE > df_3a4.LT_TARGET_FCD,
                                                df_3a4.LT_TARGET_FCD,
                                                df_3a4.ORIGINAL_FCD_NBD_DATE))

    #更改列名
    df_3a4.rename(columns={'CTB_STATUS':'CTB_STATUS(CTB_UI)'},inplace=True)

    return df_3a4

def exclude_unneeded_and_missing_ct2r(df_supply_df,df_ct2r):
    """
    Exclude the unneeded CT2R data and also identify the missing CT2R PN for the DF materials. Do this with the versionless PN.
    :param df_supply_df:
    :param df_ct2r:
    :return:
    """
    # Missing CT2R
    pn = np.setdiff1d(df_supply_df.index,df_ct2r.index)
    df_missing_ct2r=pd.DataFrame({'Missing CT2R':pn})

    # 从df_ct2r中排除不在df_supply_df中的pn - 不需要，且否则后面会报错
    pn = np.intersect1d(df_ct2r.index, df_supply_df.index)
    df_ct2r = df_ct2r.loc[pn, :]

    return df_ct2r,df_missing_ct2r

def add_allocation_result_to_3a4(df_3a4,blg_with_allocation):
    """
    将分配结果加入到3a4中
    :param df_3a4:
    :param blg_with_allocation:key=[po_number + '_'  + tan];value= (po_supply_allocation, last_supply_date, uncovered_qty, shortage_qty to min_date)
    :return:
    """
    #df_3a4.OPTION_NUMBER = df_3a4.OPTION_NUMBER.astype(str)
    df_3a4.loc[:, 'po_pn'] = df_3a4.PO_NUMBER +  '_' + df_3a4.BOM_PN
    #print(blg_with_allocation)

    df_3a4.loc[:, 'supply_allocation'] = df_3a4.po_pn.map(
        lambda x: blg_with_allocation[x][0] if x in blg_with_allocation.keys() else None)
    df_3a4.loc[:, 'tan_supply_ready_date'] = df_3a4.po_pn.map(
        lambda x: blg_with_allocation[x][1] if x in blg_with_allocation.keys() else None)
    df_3a4.loc[:, 'tan_qty_wo_supply'] = df_3a4.po_pn.map(
        lambda x: blg_with_allocation[x][2] if x in blg_with_allocation.keys() else None)
    df_3a4.loc[:, 'shortage_to_min_date'] = df_3a4.po_pn.map(
        lambda x: blg_with_allocation[x][3] if x in blg_with_allocation.keys() else None)

    #df_3a4.OPTION_NUMBER = df_3a4.OPTION_NUMBER.astype(int)

    return df_3a4


def calculate_earliest_packable_date(df_3a4):
    """
    Calculate an earliest allowed packout date based on if order is scheduled, and LT_target_FCD,target_SSD. This logic
    need to align with Addressable logic.
    :param df_3a4:
    :return:
    """
    packable_date_target_fcd=df_3a4.LT_TARGET_FCD - pd.Timedelta(addressable_window['LT_TARGET_FCD'],'d')
    packable_date_target_ssd = df_3a4.TARGET_SSD - pd.Timedelta(addressable_window['TARGET_SSD'], 'd')
    packable_date_14=pd.Timestamp.today().date() + pd.Timedelta(14,'d')
    packable_date_180 = pd.Timestamp.today().date() + pd.Timedelta(180, 'd')

    df_3a4.loc[:,'earliest_packable_date']=np.where(df_3a4.CURRENT_FCD_NBD_DATE.isnull(),
                                                           packable_date_14,
                                                           np.where(packable_date_target_fcd>packable_date_target_ssd,
                                                                    packable_date_target_ssd,
                                                                    packable_date_target_fcd)
                                                           )
    # if order WITH mfg_hold and earliest packable date within 180 days, further push out to 8days from today
    df_3a4.loc[:, 'earliest_packable_date'] = np.where((df_3a4.MFG_HOLD=='Y') & (df_3a4.earliest_packable_date<packable_date_180),
                                                       packable_date_180,
                                                       df_3a4.earliest_packable_date
                                                       )

    # further update for exceptional issues - push out 2 weeks (e.g. config issue)
    df_3a4.loc[:, 'earliest_packable_date']=np.where(df_3a4.EXCEPTION_NAME.notnull(),
                                                     np.where(df_3a4.earliest_packable_date>packable_date_14,
                                                            df_3a4.earliest_packable_date,
                                                            packable_date_14),
                                                     df_3a4.earliest_packable_date)

    return df_3a4

def redefine_addressable_flag_main_pip_version(df_3a4):
    '''
    Updated on Oct 27, 2020 to leveraging existing addressable definition of Y, and redefine the NO to MFG_HOLD,
    UNSCHEDULED,PACKED,PO_CANCELLED,NON_REVENUE
    :param df_3a4:
    :return:
    '''

    # Convert YES to ADDRESSABLE
    df_3a4.loc[:, 'ADDRESSABLE_FLAG'] = np.where(df_3a4.ADDRESSABLE_FLAG=='YES',
                                                 'ADDRESSABLE',
                                                 df_3a4.ADDRESSABLE_FLAG)


    # 如果没有LT_TARGET_FCD/TARGET_SSD/CURRENT_FCD_NBD_DATE,则作如下处理 - 可能是没有schedule或缺失Target LT date or Target SSD
    df_3a4.loc[:, 'ADDRESSABLE_FLAG'] = np.where(df_3a4.CURRENT_FCD_NBD_DATE.isnull(),
                                                'UNSCHEDULED',
                                                df_3a4.ADDRESSABLE_FLAG)

    # Non_revenue orders
    df_3a4.loc[:, 'ADDRESSABLE_FLAG'] = np.where(df_3a4.REVENUE_NON_REVENUE=='NO',
                                                  'NON_REVENUE',
                                                df_3a4.ADDRESSABLE_FLAG)


    #  mfg-hold
    df_3a4.loc[:, 'ADDRESSABLE_FLAG'] = np.where(df_3a4.MFG_HOLD=='Y',
                                                 'MFG_HOLD',
                                                 df_3a4.ADDRESSABLE_FLAG)

    # redefine cancellation order to PO_CANCELLED - put this after MFG_HOLD so it won't get replaced by MFG_HOLD
    df_3a4.loc[:, 'ADDRESSABLE_FLAG'] = np.where(
        (df_3a4.ORDER_HOLDS.str.contains('Cancellation', case=False)) & (df_3a4.ORDER_HOLDS.notnull()),
        'PO_CANCELLED',
        df_3a4.ADDRESSABLE_FLAG)

    #df_3a4[(df_3a4.OPTION_NUMBER==0)&(df_3a4.ORGANIZATION_CODE=='FOC')][['PO_NUMBER','ADDRESSABLE_FLAG_redefined','ADDRESSABLE_FLAG','MFG_HOLD','ORDER_HOLDS','TARGET_SSD','LT_TARGET_FCD','CURRENT_FCD_NBD_DATE']].to_excel('3a4 processed-2.xlsx',index=False)
    # OTHER non-addressable
    df_3a4.loc[:, 'ADDRESSABLE_FLAG']=np.where(df_3a4.ADDRESSABLE_FLAG=='NO',
                                               'NOT_ADDRESSABLE',
                                               df_3a4.ADDRESSABLE_FLAG)

    return df_3a4



def make_summary_build_impact(df_3a4,df_supply,output_col,qend,blg_with_allocation,FLT,cut_off='wk0'):
    """
    Calculate and add extra col to indicate build impact by the cutoff week (wk0 as current week). The backlog base is
    based on the earliest_packable_date by Saterday of the specified week. Cut off date for material is by Thursday for
    that week.
    Revenue is looking at PO level un-staged revenue.
    :param df_3a4:
    :param qend: qend date predefined in setting
    :param cut_off: cut off week - wk0, wk1, wk2, etc.
    :param output_col:
    :return:
    """
    # identify Sunday date for current week impact
    today = pd.Timestamp.now().date()

    # define the supply cut off date
    if cut_off=='qend':
        supply_cut_off = qend - pd.Timedelta(FLT, 'd') # Thur of wk13
        packable_cut_off=qend - pd.Timedelta(1, 'd') # Sat of wk13
    else:
        today_name = pd.Timestamp.today().day_name()
        offset_wk=int(cut_off[-1])

        if today_name == 'Monday':
            offset = 6 + offset_wk * 7
        elif today_name == 'Tuesday':
            offset = 5 + offset_wk * 7
        elif today_name == 'Wednesday':
            offset = 4 + offset_wk * 7
        elif today_name == 'Thursday':
            offset = 3 + offset_wk * 7
        elif today_name == 'Friday':
            offset = 2 + offset_wk * 7
        elif today_name == 'Saturday':
            offset = 1 + offset_wk * 7
        elif today_name == 'Sunday':
            offset = 0 + offset_wk * 7

        supply_cut_off=today+pd.Timedelta(offset-FLT,'d') # Thursday of the cutoff week
        packable_cut_off=today+pd.Timedelta(offset-1,'d') # Satday of the cutoff week

    # update in 3a4 if supply impact cut_off week build
    impact_rev='build_impact_'+cut_off
    impact_qty='build_shortage_'+cut_off

    # 添加rev impact 列（po unstaged revenue)
    df_3a4.loc[:,impact_rev]=np.where(df_3a4.earliest_packable_date<=packable_cut_off,
                                        np.where(df_3a4.tan_supply_ready_date>supply_cut_off,
                                                df_3a4.C_UNSTAGED_DOLLARS,
                                                None),
                                        None)

    # 添加数量列(shortage)
    dfx = df_3a4[(df_3a4.earliest_packable_date<=packable_cut_off) &
                 (df_3a4.po_pn.isin(blg_with_allocation.keys())) &
                 (df_3a4[impact_rev].notnull())]

    for row in dfx.iterrows():
        po_pn = row[1].po_pn
        allocation = blg_with_allocation[po_pn][0]
        shortage_qty = blg_with_allocation[po_pn][2]  # first is po overall shortage qty
        for a in allocation:
            if pd.to_datetime(a[0]) > supply_cut_off:  # 如果allocation晚于cut off date, 分配的数量加总到总的po shortage qty上
                shortage_qty += a[1]

        df_3a4.loc[row[0], impact_qty] = shortage_qty

    # create the gating pn col and indidate whether is top gating or non-top gating
    gating_col_name = 'gating_build_' + cut_off
    df_3a4.loc[:,gating_col_name]=np.where(df_3a4[impact_rev].notnull(),# use notnull instead of >0 due to non-rev order exist
                                                    np.where(df_3a4.tan_supply_ready_date==df_3a4.po_supply_ready_date,
                                                             'Top-gating',
                                                             'Non top-gating'),
                                                     None)


    # For top gating: when multiple top gating in one PO, remove the rest and keep the first one as top gating.
    df_top_gating=df_3a4[df_3a4[gating_col_name]=='Top-gating']
    df_top_gating_duplicated_po_pn=df_top_gating[df_top_gating.duplicated('PO_NUMBER')].po_pn
    df_3a4.loc[:,gating_col_name]=np.where(df_3a4.po_pn.isin(df_top_gating_duplicated_po_pn),
                                                         'Non top-gating',
                                                         df_3a4[gating_col_name])

    # For both top gating and NON top gating: when same PN in one PO, remove the rest and keep the first one to avoid duplicate on rev
    df_gating = df_3a4[df_3a4[gating_col_name].notnull()] # both top and non top gating
    df_gating_duplicated_po_pn = df_gating[df_gating.duplicated(['PO_NUMBER','BOM_PN'])].po_pn
    df_3a4.loc[:, gating_col_name] = np.where(
        df_3a4.po_pn.isin(df_gating_duplicated_po_pn),
        None,
        df_3a4[gating_col_name])

    # add the col to the list
    output_col.append(impact_rev)
    output_col.append(impact_qty)
    output_col.append(gating_col_name)

    # 制作汇总数据表 - 不考虑'BUSINESS_UNIT','PRODUCT_FAMILY'，后面处理后加入
    df_impact_rev_summary = df_3a4[(df_3a4[gating_col_name].notnull())].pivot_table(index=['ORGANIZATION_CODE', 'BOM_PN'],
                                                                                    columns=gating_col_name,
                                                                                    values=impact_rev,
                                                                                    aggfunc=sum)
    df_impact_rev_summary.loc[:, 'Total'] = df_impact_rev_summary.sum(axis=1)
    df_impact_rev_summary = df_impact_rev_summary.applymap(lambda x: round(x / 1000000, 1))

    df_impact_qty_summary = df_3a4[(df_3a4[gating_col_name].notnull())].pivot_table(
                                                                                    index=['ORGANIZATION_CODE', 'BOM_PN'],
                                                                                    columns=gating_col_name,
                                                                                    values=impact_qty,
                                                                                    aggfunc=sum)
    df_impact_qty_summary.loc[:, 'Total'] = df_impact_qty_summary.sum(axis=1)

    # combine the rev and qty summaries
    df_build_impact_summary=pd.merge(df_impact_rev_summary,df_impact_qty_summary,left_index=True,right_index=True,
                                     sort=False,suffixes=('_x','_y'))

    df_build_impact_summary.rename(columns={'Non top-gating_x':'Rev impact (non-gating)',
                                            'Top-gating_x':'Rev impact (top-gating)',
                                            'Total_x':'Rev impact (total)',
                                            'Non top-gating_y':'Short qty (non-gating)',
                                            'Top-gating_y':'Short qty (top-gating)',
                                            'Total_y':'Short qty (total)'},
                                   inplace=True)

    # 把BU/PF信息加入summary (针对重复的org_PN (due to reporting to different BU/PF),对BU/PF做相应的汇总)
    df_build_impact_summary.reset_index(inplace=True)
    df_build_impact_summary.loc[:,'org_pn']=df_build_impact_summary.ORGANIZATION_CODE + '_' + df_build_impact_summary.BOM_PN

    df_3a4.loc[:,'org_pn']=df_3a4.ORGANIZATION_CODE + '_' + df_3a4.BOM_PN
    dfx=df_3a4[df_3a4.org_pn.isin(df_build_impact_summary.org_pn.unique())][['org_pn','BUSINESS_UNIT','PRODUCT_FAMILY']]
    dfx.drop_duplicates(['org_pn','BUSINESS_UNIT','PRODUCT_FAMILY'],inplace=True)
    dfx.sort_values(by='org_pn',inplace=True) # important to sort first
    org_pn_dic={} #生成一个org_pn：bu,pf的字典
    org_pn=dfx.iloc[0,0]
    bu = dfx.iloc[0,1]
    pf = dfx.iloc[0,2]
    org_pn_dic[org_pn]=[bu,pf]
    for row in dfx[1:].itertuples(index=False):
        org_pn_new=row.org_pn
        bu_new=row.BUSINESS_UNIT
        pf_new=row.PRODUCT_FAMILY
        if org_pn_new!=org_pn:
            org_pn=org_pn_new
            bu=bu_new
            pf=pf_new
        else:
            if bu_new not in bu:
                bu=bu+'/'+bu_new

            if pf_new not in pf:
                pf=pf+'/'+pf_new
        org_pn_dic[org_pn] = [bu, pf]

    df_build_impact_summary.loc[:,'BUSINESS_UNIT']=df_build_impact_summary.org_pn.map(lambda x: org_pn_dic[x][0])
    df_build_impact_summary.loc[:, 'PRODUCT_FAMILY'] = df_build_impact_summary.org_pn.map(lambda x: org_pn_dic[x][1])

    df_build_impact_summary.drop('org_pn',axis=1,inplace=True)
    df_build_impact_summary.loc[:,'Recovery date']=None

    # 把supply (future supply)合并入df_build_impact_summary
    supply_col=pd.to_datetime(df_supply.columns).tolist()
    supply_col=[dt.date() for dt in supply_col]
    ind=0
    for dt in supply_col:
        if dt>supply_cut_off:
            ind=supply_col.index(dt)
            #print(supply_cut_off,dt,ind)
            break

    #TODO: should use org_pn as merging key instead - currently only for single site so it's OK
    df_supply_x=df_supply.iloc[:,ind:]
    df_build_impact_summary=pd.merge(df_build_impact_summary,df_supply_x,left_on='BOM_PN',right_on='TAN',how='left')

    df_build_impact_summary.set_index(['ORGANIZATION_CODE', 'BOM_PN', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'], inplace=True)
    df_build_impact_summary.sort_values('Rev impact (total)', ascending=False, inplace=True)

    return df_3a4,df_build_impact_summary,output_col



def update_ss_status(df_3a4):
    #print(df_3a4.ss_ctb.unique())
    df_3a4.loc[:, 'ss_updated_status'] = np.where(df_3a4.CURRENT_FCD_NBD_DATE.notnull(),
                                                  np.where(df_3a4.CURRENT_FCD_NBD_DATE > df_3a4.ss_ctb+ pd.Timedelta(3,'d'),
                                                           'Pull in opportunity(>3 days)', # >3 days considered as pull in opportunity
                                                           np.where(df_3a4.CURRENT_FCD_NBD_DATE <
                                                               df_3a4.ss_ctb,
                                                                    'Decommit risk',
                                                                    'On schedule(0~3 days)')),
                                                'Not scheduled')

    return df_3a4

def write_data_to_spreadsheet(base_dir_output,output_filename,data_to_write):
    """
    Write the data to spreadsheet in multiple sheets.
    :param output_filename: File path and name to save the file
    :param data_to_write: A dict: {sheet_name:df}
    :return:
    """
    output_path=os.path.join(base_dir_output,output_filename)

    write_excel_file(output_path, data_to_write)

#??????????? Below was old and discarded - refer to ss_ranking_overall_new
def ss_ranking_overall(df_3a4,ranking_col, qend, order_col='SO_SS', new_col='ss_overall_rank'):
    """
    根据priority_cat,partial_pack, min_date, ss_rev_rank(或po_rev_rank),按照ranking_col的顺序对SS进行排序。
    CRSD 不在本quarter放在后面；最后放MFG_HOLD订单; non-revenue 订单revenue值当成0，且只考虑current FCD(包括priority订单）。
    :param df_3a4:
    :param ranking_col:e.g. ['priority_rank', 'ss_rev_rank', 'min_date', 'SO_SS', 'PO_NUMBER']
    :param qend: qend date,用来判断CRSD是否在本季度
    :param order_col:
    :param new_col:
    :return:
    """

    ### Step0: change non-rev orders unstaged $ to 0
    df_3a4.loc[:,'C_UNSTAGED_DOLLARS']=np.where(df_3a4.REVENUE_NON_REVENUE == 'NO',
                                                0,
                                                df_3a4.C_UNSTAGED_DOLLARS)

    #### Step1: 生成ss_unstg_rev并据此排序
    # 计算ss_unstg_rev
    ss_unstg_rev = {}
    df_rev = df_3a4.pivot_table(index='SO_SS', values='C_UNSTAGED_DOLLARS', aggfunc=sum)
    for ss, rev in zip(df_rev.index, df_rev.values):
        ss_unstg_rev[ss] = rev[0]
    df_3a4.loc[:, 'ss_unstg_rev'] = df_3a4.SO_SS.map(lambda x: ss_unstg_rev[x])

    """
    # 计算po_rev_unit - non revenue change to 0
    df_3a4.loc[:, 'po_rev_unit'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'YES',
                                            df_3a4.SOL_REVENUE / df_3a4.ORDERED_QUANTITY,
                                            0)

    # 计算ss_rev_unit: 通过po_rev_unit汇总
    ss_rev_unit = {}
    dfx_rev = df_3a4.pivot_table(index='SO_SS', values='po_rev_unit', aggfunc=sum)
    for ss, rev in zip(dfx_rev.index, dfx_rev.values):
        ss_rev_unit[ss] = rev[0]
    df_3a4.loc[:, 'ss_rev_unit'] = df_3a4.SO_SS.map(lambda x: int(ss_rev_unit[x]))
    """

    # create rank#
    rank = {}
    order_list = df_3a4.sort_values(by='ss_unstg_rev', ascending=False).SO_SS.unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, 'ss_rev_rank'] = df_3a4.SO_SS.map(lambda x: rank[x])

    #### Step2: 取最小日期作为排序依据
    #(Normal time)
    df_3a4.loc[:, 'min_date'] = np.where(df_3a4.ORIGINAL_FCD_NBD_DATE > df_3a4.LT_TARGET_FCD,
                                        df_3a4.LT_TARGET_FCD,
                                        df_3a4.ORIGINAL_FCD_NBD_DATE)
    #(special time - QED - non_rev only consider current FCD)
    '''
    df_3a4.loc[:, 'min_date'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'NO',
                                         # !!!!non rev orders consider CURRENT FCD instead and only
                                         df_3a4.CURRENT_FCD_NBD_DATE,
                                         np.where(df_3a4.ORIGINAL_FCD_NBD_DATE > df_3a4.LT_TARGET_FCD,
                                                  df_3a4.LT_TARGET_FCD,
                                                  df_3a4.ORIGINAL_FCD_NBD_DATE))
    '''

    #### Step3: 重新定义priority order及排序
    df_3a4.loc[:, 'priority_cat'] = np.where(df_3a4.SECONDARY_PRIORITY.isin(['PR1', 'PR2', 'PR3']),
                                             df_3a4.SECONDARY_PRIORITY,
                                             np.where(df_3a4.FINAL_ACTION_SUMMARY == 'TOP 100',
                                                      'TOP 100',
                                                      np.where(
                                                          df_3a4.FINAL_ACTION_SUMMARY == 'LEVEL 4 ESCALATION PRESENT',
                                                          'L4',
                                                          np.where(df_3a4.BUP_RANK.notnull(),
                                                                   'BUP',
                                                                   None)
                                                          )
                                                      )
                                             )
    """
    # change non_rev ones to different so they don't have priority
    df_3a4.loc[:, 'priority_cat'] = np.where((df_3a4.REVENUE_NON_REVENUE=='NO')&(df_3a4.priority_cat.notnull()),
                                             'Non_rev '+df_3a4.priority_cat,
                                             df_3a4.priority_cat)
    """

    # Update below to PR3 due to current PR1/2/3 not updated when order change to DPAS from others
    df_3a4.loc[:, 'priority_cat']=np.where((df_3a4.DPAS_RATING.isin(['DO','DX','TAA-DO','TAA-DX']))&(df_3a4.priority_cat.isnull()),
                                           'PR1',
                                           df_3a4.priority_cat)

    df_3a4.loc[:, 'priority_rank'] = np.where(df_3a4.priority_cat=='PR1',
                                            1,
                                            np.where(df_3a4.priority_cat =='PR2',
                                                     2,
                                                     np.where(df_3a4.priority_cat =='PR3',
                                                              3,
                                                              np.where(df_3a4.priority_cat == 'TOP 100',
                                                                        4,
                                                                        np.where(df_3a4.priority_cat == 'L4',
                                                                                5,
                                                                                np.where(df_3a4.priority_cat=='BUP',
                                                                                         6,
                                                                                         None)
                                                                                )
                                                                        )
                                                                )
                                                     )
                                              )

    #### Step4: create partial packed ranking
    dfx=df_3a4[(df_3a4.PACKOUT_QUANTITY.notnull())&(df_3a4.PACKOUT_QUANTITY!='Packout Completed')][['SO_SS','PACKOUT_QUANTITY']]
    dfx=dfx[~dfx.PACKOUT_QUANTITY.str.contains('0 of',case=False)]
    partial_ss=dfx.SO_SS.unique()
    df_3a4.loc[:,'partial_rank']=np.where(df_3a4.SO_SS.isin(partial_ss),1,2)

    ##### Step5: sort the SS and Put CRSD outside quarter and MFG hold orders at the back
    df_3a4.sort_values(by=ranking_col, ascending=True, inplace=True)
    # Put CRSD outside quarter and MFG hold orders at the back
    df_crsd_out=df_3a4[df_3a4.CUSTOMER_REQUESTED_SHIP_DATE>qend+pd.Timedelta(days=1)].copy()
    df_hold=df_3a4[df_3a4.ADDRESSABLE_FLAG=='MFG_HOLD'].copy()
    df_3a4=df_3a4[(df_3a4.ADDRESSABLE_FLAG!='MFG_HOLD')&(df_3a4.CUSTOMER_REQUESTED_SHIP_DATE<=qend+pd.Timedelta(days=1))].copy()
    df_3a4=pd.concat([df_3a4,df_crsd_out,df_hold],sort=False)

    # create rank# and put in 3a4
    rank = {}
    order_list = df_3a4[order_col].unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, new_col] = df_3a4[order_col].map(lambda x: rank[x])

    return df_3a4

## This is also discarded!!!!!
def ss_ranking_overall_new(df_3a4,ss_exceptional_priority, ranking_col, order_col='SO_SS', new_col='ss_overall_rank'):
    """
    根据priority_cat,OSSD,FCD, REVENUE_NON_REVENUE,C_UNSTAGED_QTY,按照ranking_col的顺序对SS进行排序。最后放MFG_HOLD订单.
    CTB和PCBA allocation用相同的方式在开始处删除cancelled的订单；summary_3a4不删除cancelled订单，不过在结尾处清除cancelled订单的ranking#
    :param df_3a4:
    :param ss_exceptional_priority: the exceptional priority from smartsheet
    :param ranking_col:e.g. ['priority_rank', 'ORIGINAL_FCD_NBD_DATE', 'CURRENT_FCD_NBD_DATE','rev_non_rev_rank',
                        'C_UNSTAGED_QTY', 'SO_SS','PO_NUMBER']
    :param order_col:'SO_SS'
    :param new_col:'ss_overall_rank'
    :return: df_3a4
    """

    # Below create a rev_rank for reference -  currently not used in overall ranking
    ### change non-rev orders unstaged $ to 0
    df_3a4.loc[:,'C_UNSTAGED_DOLLARS']=np.where(df_3a4.REVENUE_NON_REVENUE == 'NO',
                                                0,
                                                df_3a4.C_UNSTAGED_DOLLARS)

    #### 生成ss_unstg_rev并据此排序
    # 计算ss_unstg_rev
    ss_unstg_rev = {}
    df_rev = df_3a4.pivot_table(index='SO_SS', values='C_UNSTAGED_DOLLARS', aggfunc=sum)
    for ss, rev in zip(df_rev.index, df_rev.values):
        ss_unstg_rev[ss] = rev[0]
    df_3a4.loc[:, 'ss_unstg_rev'] = df_3a4.SO_SS.map(lambda x: ss_unstg_rev[x])

    """
    # 计算po_rev_unit - non revenue change to 0
    df_3a4.loc[:, 'po_rev_unit'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'YES',
                                            df_3a4.SOL_REVENUE / df_3a4.ORDERED_QUANTITY,
                                            0)

    # 计算ss_rev_unit: 通过po_rev_unit汇总
    ss_rev_unit = {}
    dfx_rev = df_3a4.pivot_table(index='SO_SS', values='po_rev_unit', aggfunc=sum)
    for ss, rev in zip(dfx_rev.index, dfx_rev.values):
        ss_rev_unit[ss] = rev[0]
    df_3a4.loc[:, 'ss_rev_unit'] = df_3a4.SO_SS.map(lambda x: int(ss_rev_unit[x]))
    """

    # create rev rank#
    rank = {}
    order_list = df_3a4.sort_values(by='ss_unstg_rev', ascending=False).SO_SS.unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, 'ss_rev_rank'] = df_3a4.SO_SS.map(lambda x: rank[x])

    # below creates overall ranking col
    ### Step1: 重新定义priority order及排序
    df_3a4.loc[:, 'priority_cat'] = np.where(df_3a4.SECONDARY_PRIORITY.isin(['PR1', 'PR2', 'PR3']),
                                             df_3a4.SECONDARY_PRIORITY,
                                             np.where(df_3a4.FINAL_ACTION_SUMMARY == 'TOP 100',
                                                      'TOP 100',
                                                      np.where(
                                                          df_3a4.FINAL_ACTION_SUMMARY == 'LEVEL 4 ESCALATION PRESENT',
                                                          'L4',
                                                          np.where(df_3a4.BUP_RANK.notnull(),
                                                                   'BUP',
                                                                   None)
                                                          )
                                                      )
                                             )
    #### Update below DO/DX orders to PR1 due to current PR1/2/3 not updated when order change to DPAS from others
    df_3a4.loc[:, 'priority_cat']=np.where((df_3a4.DPAS_RATING.isin(['DO','DX','TAA-DO','TAA-DX']))&(df_3a4.priority_cat.isnull()),
                                           'PR1',
                                           df_3a4.priority_cat)
    #### Give them a rank
    df_3a4.loc[:, 'priority_rank'] = np.where(df_3a4.priority_cat=='PR1',
                                            1,
                                            np.where(df_3a4.priority_cat =='PR2',
                                                     2,
                                                     np.where(df_3a4.priority_cat =='PR3',
                                                              3,
                                                              np.where(df_3a4.priority_cat == 'TOP 100',
                                                                        4,
                                                                        np.where(df_3a4.priority_cat == 'L4',
                                                                                5,
                                                                                np.where(df_3a4.priority_cat=='BUP',
                                                                                         6,
                                                                                         None)
                                                                                )
                                                                        )
                                                                )
                                                     )
                                              )

    ##### Step2: Give revenue/non-revenue a rank
    df_3a4.loc[:,'rev_non_rev_rank']=np.where(df_3a4.REVENUE_NON_REVENUE=='YES', 0, 1)

    #### Step3: Integrate ranking from smartsheet
    df_3a4.loc[:, 'priority_rank'] = np.where(df_3a4.SO_SS.isin(ss_exceptional_priority.keys()),
                                              df_3a4.SO_SS.map(lambda x: ss_exceptional_priority.get(x)),
                                              df_3a4.priority_rank)

    #df_3a4.loc[:,'priority_rank']=df_3a4.SO_SS.map(lambda x: ss_exceptional_priority.get(x))
    #df_3a4.loc[:, 'priority_rank_temp'] = df_3a4.SO_SS.map(lambda x: ss_exceptional_priority.get(x))

    ##### Step4: sort the SS per ranking columns and Put MFG hold orders at the back
    df_3a4.sort_values(by=ranking_col, ascending=True, inplace=True)
   # Put MFG hold orders at the back - the 3a4 here has no option so can also use mfg_hold directly alternatively
    df_hold = df_3a4[df_3a4.ADDRESSABLE_FLAG == 'MFG_HOLD'].copy()
    df_3a4 = df_3a4[df_3a4.ADDRESSABLE_FLAG != 'MFG_HOLD'].copy()
    df_3a4 = pd.concat([df_3a4, df_hold], sort=False)

    ##### Step5: create rank# and put in 3a4
    rank = {}
    order_list = df_3a4[order_col].unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, new_col] = df_3a4[order_col].map(lambda x: rank[x])

    return df_3a4


def ss_ranking_overall_new_december(df_3a4, ss_exceptional_priority, ranking_col, order_col='SO_SS', new_col='ss_overall_rank'):
    """
    根据priority_cat,OSSD,FCD, REVENUE_NON_REVENUE,C_UNSTAGED_QTY,按照ranking_col的顺序对SS进行排序。最后放MFG_HOLD订单.
    """

    # Below create a rev_rank for reference -  currently not used in overall ranking
    ### change non-rev orders unstaged $ to 0
    df_3a4.loc[:, 'C_UNSTAGED_DOLLARS'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'NO',
                                                   0,
                                                   df_3a4.C_UNSTAGED_DOLLARS)

    #### 生成ss_unstg_rev - 在这里不参与排序
    # 计算ss_unstg_rev
    ss_unstg_rev = {}
    df_rev = df_3a4.pivot_table(index='SO_SS', values='C_UNSTAGED_DOLLARS', aggfunc=sum)
    for ss, rev in zip(df_rev.index, df_rev.values):
        ss_unstg_rev[ss] = rev[0]
    df_3a4.loc[:, 'ss_unstg_rev'] = df_3a4.SO_SS.map(lambda x: ss_unstg_rev[x])

    """
    # 计算po_rev_unit - non revenue change to 0
    df_3a4.loc[:, 'po_rev_unit'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'YES',
                                            df_3a4.SOL_REVENUE / df_3a4.ORDERED_QUANTITY,
                                            0)

    # 计算ss_rev_unit: 通过po_rev_unit汇总
    ss_rev_unit = {}
    dfx_rev = df_3a4.pivot_table(index='SO_SS', values='po_rev_unit', aggfunc=sum)
    for ss, rev in zip(dfx_rev.index, dfx_rev.values):
        ss_rev_unit[ss] = rev[0]
    df_3a4.loc[:, 'ss_rev_unit'] = df_3a4.SO_SS.map(lambda x: int(ss_rev_unit[x]))
    """

    # create rank#
    rank = {}
    order_list = df_3a4.sort_values(by='ss_unstg_rev', ascending=False).SO_SS.unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, 'ss_rev_rank'] = df_3a4.SO_SS.map(lambda x: rank[x])

    # below creates overall ranking col
    ### Step1: 重新定义priority order及排序
    df_3a4.loc[:, 'priority_cat'] = np.where(df_3a4.SECONDARY_PRIORITY.isin(['PR1', 'PR2', 'PR3']),
                                             df_3a4.SECONDARY_PRIORITY,
                                             np.where(df_3a4.FINAL_ACTION_SUMMARY == 'LEVEL 4 ESCALATION PRESENT',
                                                      'L4',
                                                      np.where(df_3a4.BUP_RANK.notnull(),
                                                               'BUP',
                                                                np.where(df_3a4.PROGRAM.notnull(),
                                                                        'YE',
                                                                         None))))

    #### Update below DO/DX orders to PR1 due to current PR1/2/3 not updated when order change to DPAS from others
    df_3a4.loc[:, 'priority_cat'] = np.where(
        (df_3a4.DPAS_RATING.isin(['DO', 'DX', 'TAA-DO', 'TAA-DX'])) & (df_3a4.priority_cat.isnull()),
        'PR1',
        df_3a4.priority_cat)

    #### Step2: Generate rank for priority orders
    df_3a4.loc[:, 'priority_rank_top'] = np.where(df_3a4.priority_cat == 'PR1',
                                              1,
                                              np.where(df_3a4.priority_cat == 'PR2',
                                                       2,
                                                       np.where(df_3a4.priority_cat == 'PR3',
                                                                3,
                                                                None)))

    df_3a4.loc[:, 'priority_rank_mid'] =np.where(df_3a4.priority_cat == 'L4',
                                            4,
                                            np.where(df_3a4.priority_cat == 'BUP',
                                                    5,
                                                    np.where(df_3a4.priority_cat == 'YE',
                                                             6,
                                                             None)))

    #### update ranking based on exception priority setting
    df_3a4.loc[:, 'priority_rank_top'] = np.where(df_3a4.SO_SS.isin(ss_exceptional_priority['priority_top'].keys()),
                                                  df_3a4.SO_SS.map(lambda x: ss_exceptional_priority['priority_top'].get(x)),
                                                  np.where(df_3a4.SO_SS.isin(ss_exceptional_priority['priority_mid'].keys()),
                                                            None,
                                                            df_3a4.priority_rank_top))
    df_3a4.loc[:, 'priority_rank_mid'] = np.where(df_3a4.SO_SS.isin(ss_exceptional_priority['priority_mid'].keys()),
                                                  df_3a4.SO_SS.map(lambda x: ss_exceptional_priority['priority_mid'].get(x)),
                                                  np.where(df_3a4.SO_SS.isin(ss_exceptional_priority['priority_top'].keys()),
                                                            None,
                                                            df_3a4.priority_rank_mid))


    # Create a new col to indicate the rank - in ranking, actually use priority_rank_top and priority_rank_mid
    df_3a4.loc[:, 'priority_rank'] = np.where(df_3a4.priority_rank_top.notnull(),
                                              df_3a4.priority_rank_top,
                                              df_3a4.priority_rank_mid)

    ##### Step3: Give revenue/non-revenue a rank
    df_3a4.loc[:, 'rev_non_rev_rank'] = np.where(df_3a4.REVENUE_NON_REVENUE == 'YES', 0, 1)


    ##### Step4: sort the SS per ranking columns and Put MFG hold orders at the back
    df_3a4.sort_values(by=ranking_col, ascending=True, inplace=True)
    # Put MFG hold orders at the back - the 3a4 here has no option so can also use mfg_hold directly alternatively
    df_hold = df_3a4[df_3a4.ADDRESSABLE_FLAG == 'MFG_HOLD'].copy()
    df_3a4 = df_3a4[df_3a4.ADDRESSABLE_FLAG != 'MFG_HOLD'].copy()
    df_3a4 = pd.concat([df_3a4, df_hold], sort=False)

    ##### Step5: create rank# and put in 3a4
    rank = {}
    order_list = df_3a4[order_col].unique()
    for order, rk in zip(order_list, range(1, len(order_list) + 1)):
        rank[order] = rk
    df_3a4.loc[:, new_col] = df_3a4[order_col].map(lambda x: rank[x])

    return df_3a4

def df_pn_ct2r_date_judgement(df_ct2r, df_supply_df):
    """
    根据df_ct2r和df_supply_df判断假定df supply OK的日期
    :param df_ct2r:
    :param df_supply_df:
    :return:
    """
    pn_ct2r_lastsupply_max = {}
    max_date = df_supply_df.columns[-1]

    # 先把ct2r对应的PN及date加入字典 - 仅在ct2r date不大于supply df 中的最大日期时才添加（超出则认为没有supply）
    for row in df_ct2r.itertuples(index=True):
        # print(row)
        ct2r_date = pd.Timestamp.now().date() + pd.Timedelta(row.CT2R, 'd')
        if ct2r_date <= max_date:
            pn_ct2r_lastsupply_max[row.Index] = ct2r_date

    # iter df_supply_df,从后向前检查每一个PN的supply,如果>0的值对应的date>pn_ct2r_lastsupply_max,则将日期改成对应的date;否则不变
    df_with_ct2r_to_fix = df_supply_df.loc[pn_ct2r_lastsupply_max.keys(),:]  # df_ct2r中不应包含不存在于df_supply_df中的pn,否则报错
    date_col = df_with_ct2r_to_fix.columns

    for row in df_with_ct2r_to_fix.itertuples(index=True):
        row_data = row[::-1][:-1]  # 将每一行数据倒置检查
        # print(row[::-1])
        supply_index = next((-row_data.index(x) for x in row_data if x > 0), '0_supply')  # 找到>0的第一个值，存储相应的负值（倒数位置）

        if supply_index != '0_supply':  # 有>0的值，取大的date并加1
            pn_ct2r_lastsupply_max[row.Index] = max(date_col[supply_index - 1],
                                                    pn_ct2r_lastsupply_max[row.Index]) + pd.Timedelta(1, 'd')
        else:  # 没有>0的值，保留ct2r date并加1
            pn_ct2r_lastsupply_max[row.Index] = pn_ct2r_lastsupply_max[row.Index] + pd.Timedelta(1, 'd')

    return pn_ct2r_lastsupply_max


# 修改df_supply_df：对应日期赋值100000
def update_supply_for_df_w_ct2r(df_supply_df, pn_ct2r_lastsupply_max):
    for pn, date in pn_ct2r_lastsupply_max.items():
        df_supply_df.loc[pn, date] = 100000

    # 按日期排序
    col = df_supply_df.columns
    df_supply_df = df_supply_df[sorted(col)]

    return df_supply_df


def update_order_bom_to_3a4(df_3a4, df_order_bom,df_supply):
    """
    Add PN into 3a4 based on BOM; and remove the rows based on df_supply PN.
    :param df_3a4:
    :param df_bom:
    :return: df_3a4, df_missing_bom_pid
    """
    # add the BOM PN through merge method
    df_3a4 = pd.merge(df_3a4, df_order_bom, left_on='PO_NUMBER', right_on='PO_NUMBER', how='left')

    # remove the BOM_PN not in supply file
    df_3a4=df_3a4[df_3a4.BOM_PN.isin(df_supply.index.tolist())].copy()

    """
    # PID missing BOM data
    missing_bom_pid = df_3a4[df_3a4.TAN.notnull() & df_3a4.PN.isnull()].PRODUCT_ID.unique()
    df_missing_bom_pid = pd.DataFrame({'Missing BOM PID': missing_bom_pid})

    # 对于BOM missing 的采用3a4中已有的TAN
    df_3a4.loc[:, 'PN'] = np.where(df_3a4.TAN.notnull() & df_3a4.PN.isnull(),
                                   df_3a4.TAN,
                                   df_3a4.PN)
    """
    # correct the quantity by multiplying BOM Qty
    df_3a4.loc[:, 'C_UNSTAGED_QTY']=df_3a4.C_UNSTAGED_QTY * (df_3a4.TAN_QTY/df_3a4.ORDERED_QUANTITY)
    df_3a4.loc[:, 'ORDERED_QUANTITY'] = df_3a4.TAN_QTY

    # add indicator for distinct PO filtering
    df_3a4.loc[:,'distinct_po_filter']=np.where(~df_3a4.duplicated('PO_NUMBER'),
                                              'YES',
                                                '')


    return df_3a4



def resample_columns_and_agg_pastdue(df,method='W-SAT',agg_col_name='Current week',total_col=None,total_row=None,convert_num=False):
    """
    Resample the data based on method specified, aggregate the pastdue, add total
    :param df: data df to process
    :param method: 'W-MON' for weekly or 'M', etc.
    :return:
    """
    #重采样，并改回日期格式
    df.columns = pd.to_datetime(df.columns)
    df = df.resample(method, label='right',axis=1).sum() # when using 'W-SAT', it aggregates Sun to Sat, label as Sat
    df.columns=df.columns.map(lambda x:x.date() - pd.to_timedelta(6,'D'))  # -6 to use Start date as the label instead

   # 生成past_due并删除相应的日期列
    today=pd.Timestamp.now().date()
    today_name=pd.Timestamp.today().day_name()
    col_dates=list(df.columns)

    # cut off by Sat as current week
    if today_name=='Monday':
        offset=5
    elif today_name=='Tuesday':
        offset=4
    elif today_name=='Wednesday':
        offset=3
    elif today_name=='Thursday':
        offset=2
    elif today_name=='Friday':
        offset=1
    else:
        offset=0
    past_due_or_current_week_date = []
    for dt in col_dates:
        if dt<=today+pd.Timedelta(offset,'d'): # 过去及当前周
            past_due_or_current_week_date.append(dt)

    # 汇总过去及当前周
    df.loc[:, agg_col_name]=df[past_due_or_current_week_date].sum(axis=1)

    # remove the past due date columns& rearrange the columns in right sequence
    df.drop(past_due_or_current_week_date,axis=1,inplace=True)
    col_dates = list(df.columns)
    col_dates=[col_dates[-1]]+col_dates[:-1]
    df=df[col_dates]

    # add the totals based on requirement
    if total_col!=None:
        df.loc[:,total_col] = df.sum(axis=1)

    if total_row!=None:
        df.loc[total_row,:] = df.sum(axis=0)

    # convert the numbers based on requirement
    if convert_num==True:
        df=df.applymap(lambda x: round(x/1000000,1))
    else:
        df = df.applymap(lambda x: float(x))

    return df

def make_summary_build_projection(df_3a4,bu_list):
    """
    Create summary to indicate the build projection (against PO CTB date); also add in addressable backlog.
    :param df_3a4:
    :return:
    """
    # summarize addressable by BU or by PF
    if bu_list==['']:
        dfx=df_3a4[(df_3a4.ADDRESSABLE_FLAG == 'ADDRESSABLE')&(~df_3a4.duplicated('PO_NUMBER'))]
    else:
        dfx = df_3a4[(df_3a4.ADDRESSABLE_FLAG == 'ADDRESSABLE') & (~df_3a4.duplicated('PO_NUMBER')) & (
            df_3a4.BUSINESS_UNIT.isin(bu_list))]

    df_addr_bu = dfx.pivot_table(
        index=['ORGANIZATION_CODE', 'BUSINESS_UNIT'],
        values='C_UNSTAGED_DOLLARS',
        aggfunc=sum)

    df_addr_bu.reset_index(inplace=True)
    df_addr_bu.loc[:, 'PRODUCT_FAMILY'] = 'All PF'
    df_addr_bu.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'], inplace=True)
    df_addr_bu.loc[('Total', 'Total', 'Total')] = df_addr_bu.sum(axis=0)

    df_addr_pf = dfx.pivot_table(
        index=['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'],
        values='C_UNSTAGED_DOLLARS',
        aggfunc=sum)
    #合并
    df_addr = pd.concat([df_addr_bu, df_addr_pf])
    df_addr=df_addr.applymap(lambda x:round(x/1000000,1))
    df_addr.rename(columns={'C_UNSTAGED_DOLLARS': 'Addressable$'}, inplace=True)

    # make ctb summary by BU
    df_po_ctb_bu = df_3a4.drop_duplicates('PO_NUMBER').pivot_table(index=['ORGANIZATION_CODE','BUSINESS_UNIT'],
                            columns='po_ctb',
                            values='C_UNSTAGED_DOLLARS',
                            aggfunc=sum)
    # test
    #df_po_ctb_bu.to_excel('ctb summary.xlsx')

    df_po_ctb_bu.reset_index(inplace=True)
    df_po_ctb_bu.loc[:,'PRODUCT_FAMILY']='All PF'
    df_po_ctb_bu.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT','PRODUCT_FAMILY'],inplace=True)

    df_po_ctb_bu=resample_columns_and_agg_pastdue(df_po_ctb_bu,method='W-SAT',agg_col_name='Current week',
                                               total_col='Total',total_row=('Total','Total','Total'),convert_num=True)
    df_po_ctb_bu.sort_values(by=['ORGANIZATION_CODE', 'Current week'], ascending=False, inplace=True)

    # test
    #df_po_ctb_bu.to_excel('ctb summary-agg.xlsx')

    # merge addressable in
    df_po_ctb_bu=pd.merge(df_po_ctb_bu,df_addr,left_index=True,right_index=True,how='left',sort=False)
    df_po_ctb_bu.reset_index(inplace=True)
    df_po_ctb_bu.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY','Addressable$'], inplace=True)

    # make ctb summary by PF
    df_po_ctb_pf = df_3a4.drop_duplicates('PO_NUMBER').pivot_table(index=['ORGANIZATION_CODE', 'BUSINESS_UNIT','PRODUCT_FAMILY'],
                                                                columns='po_ctb',
                                                                values='C_UNSTAGED_DOLLARS',
                                                                aggfunc=sum)

    df_po_ctb_pf = resample_columns_and_agg_pastdue(df_po_ctb_pf, method='W-SAT', agg_col_name='Current week',
                                                    total_col='Total', total_row=None,
                                                    convert_num=True)
    df_po_ctb_pf.sort_values(by=['ORGANIZATION_CODE','BUSINESS_UNIT', 'Current week'], ascending=False, inplace=True)

    #merge addressable to ctb summary
    df_po_ctb_pf = pd.merge(df_po_ctb_pf, df_addr, left_index=True, right_index=True, how='left', sort=False)
    df_po_ctb_pf.reset_index(inplace=True)
    df_po_ctb_pf.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY', 'Addressable$'], inplace=True)

    # concat the two summaries: by BU,by PF ctb summaries
    df_sep=pd.DataFrame(columns=df_po_ctb_pf.reset_index().columns)
    df_sep.set_index(['ORGANIZATION_CODE','BUSINESS_UNIT', 'PRODUCT_FAMILY','Addressable$'],inplace=True)
    df_sep.loc[('Below for details', '', '',''), :] = '*****'
    df_sep.loc[('ORGANIZATION_CODE','BUSINESS_UNIT','PRODUCT_FAMILY','Addressable$'),:]=df_po_ctb_pf.columns
    df_build_projection=pd.concat([df_po_ctb_bu,df_sep,df_po_ctb_pf],sort=False)

    return df_build_projection

def make_summary_decommit_vs_improve(df_3a4):
    """
    Create summary to indicate the order pull in or decommit status against current FCD based on SO_SS
    """
    df_order_status = df_3a4.drop_duplicates('SO_SS').pivot_table(index=['ORGANIZATION_CODE','BUSINESS_UNIT','PRODUCT_FAMILY','ss_updated_status'],
                            columns='CURRENT_FCD_NBD_DATE', # this will leave out the orders without FCD date - unscheduled
                            values='ss_unstg_rev',
                            aggfunc=sum)

    df_order_status=resample_columns_and_agg_pastdue(df_order_status,method='W-SAT',agg_col_name='Current week + pastdue',
                                               total_col="Total",total_row=None,convert_num=True)

    return df_order_status



def make_summary_riso(df_3a4):
    """
    Create summary to indicate current RISO and to be RISO status based on SS_CTB. This is LT RISO against LT_TARGET_FCD
    :param df_3a4:
    :return:
    """
    dfx=df_3a4.drop_duplicates('SO_SS')
    df_riso_asis = dfx[dfx['RISO (as is)']=='Yes'].pivot_table(index=['ORGANIZATION_CODE','BUSINESS_UNIT'],
                            columns='LT_TARGET_FCD',
                            values='ss_unstg_rev',
                            aggfunc=sum)

    df_riso_asis.loc[:,'RISO']='RISO (as is)'

    df_riso_tobe = dfx[(dfx['RISO (to be)'] == 'Yes')].pivot_table(index=['ORGANIZATION_CODE','BUSINESS_UNIT'],
                                                                       columns='LT_TARGET_FCD',
                                                                       values='ss_unstg_rev',
                                                                       aggfunc=sum)
    df_riso_tobe.loc[:, 'RISO'] = 'RISO (to be)'

    df_riso=pd.concat([df_riso_asis,df_riso_tobe],sort=True)
    df_riso.reset_index(inplace=True)
    df_riso.set_index(['ORGANIZATION_CODE','BUSINESS_UNIT','RISO'],inplace=True)
    df_riso.sort_index(inplace=True)

    df_riso=resample_columns_and_agg_pastdue(df_riso,method='W-SAT',agg_col_name='Current week + pastdue',
                                               total_col="Total",total_row=None,convert_num=True)

    return df_riso



def change_supply_to_versionless_and_addup_supply(df_supply,pn_col='TAN'):
    """
    Change PN in supply  into versionless. Add up the qty into the versionless PN.
    :param df_supply:
    :param pn_col: name of the PN col. In Cm supply file it's PN, in Kinaxis file it's TAN.
    :return:
    """
    regex = re.compile(r'\d{2,3}-\d{4,7}')
    """
    df_supply.to_excel('df_supply.xlsx')
    for row in df_supply.itertuples(index=True):
        try:
            regex.search(row.Index).group()
        except:
            print(row.Index)
    """

    # convert to versionless
    try:
        df_supply.index = df_supply.index.map(lambda x: regex.search(x).group())
    except:
        print("Some TAN can't be regex'ed, considered as format error!")
        # write details to error_log.txt
        log_msg = '\n\nError regex supply file TAN! ' + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M') + '\n'
        with open(os.path.join(base_dir_logs, 'error_log.txt'), 'a+') as file_object:
            file_object.write(log_msg)
        traceback.print_exc(file=open(os.path.join(base_dir_logs, 'error_log.txt'), 'a+'))
        raise ValueError('Error regex the TAN in supply data.Stops!')

    # add up the duplicate PN (due to multiple versions)
    df_supply.sort_index(inplace=True)
    df_supply.reset_index(inplace=True)
    dup_pn = df_supply[df_supply.duplicated(pn_col)][pn_col].unique()
    df_sum = pd.DataFrame(columns=df_supply.columns)

    df_sum.set_index(pn_col, inplace=True)
    df_supply.set_index(pn_col, inplace=True)

    for pn in dup_pn:
        # print(df_supply[df_supply.PN==pn].sum(axis=1).sum())
        df_sum.loc[pn, :] = df_supply.loc[pn, :].sum(axis=0)

    df_supply.drop(dup_pn, axis=0, inplace=True)
    df_supply = pd.concat([df_supply, df_sum])

    return df_supply


def change_ct2r_to_versionless(df_ct2r):
    """
    Change PN in CT2R into versionless.
    :param df_ct2r:
    :return:
    """
    regex = re.compile(r'\d{2,3}-\d{4,7}')

    # change CT2R to versionless
    df_ct2r.index=df_ct2r.index.map(lambda x: regex.search(x).group())

    return df_ct2r


def exclude_short_ct2r_from_df_supply_and_df_ct2r(df_supply_df,df_ct2r,ct2r_threshold=5):
    """
    Exclude the short CT2R PN from the DF supply so they won't be considered. Do this with the versionless data.
    :param df_supply_df:
    :param ct2r_threshold: default using <=5 days
    :return:
    """

    short_ct2r_pn=df_ct2r[df_ct2r.CT2R<=ct2r_threshold].index.tolist()
    long_ct2r_pn=df_ct2r[df_ct2r.CT2R>ct2r_threshold].index.tolist()
    df_short_ct2r=df_supply_df.loc[short_ct2r_pn,:]

    # 去除重复的PN，否则会产生重复行
    long_ct2r_pn=set(long_ct2r_pn)
    df_supply_df=df_supply_df.loc[long_ct2r_pn,:].copy()
    df_ct2r=df_ct2r.loc[long_ct2r_pn,:].copy()

    return df_supply_df,df_ct2r, df_short_ct2r

def apply_transit_time_to_pcba_supply(df_supply_pcba,df_org,transit_time):
    """
    Apply transit time to PCBA supply: change from SCR commit to ETA.
    :param df_pcba:
    :param transit_time:
    :return:
    """
    ct2r=transit_time[df_org]
    current_col=df_supply_pcba.columns
    new_col=[col + pd.Timedelta(ct2r,'d') for col in current_col]
    df_supply_pcba.columns=new_col

    return df_supply_pcba

def remove_pcba_wrongly_included_in_df_supply(df_supply_pcba,df_supply_df):
    """
    Extra and temporary step for correction data: if PCBA are wrongly included in DF supply, they get removed.
    :param df_supply_pcba:
    :param df_supply_df:
    :return:
    """
    pcba_pn=df_supply_pcba.index
    df_pn=df_supply_df.index
    df_pn=np.setdiff1d(df_pn,pcba_pn)

    df_supply_df=df_supply_df.loc[df_pn,:].copy()

    return df_supply_df


def make_shortage_summary(df_3a4, col,type='revenue'):
    """
    Make summary to indicate the shortage material qty/impact$ against specified col (current FCD, LT target FCD, etc)
    :param df_3a4:
    :param col: date col to use as columns for pivot table
    :param type: revenue or qty summary
    :return:
    """

    if col == 'CURRENT_FCD_NBD_DATE':
        shortage_col = 'shortage_to_current_fcd'
        gating_col='top_gating_fcd'
    elif col == 'LT_TARGET_FCD':
        shortage_col = 'shortage_to_target_fcd'
        gating_col='top_gating_target_fcd'

    # duplicate PN (different options) may exist in same PO, using below to remove duplicates to avoid duplication
    # when using the C_UNSTAGED_DOLLARS to aggregate
    data_col = ['BUSINESS_UNIT', 'PO_NUMBER', 'BOM_PN', 'C_UNSTAGED_DOLLARS', col,shortage_col,gating_col]
    dfx = df_3a4[df_3a4[shortage_col] > 0][data_col].copy()

    if type=='revenue':
        dfx.drop_duplicates(['PO_NUMBER', 'BOM_PN'], keep='first', inplace=True)  # 对于一个PO有多个重复PN的情况，如果rev，去除重复
        value='C_UNSTAGED_DOLLARS'
        convert_rev_to_m=True
    elif type=='qty':
        value = shortage_col
        convert_rev_to_m=False

    df_shortage_summary = dfx.pivot_table(index=['BUSINESS_UNIT', 'BOM_PN',gating_col],
                                         columns=col,
                                         values=value,
                                         aggfunc=sum)


    # aggregate the summary
    df_shortage_summary = resample_columns_and_agg_pastdue(df_shortage_summary, method='W-SAT',
                                                          agg_col_name='Current week + pastdue',
                                                          total_col="Total", total_row=None, convert_num=convert_rev_to_m)

    #df_shortage_impact.sort_values(by='Total', ascending=False, inplace=True)

    return df_shortage_summary


def make_summary_shortage_material_qty(df_3a4, col):
    """
    Make summary to indicate the shortage qty against specified col (current FCD, LT target FCD, etc)
    :param df_3a4:
    :param col: date col to use as columns for pivot table
    :return:
    """

    if col == 'CURRENT_FCD_NBD_DATE':
        shortage_col = 'shortage_to_current_fcd'
        gating_col = 'top_gating_fcd'
    elif col == 'LT_TARGET_FCD':
        shortage_col = 'shortage_to_target_fcd'
        gating_col = 'top_gating_target_fcd'

    # duplicate PN (different options) may exist in same PO - keep as is and qty need to sum up
    data_col = ['BUSINESS_UNIT', 'PO_NUMBER', 'BOM_PN', 'C_UNSTAGED_DOLLARS', col, shortage_col, gating_col]
    dfx = df_3a4[df_3a4[shortage_col] > 0][data_col].copy()
    # dfx.drop_duplicates(['PO_NUMBER', 'BOM_PN'], keep='first', inplace=True)

    df_shortage_qty = dfx.pivot_table(index=['BUSINESS_UNIT', 'BOM_PN',gating_col],
                                      columns=col,
                                      values=shortage_col,
                                      aggfunc=sum)

    # Above pivot duplicated revenue when multiple PN exist in on PO. Use below to deduct

    df_shortage_qty = resample_columns_and_agg_pastdue(df_shortage_qty, method='W-SAT',
                                                       agg_col_name='Current week + pastdue',
                                                       total_col="Total", total_row=None, convert_num=False)

    df_shortage_qty.sort_values(by='Total', ascending=False, inplace=True)

    return df_shortage_qty


def make_sd_summary(df_3a4,df_supply,supply_source,date_col='min_date'):
    """
    Make supply/demand type of summary against LT target FCD or other specified date
    :param df_3a4:
    :param df_supply:
    :return:
    """
    if supply_source=='CM':
        pn_col='PN'
    elif supply_source=='Kinaxis':
        pn_col='TAN'

    # 可以根据以下date做不同的summary
    if date_col=='LT_TARGET_FCD':
        shortage_col='shortage_to_target_fcd'
    elif date_col=='min_date':
        shortage_col='shortage_to_min_date'

    # demand summary - 选出有缺料以及LT_TARGET_FCD（或其他）不为空的PN
    today = pd.Timestamp.today().date()
    pn_shortage=df_3a4[(df_3a4[shortage_col]>0)&(df_3a4[date_col]>today)].BOM_PN.unique()

    pn_demand_qty_combined=[]
    pn_demand_rev_combined = []
    print(pn_shortage)
    for pn in pn_shortage:
        # TODO: date_col. notnull() potentially miss out some demand wrongly scheduled or not schedule - need to further reivew how to deal with this
        dfx=df_3a4[df_3a4.BOM_PN==pn][['BOM_PN','C_UNSTAGED_QTY',date_col,'C_UNSTAGED_DOLLARS']]
        dfx.sort_values(by=date_col,inplace=True)
        date=dfx.iloc[0,2]
        qty=dfx.iloc[0,1]
        rev=dfx.iloc[0,3]
        date_list=[]
        qty_list=[]
        rev_list=[]

        if dfx.shape[0]==1:  # dfx只有一行记录
            date_list.append(date)
            qty_list.append(qty)
            rev_list.append(rev)
        else:
            row_num=1
            for row in dfx[1:].itertuples(index=False):
                row_num+=1
                if row[2]==date: #date_col - 日期相同，加总数量
                    qty+=row[1] #C_UNSTAGED_QTY 加总
                    rev+=row[3] # po unstaged rev
                    if row_num==dfx.shape[0]: # 如果已经是最后一行，计入list
                        date_list.append(date)
                        qty_list.append(qty)
                        rev_list.append(rev)
                else: # new date
                    date_list.append(date) # 将上一个date计入list
                    qty_list.append(qty)
                    rev_list.append(rev)

                    date=row[2]
                    qty=row[1]
                    rev=row[3]
                    if row_num==dfx.shape[0]: # 已经是最后一行,将新的date也计入list
                        date_list.append(date)
                        qty_list.append(qty)
                        rev_list.append(rev)

        #生成df并合并到list
        df_qty=pd.DataFrame(data=np.array(qty_list).reshape(1,-1),columns=date_list,index=[pn])
        df_rev = pd.DataFrame(data=np.array(rev_list).reshape(1, -1), columns=date_list, index=[pn])

        #print(df.shape,pn)
        pn_demand_qty_combined.append(df_qty)
        pn_demand_rev_combined.append(df_rev)

    # demand qty for all pn合并成df
    df_demand_qty_combined=pd.concat(pn_demand_qty_combined,sort=True)
    df_demand_qty_combined.loc[:,'Items']='Demand'
    df_demand_qty_combined.index.name='PN'
    df_demand_qty_combined.reset_index(inplace=True)
    df_demand_qty_combined.set_index(['PN','Items'],inplace=True)

    # demand rev for all pn合并成df
    df_demand_rev_combined = pd.concat(pn_demand_rev_combined, sort=True)
    df_demand_rev_combined.loc[:, 'Items'] = 'Demand_rev'
    df_demand_rev_combined.index.name = 'PN'
    df_demand_rev_combined.reset_index(inplace=True)
    df_demand_rev_combined.set_index(['PN', 'Items'], inplace=True)

    # supply
    df_supply.loc[:,'Items']='Supply'
    df_supply_short_pn=df_supply.loc[pn_shortage]
    df_supply_short_pn.reset_index(inplace=True)
    df_supply_short_pn.set_index([pn_col, 'Items'], inplace=True)

    # 合并demand和supply
    df_sd_combined=pd.concat([df_supply_short_pn,df_demand_qty_combined,df_demand_rev_combined],join='outer',sort=True)
    df_sd_combined.sort_index(inplace=True)

    # resample to weekly
    df_sd_combined=resample_columns_and_agg_pastdue(df_sd_combined, method='W-SAT', agg_col_name='Current week&pastdue',
                                                   total_col='Total', total_row=None,convert_num=False)

    # create the cum delta row
    pn_not_recovered=[] # PN that have not covered even consider current OH... previously shortage may happen in the past due
    for pn in pn_shortage:
        demand_qty=df_sd_combined.loc[(pn,'Demand'),:]
        demand_rev=df_sd_combined.loc[(pn,'Demand_rev'),:]
        supply = df_sd_combined.loc[(pn, 'Supply'), :]

        delta=supply-demand_qty
        cum_delta=delta.cumsum()

        cum_demand_qty=demand_qty.cumsum()
        cum_demand_rev=demand_rev.cumsum()
        cum_demand_qty=[x if x>0 else 1 for x in cum_demand_qty ] # 避免除数为0的情况
        cum_rev_impact=(cum_delta / cum_demand_qty) * cum_demand_rev
        cum_rev_impact=[round(x/1000000,1) if x<0 else None for x in cum_rev_impact] # 只取负值

        if not np.all(cum_delta>0):
            pn_not_recovered.append(pn)
            df_sd_combined.loc[(pn,'X_Cum_delta'),:]=cum_delta
            df_sd_combined.loc[(pn, 'X_Cum_impact(M$)'), :] = cum_rev_impact
            #df_sd_combined.loc[(pn, 'X_Cum_demand_qty'), :] = cum_demand_qty
            #df_sd_combined.loc[(pn, 'X_Cum_demand_rev'), :] = cum_demand_rev

    # 通过loc去除pastdue中缺但目前current week 已经cover的部分
    df_sd_combined=df_sd_combined.loc[(pn_not_recovered,['Demand','Supply','X_Cum_delta','X_Cum_impact(M$)']),:]
    #改变Items的名称
    df_sd_combined.sort_index(inplace=True)
    df_sd_combined.reset_index(inplace=True)
    df_sd_combined.Items=df_sd_combined.Items.map({'X_Cum_delta':'Cum_delta','X_Cum_impact(M$)':'Cum_impact(M$)','Demand':'Demand','Supply':'Supply'})
    df_sd_combined.set_index([pn_col,'Items'],inplace=True)

    return df_sd_combined

def generate_df_order_bom_from_flb_tan_col(df_3a4):
    """
    Generate the BOM usage file from the FLB_TAN col
    :param df_3a4:
    :return:
    """
    regex_pn = re.compile(r'\d{2,3}-\d{4,7}')
    regex_usage = re.compile(r'\([0-9.]+\)')

    df_flb_tan = df_3a4[df_3a4.FLB_TAN.notnull()][['PO_NUMBER','PRODUCT_ID','ORDERED_QUANTITY','FLB_TAN']].copy()
    #df_flb_tan.drop_duplicates(['PRODUCT_ID'], keep='first', inplace=True)

    po_list=[]
    pn_list = []
    usage_list = []
    for row in df_flb_tan.itertuples(index=False):
        po=row.PO_NUMBER
        flb_tan = row.FLB_TAN
        #order_qty = row.ORDERED_QUANTITY

        pn = regex_pn.findall(flb_tan)
        usage = regex_usage.findall(flb_tan)
        usage = [float(u[1:-1]) for u in usage]
        po_list += [po] * len(pn)
        pn_list += pn
        usage_list += usage
        """
        if len(pn)!=len(usage):# 检查错误
            print('Extracting FLB_TAN error:',po,'# of PN:',len(pn),'; # of usage:',len(usage))
        """
    df_order_bom_from_flb = pd.DataFrame({'PO_NUMBER': po_list, 'BOM_PN': pn_list, 'TAN_QTY': usage_list})

    return df_order_bom_from_flb


def exclude_pn_no_need_to_consider_from_kinaxis_supply(df_supply,class_code_exclusion):
    """
    Exlude supply based on class code exclusion setting.
    :param df_supply:
    :param class_code_exclusion:
    :return:
    """
    df_supply.reset_index(inplace=True)

    # Exclue PID - TAN的前10个字符中包含以上任意字母 (部分TAN后面带字母)
    letter_list=[x for x in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ']
    df_supply.loc[:,'temp']=df_supply.TAN.map(lambda x: 'yes' if any(letter in x[:10] for letter in letter_list) else 'no')

    df_supply.loc[:,'temp']=np.where(~df_supply.TAN.str.contains('-'),
                                     'yes',df_supply.temp)

    df_supply=df_supply[df_supply.temp=='no'].copy()

    #df_supply.to_excel('test.xlsx')
    df_supply.drop('temp', axis=1, inplace=True)

    # Exclude by class_code_exclusion
    pn_supply = df_supply.TAN.values
    pn_to_consider=[pn for pn in pn_supply if pn[:3] not in class_code_exclusion and pn[:4] not in class_code_exclusion]

    df_supply=df_supply[df_supply.TAN.isin(pn_to_consider)].copy()
    df_supply.set_index(['TAN'],inplace=True)

    return df_supply

def get_packed_or_cancelled_ss_from_3a4(df_3a4):
    """
    Get the fully packed or canceleld SS from 3a4 - for deleting exceptional priority smartsheet purpose.
    """
    ss_cancelled=df_3a4[df_3a4.ADDRESSABLE_FLAG=='PO_CANCELLED'].SO_SS.unique()

    ss_with_po_packed=df_3a4[df_3a4.PACKOUT_QUANTITY=='Packout Completed'].SO_SS.unique()
    ss_wo_po_packed = df_3a4[df_3a4.PACKOUT_QUANTITY != 'Packout Completed'].SO_SS.unique() # some PO may not be packed in one SS
    ss_fully_packed=np.setdiff1d(ss_with_po_packed,ss_wo_po_packed)

    ss_packed_not_cancelled=np.setdiff1d(ss_fully_packed,ss_cancelled)

    ss_cancelled_or_packed_3a4=ss_cancelled.tolist()+ss_packed_not_cancelled.tolist()

    return ss_cancelled_or_packed_3a4

def read_and_add_exception_po_to_3a4(df_3a4):
    """
    Read Exceptional PO (GIMS, Config, etc) from smartsheet and add to 3a4
    """
    # 从smartsheet读取backlog
    token = os.getenv('PRIORITY_TOKEN')
    sheet_id = os.getenv('EXCEPTION_ID')
    proxies = None  # for proxy server
    smartsheet_client = SmartSheetClient(token, proxies)
    df_smart = smartsheet_client.get_sheet_as_df(sheet_id, add_row_id=True, add_att_id=False)

    #df_smart.drop_duplicates(['PO_NUMBER'],keep='last',inplace=True)
    #df_smart=df_smart[df_smart.ACTION=='Issue Open'].copy()

    exception_po={}

    for row in df_smart.itertuples():
        exception_po[row.PO_NUMBER]=row.EXCEPTION_NAME

    df_3a4.loc[:,'EXCEPTION_NAME']=df_3a4.PO_NUMBER.map(lambda x: exception_po.get(x,None))

    return df_3a4


def read_backlog_priority_from_smartsheet(df_3a4,login_user):
    '''
    Read backlog priorities from smartsheet; remove SS showing packed/cancelled, or created by self but disappear from 34(if the org/BU also exist in 3a4.);
     create and segregate to top priority and mid priority
    :return:
    '''
    # 从smartsheet读取backlog
    token = os.getenv('PRIORITY_TOKEN')
    sheet_id = os.getenv('PRIORITY_ID')
    proxies = None  # for proxy server
    smartsheet_client = SmartSheetClient(token, proxies)
    df_smart = smartsheet_client.get_sheet_as_df(sheet_id, add_row_id=True, add_att_id=False)

    # Identify SS not in df_3a4 that can be removed - SS created by self and is disappeared from 3a4 - if the 3a4 include the org and BU
    df_smart_self=df_smart[df_smart['Created By']==login_user+'@cisco.com']
    df_smart_w_org_bu_in_3a4 = df_smart_self[
        (df_smart_self.ORG.isin(df_3a4.ORGANIZATION_CODE.unique())) & (df_smart_self.BU.isin(df_3a4.BUSINESS_UNIT.unique()))]
    ss_not_in_3a4 = np.setdiff1d(df_smart_w_org_bu_in_3a4.SO_SS.values, df_3a4.SO_SS.values)

    # SS showing as packed or cancelled in 3a4 - discard below
    #ss_cancelled_or_packed_3a4 = get_packed_or_cancelled_ss_from_3a4(df_3a4)

    # total ss to remove - discard below
    #df_removal = df_smart[(df_smart.SO_SS.isin(ss_cancelled_or_packed_3a4)) | (df_smart.SO_SS.isin(ss_not_in_3a4))]

    # create the priority dict
    df_smart.drop_duplicates('SO_SS', keep='last', inplace=True)
    df_smart = df_smart[(df_smart.SO_SS.notnull()) & (df_smart.Ranking.notnull())]
    ss_exceptional_priority = {}
    priority_top = {}
    priority_mid = {}
    for row in df_smart.itertuples():
        try: # in case error input of non-num ranking
            if float(row.Ranking)<4:
                priority_top[row.SO_SS] = float(row.Ranking)
            else:
                priority_mid[row.SO_SS] = float(row.Ranking)
        except:
            print('{} has a wrong ranking#: {}.'.format(row.SO_SS,row.Ranking) )

        ss_exceptional_priority['priority_top'] = priority_top
        ss_exceptional_priority['priority_mid'] = priority_mid

    return ss_exceptional_priority

def remove_priority_ss_from_smtsheet_and_notify(df_removal,login_user,sender='APJC DFPM'):
    """
    Remove the packed/cancelled SS from priority smartsheet and send email to corresponding people for whose SS are removed from the priority smartsheet
    """
    if df_removal.shape[0]>0:
        token = os.getenv('PRIORITY_TOKEN')
        sheet_id = os.getenv('PRIORITY_ID')
        proxies = None  # for proxy server
        smartsheet_client = SmartSheetClient(token, proxies)

        removal_row_id = df_removal.row_id.values.tolist()
        removal_ss_email = list(set(df_removal['Created By'].values.tolist()))
        if len(removal_row_id) > 0:
            smartsheet_client.delete_row(sheet_id=sheet_id, row_id=removal_row_id)

        to_address = removal_ss_email
        to_address = to_address + [login_user+'@cisco.com']
        bcc=['kwang2@cisco.com']
        html_template='priority_ss_removal_email.html'
        subject='SS auto removal from exceptional priority smartsheet - by {}'.format(login_user)

        send_attachment_and_embded_image(to_address, subject, html_template, att_filenames=None,
                                         embeded_filenames=None, sender=sender,bcc=bcc,
                                         removal_ss_header=df_removal.columns,
                                         removal_ss_details=df_removal.values,
                                         user=login_user)


def make_summary_fcd_vs_ctb(df_3a4):
    """
    Make summary by PF to indicate the difference between current FCD and PO_CTB date.
    """

    dfx = df_3a4[(df_3a4.distinct_po_filter == 'YES')].copy()

    # add in date_180 for unscheduled order to avoid potential error if all PO are unscheduled
    date_180 = pd.Timestamp.today().date() + pd.Timedelta(180, 'd')
    dfx.CURRENT_FCD_NBD_DATE.fillna(date_180,inplace=True)

    # create summary by FCD date
    df10 = dfx.pivot_table(index=['ORGANIZATION_CODE'], columns='CURRENT_FCD_NBD_DATE', values='C_UNSTAGED_DOLLARS',
                           aggfunc=sum).reset_index()
    df11 = dfx.pivot_table(index=['ORGANIZATION_CODE', 'BUSINESS_UNIT'], columns='CURRENT_FCD_NBD_DATE',
                           values='C_UNSTAGED_DOLLARS', aggfunc=sum).reset_index()
    df12 = dfx.pivot_table(index=['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'],
                           columns='CURRENT_FCD_NBD_DATE', values='C_UNSTAGED_DOLLARS', aggfunc=sum).reset_index()

    df1 = pd.concat([df10, df11, df12], sort=False)

    df1 = df1.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY']).sort_values(
        ['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'])

    df1 = df1.reset_index()

    df1.BUSINESS_UNIT.fillna('zTotal - all BU', inplace=True)

    df1.loc[:, 'PRODUCT_FAMILY'] = df1.apply(
        lambda x: 'zTotal - ' + x.BUSINESS_UNIT if pd.isnull(x.PRODUCT_FAMILY) else x.PRODUCT_FAMILY, axis=1)

    # create summary by po_ctb date
    df20 = dfx.pivot_table(index=['ORGANIZATION_CODE'], columns='po_ctb', values='C_UNSTAGED_DOLLARS',
                           aggfunc=sum).reset_index()
    df21 = dfx.pivot_table(index=['ORGANIZATION_CODE', 'BUSINESS_UNIT'], columns='po_ctb', values='C_UNSTAGED_DOLLARS',
                           aggfunc=sum).reset_index()
    df22 = dfx.pivot_table(index=['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'], columns='po_ctb',
                           values='C_UNSTAGED_DOLLARS', aggfunc=sum).reset_index()

    df2 = pd.concat([df20, df21, df22], sort=False)

    df2 = df2.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY']).sort_values(
        ['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'])

    df2 = df2.reset_index()

    df2.BUSINESS_UNIT.fillna('zTotal - all BU', inplace=True)

    df2.loc[:, 'PRODUCT_FAMILY'] = df2.apply(
        lambda x: 'zTotal - ' + x.BUSINESS_UNIT if pd.isnull(x.PRODUCT_FAMILY) else x.PRODUCT_FAMILY, axis=1)

    # Add FCD and CTB mark respectively and concat the df
    df1.loc[:, 'Item'] = 'FCD'
    df2.loc[:, 'Item'] = 'CTB'
    dfc = pd.concat([df1, df2], sort=False)

    dfc.reset_index(inplace=True)
    dfc.drop('index', axis=1, inplace=True)

    # Resample the data to weekly
    dfc.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY', 'Item'], inplace=True)
    dfc=resample_columns_and_agg_pastdue(dfc, method='W-SAT', agg_col_name='Current week', total_col=None, total_row=None,
                                     convert_num=True)

    dfc.loc[:, 'Total'] = dfc.sum(axis=1)
    # sort the data
    dfc.sort_values(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'], inplace=True)

    dfc.to_excel('test.xlsx')
    # calculate Cum delta between FCD and CTB
    dfc.fillna(0, inplace=True)
    for row in dfc.itertuples():
        if row.Index[3] == 'CTB':
            pre_col_value = 0
            for col in dfc.columns:
                if col != dfc.columns[-1]:
                    dfc.loc[(row.Index[0], row.Index[1], row.Index[2], 'Cum delta'), col] = pre_col_value + \
                                                                                            dfc.loc[(row.Index[0], row.Index[1], row.Index[2], 'CTB'), col] - \
                                                                                            dfc.loc[(row.Index[0], row.Index[1],row.Index[2], 'FCD'), col]
                    pre_col_value = dfc.loc[(row.Index[0], row.Index[1], row.Index[2], 'Cum delta'), col]
                else:
                    dfc.loc[(row.Index[0], row.Index[1], row.Index[2], 'Cum delta'), col] = pre_col_value

    # Sort the data by index
    dfc.sort_values(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY'], inplace=True)

    # correct the total label after sorting
    dfc.reset_index(inplace=True)
    dfc.loc[:, 'BUSINESS_UNIT'] = np.where(dfc.BUSINESS_UNIT == 'zTotal - all BU',
                                           'Total - all BU',
                                           dfc.BUSINESS_UNIT)
    dfc.loc[:, 'PRODUCT_FAMILY'] = np.where(dfc.PRODUCT_FAMILY == 'zTotal - zTotal - all BU',
                                            'Total - all PF',
                                            dfc.PRODUCT_FAMILY)
    dfc.loc[:, 'PRODUCT_FAMILY'] = np.where(dfc.PRODUCT_FAMILY.str.contains('zTotal'),
                                            dfc.PRODUCT_FAMILY.str[1:],
                                            dfc.PRODUCT_FAMILY)
    dfc.set_index(['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY', 'Item'], inplace=True)

    return dfc


def main_program_all(df_3a4,org_list, bu_list, description,ranking_col,df_supply,qend_list,output_col,login_user):
    """
    Consolidated main functions for this programs.
    :param df_3a4:
    :return:
    """

    qend=decide_qend_date(qend_list)
    # Do basic data processing for 3a4
    df_3a4 = basic_data_processing_3a4(df_3a4)

    # split out the 0 ordered quantity orders
    df_3a4, df_3a4_zero_qty = pick_out_zero_qty_order(df_3a4)

    # redefine addressable flag
    df_3a4=redefine_addressable_flag_main_pip_version(df_3a4)

    # read smartsheet priorities
    ss_exceptional_priority = read_backlog_priority_from_smartsheet(df_3a4, login_user)

    # Remove and send email notification for ss removal from exceptional priority smartsheet - discard below
    #remove_priority_ss_from_smtsheet_and_notify(df_removal, login_user, sender='APJC DF - auto CTB')

    # remove cancelled/packed orders - remove the record from 3a4 (in creating blg dict it's double removed - together with packed orders)
    df_3a4 = df_3a4[(df_3a4.ADDRESSABLE_FLAG != 'PO_CANCELLED') & (df_3a4.PACKOUT_QUANTITY != 'Packout Completed')].copy()

    # calculate the earliest packout date based on LT target fcd
    df_3a4=calculate_earliest_packable_date(df_3a4)

    # sort and rank the orders with overall priority (e.g.: ranking_col = ['priority_rank', 'partiak_rank','ss_rev_rank', 'min_date', 'SO_SS', 'PO_NUMBER'])
    df_3a4=ss_ranking_overall_new_december(df_3a4, ss_exceptional_priority, ranking_col,order_col='SO_SS', new_col='ss_overall_rank')

   # (do below after ranking)Add TAN into 3a4 based on BOM
    df_bom = generate_df_order_bom_from_flb_tan_col(df_3a4)
    df_3a4 = update_order_bom_to_3a4(df_3a4, df_bom,df_supply)

     #  生成supply_dic_tan
    supply_dic_tan = created_supply_dict_per_df_supply(df_supply)

    # create backlog dict for Tan require allocation
    blg_dic_tan = create_blg_dict_per_sorted_3a4_and_selected_tan(df_3a4, supply_dic_tan)

    # 生成订单分配supply的结果
    blg_with_allocation = allocate_supply_to_backlog_and_calculate_shortage(supply_dic_tan, blg_dic_tan)

    # 将分配结果加入到df_3a4中,并生成shortage columns (against target FCD and current FCD)
    df_3a4 = add_allocation_result_to_3a4(df_3a4, blg_with_allocation)
    #df_3a4[['PO_NUMBER','OPTION_NUMBER','PRODUCT_ID','PN','C_UNSTAGED_QTY','tan_supply_ready_date','supply_allocation']].to_excel('test.xlsx')

    # calculate the PO supply ready date and add to 3a4 - based on 3a4 及 supply中共同包含的tan
    pn_to_consider = np.intersect1d(df_supply.index.tolist(), df_3a4[df_3a4.C_UNSTAGED_QTY > 0].BOM_PN.unique().tolist())
    df_3a4 = calculate_po_supply_ready_date_and_add_to_3a4(df_3a4, pn_to_consider)
    # 去除不考虑的BOM_PM (不能去除 mark "yes" distinct_po_filter)
    #df_3a4=df_3a4[df_3a4.BOM_PN.isin(pn_to_consider)].copy()

    # TODO: allocate capacity to PO

    # calculate PO CTB date (curretly oly consider simple FLT in settings ... capcity not cosidered yet and need update)
    df_3a4 = calculate_po_ctb_in_3a4(df_3a4)

    # calculate SS CTB date -- used on pull in/decommit calculation
    df_3a4 = calculate_ss_ctb_and_add_to_3a4(df_3a4)

    # update_ss_status (pull in opportunity/recommit risk)
    df_3a4 = update_ss_status(df_3a4)

    # 在3a4中生成新列判断PN是否是top gating (分别针对target FCD 和 FCD)
    #df_3a4=identify_top_gating_pn(df_3a4)

    # calculate the as is (per FCD) and to be RISO status (per SS_CTB)
    df_3a4=calculate_riso_status(df_3a4)

    # make summaries
    df_po_ctb=make_summary_build_projection(df_3a4,bu_list)
    df_decommit_improve_summary=make_summary_decommit_vs_improve(df_3a4)
    #df_riso = make_summary_riso(df_3a4)
    df_fcd_ctb_summary=make_summary_fcd_vs_ctb(df_3a4)

    # make supply/demand summaries:for shortage PN
    #df_sd_combined_short=make_sd_summary(df_3a4, df_supply,supply_source, date_col='min_date')

    # make build impact summaries
    df_3a4,build_impact_summary_wk0,output_col=make_summary_build_impact(df_3a4, df_supply,output_col, qend,blg_with_allocation,FLT,cut_off='wk0')
    #df_3a4,build_impact_summary_wk1,output_col = make_summary_build_impact(df_3a4, df_supply,output_col,qend,blg_with_allocation,cut_off='wk1')
    #df_3a4,build_impact_summary_wk2,output_col = make_summary_build_impact(df_3a4, df_supply,output_col,qend,blg_with_allocation,cut_off='wk2')
    df_3a4, build_impact_qend,output_col = make_summary_build_impact(df_3a4, df_supply,output_col,qend,blg_with_allocation,FLT,cut_off='qend')

    # output the file
    data_to_write={
                    'build_projection':df_po_ctb,
                   'agg_fcd_vs_ctb':df_fcd_ctb_summary,
                  'ss_decommit_vs_pullin':df_decommit_improve_summary,
                   #'riso_status(vs LT target)':df_riso,
                   #'sd_shortage_pn(vs LT target)':df_sd_combined_short,
                   'build_impact_wk0':build_impact_summary_wk0,
                    #'build_impact_wk1':build_impact_summary_wk1,
                    #'build_impact_wk2':build_impact_summary_wk2,
                   'build_impact_qend': build_impact_qend,
                    #'shortage_impact(vs FCD)':df_shortage_impact_fcd,
                   # 'shortage_qty(vs FCD)':df_shortage_qty_fcd,
                   #'shortage_impact(vs LT target)': df_shortage_impact_lt_target_fcd,
                   #'shortage_qty(vs LT target)': df_shortage_qty_lt_target_fcd,
                    '3a4': df_3a4[output_col].set_index('ORGANIZATION_CODE'),
                    'supply': df_supply,
                   }
    # save the output file
    orgs='_'.join(org_list)
    bu='_'.join(bu_list)
    dt=(pd.Timestamp.now()+pd.Timedelta(hours=8)).strftime('%m-%d %Hh%Mm') #convert from server time to local
    if bu=='' and description=='':
        output_filename = orgs + ' CTB '+dt+ ' ' + login_user + '.xlsx'
    else:
        output_filename = orgs + ' CTB (' + bu + ' ' + description + ')' + dt + ' ' + login_user + '.xlsx'

    write_data_to_spreadsheet(base_dir_output, output_filename, data_to_write)

    del df_3a4, df_supply, df_bom,blg_with_allocation,supply_dic_tan,blg_dic_tan#,df_sd_combined_short
    gc.collect()

    return output_filename

def decide_qend_date(qend_list):
    """
    Decide which date is curent qend date based on predefined qend_list.
    :param qend_list:
    :return:
    """
    qend_list=pd.to_datetime(qend_list)
    today=pd.Timestamp.today().date()

    for dt in qend_list:
        if dt>today:
            qend=dt
            break

    return qend


def remove_blank_rows_in_supply(df_supply):
    """
    Some supply have double rows - Looks caused by RTF Flag (Y/N), one row may be blank. Use total ==0 to judge and remove.
    :param df_supply:
    :return:
    """
    df_supply.loc[:,'Total']=df_supply.sum(axis=1)
    df_supply=df_supply[df_supply.Total>0].copy()
    df_supply.drop('Total',axis=1,inplace=True)

    return df_supply

def read_supply_and_process(fname_supply,class_code_exclusion):
    """
    Read supply data, CT2R (for CM collected data), and exceptional backlog (input input), and processed into the
    final format after exclude the packaging and label class codes. Also save the Kinaxis file with org in filename.
    """
    df_supply=read_kinaxis_supply_and_initial_process(fname_supply)
    df_supply=remove_blank_rows_in_supply(df_supply)
    df_supply=exclude_pn_no_need_to_consider_from_kinaxis_supply(df_supply, class_code_exclusion)
    df_supply=change_supply_to_versionless_and_addup_supply(df_supply,pn_col='TAN')

    return df_supply


def check_input_file_format(fname_supply,fname_3a4):
    """
    Check the format of Kinaxis supply and 3a4, return the missing col names.
    :param supply_source:
    :param fname_supply:
    :param fname_3a4:
    :return:
    """
    df_3a4_heading=pd.read_csv(fname_3a4,nrows=2)
    col_3a4=df_3a4_heading.columns
    required_3a4_col=['SO_SS', 'PO_NUMBER', 'ORGANIZATION_CODE', 'BUSINESS_UNIT',
       'PRODUCT_FAMILY', 'PRODUCT_ID', 'TAN', 'MFG_HOLD', 'SECONDARY_PRIORITY',
       'GLOBAL_RANK', 'BUP_RANK', 'FINAL_ACTION_SUMMARY', 'ORDER_HOLDS',
       'ORDERED_QUANTITY', 'PACKOUT_QUANTITY', 'C_STAGED_QTY',
       'C_UNSTAGED_QTY', 'BUILD_COMPLETE_DATE', 'ORDERED_DATE', 'BOOKED_DATE',
       'LINE_CREATION_DATE', 'LT_TARGET_FCD', 'TARGET_SSD',
       'CURRENT_FCD_NBD_DATE', 'ORIGINAL_FCD_NBD_DATE',
       'CUSTOMER_REQUEST_DATE', 'CUSTOMER_REQUESTED_SHIP_DATE',
       'C_UNSTAGED_DOLLARS', 'SOL_REVENUE', 'C_STAGED_DOLLARS',
       'REVENUE_NON_REVENUE', 'DPAS_RATING', 'FALLOUT_CASE_DETAILS',
       'FALLOUT_MANAGEMENT', 'FLB_TAN', 'CTB_STATUS',
       'CURRENT_QUARTER_REVENUE_ELIGIBILITY', 'END_CUSTOMER_NAME',
       'SHIP_TO_CUSTOMER_NAME', 'COMMENTS','PROGRAM']

    missing_3a4_col=np.setdiff1d(required_3a4_col,col_3a4)

    required_supply_col=['TAN','ORG','Make/Buy','Past']
    df_supply_heading=pd.read_excel(fname_supply,nrows=2,header=1)
    col_supply=df_supply_heading.columns
    supply_format_correct=np.all(np.in1d(required_supply_col,col_supply))

    return missing_3a4_col, supply_format_correct



if __name__=='__main__':
    #supply_source = 'CM'
    supply_source='Kinaxis'

    # define the file path/names - web version use the uploaded file
    if supply_source=='CM':
        fname_supply = 'input_data/PCBADF SUPPLY June 12.xlsx'  # include both pcba and DF in differetn sheet
        fname_ct2r = 'input_data/CPN CT2R.xlsx'
    elif supply_source=='Kinaxis':
        fname_supply = 'input_data/kinaxis supply - Jun 29.xlsx'
        fname_ct2r=''

    ranking_col=['priority_rank','partial_rank', 'min_date', 'ss_rev_rank', 'SO_SS','PO_NUMBER']
    #ranking_col = ['priority_rank', 'ss_rev_rank', 'min_date', 'SO_SS', 'PO_NUMBER']

    fname_3a4 = 'input_data/backlog3a4-detail (15).csv'

    #fname_exception = 'input_data/Exception order Jun 5.xlsx'
    fname_exception=None

    org_list=['FOC']
    bu_list=['PABU','CRBU','ERBU']

    # exclude class code without consideration of supply and other actions
    class_code_exclusion = ['47-', '471-', '501-', '502-', '503-', '504-', '55-', '84-','83-']

    # 读取3a4,选择相关的org/bu,并添加exception
    df_3a4=read_3a4_and_limit_org_bu_and_add_exception(fname_3a4,bu_list,org_list,fname_exception=fname_exception)

    # 读取supply及相关数据并处理
    df_supply=read_supply_and_process(supply_source, fname_supply, fname_ct2r,class_code_exclusion)

    # Rank backlog，allocate supply, and make the summaries
    main_program_all(df_3a4,org_list, bu_list, ranking_col,df_supply,qend_list,output_col,login_user)
