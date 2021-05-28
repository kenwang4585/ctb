import os
import getpass

allowed_user = ['unknown', 'kwang2', 'anhao', 'cagong', 'hiung', 'julzhou', 'julwu', 'rachzhan', 'alecui', 'daidai',
                'raeliu', 'karzheng','seyang']


# Quarter end cutoff
qend_list=['2020-7-26','2020-10-24','2021-1-23','2021-5-1','2021-7-31']

# Use 2 days as FLT for all
FLT = 2

# PCBA to DF transit time - kinnaxis supply not use this-
transit_time={'FOC':2, # from FOL to FOC
              'FTX':7, # from FOL to FTX
              }

required_kinaxis_supply_col=['TAN','ORG','Make/Buy','Past']

required_3a4_col=['SO_SS', 'PO_NUMBER', 'ORGANIZATION_CODE', 'BUSINESS_UNIT',
       'PRODUCT_FAMILY', 'PRODUCT_ID', 'TAN', 'MFG_HOLD', 'SECONDARY_PRIORITY',
       'GLOBAL_RANK', 'BUP_RANK', 'FINAL_ACTION_SUMMARY', 'ORDER_HOLDS',
       'ORDERED_QUANTITY', 'PACKOUT_QUANTITY', 'C_STAGED_QTY',
       'C_UNSTAGED_QTY', 'BUILD_COMPLETE_DATE', 'ORDERED_DATE', 'BOOKED_DATE',
       'LINE_CREATION_DATE', 'LT_TARGET_FCD', 'TARGET_SSD',
       'CURRENT_FCD_NBD_DATE', 'ORIGINAL_FCD_NBD_DATE',
       'CUSTOMER_REQUEST_DATE', 'CUSTOMER_REQUESTED_SHIP_DATE',
       'C_UNSTAGED_DOLLARS', 'SOL_REVENUE', 'C_STAGED_DOLLARS',
       'REVENUE_NON_REVENUE', 'DPAS_RATING', 'FLB_TAN', 'CTB_STATUS','PROGRAM']

# below are the base col...some additional ones related to revenue impact will be added.
output_col=['ORGANIZATION_CODE', 'BUSINESS_UNIT', 'PRODUCT_FAMILY', 'SO_SS', 'PO_NUMBER','distinct_po_filter', 'EXCEPTION_NAME','PRODUCT_ID',
           'BOM_PN','ADDRESSABLE_FLAG','priority_cat','priority_rank','ss_overall_rank','LINE_CREATION_DATE',
           'CURRENT_FCD_NBD_DATE','ORIGINAL_FCD_NBD_DATE', 'LT_TARGET_FCD','min_date','CUSTOMER_REQUESTED_SHIP_DATE','C_UNSTAGED_DOLLARS',
            'ORDERED_QUANTITY','C_UNSTAGED_QTY','PACKOUT_QUANTITY','supply_allocation','tan_qty_wo_supply',
            'tan_supply_ready_date','po_supply_ready_date', 'earliest_allowed_pack_date','earliest_allowed_pack_date_factor','po_ctb', 'ctb_comment','ss_ctb',
            'ss_updated_status', 'CTB_STATUS(CTB_UI)','GLOBAL_RANK','RISO (as is)','RISO (to be)']

# used temporarily when output testing file
test_col=['PO_NUMBER','OPTION_NUMBER','PRODUCT_ID','PN','C_UNSTAGED_QTY','tan_supply_ready_date','supply_allocation']

# hold types
mfg_holds=['Booking Validation Hold','Cancellation','CFOP Product Hold','CMFS-Credit Check Pending','CMFS-Scheduled, Booked',
 'Compliance Hold','CONDITIONAL HOLD','Config Problem Hold','Configuration Hold','Conversion Dispatch Hold',
 'CSC-Credit Check Pending','CSC-Not Scheduled, Booked','Export','Localization Change Hold','New Product',
 'Non-FCC Compliant Hold','Order Aging Hold','Order Change','Order Transfer Changes (OTC) Hold','Order Validation Hold',
 'Pending Trade Collaborator Response','Quantity Validation','SCORE Chg Parameter','Scheduling COO','TCH Order Validation',
           'Country Certification Hold']

#Logic to be clarified: both below should be met (take the min value of below)
#addressable_window={
#                    'LT_TARGET_FCD':28, # need to change code if these 2 dates change
#                    'TARGET_SSD':21
#                    }

addressable_window = 30

if getpass.getuser()=='ubuntu': # if it's on crate server
    base_dir_output = '/home/ubuntu/ctb_output'
    base_dir_upload='/home/ubuntu/upload_file'
    base_dir_trash='/home/ubuntu/trash_file'
    base_dir_logs = '/home/ubuntu/logs'
    db_uri='sqlite:////home/ubuntu/database/foo.db'
else:
    base_dir_output = os.path.join(os.getcwd(),'ctb_output')
    base_dir_upload = os.path.join(os.getcwd(),'upload_file')
    base_dir_trash = os.path.join(os.getcwd(),'trash_file')
    base_dir_logs = os.path.join(os.getcwd(), 'logs')
    db_uri= 'sqlite:///' + os.path.join(os.getcwd(), 'database') + '/foo.db'


# rank sequences
ranking_col_cust = ['priority_rank_top', 'CURRENT_FCD_NBD_DATE', 'priority_rank_mid','ORIGINAL_FCD_NBD_DATE',
                           'C_UNSTAGED_QTY', 'rev_non_rev_rank', 'SO_SS', 'PO_NUMBER']
