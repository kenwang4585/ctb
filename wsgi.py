'''
Written by Ken, 2020
Flask web service app for CTB orchestration
'''

# add below matplotlib.use('Agg') to avoid this error: Assertion failed: (NSViewIsCurrentlyBuildingLayerTreeForDisplay()
# != currentlyBuildingLayerTree), function NSViewSetCurrentlyBuildingLayerTreeForDisplay
import matplotlib
matplotlib.use('Agg')

import time
from werkzeug.utils import secure_filename
from flask import flash,send_from_directory,render_template,redirect,request,url_for
from flask_settings import *
from functions import *
from setting import *
from db_add import add_user_log
from db_read import read_table
#from db_delete import delete_record
import traceback
#from flask_bootstrap import Bootstrap

#Bootstrap(app)



@app.route('/ctb_download', methods=['GET', 'POST'])
def ctb_download():
    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    allowed_user=['unknown','kwang2','anhao','cagong','hiung','julzhou','julwu','rachzhan','alecui','daidai','raeliu','karzheng']
    if login_user not in allowed_user:
        raise ValueError

    form = FileDownloadForm()

    # get file info
    output_record_hours = 360
    upload_record_hours = 240
    df_output = get_file_info_on_drive(base_dir_output, keep_hours=output_record_hours)
    df_upload = get_file_info_on_drive(base_dir_upload, keep_hours=upload_record_hours)


    if form.validate_on_submit():
        fname = form.file_name_delete.data
        if login_user in fname:
            if fname in df_output.File_name.values:
                f_path = df_output[df_output.File_name == fname].File_path.values[0]
                os.remove(f_path)
                msg = '{} removed!'.format(fname)
                flash(msg, 'success')
            elif fname in df_upload.File_name.values:
                f_path = df_upload[df_upload.File_name == fname].File_path.values[0]
                os.remove(f_path)
                msg = '{} removed!'.format(fname)
                flash(msg, 'success')
            else:
                add_user_log(user=login_user, location='Download', user_action='Delete file',
                             summary='Fail: {}'.format(fname))
                msg = 'Verify file name you put in and ensure a correct file name here: {}'.format(fname)
                flash(msg, 'warning')
                return redirect(url_for('ctb_download',_external=True,_scheme='http',viewarg1=1))
            add_user_log(user=login_user, location='Download', user_action='Delete file',
                         summary='Success: {}'.format(fname))
        else:
            msg = 'You are not allowed to delete this file created by others: {}'.format(fname)
            flash(msg, 'warning')
            return redirect(url_for('ctb_download',_external=True,_scheme='http',viewarg1=1))


    return render_template('ctb_download.html',form=form,
                           files_output=df_output.values,
                           output_record_days=int(output_record_hours / 24),
                           files_uploaded=df_upload.values,
                           upload_record_days=int(upload_record_hours / 24),
                           user=login_name)

#@app.route('/ctb_about', methods=['GET', 'POST'])
#def ctb_about():
#    login_user=request.headers.get('Oidc-Claim-Sub')
#    login_name = request.headers.get('Oidc-Claim-Fullname')
#    if login_user == None:
#        login_user = ''
#        login_name = ''

#    return render_template('ctb_about.html',user=login_name)

@app.route('/ctb', methods=['GET', 'POST'])
def ctb_run():
    form = UploadForm()
    # as these email valiable are redefined below in email_to_only check, thus have to use global to define here in advance
    # otherwise can't be used. (as we create new vaiables with _ suffix thus no need to set global variable)
    # global backlog_dashboard_emails
    program_log = []
    user_selection = []
    time_details=[]

    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    allowed_user=['unknown','kwang2','anhao','cagong','hiung','julzhou','julwu','rachzhan','alecui','daidai','raeliu','karzheng']
    if login_user not in allowed_user:
        raise ValueError

    if login_user!='unknown' and login_user!='kwang2':
        add_user_log(user=login_user, location='Run', user_action='Visit', summary='')

    if form.validate_on_submit():
        start_time = pd.Timestamp.now()
        print('start to run: {}'.format(start_time.strftime('%Y-%m-%d %H:%M')))
        log_msg = []
        log_msg.append('\n\n[' + login_user + '] ' + start_time.strftime('%Y-%m-%d %H:%M'))

        # 通过条件判断及邮件赋值，开始执行任务
        org=form.org.data # currently only support one org one time
        org_list=org.strip().upper().split('/')

        bu=form.bu.data
        bu_list=bu.strip().upper().split('/')

        log_msg.append('Org: ' + org)
        log_msg.append('BU: ' + bu)
        class_code_exclusion=form.class_code_exclusion.data
        class_code_exclusion=class_code_exclusion.strip().split('/')
        class_code_exclusion=[x+'-' for x in class_code_exclusion]

        f_3a4 = form.file_3a4.data
        f_supply= form.file_supply.data

        # 存储文件 - will save again with Org name in file name later
        #file_path_3a4 = os.path.join(app.config['UPLOAD_PATH'],'3a4.csv')
        #file_path_supply = os.path.join(app.config['UPLOAD_PATH'],'supply.xlsx')
        file_path_3a4 = os.path.join(base_dir_upload, login_user + '_'+ secure_filename(f_3a4.filename))
        file_path_supply = os.path.join(base_dir_upload, login_user + '_'+ secure_filename(f_supply.filename))

        f_3a4.save(file_path_3a4)
        f_supply.save(file_path_supply)

        # 判断并定义ranking_col
        ranking_col = ranking_col_cust

        try:
            # check file format by reading headers
            module='checking file format'
            missing_3a4_col, supply_format_correct = check_input_file_format(file_path_supply,file_path_3a4)
            if len(missing_3a4_col) >0:
                flash('Check your 3a4 file! Missing columns in the file you uploaded: {}'.format(missing_3a4_col),'warning')
                return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))
            if not supply_format_correct:
                flash('Supply file format error! Ensure header in 2nd row which is default format of Kinaxis file.','warning')
                return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))

            # 读取3a4,选择相关的org/bu
            module='read_3a4_and_limit_org_bu'
            df_3a4 = read_3a4_and_limit_org_bu(file_path_3a4, bu_list, org_list)

            # 读取Exceptional PO from smartsheet and add in 3a4
            df_3a4 = read_and_add_exception_po_to_3a4(df_3a4)

            # 读取supply及相关数据并处理
            module = 'read_supply_and_process'
            df_supply = read_supply_and_process(file_path_supply,class_code_exclusion)

            # Rank backlog，allocate supply, and make the summaries
            module='main_program_all'
            output_filename=main_program_all(df_3a4, org_list,bu_list,ranking_col, df_supply, qend_list, output_col,login_user)
            flash('CTB file created:{}! You can download accordingly.'.format(output_filename), 'success')

            finish_time = pd.Timestamp.now()
            processing_time = round((finish_time - start_time).total_seconds() / 60, 1)
            log_msg.append('Processing time: ' + str(processing_time) + ' min')
            print('Finish run:',finish_time.strftime('%Y-%m-%d %H:%M'))

            # Write the log file
            summary='; '.join(log_msg)
            add_user_log(user=login_user, location='Run', user_action='Run CTB', summary=summary)

        except Exception as e:
            try:
                del df_supply, df_3a4
                gc.collect()
            except:
                pass

            print(module,': ', e)
            traceback.print_exc()
            flash('Error encountered in module : {} - {}'.format(module,e),'warning')
            # Write the log file
            summary = 'Error: ' + str(e)
            add_user_log(user=login_user, location='Run', user_action='Run CTB', summary=summary)

            # write details to error_log.txt
            log_msg = '\n\n' + login_user + ' ' + pd.Timestamp.now().strftime('%Y-%m-%d %H:%M') + '\n'
            with open(os.path.join(base_dir_logs, 'error_log.txt'), 'a+') as file_object:
                file_object.write(log_msg)
            traceback.print_exc(file=open(os.path.join(base_dir_logs, 'error_log.txt'), 'a+'))

        # clear memory
        try:
            del df_supply,df_3a4
            gc.collect()
        except:
            pass

        return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))

    return render_template('ctb_run.html', form=form,user=login_name)


@app.route('/o/<filename>',methods=['GET'])
def download_file_output(filename):
    f_path=base_dir_output
    print(f_path)
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != None:
        add_user_log(user=login_user, location='Download', user_action='Download file',
                 summary=filename)
    return send_from_directory(f_path, filename=filename, as_attachment=True)

@app.route('/u/<filename>',methods=['GET'])
def download_file_upload(filename):
    f_path=base_dir_upload
    print(f_path)
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != None:
        add_user_log(user=login_user, location='Download', user_action='Download file',
                 summary=filename)
    return send_from_directory(f_path, filename=filename, as_attachment=True)

@app.route('/l/<filename>',methods=['GET'])
def download_file_logs(filename):
    f_path=base_dir_logs
    print(f_path)
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != None:
        add_user_log(user=login_user, location='Download', user_action='Download file',
                 summary=filename)
    return send_from_directory(f_path, filename=filename, as_attachment=True)



@app.route('/admin', methods=['GET','POST'])
def ctb_admin():
    form = AdminForm()
    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    allowed_user=['unknown','kwang2','anhao','cagong','hiung','julzhou','julwu','rachzhan','alecui','daidai','raeliu','karzheng']
    if login_user not in allowed_user:
        raise ValueError

    if login_user!='unknown' and login_user!='kwang2':
        return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))
        add_user_log(user=login_user, location='Admin', user_action='Visit', summary='Why happens?')

    # get file info
    output_record_hours=360
    upload_record_hours=240
    df_output=get_file_info_on_drive(base_dir_output,keep_hours=output_record_hours)
    df_upload=get_file_info_on_drive(base_dir_upload,keep_hours=upload_record_hours)
    df_logs=get_file_info_on_drive(base_dir_logs,keep_hours=10000)

    # read logs
    df_log_detail = read_table('user_log')
    df_log_detail.sort_values(by=['DATE','TIME'],ascending=False,inplace=True)

    if form.validate_on_submit():
        fname=form.file_name_delete.data
        print(fname)
        if fname in df_output.File_name.values:
            f_path=df_output[df_output.File_name==fname].File_path.values[0]
            os.remove(f_path)
            msg='{} removed!'.format(fname)
            flash(msg,'success')
        elif fname in df_upload.File_name.values:
            f_path = df_upload[df_upload.File_name == fname].File_path.values[0]
            os.remove(f_path)
            msg = '{} removed!'.format(fname)
            flash(msg, 'success')
        else:
            msg = 'Error file name! Ensure it is in output folder,upload folder or supply folder: {}'.format(fname)
            flash(msg, 'warning')
            return redirect(url_for('ctb_admin',_external=True,_scheme='http',viewarg1=1))

    return render_template('ctb_admin.html',form=form,
                           files_output=df_output.values,
                           files_uploaded=df_upload.values,
                           files_log=df_logs.values,
                           log_details=df_log_detail.values,
                           user=login_name)


@app.route('/resume', methods=['GET'])
def resume():
    return render_template('resume.html')
