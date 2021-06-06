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
from db_add import *
from db_read import *
#from db_delete import delete_record
import traceback
#from flask_bootstrap import Bootstrap

#Bootstrap(app)



@app.route('/ctb_result', methods=['GET'])
def ctb_result():
    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    if login_user not in allowed_user:
        add_log_summary(user=login_user, location='Result', user_action='Visit-trying', summary='')
        raise ValueError

    form = FileDownloadForm()

    # get file info
    output_record_hours = 360
    upload_record_hours = 240
    trash_record_hours = 240
    df_output = get_file_info_on_drive(base_dir_output, keep_hours=output_record_hours)
    df_upload = get_file_info_on_drive(base_dir_upload, keep_hours=upload_record_hours)
    df_trash = get_file_info_on_drive(base_dir_trash, keep_hours=trash_record_hours)

    return render_template('ctb_result.html',form=form,
                           files_output=df_output.values,
                           output_record_days=int(output_record_hours / 24),
                           files_uploaded=df_upload.values,
                           upload_record_days=int(upload_record_hours / 24),
                           files_trash=df_trash.values,
                           trash_record_days=int(trash_record_hours / 24),
                           user=login_user,
                           login_user=login_user)

@app.route('/user-guide')
def user_guide():
    login_user = request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    return render_template('ctb_userguide.html',user=login_user, subtitle=' - FAQ')

@app.route('/ctb', methods=['GET', 'POST'])
def ctb_run():
    form = UploadForm()
    # as these email valiable are redefined below in email_to_only check, thus have to use global to define here in advance
    # otherwise can't be used. (as we create new vaiables with _ suffix thus no need to set global variable)
    # global backlog_dashboard_emails

    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    if login_user not in allowed_user:
        add_log_summary(user=login_user, location='Home-RUN', user_action='Visit-trying', summary='')
        raise ValueError

    if login_user not in ['kwang2','unknown']:
        add_log_summary(user=login_user, location='Home-RUN', user_action='Visit', summary='')

    if form.validate_on_submit():
        log_msg_main = []
        start_time = pd.Timestamp.now()
        print('start to run for {}: {}'.format(login_user,start_time.strftime('%Y-%m-%d %H:%M')))

        log_msg = '\n\n[' + login_user + '] ' + start_time.strftime('%Y-%m-%d %H:%M')
        add_log_details(msg=log_msg)

        # 通过条件判断及邮件赋值，开始执行任务
        org=form.org.data.strip().upper() # currently only support one org one time
        #org_list=org.strip().upper().split('/')

        bu=form.bu.data
        bu_list=bu.strip().upper().split('/')

        description=form.description.data.strip()

        class_code_exclusion=form.class_code_exclusion.data
        class_code_exclusion=class_code_exclusion.strip().split('/')
        class_code_exclusion=[x+'-' for x in class_code_exclusion]
        log_msg_main.append('Org: {}; BU: {}; Exclusion code: {}'.format(org,bu,class_code_exclusion))
        log_msg = '\nOrg: {}; BU: {}; \nexclusion class: {}'.format(org,bu, class_code_exclusion)
        add_log_details(msg=log_msg)

        f_3a4 = form.file_3a4.data
        f_kinaxis_supply= form.file_kinaxis_supply.data
        f_allocation_supply = form.file_allocation_supply.data

        # check and save supply file
        if f_kinaxis_supply.filename=='' and f_allocation_supply.filename=='':
            msg = 'Pls upload either or both of the supply file: Kinaxis supply file, PCBA allocation supply file(s).'
            flash(msg,'warning')
            return render_template('ctb_run.html', form=form, user=login_user)

        if f_kinaxis_supply.filename!='':
            file_path_kinaxis_supply = os.path.join(base_dir_upload, login_user + '_'+ secure_filename(f_kinaxis_supply.filename))
            f_kinaxis_supply.save(file_path_kinaxis_supply)
            add_log_details(msg='\nSupply Kinaxis: ' + f_kinaxis_supply.filename)
            log_msg_main.append('Kinaxis supply')

        if f_allocation_supply.filename!='':
            file_path_allocation_supply = os.path.join(base_dir_upload,login_user + '_' + secure_filename(f_allocation_supply.filename))
            f_allocation_supply.save(file_path_allocation_supply)
            add_log_details(msg='\nSupply Kinaxis: ' + f_allocation_supply.filename)
            log_msg_main.append('PCBA allocation supply')

        # save 3a4 file
        file_path_3a4 = os.path.join(base_dir_upload, login_user + '_'+ secure_filename(f_3a4.filename))
        f_3a4.save(file_path_3a4)
        add_log_details(msg='\n3a4: ' + f_3a4.filename)

        # 判断并定义ranking_col
        ranking_col = ranking_col_cust

        try:
            # read 3a4 and check format
            df_3a4,error_msg=read_3a4_and_check_format(file_path_3a4, required_3a4_col)
            if error_msg!='':
                add_log_details(msg=error_msg)
                flash(error_msg,'warning')
                return redirect(url_for('ctb_run', _external=True, _scheme='http', viewarg1=1))

            # read Kinaxis supply and check format
            if f_kinaxis_supply.filename!='':
                df_supply_kinaxis, error_msg=read_kinaxis_supply_and_check_format(file_path_kinaxis_supply, required_kinaxis_supply_col)
                if error_msg != '':
                    add_log_details(msg=error_msg)
                    flash(error_msg, 'warning')
                    return redirect(url_for('ctb_run', _external=True, _scheme='http', viewarg1=1))

                # 处理 kinaxis supply data (oh_date used in allocation supply OH to use same date)
                df_supply_kinaxis = process_kinaxis_supply(df_supply_kinaxis, class_code_exclusion)
            else:
                df_supply_kinaxis=pd.DataFrame()

            # read pcba allocation supply and check format
            # TODO: update inside of read_pcba_allocation.... multi-files case with multi-allocation files
            if f_allocation_supply.filename!='':
                df_supply_allocation, df_supply_allocation_transit, df_supply_tan_transit_time, error_msg=read_pcba_allocation_supply_and_check_format(file_path_allocation_supply)
                if error_msg != '':
                    add_log_details(msg=error_msg)
                    flash(error_msg, 'warning')
                    return redirect(url_for('ctb_run', _external=True, _scheme='http', viewarg1=1))

                # process supply and update into kinaxis supply file (Replace by TAN)
                df_supply_allocation_combined=consolidate_pcba_allocation_supply(df_supply_allocation, df_supply_tan_transit_time,
                                               df_supply_allocation_transit, org)

            else:
                df_supply_allocation_combined=pd.DataFrame()

            # concat allocation PCBA data to Kinaxis data (remove same TAN from kinaxis file if exist)
            df_supply = consolidate_allocated_pcba_and_kinaxis(df_supply_allocation_combined, df_supply_kinaxis)

            # 选择相关的org/bu
            df_3a4 = limit_3a4_org_and_bu(df_3a4,org,bu_list)

            # Rank backlog，allocate supply, and make the summaries
            output_filename=main_program_all(df_3a4, org,bu_list,description, ranking_col, df_supply, qend_list, output_col,login_user)
            log_msg='CTB file created:{}! You can download accordingly.'.format(output_filename)
            flash(log_msg, 'success')
            add_log_details(msg=log_msg)

            finish_time = pd.Timestamp.now()
            processing_time = round((finish_time - start_time).total_seconds() / 60, 1)
            log_msg='\nProcessing time: ' + str(processing_time) + ' min'
            log_msg_main.append(log_msg)
            add_log_details(msg=log_msg)
            print('Finish run for {}: {}'.format(login_user,finish_time.strftime('%Y-%m-%d %H:%M')))

            # Write the log summary
            summary='; '.join(log_msg_main)
            add_log_summary(user=login_user, location='Run', user_action='Run CTB', summary=summary)

            # clear memory
            try:
                del df_3a4, df_supply_kinaxis, df_supply, df_supply_allocation_combined
                gc.collect()
            except:
                pass
        except Exception as e:
            try:
                del df_3a4, df_supply_kinaxis, df_supply, df_supply_allocation_combined
                gc.collect()
            except:
                pass

            traceback.print_exc()
            flash(e,'warning')
            # Write the log file
            log_msg_main.append('Error: ' + str(e))
            summary='; '.join(log_msg_main)
            add_log_summary(user=login_user, location='Run', user_action='Run CTB - error', summary=summary)

            # write error log details
            traceback.print_exc(file=open(os.path.join(base_dir_logs, 'log_details.txt'), 'a+'))

        return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))

    return render_template('ctb_run.html', form=form,user=login_user)


@app.route('/o/<login_user>/<filename>',methods=['GET'])
def delete_file_output(login_user,filename):
    if login_user == 'unknown':
        http_scheme = 'http'
    else:
        http_scheme = 'https'

    if login_user in filename:
        os.rename(os.path.join(base_dir_output,filename),os.path.join(base_dir_trash,filename))
        msg='File removed: {}'.format(filename)
        flash(msg,'success')
    else:
        msg='You can only delete file created by yourself!'
        flash(msg,'warning')

    return redirect(url_for("ctb_result", _external=True, _scheme='http', viewarg1=1))

@app.route('/u/<login_user>/<filename>',methods=['GET'])
def delete_file_upload(login_user,filename):
    if login_user in filename:
        os.rename(os.path.join(base_dir_upload,filename),os.path.join(base_dir_trash,filename))
        msg='File removed: {}'.format(filename)
        flash(msg,'success')
    else:
        msg='You can only delete file uploaded by yourself!'
        flash(msg,'warning')

    return redirect(url_for("ctb_result", _external=True, _scheme='http', viewarg1=1))

@app.route('/recover/<login_user>/<filename>', methods=['GET'])
def recover_file_trash(login_user, filename):
    if login_user == 'unknown':
        http_scheme = 'http'
    else:
        http_scheme = 'https'

    if 'CTB' in filename:
        dest_path=base_dir_output
    else:
        dest_path=base_dir_upload

    if login_user in filename:
        os.rename(os.path.join(base_dir_trash, filename), os.path.join(dest_path, filename))
        msg = 'File put back to original place: {}'.format(filename)
        flash(msg, 'success')
    else:
        msg = 'You can only operate file created by yourself!'
        flash(msg, 'warning')

    return redirect(url_for("ctb_result", _external=True, _scheme='http', viewarg1=1))

@app.route('/t/<filename>',methods=['GET'])
def download_file_trash(filename):
    f_path=base_dir_trash
    login_user = request.headers.get('Oidc-Claim-Sub')

    if login_user != 'unknown':
        add_log_summary(user=login_user, location='Result', user_action='Download file', summary=filename)

    return send_from_directory(f_path, filename, as_attachment=True)

@app.route('/o/<filename>',methods=['GET'])
def download_file_output(filename):
    f_path=base_dir_output
    print(f_path)
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != 'unknown':
        add_log_summary(user=login_user, location='Result', user_action='Download file', summary=filename)
    return send_from_directory(f_path, filename, as_attachment=True)

@app.route('/u/<filename>',methods=['GET'])
def download_file_upload(filename):
    f_path=base_dir_upload
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != 'unknown':
        add_log_summary(user=login_user, location='Result', user_action='Download file',summary=filename)
    return send_from_directory(f_path, filename, as_attachment=True)

@app.route('/l/<filename>',methods=['GET'])
def download_file_logs(filename):
    f_path=base_dir_logs
    print(f_path)
    login_user = request.headers.get('Oidc-Claim-Sub')
    if login_user != None:
        add_log_summary(user=login_user, location='Result', user_action='Download file',summary=filename)
    return send_from_directory(f_path, filename, as_attachment=True)



@app.route('/admin', methods=['GET','POST'])
def ctb_admin():
    form = AdminForm()
    login_user=request.headers.get('Oidc-Claim-Sub')
    login_name = request.headers.get('Oidc-Claim-Fullname')
    if login_user == None:
        login_user = 'unknown'
        login_name = 'unknown'

    if login_user!='unknown' and login_user!='kwang2':
        return redirect(url_for('ctb_run',_external=True,_scheme='http',viewarg1=1))
        add_log_summary(user=login_user, location='Admin', user_action='Visit', summary='Why happens?')

    # get file info
    output_record_hours=360
    upload_record_hours=240
    df_output=get_file_info_on_drive(base_dir_output,keep_hours=output_record_hours)
    df_upload=get_file_info_on_drive(base_dir_upload,keep_hours=upload_record_hours)
    df_logs=get_file_info_on_drive(base_dir_logs,keep_hours=10000)

    # read logs
    df_log_detail = read_table('ctb_user_log')
    df_log_detail.sort_values(by=['id'],ascending=False,inplace=True)

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
                           user=login_user)


@app.route('/resume', methods=['GET'])
def resume():
    return render_template('resume.html')
