# ken, 2019

from flask import Flask, render_template
from flask_mail import Mail, Message
import os
import pandas as pd
from send_sms import send_me_sms


def send_attachment_and_embded_image(to_address,subject,html_template,att_filenames=None, embeded_filenames=None,sender='APJC DF',**kwargs):
    '''
    Use Flask_mail to send the result to defined emails
    :param to_address:
    :param subject:
    :param html_template: under templates folder in same directory
    :param att_filenames: List of tuples (path, file_name)
    :param embeded_filenames: dictionary {chart_name: chart_file_path}
    :param kwargs:
    :return:
    '''

    app = Flask(__name__)
    app.config.update(
        #MAIL_SERVER='smtp.office365.com',
        #MAIL_PORT=587,
        MAIL_SERVER='outbound.cisco.com',
        MAIL_PORT=25,
        MAIL_USE_TLS=True,
        #MAIL_USERNAME='kwang2@cisco.com',
        #MAIL_PASSWORD=os.getenv('MAIL_PASSWORD'),
        MAIL_DEFAULT_SENDER=(sender, 'noreply@cisco.com')
    )

    mail = Mail(app)

    # app.config['SERVER_NAME'] = 'example.com' #必须设置此项如果用到url_for
    with app.app_context():
        time_stamp1 = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')
        time_stamp2 = pd.Timestamp.now().strftime('%Y-%m-%d')

        message = Message(subject= subject + ' - ' + time_stamp1,
                          recipients=to_address,
                          body='',
                          # body=render_template('email.txt',data=data),
                          html=render_template(html_template, **kwargs)
                          )
        #添加excel附件
        att_size = 0
        size_over_limit=False
        if att_filenames!=None:
            for item in att_filenames:
                full_fname=os.path.join(item[0],item[1])
                short_fname=item[1]

                file_size=os.path.getsize(full_fname)
                att_size+=file_size

                if att_size<=20000000:
                    with app.open_resource(full_fname) as at:
                        dated_fname=short_fname[:-5]+' ('+time_stamp2 + ').xlsx'
                        message.attach(dated_fname, 'excel/xlsx', at.read())
                else:
                    size_over_limit=True
                    message.body='Attachment not added due to over size limit (20Mb).'

        #添加embeded images
        if embeded_filenames!=None:
            for chart_id, fname in embeded_filenames.items():
                with app.open_resource(fname) as fp:
                    message.attach('image.png', "image/png", fp.read(), 'inline', headers=[('Content-ID', chart_id)])

        mail.send(message)

        if att_size<=20000000:
            msg = '{} sent to defined email address!\n'.format(subject)
        else:
            msg = 'Alert***{} sent to {} w/o attachment due to over 20Mb limit (file size: {}Mb)!\n'.format(subject, to_address, str(round(att_size/1024000,1)))

        print(msg)

    return msg,size_over_limit

if __name__=='__main__':
    # Send program report to Ken wang
    to_address=['kwang2@cisco.com']
    subject = 'Config error identification result'
    html_template = 'config_prediction_result.html'

    send_attachment_and_embded_image(to_address, subject, html_template,
                                     summary=summary.values,
                                     new_error=new_error.values,
                                     unval_error=unval_error.values,
                                     old_error=old_error.values,
                                     log_messages=log_msg,
                                     alert_messages=alert_msg,

                                     att_filenames=None, embeded_filenames=None)
