from flask import Flask
from flask_wtf.file import FileField, FileRequired
from flask_wtf import FlaskForm
from wtforms.validators import Email, DataRequired,input_required
from wtforms import SubmitField, BooleanField, StringField,IntegerField,SelectField,PasswordField,TextAreaField,RadioField
import os
from flask_sqlalchemy import SQLAlchemy
from setting import db_uri

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'secret string')
#app.config['UPLOAD_PATH'] = os.path.join(app.root_path, 'upload_file')

app.config['SQLALCHEMY_DATABASE_URI'] = db_uri
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Flask forms
class UploadForm(FlaskForm):
    # 创建各种表单对象
    org=StringField('Organization code:',validators=[DataRequired()],render_kw={'placeholder':'e.g. FOC; currently only support one org at a time'})
    bu=StringField('Business units: ',render_kw={'placeholder':'e.g. PABU/ERBU; recommend leave blank to run at org level due to common parts across BU.'})
    class_code_exclusion=StringField('Class code to exclude for supply:',
                                     default='47/471/501/502/503/504/55/83/84/90')
    customer=StringField('(optional)Input customer name to flag(e.g. Google/NTT):')
    description=StringField('Description:',render_kw={'placeholder':'Short description show in output file name'})

    file_3a4 = FileField('Upload 3A4 file (.csv):',validators=[DataRequired()])
    file_supply=FileField('Upload supply file (.xlsx):',validators=[DataRequired()])
    submit_ctb=SubmitField(' RUN CTB ')

class FileDownloadForm(FlaskForm):
    # for deleting filename created by user self
    file_name_delete = StringField('File to delete', validators=[DataRequired()],
                                   default='put in filename here')
    submit_delete = SubmitField('   Delete   ')



class AdminForm(FlaskForm):
    file_name_delete=StringField('File to delete', validators=[DataRequired()])
    submit_delete=SubmitField('   Delete   ')


# Database tables
class UserLog(db.Model):
    '''
    User logs db table
    '''
    id=db.Column(db.Integer,primary_key=True)
    USER_NAME=db.Column(db.String(10))
    DATE=db.Column(db.Date)
    TIME=db.Column(db.String(8))
    LOCATION=db.Column(db.String(10))
    USER_ACTION=db.Column(db.String(20))
    SUMMARY=db.Column(db.Text)

