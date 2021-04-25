from flask import Flask
from flask_wtf.file import FileField, FileRequired
from flask_wtf import FlaskForm
from wtforms.validators import Email, DataRequired,input_required
from wtforms import SubmitField, BooleanField, StringField,IntegerField,SelectField,PasswordField,TextAreaField,RadioField
import os
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'secret string')
#app.config['UPLOAD_PATH'] = os.path.join(app.root_path, 'upload_file')

app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DB_URI') #os.getenv('DB_URI')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# Flask forms
class UploadForm(FlaskForm):
    # 创建各种表单对象
    org=StringField('Organization code (e.g. FOC):',validators=[DataRequired()],default='FOC')
    bu=StringField('Business units (e.g. PABU/ERBU; leave blank for all BU): ',default='')
    class_code_exclusion=StringField('Class code to exclude for supply:',
                                     default='47/471/501/502/503/504/55/83/84/90')
    customer=StringField('(optional)Input customer name to flag(e.g. Google/NTT):')

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

