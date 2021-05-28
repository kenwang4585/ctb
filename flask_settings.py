from flask import Flask
#from flask_wtf.file import FileField
from flask_wtf import FlaskForm
from wtforms.validators import DataRequired
from wtforms import SubmitField, StringField, MultipleFileField, FileField
import os
from flask_sqlalchemy import SQLAlchemy

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'secret string')
#app.config['UPLOAD_PATH'] = os.path.join(app.root_path, 'upload_file')

#app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + base_dir_db + os.getenv('DB_URI')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DB_URI')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
#app.config['MAX_CONTENT_LENGTH']=150*1024*1024


db = SQLAlchemy(app)

# Flask forms
class UploadForm(FlaskForm):
    # 创建各种表单对象
    org=StringField('Organization code:',validators=[DataRequired()],render_kw={'placeholder':'e.g. FOC; currently only support one org at a time'})
    bu=StringField('Business units: ',render_kw={'placeholder':'e.g. PABU/ERBU; recommend leave blank to run at org level due to common parts across BU.'})
    class_code_exclusion=StringField('Class code to exclude for supply:',
                                     default='47/471/501/502/503/504/55/83/84/90')
    #customer=StringField('(optional)Input customer name to flag(e.g. Google/NTT):')
    description=StringField('Description:',render_kw={'placeholder':'Short description show in output file name'})

    file_3a4 = FileField('Upload 3A4 file (.csv):',validators=[DataRequired()])
    file_kinaxis_supply=FileField('Upload Kinaxis supply file (.xlsx):')
    file_allocation_supply=FileField('Upload PCBA allocation file (.xlsx):')

    file_supply_multiple=MultipleFileField('Supply files')
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
class CtbUserLog(db.Model):
    '''
    User logs db table
    '''
    id=db.Column(db.Integer,primary_key=True)
    USER_NAME=db.Column(db.String(10))
    DATE=db.Column(db.Date)
    TIME=db.Column(db.String(8))
    LOCATION=db.Column(db.String(25))
    USER_ACTION=db.Column(db.String(35))
    SUMMARY=db.Column(db.Text)

