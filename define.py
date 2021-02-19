from flask import Flask,render_template,flash,request,redirect,url_for, send_from_directory
from flask_sqlalchemy import SQLAlchemy
import flask_mysqldb
import sqlalchemy
from sqlalchemy.sql import func
import pymysql
from flask_wtf import FlaskForm
from flask_wtf.file import FileRequired,FileAllowed
from wtforms import StringField,SubmitField,DateField,SelectField,RadioField,FileField
from wtforms.fields.html5 import DateField
from wtforms.validators import DataRequired
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl import *
import os
import datetime
import random
from collections import Counter
import pandas as pd

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI']= 'mysql://root:hqmo8888@127.0.0.1:3306/hqmo?charset=utf8mb4'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
app.secret_key = '12345678'
dirpath = os.path.dirname(__file__)


class Department(db.Model):
    #表名
    __tablename__ = "department"
    #字段
    # __table_args__ = {'mysql_autoincrement': True}
    id = db.Column(db.String(8))
    name = db.Column(db.String(36), primary_key = True, unique=True, nullable=False)
    classname = db.Column(db.String(16),unique=False, nullable=False)
    bednumber = db.Column(db.Integer,unique=False, nullable=False)
    #关系引用
    datas = db.relationship('Recdata', backref='data')
    anes = db.relationship('Anestdata', backref = 'ane')
    exams = db.relationship('ExamTestData', backref = 'exam')
    incomes = db.relationship('Income', backref='income')
    inps = db.relationship('InPatient', backref = 'inp')
    outps = db.relationship('OutPatient', backref = 'outp')
    phams = db.relationship('PharmacyData', backref = 'pham')
    quas = db.relationship('Quality', backref = 'qua')
    surgs = db.relationship('Surgery', backref='surg')
    staffs = db.relationship('Staff', backref='staff')
    def __repr__(self):
        return '科室名称: %s' % self.name


class Staff(db.Model):
    #表名,员工数据表
    #字段
    # __table_args__ = {'mysql_autoincrement': True}
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(12),nullable=False)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    identity = db.Column(db.String(16))
    def __repr__(self):
        return self.deptname


class Anestdata(db.Model):
    #表名,麻醉数据表
    __tablename__ = "anestdata"
    #字段
    # __table_args__ = {'mysql_autoincrement': True}
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    ratepreaneseval = db.Column(db.Float)
    ratesurgsafeveri = db.Column(db.Float)
    rateanesresu = db.Column(db.Float)
    analysisanescomp = db.Column(db.Integer)
    ratepostsurgvisit = db.Column(db.Float)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class ExamTestData(db.Model):
    # #表名,医技数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    MedEquipTime = db.Column(db.Integer)
    MedEquipMalFunc = db.Column(db.Integer)
    RateIQC = db.Column(db.Float)
    NumPassEQA = db.Column(db.Integer)
    AccuRateInspRept = db.Column(db.Float)
    TimelyRateInspRept = db.Column(db.Float)
    NumForumwithClin = db.Column(db.Integer)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class Income(db.Model):
    # #表名,收入结构数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    totalincome = db.Column(db.Float)
    drugincome = db.Column(db.Float)
    adjDrugincome = db.Column(db.Float)
    consumIncome = db.Column(db.Float)
    examIncome = db.Column(db.Float)
    testIncome = db.Column(db.Float)
    oPIncome = db.Column(db.Float)
    iPIncome = db.Column(db.Float)
    pureIncome = db.Column(db.Float)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class InPatient(db.Model):
    # #表名,住院数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    NumIP = db.Column(db.Integer)
    NumBD = db.Column(db.Integer)
    NumRefIP = db.Column(db.Integer)
    CheckNumIPRec = db.Column(db.Integer)
    ANumIPRec = db.Column(db.Integer)
    BNumIPRec = db.Column(db.Integer)
    CNumIPRec = db.Column(db.Integer)
    OANumIPRec = db.Column(db.Integer)
    DRGGrp = db.Column(db.Integer)
    CMI = db.Column(db.Float)
    NumPath = db.Column(db.Integer)
    IPNumInDist = db.Column(db.Float)
    def __repr__(self):
        return '科室名称: %s%d%d%d' % (self.deptname,self.CheckNumIPRec,self.ANumIPRec,self.BNumIPRec)


class OutPatient(db.Model):
    # #表名,门诊数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    NumOP = db.Column(db.Integer)
    NumRefOP = db.Column(db.Integer)
    NumOPRec = db.Column(db.Integer)
    CheckNumOPRec = db.Column(db.Integer)
    ANumOPRec = db.Column(db.Integer)
    BNumOPRec = db.Column(db.Integer)
    CNumOPRec = db.Column(db.Integer)
    NumAppt = db.Column(db.Integer)
    WaitTimeAppt = db.Column(db.Float)
    OPNumInDist = db.Column(db.Float)
    def __repr__(self):
        return '科室名称: %s' % self.deptname

class PharmacyData(db.Model):
    # #表名,药剂数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    NumRp = db.Column(db.Integer)
    NumRpRev = db.Column(db.Integer)
    NumQuaRp = db.Column(db.Integer)
    NumEssDrugOP = db.Column(db.Integer)
    NumEssDrugIP = db.Column(db.Integer)
    DDD = db.Column(db.Integer)
    NumDrugConsultation = db.Column(db.Integer)
    NumPharmCliRd = db.Column(db.Integer)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class Quality(db.Model):
    # #表名,质量数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    NumDie = db.Column(db.Integer)
    QCCScore = db.Column(db.Integer)
    CriValRepNum = db.Column(db.Integer)
    CriValHPNum = db.Column(db.Integer)
    NumAdvEvtRep = db.Column(db.Integer)
    NumCTMri = db.Column(db.Integer)
    NumCTMriPositive = db.Column(db.Integer)
    ScoreEDM = db.Column(db.Integer)
    ScoreHandHygi = db.Column(db.Integer)
    NumVioCoreSys = db.Column(db.Float)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class Surgery(db.Model):
    # #表名,质量数据表
    # __tablename__ = "anestdata"
    #字段
    id = db.Column(db.Integer, primary_key=True)
    deptname = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.Date)
    NumSurg = db.Column(db.Integer)
    NumEleSurg = db.Column(db.Integer)
    NumDaySurg = db.Column(db.Integer)
    NumMiroInvaSurg = db.Column(db.Integer)
    Num4thSurg = db.Column(db.Integer)
    NumSurgComp = db.Column(db.Integer)
    NumType1Surg = db.Column(db.Integer)
    NumType1Infect = db.Column(db.Integer)
    def __repr__(self):
        return '科室名称: %s' % self.deptname


class Recdata(db.Model):
    # 字段
    id = db.Column(db.Integer, primary_key=True,autoincrement=True)
    dept_name = db.Column(db.String(36), db.ForeignKey('department.name'))
    date = db.Column(db.DATE)
    opnum = db.Column(db.Integer)
    oprecnum = db.Column(db.Integer)
    oprecchecknum =db.Column(db.Integer)
    oprecanum = db.Column(db.Integer)
    oprecbnum = db.Column(db.Integer)
    opreccnum = db.Column(db.Integer)
    iprecchecknum =db.Column(db.Integer)
    iprecanum = db.Column(db.Integer)
    iprecbnum = db.Column(db.Integer)
    ipreccnum = db.Column(db.Integer)
    def __repr__(self):
        return '年月: %s %s %s' % (self.date, self.dept_id, self.opnum)


class Deptform(FlaskForm):
    deptname = StringField('科室名称:',validators=[DataRequired])
    deptclass = StringField('科室类别:', validators=[DataRequired])
    deptbed = StringField('床位数:', validators=[DataRequired])
    submit = SubmitField('增加科室')
    submit2 = SubmitField('删除科室')
    submit3 = SubmitField('修改科室')


class Queryform(FlaskForm):
    startdate = DateField('输入开始年月', format='%Y-%m-%d', validators=[DataRequired])
    overdate = DateField('输入结束年月', format='%Y-%m-%d', validators=[DataRequired])
    deptname = SelectField('选择科室类别', default='手术', choices=[('手术', '手术科室'), ('住院', '住院科室'),
                                              ('门诊', '门诊科室'),('医技','医技科室'),('麻醉','麻醉科'),
                                              ('药剂','药剂科'),('重症','重症医学科')])
    submit = SubmitField('查询数据')
    toexcel = SubmitField('下载表格')


class Dataprocess(FlaskForm):
    date_sel = DateField('DatePicker', format='%Y-%m-%d')
    func_sel = RadioField(label="选择功能:",
                          choices=((1,'病历质控统计'),(2,'病历质控结果数据导入'),
                                   (3,'危急值数据导入'),(4,'不良事件数据导入'),
                                   (5,'HIS数据导入'),(6,'QCC数据导入'),
                                   (7,'院感药剂数据导入'),(8,'病案系统数据导入'),
                                   (9,'DRG数据导入'),(10,'处理医技数据'),
                                   (11,'所有空数据处理')))
    file_upload = FileField(label='文件1上传：', validators=[FileRequired(), FileAllowed(['xlsx'])])
    file2_upload = FileField(label='文件2上传：', validators=[FileRequired(), FileAllowed(['xlsx'])])
    submit = SubmitField('导入数据')

class Assign(FlaskForm):
    date_sel = DateField('DatePicker', format='%Y-%m-%d')
    file_upload = FileField(label='出院病历列表文件上传：', validators=[FileRequired(), FileAllowed(['xlsx'])])
    submit = SubmitField('分配病历')

# class Upload(FlaskForm):
#     file_upload = FileField(label='文件上传', validators=[FileRequired(), FileAllowed(['xlsx'])])
#     file2_upload = FileField(label='文件2上传', validators=[FileRequired(), FileAllowed(['xlsx'])])
#     submit = SubmitField('上传')