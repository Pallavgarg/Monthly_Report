import os
from decimal import Decimal

from dateutil.relativedelta import relativedelta
from flask_mail import Mail
import xlsxwriter

from flask import Flask, render_template, request, flash, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date

from sqlalchemy import func
today = date.today()
app = Flask(__name__)  # creating the Flask class object
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql+pymysql://root:pallav123@localhost/mps'
db = SQLAlchemy(app)
app.secret_key="Don't tell"
app.config.update(
    MAIL_SERVER='smtp.gmail.com',
    MAIL_PORT='465',
    MAIL_USE_SSL= True,
    MAIL_USERNAME="pallavgarg04@gmail.com",
    MAIL_PASSWORD="samsung@412"
)
mail=Mail(app)
class Contacts(db.Model):
    __tablename__ = 'contacts'
    sno = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False)
    phone_num = db.Column(db.String(12), nullable=False)
    msg = db.Column(db.String(200), nullable=False)
    date = db.Column(db.String(12), nullable=False)
    email = db.Column(db.String(45), nullable=False)

class Accounts(db.Model):
    __tablename__ = 'accounts'
    sno = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), nullable=False)
    username = db.Column(db.String(80), nullable=False)
    password = db.Column(db.String(80), nullable=False)
    email = db.Column(db.String(45), nullable=False)
    phone_num = db.Column(db.String(12), nullable=False)
    date_of_join = db.Column(db.String(12), nullable=False)

class Report(db.Model):
    __tablename__ = 'report'
    sno = db.Column(db.Integer, primary_key=True)
    serial = db.Column(db.Integer, nullable=True)
    mis = db.Column(db.Integer, nullable=False)
    nsc = db.Column(db.Integer, nullable=False)
    kvp = db.Column(db.Integer, nullable=False)
    td = db.Column(db.Integer, nullable=False)
    amount = db.Column(db.Integer, nullable=False)
    comm= db.Column(db.Integer, nullable=False)
    tds = db.Column(db.Integer, nullable=False)
    netcomm = db.Column(db.Integer, nullable=False)
    username = db.Column(db.Integer, nullable=False)
    date = db.Column(db.Integer, nullable=False)



@app.route('/')  # decorator drfines the
def home():
    return render_template('home.html',title='Moneyrep')


@app.route('/about')  # decorator drfines the
def about():
    return render_template('about.html',title='Moneyrep')

@app.route('/services')  # decorator drfines the
def services():
    return render_template('services.html',title='Moneyrep')

@app.route('/contact',methods=['GET','POST'])  # decorator drfines the
def contact():
    if(request.method=='POST'):
        '''Add entry to the database'''
        name = request.form.get('name')
        email = request.form.get('email')
        phone = request.form.get('phone')
        message = request.form.get('message')
        entry = Contacts(name=name,phone_num=phone,date=datetime.now(),msg=message,email=email)
        if name!="" and email!="" and phone!="":
            db.session.add(entry)
            flash("Send :)")
            mail.send_message('New Message from '+name+"@Moneyrep",
                              sender=email,
                              recipients=["pallavgarg04@gmail.com"],
                              body=message+"\n"+phone
                              )
        else:
            flash("Please!! Fill all the details")
        db.session.commit()
    return render_template('contact.html',title='Moneyrep')
#########################################################################################################
@app.route('/signin',methods=['GET','POST'])  # decorator drfines the
def signin():
    if (request.method == 'POST'):
        '''Add entry to the database'''
        user = request.form.get('username')
        passw = request.form.get('password')
        if db.session.query(Accounts).filter(Accounts.username == user, Accounts.password == passw).all():
            #mypage(username=user)
            qry = [(Decimal('0'), Decimal('0'), Decimal('0'), Decimal('0'), Decimal('0'),
                    Decimal('0'), Decimal('0'), Decimal('0'))]
            return render_template('mypage.html', username=user,qry=qry)
        else:
            flash("Wrong Credentials !!!\n Please Check again")

        db.session.commit()


    return render_template('signin.html',title='Moneyrep')
#######################################################################################################
@app.route('/signup',methods=['GET','POST'])  # decorator drfines the
def signup():
    if (request.method == 'POST'):
        '''Add entry to the database'''
        name = request.form.get('name')
        username = request.form.get('username')
        password = request.form.get('password')
        email = request.form.get('email')
        phone = request.form.get('phone')
        entry = Accounts(name=name, phone_num=phone, date_of_join=datetime.now(),username=username,password=password,email=email)
        if name!="" and email!="" and phone!="" and password!="" and username!="":
            db.session.add(entry)
            flash("Added :)")
        else:
            flash("Please!! Fill all the details")
        db.session.commit()
    return render_template('signup.html',title='Moneyrep')
###################################################################################################



@app.route('/mypage',methods=['GET','POST'])  # decorator drfines the
def mypage():
    if (request.method == 'POST'):
        serial = request.form.get('serial')
        mis = request.form.get('mis')
        nsc = request.form.get('nsc')
        kvp = request.form.get('kvp')
        td = request.form.get('td')
        user = request.form.get('username')
        date = request.form.get('date')

        if serial!="" and mis!="" and nsc!="" and kvp!="" and date!="" and td!="":
            amount = int(mis) + int(nsc) + int(kvp) + int(td)
            comm = round(amount * 0.005)
            tds = round(comm * 0.05)
            netcomm = comm - tds
            entry = Report(serial=serial, mis=mis, nsc=nsc, kvp=kvp, td=td, amount=amount, comm=comm, tds=tds,
                           netcomm=netcomm, username=user, date=date)
            db.session.add(entry)
            flash("Added :)")
        elif serial=="" or mis=="" or nsc=="" or kvp=="" or date=="" or td=="":
            flash("Please!! Fill all the details")
        else:
            flash("Please!! Fill all the details")
        db.session.commit()


        first_day_of_month = today - relativedelta(months=1)
        last_day_of_month = today

        '''
        first = request.form.get('first_date')
        last = request.form.get('last_date')
        if first=="":
            first_day_of_month=datetime.today()
        else:
            first_day_of_month = first

        if last=="":
            last_day_of_month=datetime.today()
        else:
            last_day_of_month = last
        '''
        Reports= db.session.query(Report).\
            filter(Report.username == user,Report.date >= first_day_of_month,Report.date <= last_day_of_month).\
            order_by(Report.serial).all()

        qry = db.session.query(func.sum(Report.mis).label("mis"),func.sum(Report.nsc).label("nsc"),\
                func.sum(Report.kvp).label("kvp"),func.sum(Report.td).label("td"),func.sum(Report.amount).label("amount"), \
                func.sum(Report.comm).label("comm"), func.sum(Report.tds).label("tds"),func.sum(Report.netcomm).label("netcomm")).\
              group_by(Report.username). \
              filter(Report.username == user, Report.date >= first_day_of_month,Report.date <= last_day_of_month).all()
        if qry==[]:
            qry = [(Decimal('0'), Decimal('0'), Decimal('0'), Decimal('0'), Decimal('0'),
                    Decimal('0'), Decimal('0'), Decimal('0'))]
            


        return render_template("mypage.html",username=user,Report=Reports,qry=qry )
###################################################################################################



@app.route('/uploader', methods = ['GET', 'POST'])


def upload():
    if request.method == 'POST':
        today = datetime.today()
        user = request.form.get('username')
        filename= user+"_"+str(today.strftime("%d_%m_%Y_%H_%M"))+'.xlsx'
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        first = request.form.get('first_date')
        last = request.form.get('last_date')
        if first == "":
            first_day_of_month = date.today()
        else:
            first_day_of_month = first

        if last == "":
            first_day_of_month = date.today()
        else:
            last_day_of_month = last



        merge_format = workbook.add_format({
            'bold': 2,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'fg_color': 'white'})


        worksheet.merge_range('A1:J5', 'Monthly Report', merge_format)
        worksheet.write("A7", "Serial", bold)
        worksheet.write("B7", "Date", bold)
        worksheet.write("C7", "MIS", bold)
        worksheet.write("D7", "NSC", bold)
        worksheet.write("E7", "KVP", bold)
        worksheet.write("F7", "TD", bold)
        worksheet.write("G7", "Amount", bold)
        worksheet.write("H7", "Comm.", bold)
        worksheet.write("I7", "TDS", bold)
        worksheet.write("J7", "Net Comm", bold)


        format2 = workbook.add_format({'num_format': 'dd/mm/yy'})

        row = 7
        col = 0

        Reports = db.session.query(Report). \
            filter(Report.username == user, Report.date >= first_day_of_month, Report.date <= last_day_of_month). \
            order_by(Report.serial).all()
        print(Reports)

        # Iterate over the data and write it out row by row.
        for report in Reports:
            worksheet.write(row, col + 0, report.serial)
            worksheet.write(row, col + 1, report.date,format2)
            worksheet.write(row, col + 2, report.mis)
            worksheet.write(row, col + 3, report.nsc)
            worksheet.write(row, col + 4, report.kvp)
            worksheet.write(row, col + 5, report.td)
            worksheet.write(row, col + 6, report.amount)
            worksheet.write(row, col + 7, report.comm)
            worksheet.write(row, col + 8, report.tds)
            worksheet.write(row, col + 9, report.netcomm)

            row += 1

        worksheet.write(row, 1, 'Total', bold)

        qry = db.session.query(func.sum(Report.mis).label("mis"), func.sum(Report.nsc).label("nsc"), \
                               func.sum(Report.kvp).label("kvp"), func.sum(Report.td).label("td"),
                               func.sum(Report.amount).label("amount"), \
                               func.sum(Report.comm).label("comm"), func.sum(Report.tds).label("tds"),
                               func.sum(Report.netcomm).label("netcomm")). \
                                group_by(Report.username). \
                                filter(Report.username == user, Report.date >= first_day_of_month, Report.date <= last_day_of_month).all()

        worksheet.write(row, 2, qry[0][0], bold)
        worksheet.write(row, 3, qry[0][1], bold)
        worksheet.write(row, 4, qry[0][2], bold)
        worksheet.write(row, 5, qry[0][3], bold)
        worksheet.write(row, 6, qry[0][4], bold)
        worksheet.write(row, 7, qry[0][5], bold)
        worksheet.write(row, 8, qry[0][6], bold)
        worksheet.write(row, 9, qry[0][7], bold)

        workbook.close()
        return send_file(os.path.join(r"C:\Users\ADMIN\PycharmProjects\Monthly_Report",filename), attachment_filename=filename)


####################################################################################################

@app.route('/layout')  # decorator drfines the
def layout():
    return render_template('layout.html',title='Moneyrep')

if __name__ == '__main__':
    app.run(debug=True)