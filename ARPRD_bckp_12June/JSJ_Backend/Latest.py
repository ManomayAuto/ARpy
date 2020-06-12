import gc
import random
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from flask_cors import CORS
import pyodbc as pyodbc
from flask import Flask, render_template, url_for, request, redirect, Response, session, flash, send_from_directory, \
    send_file, abort, jsonify, logging, make_response
from flask_cors import CORS, cross_origin
import json
import firebase_admin
from firebase_admin import db, credentials, auth, firestore
import jwt
from datetime import datetime, timedelta
from functools import wraps
from motor import getpremiummotor
from private import getpremiumprivate
from encryption import DataEncryption

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=192.193.194.40;'
                      'Database=MantraDB;'
                      'Trusted_Connection=no;')

# cred = credentials.Certificate("secret/portals-2edf2-firebase-adminsdk-aljih-9a4318b0c8.json")
# firebase_admin.initialize_app(cred)
# db1 = firestore.client()

app = Flask(__name__)
CORS(app)
permissions = ["QI", "QRU", "QRM", "QRP", "DF", "DR", "ARRP"]
de = DataEncryption()

def login_required(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'logged_in' in session:
            return f(*args, **kwargs)
        else:
            flash("You need to login first")
            return redirect(url_for('login'))

    return wrap

@app.route('/calculation', methods=['POST', 'OPTIONS'])
@cross_origin(
    origin=['*'], allow_headers=['Content-Type', 'Authorization', 'User'])
def cal():
    data = request.json
    print(data)
    product = data['prod']
    if product is not None and product != '':
        if product == 'Private':
            annualgpwp, annualgpwpd, annualnet, taxamt, nc = getpremiumprivate(data['cc'], data['vehiclevalue'], data['softd'], data['vehicletype'], data['manloadp'], data['claimfre'], data['manualdisc'], data['fleet'], data['promotion'], data['ct'])
            return jsonify({
                'annualgpwp': annualgpwp,
                'annualgpwpd': annualgpwpd,
                'annualnet': annualnet,
                'taxamt': taxamt,
                'nc': nc,
            })
        elif product == 'Motor':
            annualgpwp, annualgpwpd, annualnet, taxamt, nc = getpremiummotor(data['cc'], data['vehiclevalue'], data['softd'], data['vehicletype'], data['manloadp'], data['claimfre'], data['manualdisc'], data['fleet'], data['promotion'], data['ct'])
            return jsonify({
                'annualgpwp': annualgpwp,
                'annualgpwpd': annualgpwpd,
                'annualnet': annualnet,
                'taxamt': taxamt,
                'nc': nc,
            })
        else:
            return jsonify({"result": "please send a product"}),400
    else:
        return jsonify({"result": "please send a product"}),422
    return jsonify({"result": "success"}),200

@app.route('/checktoken', methods=['POST', 'OPTIONS'])
# @cross_origin(
#     origin=['*'], allow_headers=['Content-Type', 'Authorization', 'User'])
def checktoken():
    if request.method == "POST":
        result = False
        print(request.headers)
        try:
            token = ""
            request_keys = request.headers.keys()
            if "Authorization" in request_keys and "Ipaddress" in request_keys and "User" in request_keys:
                token = request.headers["Authorization"]
                print(token)
            print(request.get_data())
            data = json.loads(str(request.data, encoding='utf-8'))
            decoded = jwt.decode(token, key='secret')
            print('-------------------------')
            print(decoded)
            print(decoded["user"])
            if decoded["user"] == request.headers["User"] and decoded["ip"] == request.headers["Ipaddress"]:
                result = True
        except jwt.DecodeError:
            print("decode error")
            result = False
        except jwt.ExpiredSignatureError:
            print("sign")
            result = False
        except KeyError:
            print("key error")
            print(result)
        if result:
            return jsonify({"result": result}), 200
        else:
            return jsonify({"result": result}), 401

# this is for admin portal
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == "GET":
        return render_template("login.html")
    elif request.method == "POST":
        email = request.form["email"]
        password = request.form["password"]
        encryptpas = de.encrypt(password)
        print(email)
        print(password)
        cursor = conn.cursor()

        cursor.execute( "select Permissions,temppass,Name,Email from Users_man where Email = '%s' and Password = '%s'" % (email,encryptpas))

        temp_pass = 0
        permissions = None
        name = None

        for row in cursor:
            permissions = row[0]
            temp_pass = row[1]
            name = row[2]
            email = row[3]
        if temp_pass == 1:
            return redirect(url_for("resetpass"))
        if permissions is not None and name is not None and "Admin" in permissions:
            session["logged_in"] = True
            session["email"] = email
            return redirect(url_for('usermanagement'))
        else:
            flash("You have no authorization")
            return redirect(url_for('login'))

#t this is for angular
@app.route('/login', methods=['POST'])
@cross_origin(origins="*", headers="*")
def login_form():
    print(request.data)
    data = json.loads(request.data, encoding='utf-8')
    # email = request.form["email"]
    # password = request.form["password"]
    email = data["email"]
    password = data["password"]
    print(email)
    print(password)
    encryptpas = de.encrypt(password)
    print(de.encrypt(password))
    cursor = conn.cursor()
    cursor.execute(
        "select Permissions,temppass,Name,Email,Branch from Users_man where Email = '%s' and Password = '%s'" % (
            email,encryptpas))

    temp_pass = True
    permissions = None
    name = None

    for row in cursor:
        permissions = row[0]
        temp_pass = row[1]
        name = row[2]
        email = row[3]
        branch = row[4]

    if permissions is not None and name is not None:
        current_time = datetime.now().timestamp()
        #current_time = datetime.now()
        print("current time -----------------------------------------------")
        print(current_time)
        #exp = (current_time + timedelta(hours=3)).timestamp()
        exp = datetime.utcnow() + timedelta(hours=3, minutes=0, seconds=0)
        print("Expired time ---------------------------------------------")
        print(exp)
        token = jwt.encode({"exp": exp, "iat": current_time, "permissions": permissions,"Name": name, "Email": email,"Branch": branch}, algorithm='HS256', key='secret')
        print(token)
        return jsonify({"permissions": permissions, "Email": email, "Name": name,"Branch": branch, "Token": token.decode("utf-8")}), 200
    else:
        print("Not authorized")
        return jsonify({"error": "Not Authorized"}), 401



@app.route('/Usermanage', methods=['GET', 'POST'])
@login_required
def usermanagement():
    cursor = conn.cursor()
    cursor.execute(
        "select Name,Permissions,Email,temppass,lastupdated,branch from Users_man where Email <> '%s'" % (
            session.get('email')))
    names = []
    permissions = []
    emails = []
    branches = []
    c = conn.cursor()
    for row in cursor:
        names.append(row[0])
        permissions.append(row[1])
        emails.append(row[2])
        branches.append(row[5])

    count = c.execute("select count(*) from Users_man where Email <> '%s'" % (
        session.get('email')))
    print(count)
    count = [int(i[0]) for i in c.fetchall()]
    print(count[0])
    cp = count[0]
    return render_template("Usermanage.html", u=names, u1=permissions, u2=emails, u3=branches, cp=cp, level=session.get('level'))


@app.route("/createuser", methods=['POST'])
@login_required
def add_user():
    if request.method == "POST":
        email = request.form.get("email")
        name = request.form.get("Name")
        password = request.form.get("Password")
        branch = request.form.get("branch")
        permissions_to_db = ""
        for permission in permissions:
            if request.form.get(permission):
                permissions_to_db += permission + ","
        permissions_to_db = permissions_to_db[:-1]
        print(email)
        print(name)
        print(password)
        print(permissions_to_db)
        encryptpas = de.encrypt(password)
        print("The encrypted password is  -----------------------------------------")
        print(encryptpas)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO Users_man (Name,Permissions,Email,Password,Branch,Createdby,temppass) VALUES('%s','%s',"
                       "'%s','%s','%s','%s',%s)" % (name, permissions_to_db, email, encryptpas,branch, session.get("email"), 1))
        conn.commit()
        return redirect(url_for('usermanagement'))


@app.route('/updateuser', methods=['GET', 'POST'])
@login_required
def updateuser():
    email = request.form.get("Email")
    name = request.form.get("Name")
    permissions_to_db = ""
    for permission in permissions:
        if request.form.get(permission):
            permissions_to_db += permission + ","
    permissions_to_db = permissions_to_db[:-1]
    print(email)
    print(name)
    print(permissions_to_db)
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE Users_man SET Name='%s',Permissions='%s' WHERE Email='%s'" % (
            name, permissions_to_db, email))
    conn.commit()
    flash("User details updated Successfully")
    return redirect(url_for('usermanagement'))


@app.route('/deluser', methods=['GET', 'POST'])
@login_required
def deleteuser():
    try:
        email = request.form["Email"]
        print(email)
        cursor = conn.cursor()
        cursor.execute("delete from Users_man WHERE Email='%s'" % email)
        conn.commit()
        flash("User deleted Successfully")

    except Exception as e:

        write_to_file(e)
        flash("Can't delete the user")
        conn.rollback()
    finally:
        return redirect(url_for('usermanagement'))


# """Function to change the password"""
@app.route('/pasuser', methods=['GET', 'POST'])
@login_required
def passuser():
    try:
        passwa = request.form["Password1"]
        encryptpas = de.encrypt(password)
        print(encryptpas)
        email = request.form["Email"]
        cursor = conn.cursor()
        cursor.execute("UPDATE Users_man SET Password='%s' WHERE Email='%s'" % (encryptpas, email))
        conn.commit()
        flash("Password changed Successfully")
    except Exception as e:
        flash("Cant change the password check the database connectivity")
        write_to_file(e)
        conn.rollback()
    finally:
        return redirect(url_for('usermanagement'))


@app.route('/emailval', methods=['GET', 'POST'])
@login_required
def emailval():
    email = request.form.get("email")
    print(email)
    response = "OK"
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM Users_man where Email = '%s'" % email)
    if len(list(cursor)) > 0:
        response = "The username is already taken"
        print("username already exists")
    return response


@app.route('/checkotp', methods=['GET', 'POST'])
def checkotp():
    if request.method == "POST":
        otp = session.get("jksadfjhk")
        print("check")
        zero = request.form["0"]
        one = request.form["1"]
        two = request.form["2"]
        three = request.form["3"]

        if str(otp) == zero + one + two + three and session.get('reset_email') == request.form["emailid"]:
            print(otp)
            session['reset'] = True
            return render_template('newpass.html')
        else:
            return """<script>alert("The OTP entered is not valid");window.location.href="/"</script>"""
    return '', 202


# last stage of resetting password

@app.route('/password_reset', methods=['GET', 'POST'])
def password_reset():
    if session.get('reset') and request.method == "POST":
        email = session.get('reset_email')
        print(email)
        db1 = conn.cursor()
        password = request.form["password"]
        encryptpas = de.encrypt(password)
        print(encryptpas)
        time = (datetime.now() - session.get('time')).total_seconds()
        print(time)
        print(session.get('time'))
        if time < 120:
            if len(password) >= 6:  # and not str(email).__contains__("admin"):
                db2 = conn.cursor()
                db2.execute("SELECT Name from Users_man where Email='%s'" % email)
                if db2 is None:
                    return "User doesn't exists. please provide a valid email"
                db2.close()

                db1.execute("UPDATE Users_man SET Password='%s',lastupdated='%s',temppass=0 WHERE Email='%s'" % (encryptpas, datetime.today(), email))
                db1.commit()
                return "OK"

            else:
                print("password strength not adequate")
                return "Not a Strong Password"

        else:
            return """<script type="text/javascript">alert("The OTP has been 
            expired");window.location.href="/resetpass"</script> """


"""Function to render the template"""


@app.route('/resetpass')
def resetpass():
    return render_template('resetpass.html')


"""Function to take the action resetting the password 
This function takes the email id and sends an OTP to it"""


@app.route('/reset_pass', methods=['GET', 'POST'])
def reset_pass():
    if request.method == "POST":
        email = request.form["emailid"]
        session["reset_email"] = email
        print(email)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM Users_man where Email = '%s'" % email)
        if len(list(cursor)) > 0:
            otp = random.randrange(1111, 9999, 5)
            session['jksadfjhk'] = otp
            print(otp)
            smtpObj = smtplib.SMTP_SSL('smtp.gmail.com', port=465)
            smtpObj.login('icbluat1@gmail.com', "manomay@icbl!")
            msg = MIMEMultipart()
            msg['subject'] = "OTP for reset password"
            msg['from'] = "icbluat1@gmail.com"
            msg['to'] = "icbluat1@gmail.com"
            msgtext1 = MIMEText(
                'your otp is %s' % str(otp), 'html')
            msg.attach(msgtext1)
            smtpObj.sendmail('icbluat1@gmail.com', email, msg.as_string())
            session["time"] = datetime.now()
            return render_template('optgenerate.html', email=email)
        else:
            return """<script type="text/javascript">alert("please enter a valid 
            email");window.location.href="/resetpass"</script> """
    else:
        return url_for('resetpass')


@app.route("/logout")
def logout():
    session.clear()
    gc.collect()
    return redirect(url_for('login'))


@app.route('/ae', methods=['GET', 'OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def ae():
    cur = conn.cursor()
    print("---------------------------------------------------------------")
    datas = request.headers.get('Authorization')
    print("-----------------",datetime.now())
    print(datas)
    datas = datas.replace('"','')
    print(datas)
    result = False
    try:
        decoded = jwt.decode(datas, str('secret'), 'utf-8')
        perm = decoded.get('permissions')
        if perm == "DF,DR":
            AEName = decoded.get('Name')
            print(AEName)
            cur.execute("select * from AE where AccExe = ? order by clientname,DTE",(AEName))
            row_headers=[x[0] for x in cur.description]
            records = cur.fetchall()
            json_data=[]
            for record in records:
                json_data.append(dict(zip(row_headers,record)))
        #print(records)
            result = True
            return jsonify(json_data)
    except jwt.ExpiredSignatureError:
        print("sign")
        result = False
    if result:
        return jsonify({"result": result}), 200
    else:
        return jsonify({"result": result}), 401

@app.route('/nonae', methods=['GET', 'OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def nonae():
    cur = conn.cursor()
    datas = request.headers.get('Authorization')
    datas = datas.replace('"', '')
    print("------------------------IN HO CALL  ATOKEN ISSSSSSSS-----------------------------")
    print(datas)
    result = False
    try:
        decoded = jwt.decode(datas, str('secret'), 'utf-8')
        perm = decoded.get('permissions')
        if perm == "DF":
            BranchName = decoded.get('Branch')
            print(BranchName)
            cur.execute("select * from NONAE where DTE between 35 and 45 and Branch = ? order by clientname", BranchName)
            row_headers = [x[0] for x in cur.description]
            records = cur.fetchall()
            json_data=[]
            for record in records:
                json_data.append(dict(zip(row_headers,record)))
            result = True
            return jsonify(json_data)
    except jwt.ExpiredSignatureError:
        print("sign")
        result = False
    if result:
        return jsonify({"result": result}), 200
    else:
        return jsonify({"result": result}), 401

@app.route('/nonaered', methods=['GET', 'OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def nonaered():
    cur = conn.cursor()
    datas = request.headers.get('Authorization')
    datas = datas.replace('"', '')
    print(datas)
    result = False
    try:
        decoded = jwt.decode(datas, str('secret'), 'utf-8')
        perm = decoded.get('permissions')
        if perm == "DR":
            BranchName = decoded.get('Branch')
            print(BranchName)
            cur.execute("select * from NONAE where DTE between -7 and 7 and Branch = ? order by clientname", BranchName)
            row_headers=[x[0] for x in cur.description]
            records = cur.fetchall()
            json_data=[]
            for record in records:
                json_data.append(dict(zip(row_headers,record)))
            result = True
            return jsonify(json_data)
    except jwt.ExpiredSignatureError:
        print("sign")
        result = False
    if result:
        return jsonify({"result": result}), 200
    else:
        return jsonify({"result": result}), 401


# AE followup

@app.route("/followup", methods=['GET','POST','OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def data():
    Data = request.get_json()
    print(Data)
    PN = Data.get('Id')
    print(PN)
    sts = Data.get('Status')
    cur = conn.cursor()
    if sts == 'Followed up':
        n = Data.get('Note')
        m = Data.get('Mode')
        d= Data.get('Date')
        print(d)
        dto = datetime.strptime(d,'%m-%d-%Y %H:%M')
        print("----------------------------date is")
        print(dto)
        fu= int(Data.get('Followups'))
        fu=fu+1
        ed = Data.get('ExpDate')
        print(ed)
        comExpDate = '01-01-2000'
        comExpDate = datetime.strptime(comExpDate,'%m-%d-%Y')
        print(comExpDate)
        print(type(comExpDate))
        print(type(ed))
        ed = datetime.strptime(ed,'%m-%d-%Y')
        UserName = Data.get('Uname')
        if ed.date() < comExpDate.date():
            print("In Expiry Date -------------------------------------")
            ed = "null"
            cur.execute("update AE set Status = ?, mode = ?,Correspondence= ?,Followups = ?,LastFollowUp = ?, username = ?, RPAStatus = NUll where PolicyNum like ?",(sts, m, n, fu, dto,UserName, PN))
            cur.commit()
        else:
            ed = ed
            cur.execute("update AE set Status = ?, mode = ?,Correspondence= ?,Followups = ?,LastFollowUp = ?,ExpectedDate= ?, username = ?, RPAStatus = null where PolicyNum like ?",(sts, m, n, fu, dto, ed, UserName, PN))
            cur.commit()
        print(ed)


    if sts == 'Completed':
        UserName = Data.get('Uname')
        tid = Data.get('TransactionId')
        d = Data.get('Date')
        dto = datetime.strptime(d,'%m-%d-%Y %H:%M')
        c = Data.get('Checked')
        cur.execute('update AE set Status = ?,ReductionTransId = ?,DateCompleted = ?,ReductionCheckBox = ?,Username = ? where PolicyNum like ?',(sts,tid,dto,c,UserName,PN))
        cur.commit()
    return jsonify(Data), 200

# NONAE FOllow up

@app.route("/Nonfollowup", methods=['GET','POST','OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def Nonaedata():
    Data = request.get_json()
    print(Data)
    PN = Data.get('Id')
    print(PN)
    d = Data.get('Date')
    print(d)
    dto = datetime.strptime(d, '%m-%d-%Y %H:%M')
    sts = Data.get('Status')
    print(sts)
    UserName = Data.get('Uname')
    ed = Data.get('ExpDate')
    print("------------------------------Expected date is")
    print(ed)
    comExpDate = '01-01-2000'
    comExpDate = datetime.strptime(comExpDate, '%m-%d-%Y')
    ed = datetime.strptime(ed, '%m-%d-%Y')
    cur = conn.cursor()
    if sts == 'Followed up':
        n = Data.get('Note')
        print(n)
        m = Data.get('Mode')
        print(m)
        fu= int(Data.get('Followups'))
        print(fu)
        fu=fu+1
        if ed.date() < comExpDate.date():
            print("In Expiry Date -------------------------------------")
            ed = "null"
            cur.execute("update NONAE set Status = ?, mode = ?,Correspondence= ?, Followups = ?,DateCompleted = ?, username = ? where PolicyNum like ?",(sts,m,n,fu,dto,UserName,PN))
            cur.commit()
        else:
            ed = ed
            print("IN else        .....................................")
            cur.execute("update NONAE set Status = ?, mode = ?,Correspondence= ?, Followups = ?,DateCompleted = ?, username = ?, Expecteddate = ? where PolicyNum like ?",(sts, m, n, fu, dto, UserName, ed, PN))
            cur.commit()
    return jsonify(Data), 200

@app.route("/Nonaereduction", methods=['GET','POST','OPTIONS'])
@cross_origin(origins="http://192.193.194.40:4200", headers="*")
def Nonaereddata():
    Data = request.get_json()
    print(Data)
    cur = conn.cursor()
    d = Data.get('Date')
    print(d)
    dto = datetime.strptime(d, '%m-%d-%Y %H:%M')
    PN = Data.get('Id')
    sts = Data.get('Status')
    tid = Data.get('TransactionId')
    rcb = Data.get('Checked')
    UserName = Data.get('Uname')
    cur.execute("update NONAE set Status = ?, ReductionTransId = ?, ReductionCheckBox = ?,DateCompleted = ?, username = ? where PolicyNum like ?",(sts,tid,rcb,dto,UserName,PN))
    cur.commit()
    return jsonify(Data), 200


def write_to_file(e):
    file = open("log.txt", "a+")
    file.writelines(str(datetime.today()) + str(e))
    file.close()


if __name__ == '__main__':
    app.secret_key = 'random string'
    app.debug = True
    app.run(host='192.193.194.40', port=5001)
