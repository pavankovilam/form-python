import pandas as pd
import warnings
from flask import Flask, render_template, request,redirect,url_for
import openpyxl
import random 
import smtplib
glovar =['0','0']
workbook = openpyxl.load_workbook('data\credentials.xlsx')
app = Flask(__name__)


@app.route('/', methods=['GET','POST'])
def login():
    if request.method=='POST':
    # Load the Excel sheet
          
        sheet = workbook.active
        username = request.form['username']
        password = request.form['password']

    # Check if the credentials match
        success = False
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if str(row[0]) == str(username) and str(row[1]) == str(password):
                username = str(row[2])
                success = True
                break

    # Redirect to the dashboard if the credentials are correct
        if success:
            return render_template('homepage.html', uname=username)
        else:
            return render_template('login.html', error_message="Invalid username or password")
    return render_template('login.html')

@app.route('/forgotpassword',methods=['GET','POST'])
def forgotPassword():
    if request.method=='POST':
        emailid = request.form['email']
        df = pd.read_excel('data\credentials.xlsx')
        if emailid not in df['username'].values:
             return render_template('forgot-password.html', error_message="Email not registered. Please check again.")
        elif emailid in df['username'].values:
            OTP=random.randint(100000, 999999)
            glovar[0] = OTP
            glovar[1]= emailid
            otp = str(OTP) + " is your OTP" #this store message and otp as strings  
            msg= otp 
            s = smtplib.SMTP('smtp.gmail.com', 587)
            s.starttls()
            s.login("kovi.pavan.kumar@gmail.com", "tlvurhteqfowztsj")
            user="kovi.pavan.kumar@gmail.com"
            s.sendmail(user,emailid,msg)
            s.quit()
            return redirect(url_for("update"))
    return render_template('forgot-password.html')

@app.route('/update',methods=['GET','POST'])
def update():
    if request.method=='POST':
        otp = request.form['otp']
        if str(glovar[0]) == str(otp):
            new_password = request.form['password']
            df = pd.read_excel('data\credentials.xlsx')

            df.loc[df['username'] == glovar[1], 'password'] = new_password

            df.to_excel('data\credentials.xlsx', index=False)
            return render_template('otp.html', msg="Password updated successfully")
        elif str(glovar[0])==otp:
            return render_template('otp.html', msg="OTP Mismatched")            

    return render_template('otp.html', msg="OTP sent succesfully")


@app.route('/register', methods=['GET', 'POST'])
def register():
    if request.method=='POST':
         
        sheet = workbook.active
        existing_usernames = [cell.value for cell in sheet['A'][1:]]

        new_email = request.form['email']
        new_name = request.form['username']
        password = request.form['password']
        

        if new_email in existing_usernames:
            return render_template('register.html', existance="Email already exists")
        
        next_row = sheet.max_row + 1
        sheet.cell(row=next_row, column=1, value=new_email)
        sheet.cell(row=next_row, column=2, value=password)
        sheet.cell(row=next_row, column=3, value=new_name)
        workbook.save('data\credentials.xlsx')
        return render_template('register.html', existance="ACCOUNT SUCCESSFULLY CREATED")
    return render_template('register.html')

@app.route('/feedback', methods=['GET', 'POST'])
def feedback():
    
    return render_template('feedback.html')


if __name__ == '__main__':
    app.run()
