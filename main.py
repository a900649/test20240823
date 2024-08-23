from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib, ssl
import pandas as pd
from io import BytesIO
from flask import Flask, send_from_directory, request, make_response, send_file
import variable as v
import json

app = Flask(__name__)

@app.route("/",methods = ["GET"])
def home():
    return {"Status": "OK"}

@app.route("/send_mail_get/<subject>/<content>/<receiver>",methods = ["GET"])
def send_mail_get(subject,content,receiver):






    smtp_server = v.smtp_server
    port = v.port
    sender = v.sender
    mail_user = v.mail_user
    mail_password = v.mail_password

    server = smtplib.SMTP(smtp_server, port)

    # 建立連線
    context = ssl.create_default_context()
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['FROM'] = sender
    msg['To'] = receiver

    # 建立內文文字
    att2 = MIMEText(content)
    msg.attach(att2)

    server.starttls(context=context)
    server.login(mail_user, mail_password)
    server.sendmail(sender, receiver, msg.as_string())

    return {"Status": "OK"}

@app.route("/send_mail_post",methods = ["POST"])
def send_mail_post():

    smtp_server = v.smtp_server
    port = v.port
    sender = v.sender
    mail_user = v.mail_user
    mail_password = v.mail_password

    convert_string = "@1@2@3@$$#"
    def byte_string_convert(byte_string):
        return_string = byte_string.decode("utf-8").replace("\n", convert_string).replace("," + convert_string, ",").replace("{" + convert_string, "{").replace(convert_string + "}", "}")
        return return_string

    parameter_dict = json.loads(byte_string_convert(request.get_data()))
    subject = parameter_dict["Subject"].replace(convert_string, "\n")
    content = parameter_dict["Content"].replace(convert_string, "\n")
    receiver = parameter_dict["Receiver"].replace(convert_string, "\n")

    server = smtplib.SMTP(smtp_server, port)

    # 建立連線
    context = ssl.create_default_context()
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['FROM'] = sender
    msg['To'] = receiver

    # 建立內文文字
    att2 = MIMEText(content)
    msg.attach(att2)

    server.starttls(context=context)
    server.login(mail_user, mail_password)
    server.sendmail(sender, receiver, msg.as_string())

    return {"Status": "OK"}


@app.route("/download_excel",methods = ["GET","POST"])
def download_excel():

    df = pd.DataFrame([[1, 2], [3, 4]], columns=["A", "B"])
    out = BytesIO()
    writer = pd.ExcelWriter(out, engine='xlsxwriter')
    df.to_excel(excel_writer=writer, sheet_name='Sheet1',)
    writer.close()

    out.seek(0)

    return send_file(out, download_name="output.xlsx",mimetype="application/vnd.ms-excel", as_attachment=True)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=8000)

    # app.run(host="127.0.0.1", port=8000)
