import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def mail_files(bfile):
    username = "martijn.hoogendijk@inter.ikea.com"
    password = "Welkom@1989_1"
    mail_from = "martijn.hoogendijk@inter.ikea.com"
    mail_to = "martijn.hoogendijk@inter.ikea.com"
    mail_subject = "Rydoo real time data download"
    mail_body = "Rydoo real time data download"
    mail_attachment = bfile
    mail_attachment_name = bfile


    mimemsg = MIMEMultipart()
    mimemsg['From'] = mail_from
    mimemsg['To'] = mail_to
    mimemsg['Subject'] = mail_subject
    mimemsg.attach(MIMEText(mail_body, 'plain'))

    with open(mail_attachment, "rb") as attachment:
        mimefile = MIMEBase('application', 'octet-stream')
        mimefile.set_payload((attachment).read())
        encoders.encode_base64(mimefile)
        mimefile.add_header('Content-Disposition', "attachment; filename= %s" % mail_attachment_name)
        mimemsg.attach(mimefile)
        connection = smtplib.SMTP(host='smtp.office365.com', port=587)
        connection.starttls()
        connection.login(username, password)
        connection.send_message(mimemsg)
        connection.quit()


if __name__=="__main__":
    mail_files(bfile=r'rydoo_approvers_20220815-090709.txt')