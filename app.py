#!/usr/bin/python3
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import psycopg2
import xlsxwriter
import os
import sys
import smtplib

#usage python3 colum.py filename toaddress

SQL_Code = open(str(sys.argv[3]), 'r').read()

#Connecting to PostgreSQL
def main():
    conn_string = "host='db' dbname='directski' user='pgsql' password=''"
    print ("Connecting to database\n    ->%s" % (conn_string))
    conn = psycopg2.connect(conn_string)
    cursor = conn.cursor()
    print ("Connected!\n")
    cursor.execute(SQL_Code)
    filename = str(sys.argv[1]).replace(" ", "_").lower()
    workbook = xlsxwriter.Workbook(filename + ".xlsx", {'remove_timezone': True})
    worksheet = workbook.add_worksheet()
    data = cursor.fetchall()

    # Headers
    for colidx,heading in enumerate(cursor.description):
        worksheet.write(0, colidx, heading[0])

    # Writing the Rows
    for rowid, row in enumerate(data):
        for colid, col in enumerate(row):
            worksheet.write(rowid+1, colid, col)

    # Saving
    workbook.close()

    fromaddr = "temp@topflight.ie"
    toaddr = str(sys.argv[2])

    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = str(sys.argv[1])

    body = ""

    msg.attach(MIMEText(body, 'plain'))

    attachment = open(filename + ".xlsx", "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename + ".xlsx")

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, "temp123$")
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()

if __name__ == "__main__":
        main()


