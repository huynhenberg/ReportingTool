import datetime
import pypyodbc
import smtplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.utils import COMMASPACE

from openpyxl import Workbook


#
# Function to send email. The msg_from is just the message header while sender
# is the actual email of the sender used for SMTP Return-Path. msg_to is
# assumed to be an array.
#
def send_email_with_attachment(
        smtp_server, sender, msg_from, msg_to, msg_subject, msg_body,
        file_path, file_name):
    msg = MIMEMultipart()
    msg['From'] = msg_from
    msg['To'] = ",".join(msg_to)
    msg['Subject'] = msg_subject
    msg.attach(MIMEText(msg_body, "plain"))

    # Include attachment
    if file_name != "":
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(file_path + file_name, "rb").read())
        email.encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename={0}'.format(file_name))
        msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server)
        server.sendmail(sender, msg_to, msg.as_string())
        print("Successfully sent email")
    except Exception:
        print("Error: unable to send email")


def send_email(smtp_server, sender, msg_from, msg_to, msg_subject, msg_body):
    send_email_with_attachment(smtp_server, sender, msg_from, msg_to,
                               msg_subject, msg_body, "", "")

#
# Fix worksheet cell sizes
#
def fix_worksheet_columns(ws):
    dims = {}
    for row in ws.rows:
        for cell in row:
            if cell.value:
                dims[cell.column] = max((dims.get(cell.column, 0),
                                         len(str(cell.value)) + 3))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value


#
# Create a spreadsheet and populate with data
#
def create_spreadsheet(file_name, title, headings, rowData):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = title

    ws1.append(headings)

    for row in rowData:
        ws1.append(row)

    fix_worksheet_columns(ws1)
    wb.save(file_name)


#
# Fetch the data from the database
#
def get_db_data(db_config, db_query):
    connection = pypyodbc.connect(db_config)
    result = connection.cursor().execute(db_query)
    rowsReturned = result.fetchall()
    connection.close()
    return rowsReturned


#
# Write db rows to a log file.
#
def write_rows_to_file(filename, rows):
    fd = open(filename, "w")
    for row in rows:
        fd.write(str(row) + "\n")
    fd.close()


def what_date(days_back):
    today = datetime.date.today()
    return (today - datetime.timedelta(days=days_back)).strftime("%d-%b-%Y")
