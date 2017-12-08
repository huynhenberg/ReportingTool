import pypyodbc
import smtplib
import time, datetime

sender = "danny.mayer@pega.com"
TList = ["danny.mayer@pega.com", "frank.flaherty@pega.com"]
toList = ["Brian.Wainwright@pega.com", "Lisa.Cavallaro@pega.com", "Mary.Aylward@pega.com","John.Farrell@pega.com", "Jabin.Geary@pega.com","Thakur.Gyawali@pega.com", "Abhay.Thomas@pega.com", "danny.mayer@pega.com", "christopher.deverna@pega.com","Brandon.Yeh@pega.com","Sachin.Gupta@in.pega.com", "Bibhoo.Shrestha@pega.com", "Ronald.Czik@pega.com", "Murtaza.Qureshi@pega.com", "Jacob.Berg@pega.com","sapsdsupport@pega.com"]
From = "ESB SAP Error Monitoring"
#smtpServer = 'VEBS102K8.RPEGA.COM'
smtpServer = 'smtpapp.rpega.com'

days_back = 1
# Selection String
selectdbdata = """ select u.id, r.errortype, r.errormessage, u.packagename,u.classname, u.methodname, u.userid, u.starttime, u.statuscode, u.statusdescription
from wsdata.esb_usage u
LEFT JOIN wsdata.sap_err_log r
ON r.ID = u.ID
where u.StartTime > '{0}'
and u.packagename = 'EBSSAP'
and u.statuscode <> 'OK'
order by u.userid, u.starttime"""

dbInfo = "Driver={SQL Server};Server=SQLPRD04.rpega.com;Database=wsprd;uid=wsprdread;pwd=Th3Th8rdM1n!"

# Function to send email
def send_errors_mail(error_summary,PACErrors, PAMSErrors, dateSince, toList):
    receivers = toList

    receivers_string = ",".join(receivers)

    message = """From:{0}
To: {1}
Subject: SAP Error Messages since {3}

Errors logged since {3}:
{2}
{4}
{5}
""".format(From, receivers_string, error_summary, dateSince, PACInputErrors, PAMSValidationErrors)


    try:
        smtpObj = smtplib.SMTP(smtpServer)
        smtpObj.sendmail(sender, receivers, message)
        print("Successfully sent email")
    # except SMTPException:
    except Exception:
        print("Error: unable to send email")

def dump_rows_to_file(filename, rows):
    fd = open(filename, "w")
    for row in rows:
        fd.write(str(row) + "\n")
    fd.close()

# End of dump_rows_to_file

today = datetime.date.today()
afterDate = today - datetime.timedelta(days=days_back)
strAfterDate = afterDate.strftime("%d-%b-%Y")

connection = pypyodbc.connect(dbInfo)
result = connection.cursor().execute(selectdbdata.format(afterDate))
rows = result.fetchall()

currentID = ""
outbuf = "User, ErrorCode,        Time,               Service,            ESB ID,   Error\n"
outbuf += "     ErrType: ErrDescription\n"

countPACInputErrors = 0
countPAMSValidateErrors = 0

for row in rows:
    service = row[3] + "." + row[4] + "." + row[5]
    if row[0] != currentID:
        if service == "EBSSAP.Customer.GetPartners" and row[6] == "PAC":
            countPACInputErrors += 1
        elif service == "EBSSAP.Customer.ValidateAddress" and row[6] ==  "PAMS":
            countPAMSValidateErrors += 1
        else:
            outbuf += "\n{0}, {1}, {2}, {3}, {4}, {5}\n".format(row[6], row[8], row[7], service, row[0], row[9])
            if row[1] != None:
                outbuf += "    {0}: {1}\n".format(row[1], row[2])

    if row[1] != None and row[0] == currentID:
        outbuf += "    {0}: {1}\n".format(row[1],row[2])
    currentID = row[0]
#    print(row)

PACInputErrors ="{0} EBSSAP.Customer.GetPartners PAC Input Errors".format(countPACInputErrors)
PAMSValidationErrors = "{0} EBSSAP.Customer.ValidateAddress PAMS Different Postal Codes".format(countPAMSValidateErrors)
dumpfile = """C:\\SAPErrorLogs\\SAPErrors_{0}.log""".format(strAfterDate)
dump_rows_to_file(dumpfile, rows)

#print(outbuf)
#print(PACInputErrors)
#print(PAMSValidationErrors)
send_errors_mail(outbuf, PACInputErrors, PAMSValidationErrors, strAfterDate, toList)
#send_errors_mail(outbuf, PACInputErrors, PAMSValidationErrors, strAfterDate, TList)
connection.close()

