
import ESBFunctions
#
# DB Connection info
#
dbInfo = "Driver={SQL Server};Server=SQLPRD04.rpega.com;Database=wsprd;uid=wsprdread;pwd=Th3Th8rdM1n!"

# Email Details
sender = "danny.mayer@pega.com"
#toList = ["danny.mayer@pega.com", "frank.flaherty@pega.com"]
toList = ["ebsws@pega.com"]
From = "ESB Monitoring"
fileLocation = ""

days_back = 7
# Selection String
selectdbdata = """SELECT
  COUNT(*)      AS "Count",
  AVG(duration) AS "Average Duration",
  MIN(duration) AS "Min Duratation",
  MAX(duration) AS "Max Duration",
  userid,
  packagename,
  classname,
  methodname,
  tosystem
FROM
  wsdata.esb_usage
WHERE
    StartTime   > '{0}'
AND UserID NOT LIKE '%@pegasystems.com'
GROUP BY
  userid, packagename, classname, methodname, tosystem
ORDER BY
  packagename, classname, methodname
"""

# Main Code
#
afterDate = ESBFunctions.lookbackDate(days_back)
strAfterDate = afterDate.strftime("%d-%b-%Y")

rows = ESBFunctions.getDBData(dbInfo, selectdbdata.format(afterDate))

countTotal = 0
for row in rows:
    countTotal += row[0]
    print(row)

if len(rows) > 0:
    xslFile = "ServiceSummary-{0}.xlsx".format(strAfterDate)
    ESBFunctions.createSpreadsheet(fileLocation + xslFile,
                                   "Service Summary",
                                   ["Count", "Average Duration", "Min Duration","Max Duration","UserID","Package","Class","Method","To System"],
                                   rows)
    bodyMsg = "Service Summary for the last {0} days".format(days_back)
    bodyMsg += "\nTotal of {0} services called".format(countTotal)
else:
    xslFile = ""
    bodyMsg = "No Services called for the last {0} days".format(days_back)

print(bodyMsg)
subjectLine = "Web Services Summary since {0}".format(strAfterDate)
ESBFunctions.send_results_mail(sender, toList, From, subjectLine, bodyMsg, fileLocation, xslFile, toList)
