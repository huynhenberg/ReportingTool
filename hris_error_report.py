import configparser
import datetime
import esbfunc


if __name__ == '__main__':
    # Get configs
    section = "HRIS Error Reports"
    config = configparser.ConfigParser()
    config.read("C:\\ReportingTool\\config.ini")
    env = config[section]["env"]
    log_name = config[section]["log_name"]
    today = datetime.date.today().strftime("%Y%m%d")
    days_back = config.getint(section,"days_back")
    from_date = esbfunc.what_date(days_back)
    log_name = log_name.format(today)
    log_path = config[section]["log_path"]
    write_log = config.getboolean(section,"write_log")
    msg_to = config[section]["msg_to"].split(",")

    config.read("C:\\ReportingTool\\connections.ini")
    smtp_server = config["smtp"]["server"]
    driver = driver=config[env]["driver"]
    server = config[env]["server"]
    db = config[env]["db"]
    uid = config[env]["uid"]
    pwd = config[env]["pwd"]


    # Query the DB
    db_info = ("Driver={0};Server={1};Database={2};"
               "uid={3};pwd={4}").format(driver,
               server, db,
               uid, pwd)
    query = """select distinct count(*), userid, packagename, classname,
            methodname, fromsystemname, fromsystemhost, statuscode,
            statusdescription
    from wsdata.esb_usage
    where starttime > '{0}'
        and packagename = 'ebshris'
        and statuscode <> 'ok'
    group by userid, packagename, classname, methodname, fromsystemname,
            fromsystemhost, statuscode, statusdescription
    order by count(*) desc"""
    rows = esbfunc.get_db_data(db_info, query.format(from_date))


    # Write to log
    if (write_log):
        esbfunc.write_rows_to_file(log_path + log_name + ".log", rows)


    # Create spreadsheet
    esbfunc.create_spreadsheet(
        log_path + log_name + ".xlsx", "Registrants",
        ["Count", "UserID", "PackageName", "ClassName", "MethodName",
        "PegaID", "CompanyEmail", "StatusCode", "StatusDescription"], rows)


    # Send email with attachment
    summary = dict()
    for row in rows:
        status_code = row[7]
        count = summary.get(status_code, 0)
        count += row[0]
        summary.update({status_code:count})

    msg_body = "StatusCode\tOccurance\n"
    for k in summary:
        msg_body = msg_body + k + "\t\t" + str(summary[k]) + "\n"
    msg_from = "HRIS Error Reports"
    msg_subject = "HRIS Error Reports from {0}".format(from_date)
    sender = "dan.huynh@pega.com"
    esbfunc.send_email_with_attachment(
        smtp_server, sender, msg_from, msg_to, msg_subject, msg_body,
        log_path, log_name + ".xlsx")
