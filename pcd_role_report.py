import configparser
import datetime
import esbfunc


if __name__ == '__main__':
    # Get configs
    section = "PCD Registrants SOTW CLOUD"
    config = configparser.ConfigParser()
    config.read("C:\\ReportingTool\\config.ini")
    env = config[section]["env"]
    log_name = config[section]["log_name"]
    today = datetime.date.today().strftime("%Y%m%d")
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
    query = """SELECT reg.registrantID as 'Registrant ID',
        reg.firstname as 'First Name',
        reg.lastname as 'Last Name',
        reg.weblogin as 'Weblogin',
        reg.membertype as 'MemberType',
        reg.businesstitle as 'BusinessTitle',
        reg.phonenumber as 'PhoneNumber',
        acc.accountID as 'Account ID',
        rol.serviceID as 'Service ID',
        rol.roleID as 'Role ID',
        rol.startdate as 'Start Date',
        rol.expirationdate as 'Expiration Date'
    FROM pcd71data.pcd_work_registration reg,
        pcd71data.pcd_work_accountaffiliation acc,
        pcd71data.pcd_work_roleassignment rol
    WHERE reg.registrantID=acc.registrantID
        AND acc.registrantID=rol.registrantID
        AND rol.AccountID = acc.AccountID
        AND reg.RegistrationStatus IN ('Active','MustConfirm')
        AND rol.ServiceID IN ('SOTW','CLOUD')
        AND acc.Active = 'true'
        AND (acc.ExpirationDate is NULL
             OR acc.ExpirationDate > SYSUTCDATETIME())
        AND rol.Active = 'true'
        AND (rol.ExpirationDate is NULL 
             OR rol.ExpirationDate > SYSUTCDATETIME())
    ORDER BY 'Registrant ID', 'Account ID', 'Service ID', 'Role ID'"""
    rows = esbfunc.get_db_data(db_info, query)


    # Write to log
    if (write_log):
        esbfunc.write_rows_to_file(log_path + log_name + ".log", rows)


    # Create spreadsheet
    esbfunc.create_spreadsheet(
        log_path + log_name + ".xlsx", "Registrants",
        ["Registrant ID", "First Name", "Last Name", "Weblogin", "MemberType",
        "Business Title", "Phone Number", "Account ID", "Service ID",
        "Role ID", "Start Date", "Expiration Date"], rows)


    # Send email with attachment
    msg_from = "PCD Reports"
    msg_subject = "PCD Registrants with SOTW and CLOUD as of {0}".format(today)
    msg_body = "PCD Registrants with SOTW and CLOUD as of {0}".format(today)
    sender = "dan.huynh@pega.com"
    esbfunc.send_email_with_attachment(
        smtp_server, sender, msg_from, msg_to, msg_subject, msg_body,
        log_path, log_name + ".xlsx")
