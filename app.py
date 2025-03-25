import os
from datetime import datetime
from flask import Flask
import pyodbc
from dotenv import load_dotenv
import logging
import json
import requests
from plyer import notification

current_date = datetime.now()
folderName = current_date.strftime('%m-%Y')
logFileName = current_date.strftime('%m-%d-%Y')

# current_directory = os.path.dirname(__file__)
# logging.basicConfig(filename=os.path.join(current_directory, "logs", folderName, f"LOG-{logFileName}.log"),
#                     level=logging.INFO, format='%(asctime)s | %(message)s')

current_directory = os.path.dirname(__file__)
log_folder_path = os.path.join(current_directory, "logs", folderName)
# Creates the folder if it doesn't exist
os.makedirs(log_folder_path, exist_ok=True)
log_file_path = os.path.join(log_folder_path, f"LOG-{logFileName}.log")
logging.basicConfig(filename=log_file_path,
                    level=logging.INFO,
                    format='%(asctime)s | %(message)s')


icon_path = os.path.join(current_directory, "process.ico")

load_dotenv()

DB_SERVER = os.getenv('DB_SERVER')
DB_NAME = os.getenv('DB_NAME')
DB_USERNAME = os.getenv('DB_USERNAME')
DB_PASSWORD = os.getenv('DB_PASSWORD')
DB_DRIVER = '{ODBC Driver 17 for SQL Server}'

EMAIL_SENDER_API = "https://emailsender-740032229271.us-central1.run.app/send_email"


def connect_db():
    try:
        conn = pyodbc.connect(
            f'DRIVER={DB_DRIVER};SERVER={DB_SERVER};DATABASE={DB_NAME};UID={DB_USERNAME};PWD={DB_PASSWORD}')

        # print("Connected to database.")
        # logging.info(f"Connected the database.")
        return conn
    except Exception as e:
        logging.info(f"Error connecting to database: {e}")
        print(f"Error connecting to database: {e}")
        return None


def get_invalid_logout():
    try:
        conn = connect_db()
        cursor = conn.cursor()

        notification.notify(
            title='HCM Email Sender',
            message='Starting to process the data.',
            timeout=10,
            app_icon=icon_path
        )

        # query = '''select top 2 * from vw_hcm_invalid_attempts'''
        query = '''
            select * from vw_hcm_invalid_attempts
            where cast(rtrim(EMP_STAFFID) as varchar(50)) + '-' + cast(ATTENDANCE_DATE as varchar(50)) not in (
            select cast(rtrim([emp_staffid]) as varchar(50)) + '-' + cast([attendanceDate] as varchar(50)) from tbl_Invalid_logs_execution_logs_per_employee)
            and format(ATTENDANCE_DATE,'MM-dd-yyyy') = format(DATEADD(day, -1, getdate()),'MM-dd-yyyy')
            and ATT_ACCESS_IP_OUT <> ''
            '''
        cursor.execute(query)

        records = cursor.fetchall()

        for row in records:
            empid = row[0]
            attendance_date = row[1]
            # formatted_date = attendance_date.strftime(
            #     "%m-%d-%Y")

            if not is_executed(empid.rstrip(), attendance_date, "unauthorized"):
                out_date = row[2]
                ip_address = row[3]

                email_subject = "Unauthorized HCM Sign-Out Outside Company Premises"

                send_email(empid, email_subject, get_out_date(empid.rstrip(), attendance_date),
                           ip_address, "unauthorized")

                save_execution_logs(empid, attendance_date,
                                    "unauthorized", "logout")

                remove_logs(empid, attendance_date)

                result = f"Employee ID:{empid.rstrip()} | Attendance Date: {attendance_date} | Logs Time: {out_date} | IP Address: {ip_address}"
                print(result)
                logging.info(result)

        print("Done processing the data for invalid logout.")
        logging.info("Done processing the data for invalid logout.")

    except Exception as e:
        logging.info(f"ERROR: Fetching employee logout data {e}")
        print(f"ERROR: Fetching employee logout data {e}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_invalid_login_for_ip_sharing():
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''SELECT * FROM vw_hcm_ip_sharing_login where ATT_ACCESS_IP <> '' and format(ATTENDANCE_DATE,'MM-dd-yyyy') = format(DATEADD(day, -1, getdate()),'MM-dd-yyyy')'''
        cursor.execute(query)

        records = cursor.fetchall()

        for row in records:
            attendance_date = row[0]
            att_access_ip = row[1]
            emp_ids = row[2]

            ids = emp_ids.split(",")
            for emp_id in ids:
                if not is_executed(emp_id.rstrip(), attendance_date, "sharing"):
                    print(f"Processing Employee ID: {emp_id.rstrip()}")

                    in_date = get_in_date(emp_id.rstrip(), attendance_date)

                    if isinstance(in_date, str):
                        try:
                            attendance_date_obj = datetime.strptime(
                                in_date, "%m-%d-%Y %H:%M:%S")
                        except ValueError as e:
                            logging.error(
                                f"Date format error for Employee {emp_id.rstrip()}: {in_date} - {e}")
                            continue
                    else:
                        attendance_date_obj = in_date

                    formatted_date = attendance_date_obj.strftime(
                        "%m-%d-%Y %I:%M:%S %p")

                    email_subject = "Unauthorized Sign-In / Out using other devices or Sharing of a Single Device [LOGIN]"

                    send_email(emp_id.rstrip(), email_subject,
                               formatted_date, att_access_ip, "sharing")

                    save_execution_logs(
                        emp_id.rstrip(), attendance_date, "sharing", "login")

                    result = (f"Employee ID: {emp_id.rstrip()} | Attendance Date: {attendance_date} "
                              f"| Logs Time: {formatted_date} | IP Address: {att_access_ip}")
                    print(result)
                    logging.info(result)

        print("Done processing the data for login IP address sharing.")
        logging.info("Done processing the data for login IP address sharing.")

    except Exception as e:
        logging.error(f"ERROR: processing IP sharing login: {e}")
        print(f"ERROR: processing IP sharing login: {e}")

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_invalid_logout_for_ip_sharing():
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''SELECT * FROM vw_hcm_ip_sharing_logout where ATT_ACCESS_IP_OUT <> '' and format(ATTENDANCE_DATE,'MM-dd-yyyy') = format(DATEADD(day, -1, getdate()),'MM-dd-yyyy')'''
        cursor.execute(query)

        records = cursor.fetchall()

        for row in records:
            attendance_date = row[0]
            att_access_ip = row[1]
            emp_ids = row[2]

            ids = emp_ids.split(",")
            for emp_id in ids:
                if not is_executed(emp_id.rstrip(), attendance_date, "sharing"):
                    print(f"Processing Employee ID: {emp_id.rstrip()}")

                    # out_date = get_out_date(emp_id.rstrip(), attendance_date)
                    # print(out_date)

                    # if isinstance(out_date, str):
                    #     try:
                    #         attendance_date_obj = datetime.strptime(
                    #             out_date, "%Y-%m-%d %H:%M:%S")
                    #     except ValueError as e:
                    #         logging.error(
                    #             f"Date format error for Employee {emp_id.rstrip()}: {out_date} - {e}")
                    #         continue
                    # else:
                    #     attendance_date_obj = out_date

                    out_date = get_out_date(emp_id.rstrip(), attendance_date)

                    if isinstance(out_date, str):
                        try:
                            attendance_date_obj = datetime.strptime(
                                out_date, "%m-%d-%Y %H:%M:%S")
                        except ValueError as e:
                            logging.error(
                                f"Date format error for Employee {emp_id.rstrip()}: {out_date} - {e}")
                            continue
                    else:
                        attendance_date_obj = out_date

                    formatted_date = attendance_date_obj.strftime(
                        "%m-%d-%Y %I:%M:%S %p")

                    print(f"Formatted Date: {formatted_date}")

                    email_subject = "Unauthorized Sign-In / Out using other devices or Sharing of a Single Device [LOGOUT]"

                    print("Sending email...")
                    send_email(emp_id.rstrip(), email_subject,
                               formatted_date, att_access_ip, "sharing")

                    save_execution_logs(
                        emp_id.rstrip(), attendance_date, "sharing", "logout")

                    result = (f"Employee ID: {emp_id.rstrip()} | Attendance Date: {attendance_date} "
                              f"| Logs Time: {formatted_date} | IP Address: {att_access_ip}")
                    print(result)
                    logging.info(result)

        print("Done processing the data for logout IP address sharing.")
        logging.info("Done processing the data for logout IP address sharing.")

    except Exception as e:
        logging.error(f"ERROR: Processing IP sharing login: {e}")
        print(f"ERROR: Processing IP sharing login: {e}")

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_emp_name(empid):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        SELECT
            COALESCE(emp_firstname, '') + ' ' + COALESCE(emp_middlename, '') + ' ' + COALESCE(emp_lastname, '') AS EMP_FULLNAME
        FROM [dbo].[ERM_EMPLOYEE_MASTER]
        WHERE EMP_STAFFID = ?
        '''
        cursor.execute(query, (empid,))

        record = cursor.fetchone()
        employee_name = record[0]

        if record:
            return employee_name
        else:
            logging.info(f"Employee {empid} has no name.")
            return "No Name"

    except Exception as e:
        print(f"ERROR: Fetching employee name: {e}")
        logging.info(f"ERROR: Fetching employee name: {e}")
        return None

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_emp_email(empid):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        select emp_mailid from [dbo].[ERM_EMPLOYEE_MASTER] where EMP_STAFFID = ?
        '''

        cursor.execute(query, (empid,))
        record = cursor.fetchone()
        emp_email = record[0]

        if record:
            return emp_email
        else:
            logging.info(f"Employee {empid} has no email.")
            return "No Email"

    except Exception as e:
        print(f"ERROR: Fetching employee email: {e}")
        logging.info(f"ERROR: Fetching employee email: {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_emp_il_id(empid):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        select EMP_REPORTINGTO from [dbo].[ERM_EMPLOYEE_MASTER] where EMP_STAFFID = ?
        '''

        cursor.execute(query, (empid,))
        record = cursor.fetchone()
        idl_id = record[0]

        if record:
            return idl_id
        else:
            logging.info(f"Employee {empid} has no IL.")
            return "No IL ID"

    except Exception as e:
        print(f"ERROR: Fetching IL ID: {e}")
        logging.info(f"ERROR: Fetching IL ID: {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_emp_il_name(il_id):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        select emp_firstname + ' ' + emp_middlename + ' ' + emp_lastname as EMP_FULLNAME from [dbo].[ERM_EMPLOYEE_MASTER] where EMP_STAFFID = ?
        '''

        cursor.execute(query, (il_id,))
        record = cursor.fetchone()
        il_id = record[0]

        if record:
            return il_id
        else:
            logging.info(f"Employee {il_id} has no name.")
            return "No IL Name"

    except Exception as e:
        print(f"ERROR: Fetching IL Name: {e}")
        logging.info(f"ERROR: Fetching IL Name: {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_emp_il_email(il_id):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
       select emp_mailid from [dbo].[ERM_EMPLOYEE_MASTER] where EMP_STAFFID = ?
        '''

        cursor.execute(query, (il_id,))
        record = cursor.fetchone()
        il_email = record[0]

        if record:
            return il_email
        else:
            logging.info(f"Employee {il_id} has no email.")
            return "No Email"

    except Exception as e:
        print(f"ERROR: Fetching IL Email: {e}")
        logging.info(f"ERROR: Fetching IL Email: {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_out_date(empid, attendance_date):
    try:
        #     conn = connect_db()
        #     cursor = conn.cursor()

        #     query = '''
        #    select out_date from ATTENDANCE_REGISTER where EMP_STAFFID = ? and ATTENDANCE_DATE = ?
        #     '''
        #     cursor.execute(query, (empid, attendance_date))

        #     record = cursor.fetchone()
        #     out_date = record[0]

        #     # if record:
        #     #     out_date_obj = datetime.strptime(out_date, "%Y-%m-%d %H:%M:%S")
        #     #     return out_date_obj
        #     # else:
        #     #     logging.info(
        #     #         f"Employee {empid} has no out date for the given attendance date.")
        #     #     return "No out date"

        #     if isinstance(out_date, str):
        #         try:
        #             attendance_date_obj = datetime.strptime(
        #                 out_date, "%m-%d-%Y %H:%M:%S")
        #         except ValueError as e:
        #             logging.error(
        #                 f"Date format error for Employee {empid.rstrip()}: {out_date} - {e}")
        #     else:
        #         attendance_date_obj = out_date

        #     formatted_date = attendance_date_obj.strftime(
        #         "%m-%d-%Y %I:%M:%S %p")

        #     return formatted_date

        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        select out_date from ATTENDANCE_REGISTER where EMP_STAFFID = ? and ATTENDANCE_DATE = ?
        '''
        cursor.execute(query, (empid, attendance_date))

        record = cursor.fetchone()
        out_date = record[0]

        if record:
            return out_date
        else:
            logging.info(
                f"Employee {empid} has no OUT date for the given attendance date.")
            return "No out date"

    except Exception as e:
        print(f"ERROR: Fetching out date: {e}")
        logging.error(f"ERROR: Fetching out date for employee {empid}: {e}")
        return None

# def get_out_date(empid, attendance_date):
#     try:
#         conn = connect_db()
#         cursor = conn.cursor()

#         query = '''
#         select out_date from ATTENDANCE_REGISTER where EMP_STAFFID = ? and ATTENDANCE_DATE = ?
#         '''
#         cursor.execute(query, (empid, attendance_date))

#         record = cursor.fetchone()
#         out_date = record[0]

#         if out_date:
#             try:
#                 out_date_obj = datetime.strptime(out_date, "%Y-%m-%d %H:%M:%S")

#                 formatted_out_date = out_date_obj.strftime(
#                     "%Y-%m-%d %H-%M-%S %p")
#                 return formatted_out_date

#             except ValueError as e:
#                 logging.error(
#                     f"Date format error for Employee {empid}: {out_date} - {e}")
#                 return "Invalid date format"
#         else:
#             logging.info(
#                 f"Employee {empid} has no out date for the given attendance date.")
#             return "No out date"

#     except Exception as e:
#         print(f"ERROR: Fetching out date: {e}")
#         logging.error(f"ERROR: Fetching out date for employee {empid}: {e}")
#         return None

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_in_date(empid, attendance_date):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
       select in_date from ATTENDANCE_REGISTER where EMP_STAFFID = ? and ATTENDANCE_DATE = ?
        '''
        cursor.execute(query, (empid, attendance_date))

        record = cursor.fetchone()
        in_date = record[0]

        if record:
            return in_date
        else:
            logging.info(
                f"Employee {empid} has no IN date for the given attendance date.")
            return "No out date"

    except Exception as e:
        print(f"Error fetching out date: {e}")
        logging.error(f"ERROR: Fetching IN date for employee {empid}: {e}")
        return None

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def send_email(empid, subject, attendance_date, ip_address, type):
    try:
        cc = get_emp_il_email(get_emp_il_id(empid))
        bcc = os.getenv('HRVP_EMAIL')

        if type == "unauthorized":
            emailbody = get_email_body(empid, attendance_date, ip_address)
        else:
            emailbody = get_email_body_for_sharing_ip(
                empid, attendance_date, ip_address)

        # TODO: Uncomment this for production use
        payload = json.dumps({
            "receiver_email": get_emp_email(empid),
            "cc_email": [cc],
            "bcc_email": [bcc],
            "subject": subject,
            "body": emailbody,
            "is_html": 'true'
        })

        # TODO: Comment out this for production use
        # payload = json.dumps({
        #     "receiver_email": "gilbert.laman@primergrp.com",
        #     "cc_email": ["jhon.yleana@primergrp.com"],
        #     "bcc_email": ["iris.garcia@primergrp.com"],
        #     "subject": subject,
        #     "body": emailbody,
        #     "is_html": 'true'
        # })
        headers = {
            'Content-Type': 'application/json'
        }

        response = requests.request(
            "POST", EMAIL_SENDER_API, headers=headers, data=payload)

        if response.status_code == 200:
            print("Email sent successfully!")
            logging.info("Email sent successfully.")
            logging.info(
                f"Email was also sent to {get_emp_il_name(get_emp_il_id(empid))}")
        else:
            print(f"Failed to send email. Status Code: {response.status_code}")
            logging.error(
                f"Failed to send email. Status Code: {response.status_code}")

    except Exception as e:
        print(f"ERROR: Sending email: {e}")
        logging.info(f"ERROR: Sending email: {e}")


def save_execution_logs(empid, attendance_date, type, logtype):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        insert into tbl_Invalid_logs_execution_logs_per_employee (emp_staffid, dateSent, timeSent, type, attendanceDate, logtype) values (?, getdate(), format(getdate(),'hh:mm:ss tt'), ?,?,?)
            '''
        cursor.execute(query, (empid.strip(), type, attendance_date, logtype))

        conn.commit()

        print(f"Log saved successfully for {empid}.")
        logging.info(f"Log saved successfully for {empid}.")
    except Exception as e:
        print(f"ERROR: Saving execution log: {e}")
        logging.error(
            f"ERROR: Saving execution log for Employee {empid} on {attendance_date}: {e}")

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_login(empid, attendance_date):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        query = '''
        select IN_DATE from [Primer_Prod3].[dbo].[ATTENDANCE_REGISTER] where emp_staffid = ? and attendance_date = ?
        '''
        cursor.execute(query, (empid.rstrip(), attendance_date))
        record = cursor.fetchone()
        in_date = record[0]

        if record:
            return in_date
        else:
            logging.info(f"Employee {empid} has no email.")
            return "IN_DATE"

    except Exception as e:
        print(f"ERROR: Fetching employee IN_DATE: {e}")
        logging.info(f"ERROR: Fetching employee IN_DATE: {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def remove_logs(empid, attendance_date):
    try:
        conn = connect_db()
        cursor = conn.cursor()

        in_date = get_login(empid, attendance_date)

        query = f'''
        update [Primer_Prod3].[dbo].[ATTENDANCE_REGISTER] set OUT_DATE = '{in_date}' where emp_staffid = ? and attendance_date = ?
            '''
        cursor.execute(query, (empid.strip(), attendance_date))

        conn.commit()

        print(f"Log successfully removed for {empid.rstrip()}.")
        logging.info(f"Log successfully removed for {empid.rstrip()}.")
    except Exception as e:
        print(f"ERROR: {e}")
        logging.error(
            f"ERROR: {e}")
    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def is_executed(empid, attendance_date, type):
    try:
        # if isinstance(attendance_date, str):
        #     formatted_attendance_date = datetime.strptime(
        #         attendance_date, '%Y-%m-%d')
        #     print(formatted_attendance_date)

        conn = connect_db()
        cursor = conn.cursor()

        query = '''
            SELECT COUNT(*) AS [count] 
            FROM tbl_Invalid_logs_execution_logs_per_employee 
            WHERE rtrim(emp_staffid) = rtrim(?) 
            AND type = ?
            AND format(CAST(attendanceDate AS DATE),'MM/dd/yyyy') = format(CAST(? AS DATE),'MM/dd/yyyy')
        '''
        cursor.execute(query, (empid, type, attendance_date))

        record = cursor.fetchone()

        if record:
            count = record[0]
            print(
                f"Count: {count} | Employee ID: {empid} | Type: {type} | Attendance Date: {attendance_date}")
            return count > 0
        else:
            return False

    except Exception as e:
        print(f"Error in is_executed: {e}")
        return False

    finally:
        if cursor:
            cursor.close()
        if conn:
            conn.close()


def get_email_body(empid, attendance_date, emp_ip):
    body = f"""
    Dear <b>{get_emp_name(empid)}</b>,
    <p></p>
    I hope this message finds you well.
    <p></p>
    It has come to our attention that an unauthorized sign-out was recorded under your account in the HCM Web-Based Attendance <b>{attendance_date}</b> thru <b style='color:red'>{emp_ip}</b>.
    <p></p>
    Please note that signing-in or out outside company premises may be considered as an infraction of our Company's Code of Discipline - C.1 Timing in/out, logging in/out, or inappropriate use of time records, for another employee or conspiring with or arranging for another to do the same.
    <p></p>
    This sign-out will be rejected and we kindly ask you to provide clarification to your IL & HRBP on this matter within 48 hours and ensure that such actions are avoided in the future.
    <p></p>
    If this was done in error, or if there are any extenuating circumstances, please reach out to your IL and HRBP as soon as possible so we can rectify the situation.
    <p></p>
    If you require assistance with the attendance system or have any questions, please don’t hesitate to email hcmsupport@primergrp.com
    <p></p>
    Thank you for your attention to this matter.
    <p></p>
    <p></p>
    Best regards,
    <br />
    Human Resources Department
    """

    return body


def get_email_body_for_sharing_ip(empid, attendance_date, emp_ip):
    body = f"""
    Dear <b>{get_emp_name(empid)}</b>,
    <p></p>
    I hope this message finds you well.
    <p></p>
    It has come to our attention that an unauthorized log was recorded under your account in the HCM Web-Based Attendance <b>{attendance_date}</b> thru <b style='color:red'>{emp_ip}</b>.
    <p></p>
    Please note that logging thru a different device may be considered as an infraction of our Company's Code of Discipline - C.1 Timing in/out, logging in/out, or inappropriate use of time records, for another employee or conspiring with or arranging for another to do the same.
    <p></p>
    We kindly ask you to provide clarification to your IL & HRBP on this matter within 48hours and ensure that such actions are avoided in the future.
    <p></p>
    If this was done in error, or if there are any extenuating circumstances, please reach out to your IL and HRBP as soon as possible so we can rectify the situation.
    <p></p>
    If you require assistance with the attendance system or have any questions, please don’t hesitate to email hcmsupport@primergrp.com
    <p></p>
    Thank you for your attention to this matter.
    <p></p>
    <p></p>
    Best regards,
    <br />
    Human Resources Department
    """

    return body


get_invalid_logout()
get_invalid_login_for_ip_sharing()
get_invalid_logout_for_ip_sharing()
