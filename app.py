import os
from datetime import datetime
from flask import Flask
from dotenv import load_dotenv
import logging
import json
import requests
import sys
import traceback
import html
from flask import Flask, request, jsonify
from google.cloud import bigquery
from function import bq_client_reader


# New version of email sender api that receives a "display_name" to customize email display name
EMAIL_SENDER_API = "https://email-sender-api-v2-740032229271.us-west1.run.app/send_email"



def get_email_recipients(email_header):

    try:

        ## ** USE THIS CODE WHEN IN PRODUCTION/LIVE SERVER ** ##
        # query_get_email_recipients_PROD = f"""
        # SELECT email, email_header FROM `pgc-finance-and-accounting.ds_dev.tbl_NS_vs_RMS_Daily_Summary_Email_Recipients`
        # WHERE 1=1
        #     AND is_deleted = False
        #     AND email_header = @email_header
        # """

        ## ** USE THIS CODE WHEN TEST ONLY ** ##
        query_get_email_recipients_TEST = f"""
        SELECT email, email_header FROM `pgc-finance-and-accounting.ds_dev.tbl_NS_vs_RMS_Daily_Summary_Email_Recipients_TEST`
        WHERE 1=1
            AND is_deleted = False
            AND email_header = @email_header
        """

        query_parameters = [
        bigquery.ScalarQueryParameter("email_header", "STRING", email_header)]

        job_config = bigquery.QueryJobConfig(query_parameters=query_parameters)
        results = bq_client_reader.query(query_get_email_recipients_TEST, job_config=job_config).result()

# S

        email = []

        for row in results:
            email.append(row['email'])

        print(f"""These are the email for the email header: {email_header}""")
        print(email)

        # Return the values in a list
        return email

    except Exception as e:
        log_api_error_activity("NS vs RMS: Daily Email Sender - get_email_recipients function", e)


def get_daily_summary():
    try:
        query_daily_summary = f"""SELECT
                        category_name,
                        match_count,
                        not_in_ns_count,
                        not_in_rms_count
                    FROM `pgc-finance-and-accounting.ds_dev.pt_NS_vs_RMS_daily_and_weekly_summary_table`
                        WHERE captured_date = CURRENT_DATE('Asia/Manila')
                        ORDER BY category_name ASC"""

        job_config = bigquery.QueryJobConfig()
        results = bq_client_reader.query(query_daily_summary, job_config=job_config).result()

        html_lines = []

        for row in results:
            category = row['category_name']
            match_count = row['match_count']
            not_in_ns_count = row['not_in_ns_count']
            not_in_rms_count = row['not_in_rms_count']

            html_lines.append(f""" <tr><td style='text-align:left;'>{category}</td><td style='text-align:right;'>{match_count}</td><td style='text-align:right;'>{not_in_ns_count}</td><td style='text-align:right;'>{not_in_rms_count}</td>""")

        html_body = "\n".join(html_lines)
        formatted_email_body, subject = get_email_body(html_body)

        send_email(formatted_email_body, subject)
        # print(EMAIL_SENDER_API)

    except Exception as e:
        log_api_error_activity("NS vs RMS: Daily Email Sender - get_daily_summary function", e)



def send_email(hardcoded_data, subject):
    try:


        To = get_email_recipients("To")
        Cc = get_email_recipients("Cc")
        Bcc = ""
        # Bcc = get_email_recipients("Bcc")

        print(Cc)

        # Pass the email body from a function into a variable
        emailbody = hardcoded_data

        print(emailbody)

        # TODO: Uncomment this for production use
        payload = json.dumps({
            "receiver_email": To,
            "display_name": "NS vs RMS Notifier",
            "cc_email": Cc,
            "bcc_email": Bcc,
            "subject": subject,
            "body": emailbody,
            "is_html": 'true'
        })

        # TODO: Comment out this for production use
        # payload = json.dumps({
        #     "receiver_email": "jhon.yleana@primergrp.com",
        #     "From": "NS vs RMS Notifier <notification@primergrp.com>",
        #     "cc_email": ["jhon.yleana@primergrp.com"],
        #     "bcc_email": ["err.senson@primergrp.com "],
        #     "subject": subject,
        #     "body": emailbody,
        #     "is_html": 'true'
        # })
        headers = {
            'Content-Type': 'application/json'
        }
        response = requests.request(
        # Comment when testing if Cc, Bcc, To is properly retrieved to prevent confusion
            "POST", EMAIL_SENDER_API, headers=headers, data=payload
            )

        if response.status_code == 200:
            print("Email sent successfully!")
            logging.info("Email sent successfully.")
        else:
            print(f"Failed to send email. Status Code: {response.status_code}")
            logging.error(
                f"Failed to send email. Status Code: {response.status_code}")
            log_api_error_activity("NS vs RMS: Daily Email Sender - send_email function", response.text)


    except Exception as e:
        log_api_error_activity("NS vs RMS: Daily Email Sender - send_email function", e)



def get_email_body(summary_data):
    try:
        current_date = datetime.now()
        date_to_extract = current_date
        email_date =  date_to_extract.strftime("%B %d, %Y")
        subject = f"""NS vs RMS Syncing Daily Report Summary ({email_date})"""
        LOOKER_REPORT_LINK  = "https://lookerstudio.google.com/reporting/eaa32bdc-6edf-423c-ae7e-422d61922fb4"
        sanitized_looker_link = html.escape(LOOKER_REPORT_LINK)

        body = f"""
                <html>
                <head>
                </head>
                <body>
                    <b> Hi All,</b>
                    <br />
                    <br />
                    <b>This is an automated NS VS RMS: Daily Syncing Report Email Notification, as of <b style='color:red'> {email_date} </b>
                    <br />
                    <br />
                    <table border = '1' style="border-collapse: collapse; width: 50%;">
                        <tr style="padding: 7px;">
                            <td style='text-align:center; background-color: #89CFF0;'>Category Name</td><td style='text-align:center; background-color: #89CFF0;'>Match Count</td><td style='text-align:center; background-color: #89CFF0;'>Not In NS Count</td><td style='text-align:center; background-color: #89CFF0;'>Not In RMS Count</td>
                            {summary_data}
                        </tr>
                    </table>
                    <br />

                    Please continue using the tool and evaluate if the current document-to-document matching approach meets your needs or requires further improvement. Let’s coordinate on the next steps at your earliest convenience.
                    <br />
                    <br />
                    <p>
                         For detailed information, You can access the Looker Studio report via this link: <a href = "{sanitized_looker_link}"> NS vs RMS Data Sync Validation</a>
                    </p>
                    <br />
                    <br />
                    This is a system generated email. No need to reply.
                </body>
                </html>
        """.format(email_date=email_date, summary_data=summary_data, sanitized_looker_link=sanitized_looker_link)

        # Returning multiple values in this format (tuple), for it to be accessible

        return body, subject

    except Exception as e:
        log_api_error_activity("NS vs RMS: Daily Email Sender - get_email_body function", e)



# To display the error in a user-friendly way when debugging
def log_api_error_activity(title, e):
    StartDate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    exc_type, exc_obj, exc_tb = sys.exc_info()
    file_name = exc_tb.tb_frame.f_code.co_filename
    line_number = exc_tb.tb_lineno

    print(f"Error occurred in file: {file_name}")
    print(f"Error occurred on line: {line_number}")
    print(f"Error message: {e}")

    log_api_activity(StartDate, title, "Failed",
                     f"Error in file {file_name} at line {line_number}", str(e))


def log_api_activity(StartDate, LogTitle, Status, ErrorMessage, Remarks):
    try:
        url = f"https://dma-dev-job-logs-174874363586.us-west1.run.app"
        EndDate = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        print(
            f"Parameters being sent to log_api_activity: {StartDate}, {EndDate}, {LogTitle}, {Status}, {ErrorMessage}, {Remarks}")

        payload = {
            "data": {
                "StartDate": StartDate,
                "EndDate": EndDate,
                "LogTitle": LogTitle,
                "Status": Status,
                "ErrorMessage": ErrorMessage,
                "Remarks": Remarks
            }
        }

        response = requests.post(url, json=payload)

        if response.status_code == 200:
            print(f"API activity logged successfully: {response.text}")
        else:
            print(
                f"Failed to log API activity: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"Exception during logging API activity: {e}")


get_daily_summary()
