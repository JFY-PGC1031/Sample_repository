from google.cloud import bigquery
from google.oauth2 import service_account
from google.cloud import secretmanager
from dotenv import load_dotenv
from flask import request
import os
import json
import requests




load_dotenv()


bq_key_secrets_reader = os.getenv("BQ_DATA_READER")
bq_project_id = os.getenv("BQ_PROJECT_ID")

# FOR API ENDPOINT ACCESS
ep_key_project_id = os.getenv("ep_key_project_id")
ep_key_secret_id = os.getenv("ep_key_secret_id")


def get_secret(project_id: str, secret_id: str) -> str:
    client = secretmanager.SecretManagerServiceClient()
    secret_name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"
    response = client.access_secret_version(name=secret_name)
    return response.payload.data.decode("UTF-8")



api_key = get_secret(project_id=ep_key_project_id, secret_id=ep_key_secret_id)



def get_credentials_from_secret(project_id: str, secret_id: str):
    secret_client = secretmanager.SecretManagerServiceClient()
    secret_name = f"projects/{project_id}/secrets/{secret_id}/versions/latest"

    response = secret_client.access_secret_version(
        request={"name": secret_name})
    service_account_info = json.loads(response.payload.data.decode("UTF-8"))

    credentials = service_account.Credentials.from_service_account_info(
        service_account_info)
    return credentials


# FOR BQ ACCESS READER
bq_credentials_reader = get_credentials_from_secret(
    bq_project_id, bq_key_secrets_reader)


bq_client_reader = bigquery.Client(
    credentials=bq_credentials_reader, project=bq_credentials_reader.project_id)