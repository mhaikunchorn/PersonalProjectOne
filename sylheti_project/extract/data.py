"""Extracting word from Google Sheets"""
import os
import time
from pandas import read_csv, DataFrame, concat
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

start_time = time.time()

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "14rwxC2XCKDYPk_n5eP0Q8kcNSpxNRL1LPWEqGip3IoQ"
SHEETS = [
    "Grammar & Connectives",
    "Numbers & Money",
    "Weather",
    "Time & Frequency",
    "Food",
    "Phrases",
    "Possessions",
    "Pronouns",
    "Body",
    "People",
    "Affirmation & Negation",
    "Work",
    "Place & Travel",
    "Colours"
]

def extract_words():
    """
    Extracting the words and phrases from Google Sheets,
    using an API.
    """
    credentials=None
    if os.path.exists("token.json"):
        credentials=Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("sylheti_project/extract/credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    try:
        dfs = []
        for sheet in SHEETS:
            service = build("sheets","v4", credentials=credentials)
            sheets = service.spreadsheets()
            result = sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=sheet).execute()
            values = result.get("values",[])

            dfs.append(values)

        # create one big df
        df = concat([DataFrame(sheet_data) for sheet_data in dfs],axis=0)

        # --- cleaning the df
        # Clean column names
        df.columns = df.iloc[0]
        df.columns = df.columns.str.lower()
        df=df[1:]

        # if the English words begin with a a first-person pronoun then capitalise

        # make all sylheti words lowercase
        df.sylheti = df.sylheti.str.lower()

        print(df.info())
        print("My program took", time.time() - start_time, "to run")

        df.to_csv("test_df.csv",index=False)
        
        return df

    except HttpError as error:
        print(error)
    

if __name__ == "__main__":
    extract_words()
