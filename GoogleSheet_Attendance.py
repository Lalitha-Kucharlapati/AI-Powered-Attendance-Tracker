import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime

def authenticate_google_sheets(creds_file):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
    client = gspread.authorize(creds)
    return client

def update_attendance(master_sheet_url, daily_attendance_file, creds_file):
    client = authenticate_google_sheets(creds_file)
    sheet = client.open_by_url(master_sheet_url).sheet1  # Access the first sheet
    
    # Read Master Student List
    data = sheet.get_all_records()
    master_df = pd.DataFrame(data)
    
    # Read Daily Attendance List
    daily_df = pd.read_excel(daily_attendance_file, header=None, names=["PIN Number"])
    present_students = set(daily_df["PIN Number"].astype(str))
    
    # Process Attendance
    today_date = datetime.now().strftime("%d/%m/%Y")
    master_df[today_date] = master_df['PIN Number'].astype(str).apply(lambda x: "Present" if x in present_students else "Absent")
    
    # Update Google Sheet
    updated_data = [master_df.columns.tolist()] + master_df.values.tolist()
    sheet.update(updated_data)
    print(f"Attendance updated successfully for {today_date}!")

# Example Usage
creds_file = "Service_account.json"  # Replace with your service account JSON key file
master_sheet_url = "https://docs.google.com/spreadsheets/d/1yJSpnThAKKJY29vbL4Fpbc9NoeyBYmdadmyUVZxBO8o/edit?usp=sharing"  # Replace with your Google Sheet URL
daily_attendance_file = "random_20_students.xlsx"  # Replace with the daily attendance file path

update_attendance(master_sheet_url, daily_attendance_file, creds_file)
