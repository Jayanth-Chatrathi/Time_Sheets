import imaplib
import email
import os
import json
import pandas as pd
from dotenv import load_dotenv
from google import genai

# ---------------------------
# Load environment variables
# ---------------------------
load_dotenv()

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# Initialize Gemini client
client = genai.Client(api_key=GEMINI_API_KEY)

# ---------------------------
# Step 1: Fetch Email Body via IMAP
# ---------------------------
def fetch_latest_timesheet_email():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(EMAIL_USER, EMAIL_PASS)
    mail.select("inbox")

    status, messages = mail.search(None, '(SUBJECT "Timesheet")')
    email_ids = messages[0].split()

    if not email_ids:
        raise Exception("No Timesheet email found.")

    latest_email_id = email_ids[-1]

    res, msg_data = mail.fetch(latest_email_id, "(RFC822)")
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    body = ""

    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True).decode(errors="ignore")
                break
    else:
        body = msg.get_payload(decode=True).decode(errors="ignore")

    mail.logout()
    return body


# ---------------------------
# Step 2: Extract Structured Data via Gemini
# ---------------------------
def extract_timesheet_structured(email_text):

    prompt = f"""
You are a strict data extraction engine.

Extract timesheet data from the email below.

Return ONLY valid JSON.

Schema:
{{
  "employee_name": string,
  "week_start_date": string (YYYY-MM-DD),
  "week_end_date": string (YYYY-MM-DD),
  "entries": [
      {{
          "project_date": string (YYYY-MM-DD),
          "project_name": string,
          "hours": float
      }}
  ]
}}

Rules:
- Week must be Monday to Sunday.
- Convert dates to YYYY-MM-DD.
- If a day is missing include it with hours=0.
- No explanations.
- No markdown.
- Only JSON output.

Email:
{email_text}
"""


    response = client.models.generate_content(
    model="gemini-2.5-flash",  # ✅ This exists in your list
    contents=prompt,
    config={
        "response_mime_type": "application/json"
    }
)


    return response.text


# ---------------------------
# Step 3: Safe JSON Parsing
# ---------------------------
def parse_json_safely(json_text):
    try:
        return json.loads(json_text)
    except json.JSONDecodeError:
        print("Invalid JSON returned from Gemini:")
        print(json_text)
        raise


# ---------------------------
# Step 4: Save to CSV/Excel
# ---------------------------
def save_to_csv(data):

    rows = []

    for entry in data["entries"]:
        rows.append({
            "Employee Name": data["employee_name"],
            "Week Start Date": data["week_start_date"],
            "Week End Date": data["week_end_date"],
            "Project Date": entry["project_date"],
            "Project Name": entry["project_name"],
            "Hours": entry["hours"]
        })

    df = pd.DataFrame(rows)

    df.to_csv("structured_timesheet.csv", index=False)
    df.to_excel("structured_timesheet.xlsx", index=False)

    print("Timesheet saved successfully.")


# ---------------------------
# MAIN PIPELINE
# ---------------------------
if __name__ == "__main__":

    print("Fetching email...")
    email_body = fetch_latest_timesheet_email()

    print("Sending to Gemini...")
    structured_json_text = extract_timesheet_structured(email_body)

    print("Parsing JSON...")
    structured_data = parse_json_safely(structured_json_text)

    print("Saving to CSV/Excel...")
    save_to_csv(structured_data)

    print("Pipeline complete.")
