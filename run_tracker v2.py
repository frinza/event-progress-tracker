import datetime
import os.path
import re
import imaplib
import email
from email.header import decode_header
import csv
import getpass

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- CONFIGURATION ---
# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
# Regex to find Branch IDs. More flexible to find IDs like 'B123', 'B 071', or 'B-456'.
BRANCH_ID_REGEX = r'B\s*[-:]?\s*\d{3,4}'
# List of allowed sender emails to check against.
ALLOWED_SENDERS = [
    'prasert.su@spvi.co.th',
    'kankit.pu@spvi.co.th',
    'ruschapoom@spvi.co.th',
    'anuwat.ja@spvi.co.th'
]
# ---------------------

def get_google_calendar_service():
    """Shows basic usage of the Google Calendar API.
    Handles user authentication and returns a service object.
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("FATAL ERROR: 'credentials.json' not found.")
                print("Please follow the setup instructions in README.md to download it.")
                return None
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    try:
        service = build('calendar', 'v3', credentials=creds)
        return service
    except HttpError as error:
        print(f'An error occurred building the service: {error}')
        return None

def extract_and_normalize_branch_ids(text):
    """Uses regex to find all unique, normalized Branch IDs in a given string."""
    if not text:
        return []
    # Find all matches for the flexible regex
    found_ids = re.findall(BRANCH_ID_REGEX, text, re.IGNORECASE)
    # Normalize each found ID by removing extra characters and making it uppercase
    # Use a set to handle duplicates automatically
    # FIX: Escaped the hyphen in the character set to avoid a "bad character range" error.
    normalized_ids = {re.sub(r'[\s\-:]', '', an_id).upper() for an_id in found_ids}
    return list(normalized_ids)

def decode_subject(header):
    """Decodes an email subject header to a readable string."""
    if header is None:
        return ""
    decoded_parts = decode_header(header)
    subject = ""
    for part, charset in decoded_parts:
        if isinstance(part, bytes):
            subject += part.decode(charset or 'utf-8', errors='ignore')
        else:
            subject += part
    return subject

def search_and_verify_imap_subject(imap, branch_id, event_date_str):
    """
    Searches IMAP server for emails matching a branch ID, sent on or after
    a specific event date, from a list of allowed senders. Verifies using regex.
    """
    # --- NEW: Parse the event date string and format it for IMAP search ---
    try:
        # Handle both 'YYYY-MM-DD' and 'YYYY-MM-DDTHH:MM:SS...' formats
        if 'T' in event_date_str:
            date_part = event_date_str.split('T')[0]
        else:
            date_part = event_date_str
        event_date_obj = datetime.datetime.strptime(date_part, '%Y-%m-%d').date()
        # IMAP's SINCE command uses "DD-Mon-YYYY" format
        imap_date_query = event_date_obj.strftime("%d-%b-%Y")
    except (ValueError, TypeError) as e:
        print(f"  - Could not parse date '{event_date_str}'. Skipping date filter. Error: {e}")
        return False # Cannot perform search without a valid date

    # Step 1: Build the search query
    # Build the sender part: (OR (FROM "a@b.com") (FROM "c@d.com"))
    num_senders = len(ALLOWED_SENDERS)
    if num_senders == 0:
        sender_query_part = ""
    elif num_senders == 1:
        sender_query_part = f'(FROM "{ALLOWED_SENDERS[0]}")'
    else:
        sender_query_part = f'(OR (FROM "{ALLOWED_SENDERS[0]}") (FROM "{ALLOWED_SENDERS[1]}"))'
        for i in range(2, num_senders):
            sender_query_part = f'(OR {sender_query_part} (FROM "{ALLOWED_SENDERS[i]}"))'

    # --- MODIFIED: Add the SINCE criterion to the search string ---
    # Combine sender, date, and subject criteria. IMAP's AND is implicit.
    search_criteria = f'(SINCE "{imap_date_query}") (SUBJECT "{branch_id}") {sender_query_part}'

    try:
        status, messages = imap.search(None, search_criteria)
        if status != "OK" or not messages[0]:
            return False # No potential matches found
    except Exception as e:
        print(f"  - Could not perform IMAP search for {branch_id}: {e}")
        return False

    # Step 2: Verify each potential match by fetching and scanning the subject
    message_numbers = messages[0].split()
    for num in message_numbers:
        try:
            # Fetch the full email content for reliable parsing
            status, data = imap.fetch(num, '(RFC822)')
            if status == 'OK':
                msg = email.message_from_bytes(data[0][1])
                email_subject = decode_subject(msg['subject'])
                
                # Use the same normalization to verify the ID in the subject
                ids_in_subject = extract_and_normalize_branch_ids(email_subject)
                if branch_id in ids_in_subject: # Direct comparison since both are normalized
                    print(f"    -> Verified ID {branch_id} in email subject: '{email_subject[:60]}...'")
                    return True # Verified match found!
        except Exception as e:
            print(f"  - Error fetching or parsing email number {num.decode()}: {e}")
            continue # Move to the next email

    return False # No verified match found after checking all potential emails

def main():
    print("--- Google Calendar and IMAP Tracker ---")
    
    # --- Get Google Calendar Events ---
    print("\n[1/4] Authenticating with Google Calendar...")
    service = get_google_calendar_service()
    if not service:
        return
        
    print("Authentication successful.")
    
    # --- Get Calendar ID from user ---
    calendar_id = input("Enter the Google Calendar ID (e.g., yourname@gmail.com or a long ...@group.calendar.google.com ID): ")
    if not calendar_id:
        print("No Calendar ID entered. Exiting.")
        return

    print(f"\n[2/4] Fetching calendar events for the current quarter from '{calendar_id}'...")
    
    today = datetime.date.today()
    quarter_start_month = (today.month - 1) // 3 * 3 + 1
    start_of_quarter = datetime.datetime(today.year, quarter_start_month, 1)

    end_year = today.year
    end_month = quarter_start_month + 3
    if end_month > 12:
        end_month = 1
        end_year += 1
    end_of_quarter = datetime.datetime(end_year, end_month, 1)

    time_min = start_of_quarter.isoformat() + 'Z'
    time_max = end_of_quarter.isoformat() + 'Z'
    
    try:
        events_result = service.events().list(calendarId=calendar_id, timeMin=time_min, timeMax=time_max,
                                              maxResults=250, singleEvents=True,
                                              orderBy='startTime').execute()
    except HttpError as error:
        print(f"An error occurred fetching calendar events: {error}")
        print("Please check if the Calendar ID is correct and that you have access to it.")
        return

    events = events_result.get('items', [])

    if not events:
        print("No upcoming events found for the current quarter in the specified calendar.")
        return

    # --- Extract Branch IDs from Calendar ---
    print(f"Found {len(events)} events. Extracting and normalizing Branch IDs...")
    event_data = []
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        summary = event.get('summary', 'No Title')
        description = event.get('description', '')
        
        search_text = f"{summary} {description}"
        ids_found = extract_and_normalize_branch_ids(search_text)
        
        if ids_found:
            for branch_id in ids_found:
                event_data.append({
                    'Event Title': summary,
                    'Event Date': start,
                    'Branch ID': branch_id
                })

    if not event_data:
        print("No Branch IDs found in any calendar events this quarter. Exiting.")
        return
    print(f"Found {len(event_data)} Branch ID references in calendar.")

    # --- Get IMAP Credentials and Connect ---
    print("\n[3/4] Connecting to IMAP server...")
    imap_server = input("Enter your IMAP server (e.g., imap.gmail.com): ")
    imap_user = input("Enter your email address: ")
    imap_password = getpass.getpass("Enter your password: ")
    
    imap = None
    try:
        imap = imaplib.IMAP4_SSL(imap_server)
        imap.login(imap_user, imap_password)
        imap.select("inbox")
        print("IMAP connection successful.")

        # --- Process and Write to CSV ---
        print("\n[4/4] Cross-referencing with email subjects and generating report...")
        print(f"Filtering for emails from: {', '.join(ALLOWED_SENDERS)}")
        output_filename = 'event_email_report.csv'
        with open(output_filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Event Title', 'Event Date', 'Branch ID', 'Email Status']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()

            for item in event_data:
                branch_id = item['Branch ID']
                # --- MODIFIED: Get the event date to pass to the search function ---
                event_date = item['Event Date'] 
                
                # --- MODIFIED: Updated print statement for clarity ---
                event_date_simple = event_date.split('T')[0]
                print(f" - Searching for Branch ID: {branch_id} in emails on or after {event_date_simple}")
                
                # --- MODIFIED: Pass the event date to the search function ---
                email_found = search_and_verify_imap_subject(imap, branch_id, event_date)
                status = "Found" if email_found else "Waiting"
                
                writer.writerow({
                    'Event Title': item['Event Title'],
                    'Event Date': item['Event Date'],
                    'Branch ID': branch_id,
                    'Email Status': status
                })
        
        print(f"\nProcessing complete. Report saved to '{output_filename}'")

    except imaplib.IMAP4.error as e:
        print(f"\nIMAP Error: Could not log in. Please check credentials.")
        print(f"Server responded: {e}")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        if imap:
            try:
                imap.close()
                imap.logout()
                print("Logged out from IMAP server.")
            except:
                pass

if __name__ == '__main__':
    main()
