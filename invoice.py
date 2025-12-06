"""
Invoice Automation Script for Bi-Weekly Invoicing from Weekly Schedules
Handles email parsing, data storage, automatic invoice numbering, and bi-weekly invoice generation
"""

import sys
import os
from dotenv import load_dotenv
from src.colours import GREEN, RED, RESET

# Load environment variables from .env file
load_dotenv()

print(f"{GREEN}Gmail user loaded: {os.getenv('GMAIL_USER')}{RESET}")
print(f"{GREEN}Recipient loaded: {os.getenv('RECIPIENT_EMAIL')}{RESET}")

# Check if we're in a virtual environment
def check_environment():
  if not hasattr(sys, 'real_prefix') and not (hasattr(sys, 'base_prefix') and sys.base_prefix != sys.prefix):
    print(f"{RED}Warning: You don't appear to be in a virtual environment.{RESET}")
    print("   Consider running: python -m venv venv && source venv/bin/activate (Linux/Mac) or venv\\Scripts\\activate (Windows)")
    response = input("Continue anyway? (y/n): ")
    if response.lower() != 'y':
      sys.exit(1)

# Check for required files
def check_required_files():
  required_files = ['credentials.json']
  missing_files = [f for f in required_files if not os.path.exists(f)]
  
  if missing_files:
    print(f"{RED}Missing required files:{RESET}")
    for file in missing_files:
      print(f"   - {file}")
    print("\nPlease follow the setup instructions to add these files.")
    return False
  return True

# Import necessary libraries
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import re
from datetime import datetime, timedelta
import json

try:
  from googleapiclient.discovery import build
  from google.auth.transport.requests import Request
  from google.oauth2.credentials import Credentials
  from google_auth_oauthlib.flow import InstalledAppFlow
  import requests
  print(f"{GREEN}All Google API libraries imported successfully{RESET}")
except ImportError as e:
  print(f"{RED}Missing required library: {e}{RESET}")
  print("Run: pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib requests")
  sys.exit(1)

# Custom JSON encoder to handle datetime objects
class DateTimeEncoder(json.JSONEncoder):
  def default(self, obj):
    if isinstance(obj, datetime):
      return obj.isoformat()
    return super().default(obj)

# Class to handle all invoice automation tasks
class BiWeeklyInvoiceAutomator:
  def __init__(self):
    # Gmail IMAP settings
    self.GMAIL_IMAP = "imap.gmail.com"
    self.GMAIL_SMTP = "smtp.gmail.com"
    
    # Google Sheets API scopes
    self.SCOPES = [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive'
    ]
    
    # Your personal info - load from environment
    hourly_rate = os.getenv('HOURLY_RATE')
    if not hourly_rate:
      print("HOURLY_RATE not set in .env file")
      sys.exit(1)
    
    self.HOURLY_RATE = float(hourly_rate)
    self.HST_RATE = 0.13
    self.DEFAULT_HOURS_PER_SHIFT = 7.0  # Fallback for calculation errors
    self.YOUR_NAME = os.getenv('YOUR_NAME', 'Your Name')
    self.RECIPIENT_NAME = os.getenv('RECIPIENT_NAME', 'Recipient')
    
    # Pay period pattern (starts on Saturdays, bi-weekly)
    # self.FIRST_PAY_PERIOD_START = datetime(2025, 6, 21)  # Originally set to: Saturday June 21
    # self.FIRST_PAY_PERIOD_START = datetime(2025, 5, 24)  # Saturday May 24
    pay_period_start = os.getenv('FIRST_PAY_PERIOD_START')
    if not pay_period_start:
      print("FIRST_PAY_PERIOD_START not set in .env file")
      sys.exit(1)

    try:
      self.FIRST_PAY_PERIOD_START = datetime.strptime(pay_period_start, '%Y-%m-%d')
    except ValueError:
      print("FIRST_PAY_PERIOD_START must be in YYYY-MM-DD format (e.g., 2025-05-24)")
      sys.exit(1)
    self.PAY_PERIOD_LENGTH_DAYS = 14  # Bi-weekly
    
    # Spreadsheet layout constants
    self.INVOICE_START_ROW = 13
    self.INVOICE_NUMBER_CELL = 'G3'
    self.INVOICE_DATE_CELL = 'G4'
    self.DATE_COLUMN = 'B'
    self.START_TIME_COLUMN = 'C'
    self.END_TIME_COLUMN = 'D'
    self.HOURS_COLUMN = 'E'
    self.RATE_COLUMN = 'F'
    self.AMOUNT_COLUMN = 'G'
    self.TOTAL_HOURS_CELL = 'E27'
    self.SUBTOTAL_CELL = 'G27'
    self.HST_CELL = 'G28'
    self.TOTAL_CELL = 'G30'
    
    # Data storage files
    self.SHIFTS_DATA_FILE = "shifts_data.json"
    self.INVOICE_COUNTER_FILE = "invoice_counter.json"
  
  def format_period_dates(self, period_info):
    """Format period start and end dates consistently"""
    start_date = period_info.get('start', 'Unknown')
    end_date = period_info.get('end', 'Unknown')
    
    # Handle unknown dates
    if start_date == 'Unknown' or end_date == 'Unknown':
      return str(start_date), str(end_date)
    
    # Convert strings to datetime if needed
    if isinstance(start_date, str):
      try:
        start_date = datetime.fromisoformat(start_date.replace('T00:00:00', ''))
        end_date = datetime.fromisoformat(end_date.replace('T00:00:00', ''))
      except (ValueError, AttributeError):
        return str(start_date), str(end_date)
    
    # Format dates
    start_str = f"{start_date.strftime('%B')} {start_date.day}, {start_date.year}"
    end_str = f"{end_date.strftime('%B')} {end_date.day}, {end_date.year}"
    
    return start_str, end_str
  
  def convert_period_dates_to_datetime(self, period_info):
    """Convert period_info dates from strings to datetime objects if needed"""
    if isinstance(period_info['start'], str):
      period_info['start'] = datetime.fromisoformat(period_info['start'])
    if isinstance(period_info['end'], str):
      period_info['end'] = datetime.fromisoformat(period_info['end'])
    return period_info
  
  def select_period_from_list(self, periods_dict, prompt_message="Select period"):
    """Let user select a period from a dictionary of periods"""
    if len(periods_dict) == 1:
      # Only one period, return it directly
      period_key = list(periods_dict.keys())[0]
      return period_key, periods_dict[period_key]
    
    # Multiple periods - let user choose
    print(f"\n{prompt_message}:")
    period_list = list(periods_dict.items())
    
    for i, (key, data) in enumerate(period_list, 1):
      period_info = data['period_info']
      start_str, end_str = self.format_period_dates(period_info)
      
      # Simplify if same year
      current_year = str(datetime.now().year)
      if start_str.endswith(current_year) and end_str.endswith(current_year):
        start_parts = start_str.rsplit(', ', 1)
        end_parts = end_str.rsplit(', ', 1)
        date_range = f"{start_parts[0]} to {end_parts[0]}, {end_parts[1]}"
      else:
        date_range = f"{start_str} to {end_str}"
      
      print(f"{i}. {date_range}")
    
    try:
      selection = int(input(f"\nSelect period (1-{len(period_list)}): ")) - 1
      if 0 <= selection < len(period_list):
        return period_list[selection]
      else:
        print("Invalid selection")
        return None, None
    except ValueError:
      print("Invalid input")
      return None, None
    
  def load_shifts_data(self):
    """Load stored shifts data from file"""
    if os.path.exists(self.SHIFTS_DATA_FILE):
      try:
        with open(self.SHIFTS_DATA_FILE, 'r') as f:
          return json.load(f)
      except (json.JSONDecodeError, IOError) as e:
        print(f"Warning: Could not load shifts data: {e}")
        return {}
    return {}
  
  def save_shifts_data(self, data):
    """Save shifts data to file"""
    try:
      with open(self.SHIFTS_DATA_FILE, 'w') as f:
        json.dump(data, f, indent=2, cls=DateTimeEncoder)
      return True
    except Exception as e:
      print(f"Error saving shifts data: {e}")
      return False
  
  def load_invoice_counter(self):
    """Load the current invoice counter"""
    if os.path.exists(self.INVOICE_COUNTER_FILE):
      try:
        with open(self.INVOICE_COUNTER_FILE, 'r') as f:
          data = json.load(f)
          return data.get('next_invoice_number', 3)
      except (json.JSONDecodeError, IOError) as e:
        print(f"Warning: Could not load invoice counter: {e}")
        return 3
    return 3
  
  def save_invoice_counter(self, invoice_number):
    """Save the next invoice counter"""
    try:
      data = {'next_invoice_number': invoice_number + 1}
      with open(self.INVOICE_COUNTER_FILE, 'w') as f:
        json.dump(data, f, indent=2)
      return True
    except Exception as e:
      print(f"Error saving invoice counter: {e}")
      return False
  
  def get_pay_period_for_date(self, date_obj):
    """Get the pay period that contains a given date"""
    # Calculate days since first pay period start
    days_since_first = (date_obj - self.FIRST_PAY_PERIOD_START).days
    
    if days_since_first < 0:
      return None  # Date is before our tracking started
    
    # Calculate which pay period this date belongs to
    period_number = days_since_first // self.PAY_PERIOD_LENGTH_DAYS
    
    # Calculate the actual start and end dates for this period
    period_start = self.FIRST_PAY_PERIOD_START + timedelta(days=period_number * self.PAY_PERIOD_LENGTH_DAYS)
    period_end = period_start + timedelta(days=self.PAY_PERIOD_LENGTH_DAYS - 1)
    
    return {
      'start': period_start,
      'end': period_end,
      'period_number': period_number + 1,  # 1-based for display
      'period_key': f"{period_start.strftime('%Y_%m_%d')}"  # Unique identifier
    }
  
  def authenticate_gmail(self, email_address, password):
    """Connect to Gmail via IMAP"""
    try:
      mail = imaplib.IMAP4_SSL(self.GMAIL_IMAP)
      mail.login(email_address, password)
      return mail
    except Exception as e:
      print(f"Gmail authentication failed: {e}")
      return None
  
  def authenticate_google_sheets(self):
    """Authenticate with Google Sheets API"""
    creds = None
    if os.path.exists('token.json'):
      creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
    
    if not creds or not creds.valid:
      if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
      else:
        flow = InstalledAppFlow.from_client_secrets_file(
          'credentials.json', self.SCOPES)
        creds = flow.run_local_server(port=0)
      with open('token.json', 'w') as token:
        token.write(creds.to_json())
    
    return build('sheets', 'v4', credentials=creds), build('drive', 'v3', credentials=creds)
  
  # def parse_homebase_email(self, email_content):
  #   """Parse Homebase schedule email and extract shift information, excluding optional open shifts"""
  #   shifts = []
    
  #   # Find the date range (e.g., "06/30 - 07/06")
  #   date_range_match = re.search(r'(\d{2}/\d{2})\s*-\s*(\d{2}/\d{2})', email_content)
  #   if not date_range_match:
  #     print("Could not find date range in email")
  #     return []
    
  #   # Split email into sections to avoid parsing open shifts
  #   lines = email_content.split('\n')
    
  #   # Find where the "Want some extra hours?" or similar section starts
  #   # This marks the beginning of optional open shifts
  #   stop_parsing_at = len(lines)
  #   for i, line in enumerate(lines):
  #     if any(phrase in line.lower() for phrase in [
  #         'want some extra hours',
  #         'available open shift',
  #         'here is an available',
  #         'sign in to trade shifts'
  #     ]):
  #       stop_parsing_at = i
  #       print(f"Stopping shift parsing at line {i}: Found optional shifts section")
  #       break
    
  #   # Find all day/time combinations only up to the optional shifts section
  #   day_pattern = r'(\w+day)\s+(\d{2}/\d{2})'
  #   time_pattern = r'(\d{1,2}:\d{2})(am|pm)\s*-\s*(\d{1,2}:\d{2})(am|pm)'
    
  #   current_date = None
    
  #   for i in range(stop_parsing_at):  # Only parse up to optional shifts section
  #     line = lines[i]
  #     day_match = re.search(day_pattern, line)
  #     if day_match:
  #       current_date = day_match.group(2)  # e.g., "07/21"
        
  #       # Check next few lines for time pattern
  #       for j in range(i+1, min(i+4, stop_parsing_at)):  # Also respect stop limit here
  #         time_match = re.search(time_pattern, lines[j])
  #         if time_match:
  #           start_time = f"{time_match.group(1)}{time_match.group(2)}"
  #           end_time = f"{time_match.group(3)}{time_match.group(4)}"
            
  #           shifts.append({
  #             'date': current_date,
  #             'start_time': start_time,
  #             'end_time': end_time
  #           })
  #           break
    
  #   # Remove duplicates based on date (in case there are any)
  #   unique_shifts = []
  #   seen_dates = set()
  #   for shift in shifts:
  #     if shift['date'] not in seen_dates:
  #       unique_shifts.append(shift)
  #       seen_dates.add(shift['date'])
  #     else:
  #       print(f"Removing duplicate shift for date {shift['date']}")
    
  #   print(f"Parsed {len(unique_shifts)} unique regular shifts (excluded optional open shifts)")
  #   return unique_shifts
  def parse_homebase_email(self, email_content):
    """Parse Homebase schedule email and extract shift information"""
    shifts = []
    
    # Try new HTML format first
    pattern = r'<b>(\w{3})\s+(\w{3})\s+(\d+)</b>.*?(\d{1,2}:\d{2}[AP]M)-(\d{1,2}:\d{2}[AP]M)'
    matches = re.finditer(pattern, email_content, re.DOTALL)
    
    for match in matches:
        month_abbr = match.group(2)
        day_num = int(match.group(3))
        start_time = match.group(4)
        end_time = match.group(5)
        
        month_map = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                     'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
        month_num = month_map.get(month_abbr)
        
        if month_num:
            date_str = f"{month_num:02d}/{day_num:02d}"
            shifts.append({
                'date': date_str,
                'start_time': start_time.lower(),
                'end_time': end_time.lower()
            })
    
    if shifts:
        print(f"Parsed {len(shifts)} shifts from new format")
        return shifts
    
    # Fall back to old format
    date_range_match = re.search(r'(\d{2}/\d{2})\s*-\s*(\d{2}/\d{2})', email_content)
    if not date_range_match:
        print("Could not find date range in email")
        return []
    
    lines = email_content.split('\n')
    stop_parsing_at = len(lines)
    for i, line in enumerate(lines):
        if any(phrase in line.lower() for phrase in [
            'want some extra hours', 'available open shift',
            'here is an available', 'sign in to trade shifts']):
            stop_parsing_at = i
            break
    
    day_pattern = r'(\w+day)\s+(\d{2}/\d{2})'
    time_pattern = r'(\d{1,2}:\d{2})(am|pm)\s*-\s*(\d{1,2}:\d{2})(am|pm)'
    current_date = None
    
    for i in range(stop_parsing_at):
        line = lines[i]
        day_match = re.search(day_pattern, line)
        if day_match:
            current_date = day_match.group(2)
            for j in range(i+1, min(i+4, stop_parsing_at)):
                time_match = re.search(time_pattern, lines[j])
                if time_match:
                    start_time = f"{time_match.group(1)}{time_match.group(2)}"
                    end_time = f"{time_match.group(3)}{time_match.group(4)}"
                    shifts.append({'date': current_date, 'start_time': start_time, 'end_time': end_time})
                    break
    
    unique_shifts = []
    seen_dates = set()
    for shift in shifts:
        if shift['date'] not in seen_dates:
            unique_shifts.append(shift)
            seen_dates.add(shift['date'])
    
    print(f"Parsed {len(unique_shifts)} unique shifts from old format")
    return unique_shifts
  
  def normalize_time_display(self, time_str):
    """Remove leading zeros from time for cleaner display (08:00am -> 8:00am)"""
    if time_str and time_str[0] == '0':
        return time_str[1:]
    return time_str
  
  def convert_time_to_24h_simple(self, time_str):
    """Convert 12-hour format to simple 24-hour format (e.g., '8:00' not '08:00')"""
    try:
      time_obj = datetime.strptime(time_str, "%I:%M%p")
      hour = time_obj.hour
      minute = time_obj.minute
      return f"{hour}:{minute:02d}"
    except (ValueError, AttributeError):
      return ""
  
  def calculate_shift_hours(self, start_time, end_time):
    """Calculate hours worked from start and end times"""
    try:
      # Parse the times
      start_obj = datetime.strptime(start_time, "%I:%M%p")
      end_obj = datetime.strptime(end_time, "%I:%M%p")
      
      # Handle overnight shifts (though probably not applicable for your case)
      if end_obj < start_obj:
        end_obj += timedelta(days=1)
      
      # Calculate the difference
      time_diff = end_obj - start_obj
      hours = time_diff.total_seconds() / 3600
      
      return round(hours, 2)  # Round to 2 decimal places
    except Exception as e:
      print(f"Error calculating hours for {start_time} to {end_time}: {e}")
      return self.DEFAULT_HOURS_PER_SHIFT  # Fallback to default

  def sort_shifts_by_date(self, shifts):
    """Sort shifts list chronologically by date"""
    try:
      current_year = datetime.now().year
      return sorted(shifts, key=lambda s: datetime.strptime(f"{s['date']}/{current_year}", "%m/%d/%Y"))
    except (ValueError, KeyError) as e:
      print(f"Warning: Could not sort shifts: {e}")
      return shifts
  
  def store_weekly_shifts(self, shifts):
    """Store shifts from weekly email and return pay period info if ready to invoice"""
    if not shifts:
      return None
        
    # Load existing data
    all_shifts_data = self.load_shifts_data()
    
    # Determine which pay period these shifts belong to
    first_shift_date = shifts[0]['date']
    try:
      current_year = datetime.now().year
      shift_datetime = datetime.strptime(f"{first_shift_date}/{current_year}", "%m/%d/%Y")
      pay_period = self.get_pay_period_for_date(shift_datetime)
      
      if not pay_period:
        print(f"Could not determine pay period for date {first_shift_date}")
        return None
        
      period_key = pay_period['period_key']
      
      # Initialize period data if it doesn't exist
      if period_key not in all_shifts_data:
        all_shifts_data[period_key] = {
          'period_info': pay_period,
          'shifts': [],
          'weeks_received': 0,
          'complete': False,
          'invoice_generated': False
        }
      
      # Check if we already generated an invoice for this period
      if all_shifts_data[period_key].get('invoice_generated', False):
        print(f"Invoice already generated for pay period {period_key} - skipping")
        return None
      
      # Add new shifts (avoid duplicates based on date)
      existing_dates = {s['date'] for s in all_shifts_data[period_key]['shifts']}
      new_shifts = [s for s in shifts if s['date'] not in existing_dates]
      
      if new_shifts:
        all_shifts_data[period_key]['shifts'].extend(new_shifts)
        all_shifts_data[period_key]['weeks_received'] += 1
        
        print(f"Stored {len(new_shifts)} new shifts for pay period {period_key}")
        print(f"Weeks received: {all_shifts_data[period_key]['weeks_received']}/2")
        
        # Check if we have both weeks (bi-weekly period complete)
        if all_shifts_data[period_key]['weeks_received'] >= 2:
          all_shifts_data[period_key]['complete'] = True
          
          # Save updated data
          self.save_shifts_data(all_shifts_data)
          
          # Return period info for invoice generation
          return {
            'period': pay_period,
            'shifts': all_shifts_data[period_key]['shifts'],
            'period_key': period_key
          }
        else:
          # Save updated data but don't generate invoice yet
          self.save_shifts_data(all_shifts_data)
      else:
        print(f"No new shifts found for pay period {period_key}")
        
        # Still check if period is complete and hasn't been invoiced
        if (all_shifts_data[period_key]['weeks_received'] >= 2 and 
          all_shifts_data[period_key]['complete'] and 
          not all_shifts_data[period_key].get('invoice_generated', False)):
          
          print(f"Pay period {period_key} is complete but not yet invoiced")
          return {
            'period': pay_period,
            'shifts': all_shifts_data[period_key]['shifts'],
            'period_key': period_key
          }
      
    except Exception as e:
      print(f"Error storing shifts: {e}")
      
    return None

  def mark_invoice_generated(self, period_key):
    """Mark that an invoice has been generated for this period"""
    all_shifts_data = self.load_shifts_data()
    
    if period_key in all_shifts_data:
      all_shifts_data[period_key]['invoice_generated'] = True
      all_shifts_data[period_key]['hourly_rate'] = self.HOURLY_RATE
      self.save_shifts_data(all_shifts_data)
      print(f"Marked invoice as generated for period {period_key}")
      return True
    
    return False
  
  def create_invoice_data(self, shifts, pay_period):
    """Create complete invoice data for the pay period with calculated hours"""
    invoice_rows = []
    current_date = pay_period['start']
    end_date = pay_period['end']
    
    # Create a lookup dictionary for shifts
    shift_lookup = {}
    for shift in shifts:
      try:
        current_year = datetime.now().year
        shift_date = datetime.strptime(f"{shift['date']}/{current_year}", "%m/%d/%Y")
        shift_lookup[shift_date.strftime("%Y-%m-%d")] = shift
      except (ValueError, KeyError):
        continue
    
    # Generate all days in the pay period
    while current_date <= end_date:
      date_key = current_date.strftime("%Y-%m-%d")
      formatted_date = current_date.strftime("%d-%b")
      
      if date_key in shift_lookup:
        # Work day
        shift = shift_lookup[date_key]
        start_24h = self.convert_time_to_24h_simple(shift['start_time'])
        end_24h = self.convert_time_to_24h_simple(shift['end_time'])
        
        # Calculate hours - use manual override if present, otherwise calculate
        if 'hours' in shift and shift['hours'] is not None:
          calculated_hours = float(shift['hours'])
          print(f"Using manual hours for {shift['date']}: {calculated_hours}")
        else:
          calculated_hours = self.calculate_shift_hours(shift['start_time'], shift['end_time'])
          print(f"Calculated hours for {shift['date']}: {calculated_hours}")
        
        invoice_rows.append({
          'date': formatted_date,
          'start': start_24h,
          'end': end_24h,
          'hours': calculated_hours,
          'rate': self.HOURLY_RATE,
          'amount': calculated_hours * self.HOURLY_RATE
        })
      else:
        # Non-work day
        invoice_rows.append({
          'date': formatted_date,
          'start': '',
          'end': '',
          'hours': '',
          'rate': self.HOURLY_RATE,
          'amount': 0.00
        })
      
      current_date += timedelta(days=1)
    
    return invoice_rows
  
  def create_new_invoice_sheet(self, template_spreadsheet_id, pay_period):
    """Create a new invoice sheet for the pay period"""
    try:
      sheets_service, drive_service = self.authenticate_google_sheets()
      
      # Create new invoice filename
      period_end = pay_period['end'].strftime('%Y_%m_%d')
      new_title = f"{self.YOUR_NAME} Invoice {period_end}"

      # Copy the template spreadsheet
      copy_body = {'name': new_title}
      new_sheet = drive_service.files().copy(
        fileId=template_spreadsheet_id, 
        body=copy_body
      ).execute()
      
      new_spreadsheet_id = new_sheet['id']
      print(f"Created new invoice sheet: {new_title}")
      print(f"New sheet ID: {new_spreadsheet_id}")
      
      return new_spreadsheet_id
      
    except Exception as e:
      print(f"Error creating new sheet: {e}")
      return None
  
  def update_google_sheet(self, spreadsheet_id, invoice_data, pay_period):
    """Update Google Sheet with invoice data"""
    # Get the next invoice number
    invoice_number = self.load_invoice_counter()
    
    try:
      sheets_service, drive_service = self.authenticate_google_sheets()
      
      # Invoice date is the end date of the pay period
      invoice_date = pay_period['end'].strftime('%m/%d/%Y')
      
      # Prepare the data for batch update
      updates = []
      
      # Update invoice number and date
      updates.extend([
        {'range': self.INVOICE_NUMBER_CELL, 'values': [[f'INVOICE #{invoice_number:04d}']]},
        {'range': self.INVOICE_DATE_CELL, 'values': [[f'DATE: {invoice_date}']]}
      ])
      
      # Update invoice rows
      for i, row_data in enumerate(invoice_data):
        row_num = self.INVOICE_START_ROW + i
        
        updates.extend([
          {'range': f'{self.DATE_COLUMN}{row_num}', 'values': [[row_data['date']]]},
          {'range': f'{self.START_TIME_COLUMN}{row_num}', 'values': [[row_data['start']]]},
          {'range': f'{self.END_TIME_COLUMN}{row_num}', 'values': [[row_data['end']]]},
          {'range': f'{self.HOURS_COLUMN}{row_num}', 'values': [[row_data['hours']]]},
          {'range': f'{self.RATE_COLUMN}{row_num}', 'values': [[row_data['rate']]]},
          {'range': f'{self.AMOUNT_COLUMN}{row_num}', 'values': [[row_data['amount']]]}
        ])
      
      # Calculate totals
      total_hours = sum(row['hours'] for row in invoice_data if row['hours'])
      subtotal = sum(row['amount'] for row in invoice_data)
      hst_amount = subtotal * self.HST_RATE
      total_amount = subtotal + hst_amount
      
      # Update totals
      updates.extend([
        {'range': self.TOTAL_HOURS_CELL, 'values': [[total_hours]]},
        {'range': self.SUBTOTAL_CELL, 'values': [[subtotal]]},
        {'range': self.HST_CELL, 'values': [[hst_amount]]},
        {'range': self.TOTAL_CELL, 'values': [[total_amount]]}
      ])
      
      # Batch update
      body = {'valueInputOption': 'USER_ENTERED', 'data': updates}
      result = sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body).execute()
      
      # Save the updated invoice counter
      self.save_invoice_counter(invoice_number)
      
      print(f"Updated {len(updates)} ranges in spreadsheet")
      print(f"Invoice #{invoice_number:04d} created with date {invoice_date}")
      return True
      
    except Exception as e:
      print(f"Error updating Google Sheet: {e}")
      return False
      
  def list_pay_periods(self):
    """List all stored pay periods and their status"""
    all_shifts_data = self.load_shifts_data()
    
    if not all_shifts_data:
      print("No pay periods found in shifts data")
      return
    
    print("Stored Pay Periods:")
    print("-" * 50)
    
    for period_key, data in sorted(all_shifts_data.items()):
      period_info = data.get('period_info', {})
      weeks_received = data.get('weeks_received', 0)
      complete = data.get('complete', False)
      invoice_generated = data.get('invoice_generated', False)
      shift_count = len(data.get('shifts', []))
      
      # Use helper to format dates
      start_str, end_str = self.format_period_dates(period_info)
      
      if invoice_generated:
        status = "Invoice Generated"
      elif complete:
        status = "Complete but No Invoice"
      else:
        status = f"Incomplete ({weeks_received}/2 weeks)"
      
      print(f"Period: {period_key}")
      print(f"  Dates: {start_str} to {end_str}")
      print(f"  Status: {status}")
      print(f"  Shifts: {shift_count}")
      
      # Show individual shifts with hours info first
      total_hours = 0
      for shift in data.get('shifts', []):
        if 'hours' in shift and shift['hours'] is not None:
          hours = float(shift['hours'])
          hours_info = f" ({hours}h manual)"
          total_hours += hours
        else:
          calculated = self.calculate_shift_hours(shift['start_time'], shift['end_time'])
          hours_info = f" ({calculated}h calculated)"
          total_hours += calculated
        
        print(f"    {shift['date']}: {shift['start_time']}-{shift['end_time']}{hours_info}")
      
      # Then show financial summary after all shifts
      hourly_rate = data.get('hourly_rate', self.HOURLY_RATE)
      subtotal = total_hours * hourly_rate
      hst_amount = subtotal * self.HST_RATE
      total_amount = subtotal + hst_amount
      
      print(f"  Total Hours: {total_hours:.2f}h")
      print(f"  Subtotal: ${subtotal:.2f}")
      print(f"  HST (13%): ${hst_amount:.2f}")
      print(f"  Total: ${total_amount:.2f}")
      print()

  def view_tax_summary(self):
    """Display tax summary with totals for invoiced periods, grouped by year"""
    all_shifts_data = self.load_shifts_data()
    
    if not all_shifts_data:
      print("No pay periods found in shifts data")
      return
    
    # Group periods by year and status
    invoiced_by_year = {}
    pending_by_year = {}
    
    for period_key, data in all_shifts_data.items():
      period_info = data.get('period_info', {})
      
      # Convert dates if needed
      if isinstance(period_info['end'], str):
        end_date = datetime.fromisoformat(period_info['end'])
      else:
        end_date = period_info['end']
      
      year = end_date.year
      
      # Calculate totals for this period
      total_hours = 0
      for shift in data.get('shifts', []):
        if 'hours' in shift and shift['hours'] is not None:
          hours = float(shift['hours'])
        else:
          hours = self.calculate_shift_hours(shift['start_time'], shift['end_time'])
        total_hours += hours
      
      hourly_rate = data.get('hourly_rate', self.HOURLY_RATE)
      subtotal = total_hours * hourly_rate
      hst = subtotal * self.HST_RATE
      grand_total = subtotal + hst
      
      period_data = {
        'hours': total_hours,
        'subtotal': subtotal,
        'hst': hst,
        'total': grand_total,
        'period_key': period_key
      }
      
      # Add to appropriate year bucket
      if data.get('invoice_generated', False):
        if year not in invoiced_by_year:
          invoiced_by_year[year] = []
        invoiced_by_year[year].append(period_data)
      elif data.get('complete', False):
        if year not in pending_by_year:
          pending_by_year[year] = []
        pending_by_year[year].append(period_data)
    
    # Display summary
    print("\n" + "="*50)
    print("TAX SUMMARY - BY YEAR")
    print("="*50)
    
    # Show invoiced periods by year
    if invoiced_by_year:
      print("\nINVOICED PERIODS (for tax reporting)")
      print("="*50)
      
      for year in sorted(invoiced_by_year.keys()):
        periods = invoiced_by_year[year]
        year_hours = sum(p['hours'] for p in periods)
        year_subtotal = sum(p['subtotal'] for p in periods)
        year_hst = sum(p['hst'] for p in periods)
        year_total = sum(p['total'] for p in periods)
        
        print(f"\n{year}:")
        print("-"*50)
        print(f"  Invoiced periods:      {len(periods)}")
        print(f"  Total hours worked:    {year_hours:.2f}h")
        print(f"  Total income:          ${year_subtotal:.2f}")
        print(f"  Total HST collected:   ${year_hst:.2f}")
        print(f"  Grand total:           ${year_total:.2f}")
      
      # Overall invoiced totals
      all_invoiced_hours = sum(sum(p['hours'] for p in periods) for periods in invoiced_by_year.values())
      all_invoiced_subtotal = sum(sum(p['subtotal'] for p in periods) for periods in invoiced_by_year.values())
      all_invoiced_hst = sum(sum(p['hst'] for p in periods) for periods in invoiced_by_year.values())
      all_invoiced_total = sum(sum(p['total'] for p in periods) for periods in invoiced_by_year.values())
      
      if len(invoiced_by_year) > 1:
        print(f"\n  ALL YEARS COMBINED:")
        print("-"*50)
        print(f"  Total periods:         {sum(len(p) for p in invoiced_by_year.values())}")
        print(f"  Total hours:           {all_invoiced_hours:.2f}h")
        print(f"  Total income:          ${all_invoiced_subtotal:.2f}")
        print(f"  Total HST collected:   ${all_invoiced_hst:.2f}")
        print(f"  Grand total:           ${all_invoiced_total:.2f}")
    
    # Show pending periods by year
    if pending_by_year:
      print("\n\nPENDING PERIODS (not yet invoiced)")
      print("="*50)
      
      for year in sorted(pending_by_year.keys()):
        periods = pending_by_year[year]
        year_hours = sum(p['hours'] for p in periods)
        year_subtotal = sum(p['subtotal'] for p in periods)
        year_hst = sum(p['hst'] for p in periods)
        year_total = sum(p['total'] for p in periods)
        
        print(f"\n{year}:")
        print("-"*50)
        print(f"  Pending periods:       {len(periods)}")
        print(f"  Total hours:           {year_hours:.2f}h")
        print(f"  Estimated income:      ${year_subtotal:.2f}")
        print(f"  Estimated HST:         ${year_hst:.2f}")
        print(f"  Estimated total:       ${year_total:.2f}")
    
    print("\n" + "="*50)
    
    # Tax filing reminder
    if invoiced_by_year:
      print("\nTax Filing Reminder:")
      for year in sorted(invoiced_by_year.keys()):
        year_hst = sum(p['hst'] for p in invoiced_by_year[year])
        year_subtotal = sum(p['subtotal'] for p in invoiced_by_year[year])
        print(f"   {year}: Income = ${year_subtotal:.2f}, HST = ${year_hst:.2f}")

  def force_complete_period(self):
    """Manually mark or create a period as complete for invoice generation"""
    all_shifts_data = self.load_shifts_data()
    
    # Find incomplete periods
    incomplete_periods = {}
    for period_key, period_data in all_shifts_data.items():
      if (period_data.get('weeks_received', 0) > 0 and 
          not period_data.get('complete', False) and 
          not period_data.get('invoice_generated', False)):
        incomplete_periods[period_key] = period_data
    
    print("\n" + "="*50)
    print("FORCE COMPLETE PERIOD")
    print("="*50)
    
    # Show options
    print("\nOptions:")
    print("1. Mark existing incomplete period as complete")
    print("2. Create and complete a new period (for holidays/closures)")
    print("3. Cancel")
    
    try:
      choice = input("\nSelect option (1-3): ").strip()
      
      if choice == "1":
        # Handle existing incomplete periods
        if not incomplete_periods:
          print("No incomplete periods found")
          return False
        
        period_key, period_data = self.select_period_from_list(
          incomplete_periods, 
          "Select period to mark as complete"
        )
        
        if period_key is None:
          return False
        
        # Show what will be marked complete
        period_info = period_data['period_info']
        start_str, end_str = self.format_period_dates(period_info)
        
        print(f"\nPeriod: {start_str} to {end_str}")
        print(f"Weeks received: {period_data.get('weeks_received', 0)}/2")
        print(f"Shifts: {len(period_data.get('shifts', []))}")
        
        # Confirm
        if self.get_user_confirmation("\nMark this period as complete?", default="y"):
          all_shifts_data[period_key]['complete'] = True
          all_shifts_data[period_key]['weeks_received'] = 2
          
          if self.save_shifts_data(all_shifts_data):
            print("Period marked as complete!")
            print("You can now generate an invoice using option 1.")
            return True
          else:
            print("Failed to save changes")
            return False
        else:
          print("Cancelled")
          return False
      
      elif choice == "2":
        # Create a new period
        print("\nEnter a date within the period you want to create")
        date_input = input("Date (MM/DD format): ").strip()
        
        if not date_input:
          print("Date is required")
          return False
        
        # Parse date and get pay period
        try:
          current_year = datetime.now().year
          date_obj = datetime.strptime(f"{date_input}/{current_year}", "%m/%d/%Y")
          pay_period = self.get_pay_period_for_date(date_obj)
          
          if not pay_period:
            print(f"Could not determine pay period for date {date_input}")
            return False
          
          period_key = pay_period['period_key']
          
          # Check if period already exists
          if period_key in all_shifts_data:
            print(f"\nPeriod {period_key} already exists!")
            print("Use option 1 instead to mark existing period as complete.")
            return False
          
          # Show period details
          start_str, end_str = self.format_period_dates(pay_period)
          print(f"\nCreating period:")
          print(f"  Dates: {start_str} to {end_str}")
          print(f"  Period key: {period_key}")
          
          # Ask if they want to add shifts
          add_shifts = self.get_user_confirmation("\nDo you want to add any shifts to this period?", default="n")
          
          shifts = []
          if add_shifts:
            print("\nAdd shifts (leave date blank when done):")
            while True:
              shift_date = input("  Shift date (MM/DD): ").strip()
              if not shift_date:
                break
              
              start_time = input("  Start time (e.g., 8:00am): ").strip()
              end_time = input("  End time (e.g., 3:00pm): ").strip()
              hours_input = input("  Manual hours (leave blank to calculate): ").strip()
              
              shift = {
                'date': shift_date,
                'start_time': start_time,
                'end_time': end_time
              }
              
              if hours_input:
                shift['hours'] = float(hours_input)
              
              shifts.append(shift)
              print(f"  Added shift for {shift_date}")
          
          # Sort shifts chronologically
          if shifts:
            shifts = self.sort_shifts_by_date(shifts)
          
          # Confirm creation
          print(f"\nSummary:")
          print(f"  Period: {start_str} to {end_str}")
          print(f"  Shifts: {len(shifts)}")
          
          if self.get_user_confirmation("\nCreate this period as complete?", default="y"):
            # Create the period
            all_shifts_data[period_key] = {
              'period_info': pay_period,
              'shifts': shifts,
              'weeks_received': 2,  # Mark as both weeks received
              'complete': True,     # Mark as complete
              'invoice_generated': False
            }
            
            if self.save_shifts_data(all_shifts_data):
              print("Period created and marked as complete!")
              print("You can now generate an invoice using option 1.")
              return True
            else:
              print("Failed to save changes")
              return False
          else:
            print("Cancelled")
            return False
            
        except ValueError as e:
          print(f"Invalid date format: {e}")
          return False
      
      elif choice == "3":
        print("Cancelled")
        return False
      else:
        print("Invalid choice")
        return False
        
    except KeyboardInterrupt:
      print("\nCancelled")
      return False
    except Exception as e:
      print(f"Error: {e}")
      return False
  
  def export_sheet_as_pdf(self, spreadsheet_id, output_filename):
    """Export Google Sheet as PDF file without gridlines"""
    try:
      sheets_service, drive_service = self.authenticate_google_sheets()
      
      # Get credentials for authenticated request
      creds = None
      if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
      
      if not creds or not creds.valid:
        print("Error: No valid credentials for PDF export")
        return False
      
      # Create export URL for PDF
      export_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?"
        f"format=pdf&"
        f"portrait=true&"
        f"size=letter&"
        f"scale=4&"
        f"top_margin=0.3&"
        f"bottom_margin=0.3&"
        f"left_margin=0.6&"
        f"right_margin=0.6&"
        f"gridlines=false&"
        f"printtitle=false&"
        f"sheetnames=false&"
      )
      
      # Make authenticated request
      headers = {'Authorization': f'Bearer {creds.token}'}
      response = requests.get(export_url, headers=headers)
      
      if response.status_code == 200:
        with open(output_filename, 'wb') as f:
          f.write(response.content)
        print(f"PDF exported successfully: {output_filename}")
        return True
      else:
        print(f"Failed to export PDF. Status code: {response.status_code}")
        return False
        
    except Exception as e:
      print(f"Error exporting PDF: {e}")
      return False

  def export_sheet_as_excel(self, spreadsheet_id, output_filename):
    """Export Google Sheet as Excel file"""
    try:
      sheets_service, drive_service = self.authenticate_google_sheets()
      
      # Get credentials for authenticated request
      creds = None
      if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
      
      if not creds or not creds.valid:
        print("Error: No valid credentials for Excel export")
        return False
      
      # Create export URL for Excel
      export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx"
      
      # Make authenticated request
      headers = {'Authorization': f'Bearer {creds.token}'}
      response = requests.get(export_url, headers=headers)
      
      if response.status_code == 200:
        with open(output_filename, 'wb') as f:
          f.write(response.content)
        print(f"Excel file exported successfully: {output_filename}")
        return True
      else:
        print(f"Failed to export Excel. Status code: {response.status_code}")
        return False
        
    except Exception as e:
      print(f"Error exporting Excel: {e}")
      return False

  def send_invoice_email_with_attachments(self, recipient_email, pdf_path, excel_path, sender_email, sender_password, pay_period, invoice_number=None):
    """Send invoice via email with PDF and Excel attachments"""
    
    def format_date_range(start_date, end_date):
      if start_date.year == end_date.year:
        # Same year
        if start_date.month == end_date.month:
          # Same month: "July 4 to 18, 2025"
          return f"{start_date.strftime('%B')} {start_date.day} to {end_date.day}, {end_date.year}"
        else:
          # Different months, same year: "June 21 to July 4, 2025"
          return f"{start_date.strftime('%B')} {start_date.day} to {end_date.strftime('%B')} {end_date.day}, {end_date.year}"
      else:
        # Different years: "December 31, 2024 to January 6, 2025"
        return f"{start_date.strftime('%B')} {start_date.day}, {start_date.year} to {end_date.strftime('%B')} {end_date.day}, {end_date.year}"

    def format_subject_date_range(start_date, end_date):
      if start_date.year == end_date.year:
        # Same year
        if start_date.month == end_date.month:
          # Same month: "July 4-18, 2025"
          return f"{start_date.strftime('%B')} {start_date.day}-{end_date.day}, {end_date.year}"
        else:
          # Different months, same year: "June 21-July 4, 2025"
          return f"{start_date.strftime('%B')} {start_date.day}-{end_date.strftime('%B')} {end_date.day}, {end_date.year}"
      else:
        # Different years: "December 31, 2024-January 6, 2025"
        return f"{start_date.strftime('%B')} {start_date.day}, {start_date.year}-{end_date.strftime('%B')} {end_date.day}, {end_date.year}"
    
    try:
      msg = MIMEMultipart()
      msg['From'] = sender_email
      msg['To'] = recipient_email
      
      # Format the date range for subject and body
      date_range = format_date_range(pay_period['start'], pay_period['end'])
      subject_date_range = format_subject_date_range(pay_period['start'], pay_period['end'])
      msg['Subject'] = f"Invoice #{invoice_number} ({subject_date_range})"
      
      body = f"""Hi {self.RECIPIENT_NAME},

Please find attached my invoice for the pay period {date_range}.

The invoice is attached in both PDF and Excel formats.

Best regards,
{self.YOUR_NAME}
"""
      
      msg.attach(MIMEText(body, 'plain'))
      
      # Attach PDF
      if os.path.exists(pdf_path):
        with open(pdf_path, "rb") as attachment:
          part = MIMEBase('application', 'octet-stream')
          part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
          'Content-Disposition',
          f'attachment; filename={os.path.basename(pdf_path)}'
        )
        msg.attach(part)
        print(f"Attached PDF: {pdf_path}")
      else:
        print(f"Warning: PDF file not found: {pdf_path}")
      
      # Attach Excel
      if os.path.exists(excel_path):
        with open(excel_path, "rb") as attachment:
          part = MIMEBase('application', 'octet-stream')
          part.set_payload(attachment.read())
        
        encoders.encode_base64(part)
        part.add_header(
          'Content-Disposition',
          f'attachment; filename={os.path.basename(excel_path)}'
        )
        msg.attach(part)
        print(f"Attached Excel: {excel_path}")
      else:
        print(f"Warning: Excel file not found: {excel_path}")
      
      # Send email
      server = smtplib.SMTP(self.GMAIL_SMTP, 587)
      server.starttls()
      server.login(sender_email, sender_password)
      text = msg.as_string()
      server.sendmail(sender_email, recipient_email, text)
      server.quit()
      
      print("Invoice email sent successfully with attachments!")
      return True
      
    except Exception as e:
      print(f"Error sending email: {e}")
      return False

  def get_user_confirmation(self, message, default="n"):
    """Get yes/no confirmation from user"""
    valid_responses = {'yes': True, 'y': True, 'no': False, 'n': False}
    default_text = " [Y/n]" if default.lower() == "y" else " [y/N]"
    
    while True:
      try:
        response = input(f"{message}{default_text}: ").lower().strip()
        
        if response == "":
          return valid_responses[default.lower()]
        elif response in valid_responses:
          return valid_responses[response]
        else:
          print("Please answer yes or no (y/n)")
      except KeyboardInterrupt:
        print("\nOperation cancelled by user")
        return False

  def review_and_confirm_shifts(self, shifts, pay_period):
    """Review all shifts for the pay period and allow edits before invoice generation"""
    
    def display_shifts():
      """Helper function to display the shifts table"""
      print("\n" + "="*50)
      print("REVIEW SHIFTS BEFORE INVOICE GENERATION")
      print("="*50)
      print(f"Pay Period: {pay_period['start'].strftime('%B %d')} - {pay_period['end'].strftime('%B %d, %Y')}")
      print("-"*50)
      
      # Show all shifts with calculated hours
      total_hours = 0
      print("SHIFTS FOUND:")
      for i, shift in enumerate(shifts, 1):
        if 'hours' in shift and shift['hours'] is not None:
          hours = float(shift['hours'])
          hours_source = "(manual)"
        else:
          hours = self.calculate_shift_hours(shift['start_time'], shift['end_time'])
          hours_source = "(calculated)"
        
        total_hours += hours
        start_display = self.normalize_time_display(shift['start_time'])
        end_display = self.normalize_time_display(shift['end_time'])
        print(f"{i:2d}. {shift['date']}: {start_display} - {end_display} = {hours}h {hours_source}")
      
      print("-"*50)
      print(f"TOTAL HOURS: {total_hours:.2f}")
      print(f"TOTAL PAY: ${total_hours * self.HOURLY_RATE:.2f}")
      print("="*50)
      return total_hours
    
    # Initial display
    display_shifts()
    
    while True:
      print("\nOptions:")
      print("1. Looks good - Generate invoice")
      print("2. Edit a shift")
      print("3. Show shifts again")
      print("4. Cancel invoice generation")
      
      try:
        choice = input("\nSelect option (1-4): ").strip()
        
        if choice == "1":
          # Confirm invoice generation
          return True
          
        elif choice == "2":
          # Edit a shift
          print("\nAvailable shifts to edit:")
          for i, shift in enumerate(shifts, 1):
            hours = shift.get('hours') or self.calculate_shift_hours(shift['start_time'], shift['end_time'])
            print(f"{i}. {shift['date']}: {shift['start_time']} - {shift['end_time']} ({hours}h)")
          
          try:
            shift_num = int(input(f"\nEnter shift number to edit (1-{len(shifts)}): ")) - 1
            if 0 <= shift_num < len(shifts):
              shift_to_edit = shifts[shift_num]
              
              print(f"\nEditing shift for {shift_to_edit['date']}")
              print(f"Current: {shift_to_edit['start_time']} - {shift_to_edit['end_time']}")
              print("(Leave blank to keep current value)")
              
              new_start = input("New start time (e.g., 8:00am): ").strip()
              new_end = input("New end time (e.g., 3:00pm): ").strip()
              hours_input = input("Manual hours override (leave blank to calculate): ").strip()
              
              # Apply changes to the shift in memory
              if new_start:
                shift_to_edit['start_time'] = new_start
              if new_end:
                shift_to_edit['end_time'] = new_end
              if hours_input:
                shift_to_edit['hours'] = float(hours_input)
              elif new_start or new_end:
                # Remove manual hours if times changed (will recalculate)
                shift_to_edit.pop('hours', None)
              
              # Also update the stored data
              self.edit_shift_manually(shift_to_edit['date'], new_start, new_end, 
                                     float(hours_input) if hours_input else None)
              
              print("Shift updated!")
              # Show updated shifts after edit
              display_shifts()
            else:
              print("Invalid shift number")
          except (ValueError, IndexError):
            print("Invalid input")
            
        elif choice == "3":
          # Show shifts again - redisplay the table
          display_shifts()
          
        elif choice == "4":
          # Cancel invoice generation
          return False
          
        else:
          print("Invalid choice. Please select 1-4.")
          
      except KeyboardInterrupt:
        print("\nOperation cancelled")
        return False
      except Exception as e:
        print(f"Error: {e}")

  def edit_shift_manually(self, date, start_time=None, end_time=None, hours=None):
    """Manually edit a shift's details"""
    all_shifts_data = self.load_shifts_data()
    
    # Find the shift to edit
    shift_found = False
    for period_key, period_data in all_shifts_data.items():
      for shift in period_data.get('shifts', []):
        if shift['date'] == date:
          shift_found = True
          print(f"Found shift for {date}:")
          print(f"  Current: {shift['start_time']} - {shift['end_time']}")
          
          # Update fields if provided
          if start_time:
            shift['start_time'] = start_time
            print(f"  Updated start time to: {start_time}")
          
          if end_time:
            shift['end_time'] = end_time
            print(f"  Updated end time to: {end_time}")
          
          if hours is not None:
            shift['hours'] = float(hours)
            print(f"  Set manual hours to: {hours}")
          elif start_time or end_time:
            # Recalculate hours if times changed but hours not manually set
            calculated = self.calculate_shift_hours(shift['start_time'], shift['end_time'])
            shift['hours'] = calculated
            print(f"  Calculated hours: {calculated}")
          
          # Save changes
          if self.save_shifts_data(all_shifts_data):
            print("Shift updated successfully!")
          else:
            print("Failed to save changes")
          
          return True
    
    if not shift_found:
      print(f"No shift found for date {date}")
      return False
    
  def add_shift_manually(self, date, start_time, end_time, hours=None):
    """Manually add a new shift to the appropriate pay period"""
    try:
      # Parse the date to determine which pay period it belongs to
      current_year = datetime.now().year
      shift_datetime = datetime.strptime(f"{date}/{current_year}", "%m/%d/%Y")
      pay_period = self.get_pay_period_for_date(shift_datetime)
      
      if not pay_period:
        print(f"Could not determine pay period for date {date}")
        return False
      
      period_key = pay_period['period_key']
      
      # Load existing data
      all_shifts_data = self.load_shifts_data()
      
      # Initialize period data if it doesn't exist
      if period_key not in all_shifts_data:
        all_shifts_data[period_key] = {
          'period_info': pay_period,
          'shifts': [],
          'weeks_received': 0,
          'complete': False,
          'invoice_generated': False
        }
      
      # Check if shift already exists for this date
      existing_dates = {s['date'] for s in all_shifts_data[period_key]['shifts']}
      if date in existing_dates:
        print(f"A shift already exists for {date}. Use the edit option instead.")
        return False
      
      # Create new shift
      new_shift = {
        'date': date,
        'start_time': start_time,
        'end_time': end_time
      }
      
      # Add hours if provided, otherwise it will be calculated
      if hours is not None:
        new_shift['hours'] = float(hours)
      
      # Add the shift
      all_shifts_data[period_key]['shifts'].append(new_shift)

      # Sort shifts chronologically
      all_shifts_data[period_key]['shifts'] = self.sort_shifts_by_date(all_shifts_data[period_key]['shifts'])

      # Save changes
      if self.save_shifts_data(all_shifts_data):
        print(f"Successfully added shift for {date}")
        print(f"  {start_time} - {end_time}")
        if hours:
          print(f"  Hours: {hours} (manual)")
        else:
          calculated = self.calculate_shift_hours(start_time, end_time)
          print(f"  Hours: {calculated} (calculated)")
        print(f"  Pay Period: {pay_period['start'].strftime('%B %d')} - {pay_period['end'].strftime('%B %d, %Y')}")
        return True
      else:
        print("Failed to save shift")
        return False
        
    except Exception as e:
      print(f"Error adding shift: {e}")
      return False
    
  def delete_shift_manually(self, date):
    """Delete a shift from the appropriate pay period"""
    all_shifts_data = self.load_shifts_data()
    
    # Find and delete the shift
    shift_found = False
    for period_key, period_data in all_shifts_data.items():
      for i, shift in enumerate(period_data.get('shifts', [])):
        if shift['date'] == date:
          shift_found = True
          deleted_shift = period_data['shifts'].pop(i)
          
          print(f"Deleted shift for {date}:")
          print(f"  {deleted_shift['start_time']} - {deleted_shift['end_time']}")
          
          # Save changes
          if self.save_shifts_data(all_shifts_data):
            print("Shift deleted successfully!")
          else:
            print("Failed to save changes")
          
          return True
    
    if not shift_found:
      print(f"No shift found for date {date}")
      return False

  def interactive_shift_editor(self):
    """Interactive utility to edit shifts"""
    all_shifts_data = self.load_shifts_data()
    
    if not all_shifts_data:
      print("No shifts data found")
      return
    
    # Show all shifts
    print("Available shifts to edit:")
    print("-" * 50)
    
    all_shifts = []
    for period_key, period_data in all_shifts_data.items():
      for shift in period_data.get('shifts', []):
        all_shifts.append(shift)
        hours_info = f" (manual: {shift['hours']}h)" if 'hours' in shift else ""
        print(f"{shift['date']}: {shift['start_time']} - {shift['end_time']}{hours_info}")
    
    if not all_shifts:
      print("No shifts found to edit")
      return
    
    print("-" * 50)
    
    # Get user input
    try:
      date_to_edit = input("Enter date to edit (MM/DD format): ").strip()
      
      print(f"\nEditing shift for {date_to_edit}")
      print("Leave blank to keep current value")
      
      new_start = input("New start time (e.g., 8:00am): ").strip() or None
      new_end = input("New end time (e.g., 3:00pm): ").strip() or None
      
      hours_input = input("Manual hours override (leave blank to calculate): ").strip()
      new_hours = float(hours_input) if hours_input else None
      
      # Apply changes
      self.edit_shift_manually(date_to_edit, new_start, new_end, new_hours)
      
    except KeyboardInterrupt:
      print("\nEditor cancelled")
    except Exception as e:
      print(f"Error: {e}")

  def interactive_add_shift(self):
    """Interactive utility to add a new shift"""
    print("\n" + "="*50)
    print("ADD NEW SHIFT")
    print("="*50)
    
    try:
      date = input("Enter date (MM/DD format): ").strip()
      
      # Validate date format
      if not date:
        print("Date is required")
        return False
      
      start_time = input("Start time (e.g., 8:00am): ").strip()
      end_time = input("End time (e.g., 3:00pm): ").strip()
      
      hours_input = input("Manual hours override (leave blank to auto-calculate): ").strip()
      hours = float(hours_input) if hours_input else None
      
      # Validate: either (start_time AND end_time) OR hours must be provided
      if not hours and (not start_time or not end_time):
        print("\nError: You must provide either:")
        print("  1. Both start time AND end time, OR")
        print("  2. Manual hours override")
        return False
      
      # If hours provided but no times, use placeholder times
      if hours and (not start_time or not end_time):
        print("\nNote: Using placeholder times since manual hours were provided")
        start_time = "7:00am"
        end_time = "3:00pm"
      
      # Show preview
      print("\nShift to add:")
      print(f"  Date: {date}")
      print(f"  Time: {start_time} - {end_time}")
      if hours:
        print(f"  Hours: {hours} (manual)")
      else:
        calculated = self.calculate_shift_hours(start_time, end_time)
        print(f"  Hours: {calculated} (calculated)")
      
      # Confirm
      if self.get_user_confirmation("\nAdd this shift?", default="y"):
        return self.add_shift_manually(date, start_time, end_time, hours)
      else:
        print("Cancelled")
        return False
        
    except KeyboardInterrupt:
      print("\nCancelled")
      return False
    except ValueError as e:
      print(f"Error: Invalid number format - {e}")
      return False
    except Exception as e:
      print(f"Error: {e}")
      return False
    
  def interactive_delete_shift(self):
    """Interactive utility to delete a shift from any non-invoiced period"""
    print("\n" + "="*50)
    print("DELETE SHIFT")
    print("="*50)
    
    # Load shifts data
    all_shifts_data = self.load_shifts_data()
    
    if not all_shifts_data:
      print("No shifts data found")
      return False
    
    # Collect all shifts from non-invoiced periods
    available_shifts = []
    for period_key, period_data in all_shifts_data.items():
      # Only show shifts from periods that haven't been invoiced
      if not period_data.get('invoice_generated', False):
        for shift in period_data.get('shifts', []):
          available_shifts.append(shift)
    
    if not available_shifts:
      print("No shifts found in non-invoiced periods")
      return False
    
    # Show available shifts grouped by period would be nice, but let's keep it simple
    print("Shifts available to delete (from non-invoiced periods):")
    for shift in available_shifts:
      print(f"  {shift['date']}: {shift['start_time']} - {shift['end_time']}")
    
    print("-" * 50)
    
    try:
      date = input("Enter date to delete (MM/DD format): ").strip()
      
      if not date:
        print("Date is required")
        return False
      
      # Verify the shift exists in non-invoiced periods
      shift_exists = any(shift['date'] == date for shift in available_shifts)
      
      if not shift_exists:
        print(f"No shift found for {date} in non-invoiced periods")
        return False
      
      # Confirm deletion
      if self.get_user_confirmation(f"\nAre you sure you want to delete the shift on {date}?", default="n"):
        return self.delete_shift_manually(date)
      else:
        print("Cancelled")
        return False
        
    except KeyboardInterrupt:
      print("\nCancelled")
      return False
    except Exception as e:
      print(f"Error: {e}")
      return False

  def interactive_shift_editor_for_period(self, period_shifts):
    """Interactive utility to edit shifts for a specific period"""
  
    if not period_shifts:
      print("No shifts found to edit in this period")
      return
    
    # Show only shifts from this period
    print("\nAvailable shifts to edit in this period:")
    print("-" * 50)
    
    for shift in period_shifts:
      hours_info = f" (manual: {shift['hours']}h)" if 'hours' in shift else ""
      print(f"{shift['date']}: {shift['start_time']} - {shift['end_time']}{hours_info}")
    
    print("-" * 50)
    
    # Get user input
    try:
      date_to_edit = input("Enter date to edit (MM/DD format): ").strip()
      
      print(f"\nEditing shift for {date_to_edit}")
      print("Leave blank to keep current value")
      
      new_start = input("New start time (e.g., 8:00am): ").strip() or None
      new_end = input("New end time (e.g., 3:00pm): ").strip() or None
      
      hours_input = input("Manual hours override (leave blank to calculate): ").strip()
      new_hours = float(hours_input) if hours_input else None
      
      # Apply changes
      self.edit_shift_manually(date_to_edit, new_start, new_end, new_hours)
    
    except KeyboardInterrupt:
      print("\nEditor cancelled")
    except Exception as e:
      print(f"Error: {e}")

  def check_for_incomplete_invoices(self):
    """Check for pay periods that are complete but don't have invoices generated"""
    all_shifts_data = self.load_shifts_data()
    incomplete_periods = {}
    
    for period_key, period_data in all_shifts_data.items():
      # Check if period is complete but invoice not generated
      if (period_data.get('complete', False) and 
          not period_data.get('invoice_generated', False)):
        incomplete_periods[period_key] = period_data
    
    return incomplete_periods

  def generate_invoice_for_period(self, period_key, period_data, spreadsheet_id, 
                                 gmail_user, gmail_password, recipient_email, send_email):
    """Generate invoice for a specific completed period"""
    try:
      # Extract period info and shifts
      period_info = period_data['period_info']
      shifts = period_data['shifts']
      
      # Convert dates from strings to datetime if needed
      pay_period = self.convert_period_dates_to_datetime(period_info)
      
      print(f"\nGenerating invoice for period: {period_key}")
      print(f"Date range: {pay_period['start'].strftime('%B %d')} - {pay_period['end'].strftime('%B %d, %Y')}")
      print(f"Shifts: {len(shifts)}")
      
      # Review shifts before generating invoice
      if not self.review_and_confirm_shifts(shifts, pay_period):
        print("Invoice generation cancelled")
        return False
      
      print("Shifts confirmed! Generating invoice...")
      
      # Create new invoice sheet
      new_sheet_id = self.create_new_invoice_sheet(spreadsheet_id, pay_period)
      
      if not new_sheet_id:
        print("Failed to create new invoice sheet")
        return False
      
      # Create invoice data
      invoice_data = self.create_invoice_data(shifts, pay_period)
      
      # Update Google Sheet
      success = self.update_google_sheet(new_sheet_id, invoice_data, pay_period)
      
      if success:
        # Mark invoice as generated
        self.mark_invoice_generated(period_key)
        print("Invoice updated successfully!")
        
        if send_email and recipient_email:
          # Get the invoice number for the email
          invoice_number = self.load_invoice_counter() - 1
          
          # Generate filenames based on pay period
          date_str = pay_period['end'].strftime('%Y_%m_%d')
          pdf_filename = f"{self.YOUR_NAME} Invoice {date_str}.pdf"
          excel_filename = f"{self.YOUR_NAME} Invoice {date_str}.xlsx"

          # Generate the files
          print("Generating invoice files...")
          print("Exporting PDF...")
          pdf_success = self.export_sheet_as_pdf(new_sheet_id, pdf_filename)
          
          print("Exporting Excel...")
          excel_success = self.export_sheet_as_excel(new_sheet_id, excel_filename)
          
          if not (pdf_success and excel_success):
            print("Failed to export files - cannot proceed with email")
            return False
          
          # Show invoice summary and ask for confirmation
          print("\n" + "="*50)
          print("READY TO SEND INVOICE EMAIL")
          print("="*50)
          print(f"Invoice #: {invoice_number}")
          print(f"Pay Period: {pay_period['start'].strftime('%B %d')} - {pay_period['end'].strftime('%B %d, %Y')}")
          print(f"Recipient: {recipient_email}")
          print(f"PDF File: {pdf_filename}")
          print(f"Excel File: {excel_filename}")
          print("="*50)
          
          # Get user confirmation
          if self.get_user_confirmation("Send invoice email now?", default="y"):
            print("Proceeding with email...")
            
            # Send email with attachments
            email_success = self.send_invoice_email_with_attachments(
              recipient_email,
              pdf_filename,
              excel_filename,
              gmail_user,
              gmail_password,
              pay_period,
              invoice_number
            )
            
            if email_success:
              print("Invoice files exported and emailed successfully!")
              
              try:
                os.remove(pdf_filename)
                os.remove(excel_filename)
                print("Temporary files cleaned up")
              except OSError:
                print("Note: Could not clean up temporary files")
            else:
              print("Failed to send email")
              print(f"Files are still available: {pdf_filename}, {excel_filename}")
          else:
            print("Email sending cancelled by user")
            print(f"Invoice files have been generated and saved:")
            print(f"   PDF: {pdf_filename}")
            print(f"   Excel: {excel_filename}")
            print(f"   Location: {os.getcwd()}")
      
      return success
      
    except Exception as e:
      print(f"Error generating invoice for period {period_key}: {e}")
      return False

  def process_schedule_email(self, gmail_user, gmail_password, spreadsheet_id, 
                            recipient_email=None, send_email=True):
    """Main method to process latest Homebase email and update invoice"""
    
    # Check for incomplete invoices before processing new emails
    incomplete_periods = self.check_for_incomplete_invoices()
    if incomplete_periods:
      print("\n" + "="*50)
      print("WARNING: Found complete pay periods without invoices!")
      print("="*50)
      
      for period_key, period_data in incomplete_periods.items():
        period_info = period_data['period_info']
        
        # Use helper to format dates
        start_str, end_str = self.format_period_dates(period_info)
        
        print(f"\nPay Period: {start_str} to {end_str}")
        print(f"   Status: Complete but no invoice generated")
        print(f"   Shifts: {len(period_data['shifts'])}")
      
      print("\n" + "="*50)
      print("You must generate invoices for these periods before processing new emails.")
      print("="*50)
      
      while True:
        print("\nOptions:")
        print("1. Generate invoice for incomplete period")
        print("2. Edit shifts for incomplete period") 
        print("3. Continue anyway (not recommended)")
        print("4. Cancel and exit")
        
        try:
          choice = input("\nSelect option (1-4): ").strip()
          
          if choice == "1":
            # Use helper to select period
            period_key, period_data = self.select_period_from_list(
              incomplete_periods, 
              "Available periods"
            )
            
            if period_key is None:
              continue
            
            # Generate invoice for selected period
            success = self.generate_invoice_for_period(
              period_key, period_data, spreadsheet_id, 
              gmail_user, gmail_password, recipient_email, send_email
            )
            
            if success:
              print("Invoice generated successfully!")
              incomplete_periods = self.check_for_incomplete_invoices()
              if not incomplete_periods:
                print("All invoices are now up to date!")
                break
            else:
              print("Failed to generate invoice")
              
          elif choice == "2":
            # Use helper to select period
            period_key, period_data = self.select_period_from_list(
              incomplete_periods,
              "Available periods"
            )
            
            if period_key is None:
              continue
            
            # Show shifts for this period and allow editing
            print(f"\nShifts for period {period_key}:")
            for shift in period_data['shifts']:
              hours = shift.get('hours', 'calculated')
              print(f"  {shift['date']}: {shift['start_time']}-{shift['end_time']} ({hours}h)")
            
            self.interactive_shift_editor_for_period(period_data['shifts'])
            
          elif choice == "3":
            print("Continuing with incomplete invoices...")
            break
            
          elif choice == "4":
            print("Operation cancelled")
            return False
            
          else:
            print("Invalid choice. Please select 1-4.")
            
        except KeyboardInterrupt:
          print("\nOperation cancelled")
          return False
    
    # Original email processing logic continues here...
    # Connect to Gmail
    mail = self.authenticate_gmail(gmail_user, gmail_password)
    if not mail:
      return False
    
    try:
      # Search for unread Homebase emails first, then all if none unread
      mail.select('inbox')
      
      # Try new address first (unread)
      result, messages = mail.search(None, 'UNSEEN FROM "noreply@joinhomebase.com" SUBJECT "schedule"')

      # If nothing, try old address (unread)
      if result != 'OK' or not messages[0]:
          result, messages = mail.search(None, 'UNSEEN FROM "no-reply@joinhomebase.com" SUBJECT "schedule"')

      # If still nothing, try new address (all)
      if result != 'OK' or not messages[0]:
          result, messages = mail.search(None, 'FROM "noreply@joinhomebase.com" SUBJECT "schedule"')
          print("No unread Homebase emails found, checking most recent...")
          
      # If still nothing, try old address (all)
      if result != 'OK' or not messages[0]:
          result, messages = mail.search(None, 'FROM "no-reply@joinhomebase.com" SUBJECT "schedule"')
      else:
          print("Found unread Homebase email")

      if result != 'OK' or not messages[0]:
          print("No Homebase emails found")
          return False
      
      # Get the most recent email
      latest_email_id = messages[0].split()[-1]
      result, msg_data = mail.fetch(latest_email_id, '(RFC822)')
      
      if result != 'OK':
        print("Failed to fetch email")
        return False
      
      # Parse email
      email_body = email.message_from_bytes(msg_data[0][1])
      
      # Extract email content - prefer HTML for new format
      email_content = ""
      html_content = None

      if email_body.is_multipart():
          for part in email_body.walk():
              if part.get_content_type() == "text/html":
                  html_content = part.get_payload(decode=True).decode()
                  break
          
          if not html_content:
              for part in email_body.walk():
                  if part.get_content_type() == "text/plain":
                      email_content = part.get_payload(decode=True).decode()
                      break
          else:
              email_content = html_content
      else:
          email_content = email_body.get_payload(decode=True).decode()
      
      # Parse shifts from email
      shifts = self.parse_homebase_email(email_content)
      if not shifts:
        print("No shifts found in email")
        return False
      
      print(f"Found {len(shifts)} shifts in email")
      
      # Store shifts and check if bi-weekly period is complete
      complete_period = self.store_weekly_shifts(shifts)
      
      if complete_period:
        print("Bi-weekly period complete!")
        
        # Extract pay period info
        pay_period = complete_period['period']
        period_key = complete_period['period_key']
        
        # CRITICAL: Review shifts before generating invoice
        if not self.review_and_confirm_shifts(complete_period['shifts'], pay_period):
          print("Invoice generation cancelled. You can edit shifts and run the script again.")
          return False
        
        print("Shifts confirmed! Generating invoice...")
        
        # Create new invoice sheet
        new_sheet_id = self.create_new_invoice_sheet(
          spreadsheet_id, 
          pay_period
        )
        
        if not new_sheet_id:
          print("Failed to create new invoice sheet")
          return False
        
        # Create invoice data
        invoice_data = self.create_invoice_data(
          complete_period['shifts'], 
          pay_period
        )
        
        # Update Google Sheet (use the new sheet ID)
        success = self.update_google_sheet(new_sheet_id, invoice_data, pay_period)
        
        if success:
          # IMPORTANT: Mark invoice as generated to prevent duplicate processing
          self.mark_invoice_generated(period_key)
          
          print("Invoice updated successfully!")
          
          if send_email and recipient_email:
            # Get the invoice number for the email
            invoice_number = self.load_invoice_counter() - 1  # We incremented it in update_google_sheet
            
            # Generate filenames based on pay period
            date_str = pay_period['end'].strftime('%Y_%m_%d')
            pdf_filename = f"{self.YOUR_NAME} Invoice {date_str}.pdf"
            excel_filename = f"{self.YOUR_NAME} Invoice {date_str}.xlsx"

            # Generate the files FIRST, before asking for confirmation
            print("Generating invoice files...")
            print("Exporting PDF...")
            pdf_success = self.export_sheet_as_pdf(new_sheet_id, pdf_filename)
            
            print("Exporting Excel...")
            excel_success = self.export_sheet_as_excel(new_sheet_id, excel_filename)
            
            if not (pdf_success and excel_success):
              print("Failed to export files - cannot proceed with email")
              return False
            
            # Show invoice summary and ask for confirmation
            print("\n" + "="*50)
            print("READY TO SEND INVOICE EMAIL")
            print("="*50)
            print(f"Invoice #: {invoice_number}")
            print(f"Pay Period: {pay_period['start'].strftime('%B %d')} - {pay_period['end'].strftime('%B %d, %Y')}")
            print(f"Recipient: {recipient_email}")
            print(f"PDF File: {pdf_filename}")
            print(f"Excel File: {excel_filename}")
            print(f"Files saved to: {os.getcwd()}")
            print("="*50)
            
            # Get user confirmation
            if self.get_user_confirmation("Send invoice email now?", default="y"):
              print("Proceeding with email...")
              
              # Send email with attachments (files already generated)
              email_success = self.send_invoice_email_with_attachments(
                recipient_email,
                pdf_filename,
                excel_filename,
                gmail_user,
                gmail_password,
                pay_period,
                invoice_number
              )
              
              if email_success:
                print("Invoice files exported and emailed successfully!")
                
                try:
                  os.remove(pdf_filename)
                  os.remove(excel_filename)
                  print("Temporary files cleaned up")
                except OSError:
                  print("Note: Could not clean up temporary files")
              else:
                print("Failed to send email")
                print(f"Files are still available: {pdf_filename}, {excel_filename}")
            else:
              print("Email sending cancelled by user")
              print(f"Invoice files have been generated and saved:")
              print(f"   PDF: {pdf_filename}")
              print(f"   Excel: {excel_filename}")
              print(f"   Location: {os.getcwd()}")
              print("   You can send these files manually or run the script again.")
        
        return success
      else:
        # Load shifts data to provide better messaging
        all_shifts_data = self.load_shifts_data()
        
        # We need to determine which period these shifts belong to
        if shifts:
          first_shift_date = shifts[0]['date']
          current_year = datetime.now().year
          shift_datetime = datetime.strptime(f"{first_shift_date}/{current_year}", "%m/%d/%Y")
          pay_period = self.get_pay_period_for_date(shift_datetime)
          
          if pay_period:
            period_key = pay_period['period_key']
            
            if period_key in all_shifts_data:
              weeks_received = all_shifts_data[period_key]['weeks_received']
              if all_shifts_data[period_key].get('invoice_generated', False):
                print("All shifts for this pay period are already stored and invoice has been generated.")
              elif weeks_received < 2:
                print(f"Shifts already stored. Waiting for week {weeks_received + 1} of this pay period to generate invoice.")
              else:
                print("Pay period complete but invoice not yet generated.")
            else:
              print("No new shifts to process.")
          else:
            print("Could not determine pay period for these shifts.")
        else:
          print("No shifts found to process.")
        
        return True
        
    except Exception as e:
      print(f"Error processing schedule email: {e}")
      return False
    finally:
      try:
        mail.close()
        mail.logout()
      except Exception:
        pass

def main_menu():
  """Main menu with options"""
  
  # Configuration from environment variables
  GMAIL_USER = os.getenv('GMAIL_USER')
  GMAIL_PASSWORD = os.getenv('GMAIL_PASSWORD')
  SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
  RECIPIENT = os.getenv('RECIPIENT_EMAIL')
  YOUR_NAME = os.getenv('YOUR_NAME')
  RECIPIENT_NAME = os.getenv('RECIPIENT_NAME')
  FIRST_PAY_PERIOD_START = os.getenv('FIRST_PAY_PERIOD_START')
  
  # Validate required environment variables
  required_vars = {
    'GMAIL_USER': GMAIL_USER,
    'GMAIL_PASSWORD': GMAIL_PASSWORD,
    'SPREADSHEET_ID': SPREADSHEET_ID,
    'RECIPIENT_EMAIL': RECIPIENT,
    'YOUR_NAME': YOUR_NAME,
    'RECIPIENT_NAME': RECIPIENT_NAME,
    'FIRST_PAY_PERIOD_START': FIRST_PAY_PERIOD_START
  }

  missing_vars = [var for var, value in required_vars.items() if not value]
  
  if missing_vars:
    print("Missing required environment variables:")
    for var in missing_vars:
      print(f"   - {var}")
    print("\nPlease check your .env file and ensure all required variables are set.")
    print("See .env.example for reference.")
    sys.exit(1)
  
  automator = BiWeeklyInvoiceAutomator()
  
  while True:
    print("\n" + "="*50)
    print("INVOICE AUTOMATION MENU")
    print("="*50)
    print("1. Process new schedule email")
    print("2. Edit shift manually")
    print("3. Add shift manually")
    print("4. Delete shift manually")
    print("5. View all pay periods")
    print("6. View tax summary")
    print("7. Force complete period (for holidays/closures)")
    print("8. Exit")
    print("-"*50)

    try:
      choice = input("Select option (1-8): ").strip()
      
      if choice == "1":
        # Process schedule email and generate invoice
        success = automator.process_schedule_email(
          GMAIL_USER, GMAIL_PASSWORD, SPREADSHEET_ID,
          recipient_email=RECIPIENT, send_email=True
        )
        
        if success:
          print("Automation completed successfully!")
        else:
          print("Automation failed. Check the logs above.")
      
      elif choice == "2":
        # Edit shift manually
        date = input("Enter date (MM/DD): ").strip()
        start = input("Start time (e.g., 8:00am) [optional]: ").strip() or None
        end = input("End time (e.g., 3:00pm) [optional]: ").strip() or None
        hours_input = input("Hours override [optional]: ").strip()
        hours = float(hours_input) if hours_input else None
        
        automator.edit_shift_manually(date, start, end, hours)
      
      elif choice == "3":
        # Add shift manually
        automator.interactive_add_shift()
      
      elif choice == "4":
        # Delete shift manually
        automator.interactive_delete_shift()
        
      elif choice == "5":
        automator.list_pay_periods()
        
      elif choice == "6":
        automator.view_tax_summary()
        
      elif choice == "7":
        automator.force_complete_period()
        
      elif choice == "8":
        print("Goodbye!")
        break
        
      else:
        print("Invalid choice. Please select 1-8.")
        
    except KeyboardInterrupt:
      print("\nGoodbye!")
      break
    except Exception as e:
      print(f"Error: {e}")
if __name__ == "__main__":
  print("Invoice Automation Script Starting...")
  
  # Environment checks
  check_environment()
  if not check_required_files():
    sys.exit(1)
  
  # Run the menu
  main_menu()
