import pandas as pd
import random
from datetime import datetime
from openpyxl import load_workbook
import os
import sys


# --- 1. CONFIGURATION: ABSOLUTE PATHS (REQUIRED FOR YOUR ENVIRONMENT) ---
# IMPORTANT: Manually verify these absolute path strings on your PC.
# Using 'os.path.join' is safer, but hardcoding the full string is done here for a definitive fix.
PROJECT_FOLDER = 'C:\\Users\\PARAS\\Desktop\\PANDA_PROJECT' 

# Construct the full paths using os.path.join for reliability
REQUEST_FILE = os.path.join(PROJECT_FOLDER, 'Request_Queue.xlsx') 
SCHEDULE_FILE = os.path.join(PROJECT_FOLDER, 'Doctor_Availability.xlsx')

# Define the EXPECTED column names for the Request Queue file
# This list MUST match the column ORDER of your Request_Queue.xlsx data.
EXPECTED_COLUMNS = ['Patient_Name', 'DOB', 'Requested_Specialty', 'Contact_Number', 'Status']

# --- 2. CORE SCHEDULING LOGIC ---

def get_next_available_slot(specialty):
    """
    Mocks the EHR system lookup. Finds the next unbooked slot, marks it as booked, 
    and saves the updated schedule.
    """
    
    # 2.1 Check if the Doctor Availability file exists
    if not os.path.isfile(SCHEDULE_FILE):
        return f"System Error: Schedule file not found at {SCHEDULE_FILE}", None

    try:
        # Load the mock EHR schedule, assuming data is on the 'Schedule' sheet
        df_schedule = pd.read_excel(SCHEDULE_FILE, sheet_name='Schedule') 
    except ValueError:
        return "System Error: Could not find sheet named 'Schedule'.", None
    except Exception as e:
        return f"System Error loading schedule data: {e}", None

    # Ensure the 'Is_Booked' column in Doctor_Availability is clean
    if 'Is_Booked' in df_schedule.columns:
        df_schedule['Is_Booked'] = df_schedule['Is_Booked'].astype(str).str.upper().str.strip()
    
    # 1. Filter by Specialty and Availability (Case-insensitive match)
    df_filtered = df_schedule[
        (df_schedule['Specialty'].astype(str).str.lower() == specialty.lower()) & 
        (df_schedule['Is_Booked'] == 'NO')
    ].copy()

    if df_filtered.empty:
        return "No slots available for this specialty.", None
    
    # 2. Take the first unbooked slot found
    next_slot = df_filtered.iloc[0]
    
    # 3. Mark the slot as booked in the DataFrame
    schedule_index_to_update = next_slot.name
    df_schedule.loc[schedule_index_to_update, 'Is_Booked'] = 'Yes'
    
    # 4. Save the ENTIRE updated schedule back to Excel
    with pd.ExcelWriter(SCHEDULE_FILE, engine='openpyxl') as writer:
        df_schedule.to_excel(writer, sheet_name='Schedule', index=False)
    
    return (
        f"{next_slot['Day']} at {next_slot['Time_Slot']}", 
        next_slot['Doctor_Name']
    )

# --------------------------------------------------------------------------

def process_scheduling_requests():
    """
    Main function to read requests from the queue, process them, and update the status.
    """
    print(f"--- Smart Scheduler Bot Initiated at {datetime.now().strftime('%H:%M:%S')} ---")
    
    # Check for critical file existence using the hardcoded path
    if not os.path.isfile(REQUEST_FILE):
        print(f"CRITICAL ERROR: Request queue file NOT FOUND at: {REQUEST_FILE}")
        return

    try:
        # 💡 FINAL FIX: Bypass corrupted header row and specify data sheet name
        # sheet_name='Requests' tells Pandas exactly where the data is.
        # header=None tells Pandas to read ALL content, including the header row, as data.
        df_requests = pd.read_excel(REQUEST_FILE, sheet_name='Requests', header=None) 
        
        # Manually assign the correct column names (forcing the 5-column structure)
        df_requests.columns = EXPECTED_COLUMNS 

        # Drop the first row (which was the actual header row now assigned index 0)
        df_requests = df_requests.iloc[1:].copy() 
        df_requests.reset_index(drop=True, inplace=True)

    except Exception as e:
        print(f"CRITICAL ERROR loading Request Queue data: {e}")
        print("ACTION: Double-check the sheet name in Request_Queue.xlsx is 'Requests'.")
        return

    # Treat blank/NaN in the Status column as 'New'
    df_requests['Status'] = df_requests['Status'].astype(str).str.strip().fillna('New')
    
    # Filter for new requests
    df_new_requests = df_requests[
        (df_requests['Status'].astype(str).str.upper() == 'NEW')
    ].copy()

    if df_new_requests.empty:
        print("✅ No new requests found to process. System is idle.")
        return

    print(f"➡️ Found {len(df_new_requests)} new requests to process...")

    for index, row in df_new_requests.iterrows():
        try:
            # These columns are guaranteed to exist due to the fixed read logic
            patient = row['Patient_Name']
            specialty = row['Requested_Specialty'] 
        except KeyError as ke:
            print(f"CRITICAL KEY ERROR: Missing column {ke}. Skipping row.")
            continue 

        print(f"\nProcessing request for {patient} ({specialty})...")
        
        # 1. Attempt to find and book a slot
        booked_slot, doctor = get_next_available_slot(specialty)
        
        # 2. Determine the final status message
        if booked_slot.startswith("No slots available"):
            status_message = "WAITLIST: Full capacity. Added to manual review."
        elif booked_slot.startswith("System Error"):
            status_message = f"ERROR: {booked_slot}"
        else:
            # Success!
            status_message = f"BOOKED: {booked_slot} with Dr. {doctor}"

        # 3. Update the 'Status' column in the original DataFrame
        original_index = row.name
        df_requests.loc[original_index, 'Status'] = status_message
        print(f"   Status updated to: {status_message}")


    # Save the updated Request Queue back to Excel
    df_requests.to_excel(REQUEST_FILE, sheet_name='Requests', index=False)
    print("\n--- Scheduler Bot Finished. Request Queue is now up-to-date. ---")


if __name__ == '__main__':
    process_scheduling_requests()
    
    