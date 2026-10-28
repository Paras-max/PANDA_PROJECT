# import pandas as pd
# import os
# from datetime import datetime
# from tkinter import *
# from tkinter import messagebox, scrolledtext

# # --- CONFIGURATION ---
# PROJECT_FOLDER = 'C:\\Users\\PARAS\\Desktop\\PANDA_PROJECT'
# REQUEST_FILE = os.path.join(PROJECT_FOLDER, 'Request_Queue.xlsx')
# SCHEDULE_FILE = os.path.join(PROJECT_FOLDER, 'Doctor_Availability.xlsx')
# EXPECTED_COLUMNS = ['Patient_Name', 'DOB', 'Requested_Specialty', 'Contact_Number', 'Status']

# # ---------------------------------------------------------------------
# # CORE SCHEDULING LOGIC
# # ---------------------------------------------------------------------
# def get_next_available_slot(specialty):
#     if not os.path.isfile(SCHEDULE_FILE):
#         return f"System Error: Schedule file not found at {SCHEDULE_FILE}", None
#     try:
#         df_schedule = pd.read_excel(SCHEDULE_FILE, sheet_name='Schedule')
#     except ValueError:
#         return "System Error: Could not find sheet named 'Schedule'.", None
#     except Exception as e:
#         return f"System Error loading schedule data: {e}", None

#     if 'Is_Booked' in df_schedule.columns:
#         df_schedule['Is_Booked'] = df_schedule['Is_Booked'].astype(str).str.upper().str.strip()

#     df_filtered = df_schedule[
#         (df_schedule['Specialty'].astype(str).str.lower() == specialty.lower()) &
#         (df_schedule['Is_Booked'] == 'NO')
#     ].copy()

#     if df_filtered.empty:
#         return "No slots available for this specialty.", None

#     next_slot = df_filtered.iloc[0]
#     schedule_index_to_update = next_slot.name
#     df_schedule.loc[schedule_index_to_update, 'Is_Booked'] = 'YES'

#     with pd.ExcelWriter(SCHEDULE_FILE, engine='openpyxl') as writer:
#         df_schedule.to_excel(writer, sheet_name='Schedule', index=False)

#     return f"{next_slot['Day']} at {next_slot['Time_Slot']}", next_slot['Doctor_Name']


# # ---------------------------------------------------------------------
# # PROCESSING REQUESTS
# # ---------------------------------------------------------------------
# def process_scheduling_requests(log_box):
#     log_box.delete(1.0, END)
#     log_box.insert(END, f"--- Smart Scheduler Bot Started at {datetime.now().strftime('%H:%M:%S')} ---\n\n")

#     if not os.path.isfile(REQUEST_FILE):
#         messagebox.showerror("Error", f"Request queue file NOT FOUND at:\n{REQUEST_FILE}")
#         return

#     try:
#         df_requests = pd.read_excel(REQUEST_FILE, sheet_name='Requests', header=None)
#         df_requests.columns = EXPECTED_COLUMNS
#         df_requests = df_requests.iloc[1:].copy()
#         df_requests.reset_index(drop=True, inplace=True)
#     except Exception as e:
#         messagebox.showerror("Error", f"Error loading Request Queue: {e}")
#         return

#     df_requests['Status'] = df_requests['Status'].astype(str).str.strip().fillna('New')
#     df_new_requests = df_requests[df_requests['Status'].astype(str).str.upper() == 'NEW'].copy()

#     if df_new_requests.empty:
#         log_box.insert(END, "✅ No new requests found. System is idle.\n")
#         return

#     log_box.insert(END, f"➡️ Found {len(df_new_requests)} new requests...\n")

#     for index, row in df_new_requests.iterrows():
#         patient = row['Patient_Name']
#         specialty = row['Requested_Specialty']
#         log_box.insert(END, f"\nProcessing request for {patient} ({specialty})...\n")

#         booked_slot, doctor = get_next_available_slot(specialty)

#         if booked_slot.startswith("No slots available"):
#             status_message = "WAITLIST: Full capacity. Added to manual review."
#         elif booked_slot.startswith("System Error"):
#             status_message = f"ERROR: {booked_slot}"
#         else:
#             status_message = f"BOOKED: {booked_slot} with Dr. {doctor}"

#         df_requests.loc[row.name, 'Status'] = status_message
#         log_box.insert(END, f"   Status: {status_message}\n")

#     df_requests.to_excel(REQUEST_FILE, sheet_name='Requests', index=False)
#     log_box.insert(END, "\n--- Scheduler Bot Finished. Request Queue updated. ---\n")
#     messagebox.showinfo("Success", "All new requests processed successfully!")


# # ---------------------------------------------------------------------
# # GUI (Tkinter)
# # ---------------------------------------------------------------------
# def launch_gui():
#     root = Tk()
#     root.title("Smart Scheduler Bot - Medical Appointment System")
#     root.geometry("750x500")
#     root.config(bg="#F4F6F7")

#     # Heading
#     Label(root, text="🩺 Smart Scheduler Bot", font=("Helvetica", 18, "bold"), bg="#F4F6F7", fg="#1A5276").pack(pady=10)
#     Label(root, text="Automated Appointment Allocation System", font=("Helvetica", 12), bg="#F4F6F7").pack()

#     # Log box
#     log_box = scrolledtext.ScrolledText(root, width=85, height=20, font=("Consolas", 10))
#     log_box.pack(padx=10, pady=20)

#     # Buttons
#     frame = Frame(root, bg="#F4F6F7")
#     frame.pack()

#     Button(frame, text="Process Requests", font=("Helvetica", 11, "bold"), bg="#27AE60", fg="white",
#            padx=20, pady=8, relief="raised", command=lambda: process_scheduling_requests(log_box)).grid(row=0, column=0, padx=10)

#     Button(frame, text="Exit", font=("Helvetica", 11, "bold"), bg="#C0392B", fg="white",
#            padx=20, pady=8, relief="raised", command=root.destroy).grid(row=0, column=1, padx=10)

#     root.mainloop()



# if __name__ == "__main__":
#     launch_gui()
import streamlit as st
import pandas as pd
import os
from datetime import datetime

# --- CONFIGURATION ---
PROJECT_FOLDER = 'C:\\Users\\PARAS\\Desktop\\PANDA_PROJECT'
REQUEST_FILE = os.path.join(PROJECT_FOLDER, 'Request_Queue.xlsx')
SCHEDULE_FILE = os.path.join(PROJECT_FOLDER, 'Doctor_Availability.xlsx')
EXPECTED_COLUMNS = ['Patient_Name', 'DOB', 'Requested_Specialty', 'Contact_Number', 'Status']

# ---------------------------------------------------------------------
# CORE SCHEDULING LOGIC
# ---------------------------------------------------------------------
def get_next_available_slot(specialty):
    if not os.path.isfile(SCHEDULE_FILE):
        return f"System Error: Schedule file not found at {SCHEDULE_FILE}", None
    try:
        df_schedule = pd.read_excel(SCHEDULE_FILE, sheet_name='Schedule')
    except ValueError:
        return "System Error: Could not find sheet named 'Schedule'.", None
    except Exception as e:
        return f"System Error loading schedule data: {e}", None

    if 'Is_Booked' in df_schedule.columns:
        df_schedule['Is_Booked'] = df_schedule['Is_Booked'].astype(str).str.upper().str.strip()

    df_filtered = df_schedule[
        (df_schedule['Specialty'].astype(str).str.lower() == specialty.lower()) &
        (df_schedule['Is_Booked'] == 'NO')
    ].copy()

    if df_filtered.empty:
        return "No slots available for this specialty.", None

    next_slot = df_filtered.iloc[0]
    schedule_index_to_update = next_slot.name
    df_schedule.loc[schedule_index_to_update, 'Is_Booked'] = 'YES'

    with pd.ExcelWriter(SCHEDULE_FILE, engine='openpyxl') as writer:
        df_schedule.to_excel(writer, sheet_name='Schedule', index=False)

    return f"{next_slot['Day']} at {next_slot['Time_Slot']}", next_slot['Doctor_Name']


# ---------------------------------------------------------------------
# PROCESSING REQUESTS (Streamlit version)
# ---------------------------------------------------------------------
def process_scheduling_requests():
    log_messages = []

    log_messages.append(f"--- Smart Scheduler Bot Started at {datetime.now().strftime('%H:%M:%S')} ---")

    if not os.path.isfile(REQUEST_FILE):
        return ["❌ Request queue file NOT FOUND at " + REQUEST_FILE]

    try:
        df_requests = pd.read_excel(REQUEST_FILE, sheet_name='Requests', header=None)
        df_requests.columns = EXPECTED_COLUMNS
        df_requests = df_requests.iloc[1:].copy()
        df_requests.reset_index(drop=True, inplace=True)
    except Exception as e:
        return [f"❌ Error loading Request Queue: {e}"]

    df_requests['Status'] = df_requests['Status'].astype(str).str.strip().fillna('New')
    df_new_requests = df_requests[df_requests['Status'].astype(str).str.upper() == 'NEW'].copy()

    if df_new_requests.empty:
        log_messages.append("✅ No new requests found. System is idle.")
        return log_messages

    log_messages.append(f"➡️ Found {len(df_new_requests)} new requests...")

    for index, row in df_new_requests.iterrows():
        patient = row['Patient_Name']
        specialty = row['Requested_Specialty']
        log_messages.append(f"\nProcessing request for {patient} ({specialty})...")

        booked_slot, doctor = get_next_available_slot(specialty)

        if booked_slot.startswith("No slots available"):
            status_message = "WAITLIST: Full capacity. Added to manual review."
        elif booked_slot.startswith("System Error"):
            status_message = f"ERROR: {booked_slot}"
        else:
            status_message = f"BOOKED: {booked_slot} with Dr. {doctor}"

        df_requests.loc[row.name, 'Status'] = status_message
        log_messages.append(f"   Status: {status_message}")

    df_requests.to_excel(REQUEST_FILE, sheet_name='Requests', index=False)
    log_messages.append("\n--- Scheduler Bot Finished. Request Queue updated. ---")

    return log_messages


# ---------------------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Smart Scheduler Bot", layout="wide")

    st.title("🩺 Smart Scheduler Bot")
    st.subheader("Automated Medical Appointment System")
    st.divider()

    if st.button("Process Requests", use_container_width=True):
        logs = process_scheduling_requests()
        st.success("Processing Complete ✅")
        for line in logs:
            st.write(line)

    st.divider()
    st.info("Developed by Paras Shah 💻 | Powered by Streamlit")


if __name__ == "__main__":
    main()
