import os
import pandas as pd
from dotenv import load_dotenv
from google_sheets import GoogleSheetManager
from custom_functions import split_employee_data, group_by_supervisor

load_dotenv()

def main():
    creds_file = "service_account.json"
    g_sheet = GoogleSheetManager(creds_file)

    # Manage mapping
    mapping = g_sheet.get_sheet_and_tab_by_name('GH Sample Data', os.getenv('MAPPING_SHEET'))
    mapping_df = pd.DataFrame(mapping.get_all_records())

    # Manage employee data
    employees_df = split_employee_data(mapping_df, 'Direct Reports')

    # Manage supervisor data
    supervisor_data_df = group_by_supervisor(employees_df, 'Supervisor Name', 'Supervisor Email')

    # Employee-wise sheet creation
    create_employee_sheets(g_sheet, employees_df)

    # Supervisor-wise sheet creation
    create_supervisor_sheets(g_sheet, supervisor_data_df)

def create_employee_sheets(g_sheet, employees_df):
    for _, entry in employees_df.iterrows():
        print(f"Processing: {entry['Employee Name']} ({entry['Employee Email']})")
        ret_sheet, sheet_exists = g_sheet.duplicate_sheet(os.getenv('EMPLOYEE_TEMPLATE'),
                                                          entry['Employee Name'])
        if ret_sheet:
            g_sheet.update_cell_in_tab(ret_sheet, 'Orders', 'B1', [[entry['Employee Email']]])
            if sheet_exists:
                g_sheet.share_sheet(ret_sheet, entry['Employee Email'], 'writer')

def create_supervisor_sheets(g_sheet, supervisor_data_df):
    for (supervisor_name, supervisor_email), group in supervisor_data_df:
        print(f"Processing Supervisor: {supervisor_name} ({supervisor_email})")
        sheet_name = supervisor_name if supervisor_name else supervisor_email
        ret_sheet, sheet_exists = g_sheet.duplicate_sheet(os.getenv('SUPERVISOR_TEMPLATE'), sheet_name)
        if ret_sheet:
            g_sheet.update_cell_in_tab(ret_sheet, 'Orders', 'B1',
                                       [group['Employee Email'].tolist()])
            if sheet_exists:
                g_sheet.share_sheet(ret_sheet, supervisor_email, 'writer')

if __name__ == "__main__":
    main()
