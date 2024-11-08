import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class GoogleSheetManager:
    def __init__(self, credentials_file, scopes=None):
        """
        Initialize the GoogleSheetManager with the given credentials.

        Parameters:
        credentials_file (str): Path to the service account JSON file.
        scopes (list): List of OAuth scopes. Defaults to Sheets and Drive scopes if None.
        """
        self.scopes = scopes or [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive',
        ]
        self.credentials = Credentials.from_service_account_file(
            credentials_file, scopes=self.scopes)
        self.gc = gspread.authorize(self.credentials)

    def create_sheet(self, title):
        """
        Create a new Google Sheet or return an existing one with the same name.

        Parameters:
        title (str): The title of the sheet to create or retrieve.

        Returns:
        gspread.Spreadsheet: The created or existing spreadsheet.
        """
        try:
            existing_sheet = self.get_sheet_by_name(title)
            if existing_sheet:
                print(f"Spreadsheet with title '{title}' already exists.")
                return existing_sheet

            sheet = self.gc.create(title)
            print(f"Created new sheet: {title}")
            return sheet
        except Exception as e:
            print(f"Error creating sheet: {e}")
            return None

    def get_sheet_by_name(self, title):
        """
        Retrieve a spreadsheet by its title. Returns None if not found.
        """
        try:
            spreadsheet_list = self.gc.openall()
            for sheet in spreadsheet_list:
                if sheet.title == title:
                    return sheet
            print(f"No Sheet with this Name")
            return None
        except Exception as e:
            print(f"Error retrieving sheet by name: {e}")
            return None

    def update_sheet(self, sheet_id, data, range_):
        """
        Update a Google Sheet by sheet ID and range. The data must be provided as a list of lists.
        """
        try:
            sheet = self.gc.open_by_key(sheet_id)
            worksheet = sheet.sheet1
            worksheet.update(range_, data)
            print(f"Updated sheet {sheet_id} in range {range_}")
        except Exception as e:
            print(f"Error updating sheet: {e}")

    def update_cell(self, sheet_id, cell, value):
        """
        Update a specific cell in a Google Sheet.
        """
        try:
            sheet = self.gc.open_by_key(sheet_id)
            worksheet = sheet.sheet1
            worksheet.update(cell, value)
            print(f"Updated cell {cell} with value {value}")
            return True
        except Exception as e:
            print(f"Error updating cell: {e}")
            return False

    def update_cell_in_tab(self, sheet_id, tab_name, cell, value):
        """
        Updates a specific cell in a specific tab of a Google Sheet.

        Args:
            sheet_id (str): The ID of the Google Sheet.
            tab_name (str): The name of the tab (worksheet) within the sheet.
            cell (str): The cell reference, e.g., 'A1'.
            value: The value to be set in the cell.

        Returns:
            bool: True if the update was successful, False otherwise.
        """
        try:
            sheet = self.gc.open_by_key(sheet_id)
            worksheet = sheet.worksheet(tab_name)
            worksheet.update(value,cell)
            print(f"Updated cell {cell} in tab {tab_name} of sheet {sheet_id}")
            return True
        except Exception as e:
            print(f"Error updating cell: {e}")
            return False

    def share_sheet(self, sheet_id, email, role='reader'):
        """
        Share the Google Sheet with a user by email. Default role is 'writer', can be 'reader' or 'writer'.
        """
        try:
            drive_service = build('drive', 'v3', credentials=self.credentials)
            drive_service.permissions().create(
                fileId=sheet_id,
                body={
                    'type': 'user',
                    'role': role,
                    'emailAddress': email
                },
                fields='id',
            ).execute()
            print(f"Shared sheet {sheet_id} with {email} as {role}")
        except HttpError as error:
            print(f"An error occurred: {error}")

    def get_sheet_and_tab_by_name(self, sheet_name, tab_name):
        """
        Retrieve a Google Sheet by its name and get a specific tab (worksheet) by its name.
        """
        try:
            # Get the sheet by its name
            sheet = self.get_sheet_by_name(sheet_name)
            if not sheet:
                print(f"Spreadsheet '{sheet_name}' not found.")
                return None

            # Get the tab by its name
            worksheet = sheet.worksheet(tab_name)
            print(f"Found tab '{tab_name}' in sheet '{sheet_name}'")
            return worksheet
        except gspread.WorksheetNotFound:
            print(f"Tab '{tab_name}' not found in sheet '{sheet_name}'")
            return None
        except Exception as e:
            print(f"Error retrieving sheet and tab: {e}")
            return None

    def duplicate_sheet(self, original_sheet_id, new_sheet_name):
        """
        Duplicate an existing sheet with a different name.

        Parameters:
        original_sheet_id (str): The ID of the sheet to duplicate.
        new_sheet_name (str): The name for the new duplicated sheet.

        Returns:
        tuple: (str, bool) - The ID of the new sheet and whether it was newly created.
        """
        try:
            existing_sheet = self.get_sheet_by_name(new_sheet_name)
            if existing_sheet:
                print(f"Sheet with name '{new_sheet_name}' already exists.")
                return existing_sheet.id, False

            drive_service = build('drive', 'v3', credentials=self.credentials)
            copied_sheet = drive_service.files().copy(
                fileId=original_sheet_id,
                body={'name': new_sheet_name}
            ).execute()

            print(f"Duplicated sheet with ID '{original_sheet_id}' as '{new_sheet_name}'")
            return copied_sheet.get('id'), True

        except HttpError as error:
            print(f"An error occurred while duplicating the sheet: {error}")
            return None, False

    def get_sheet(self, sheet_id):
        """
        Get Google Sheet by its ID.
        """
        try:
            return self.gc.open_by_key(sheet_id)
        except Exception as e:
            print(f"Error retrieving sheet by ID: {e}")
            return None
