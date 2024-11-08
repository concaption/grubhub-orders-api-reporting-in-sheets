import pandas as pd


def split_employee_data(df, column_name):
    """
    Split the employee name and email from the given column in the DataFrame.

    Parameters:
    df (pd.DataFrame): DataFrame containing employee data in a single column.
    column_name (str): The name of the column that contains the employee data.

    Returns:
    pd.DataFrame: DataFrame with columns 'Employee Name', 'Employee Email', 'Supervisor Name', and 'Supervisor Email'.
    """
    # Lists to store employee and supervisor data
    employee_data = []

    for _, entry in df.iterrows():
        pairs = entry[column_name].split(',')
        for pair in pairs:
            name, email = pair.split('(')
            employee_data.append({
                'Employee Name': name.strip(),
                'Employee Email': email.replace(')', '').strip(),
                'Supervisor Name': entry['Supervisor Name'],
                'Supervisor Email': entry['Supervisor Email']
            })

    return pd.DataFrame(employee_data)


def group_by_supervisor(df, supervisor_name_col, supervisor_email_col):
    """
    Group records based on supervisor name and email address.

    Parameters:
    df (pd.DataFrame): DataFrame containing supervisor data.
    supervisor_name_col (str): The name of the column with supervisor names.
    supervisor_email_col (str): The name of the column with supervisor emails.

    Returns:
    pd.DataFrameGroupBy: A DataFrameGroupBy object grouped by supervisor name and email.
    """
    return df.groupby([supervisor_name_col, supervisor_email_col])
