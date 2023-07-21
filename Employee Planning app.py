import PySimpleGUI as sg
import pandas as pd
original_df = None
sg.theme('DarkTeal10')
# Import function
def import_file():
    global original_df, import_file_path
    layout = [
        [sg.Text('Import Excel File')],
        [sg.Input(key='import_file'), sg.FileBrowse()],
        [sg.Button('Import')]
    ]

    window = sg.Window('Import File', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'Import':
            import_file_path = values['import_file']
            try:
                # Perform the import and necessary operations on the DataFrame
                df = pd.read_excel(import_file_path)
                # Dropping the row where mission is nan
                df = df.dropna(subset=['Mission '], axis=0)
                pd.set_option('mode.chained_assignment', None)
                # Replacing nan, '\xa0', '      -  ', and '4' with 0
                columns_to_replace = df.columns[2:]
                df.loc[:, columns_to_replace] = df.loc[:, columns_to_replace].fillna(0)
                df.loc[:, columns_to_replace] = df.loc[:, columns_to_replace].replace('\xa0', 0)
                df.loc[:, columns_to_replace] = df.loc[:, columns_to_replace].replace('      -  ', 0)
                df.loc[:, columns_to_replace] = df.loc[:, columns_to_replace].replace('4', 0)
                # Save a copy of the imported DataFrame
                original_df = df.copy()
                window.close()
                # Call the function to move to the second screen passing the cleaned DataFrame
                main_screen(df)
            except Exception as e:
                sg.popup(f'Error: {e}')

    window.close()

# Second screen: Options
def main_screen(df):
    layout = [
        [sg.Text('Choose an option:')],
        [sg.Button('Check Availability')],
        [sg.Button('Insert & Modify')],
        [sg.Button('Delete Row')],
        [sg.Button('Remove Mission')],
        [sg.Button('Undo Changes')],
        [sg.Button('Exit')]
        ]

    window = sg.Window('Main Screen', layout)

    while True:
        event, _ = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break
        elif event == 'Check Availability':
            # Call the visualize function passing the DataFrame
            visualize(df)
        elif event == 'Insert & Modify':
            # Call the insert_modify function passing the DataFrame
            insert_modify(df)
        elif event == 'Remove Mission':
            undo_mission(df)
        elif event == 'Undo Changes':
            df = undo_changes(df)
        elif event == 'Delete Row':
            # Call the delete function passing the DataFrame
            delete(df)

    window.close()

# Visualization function
def visualize(df):
    layout = [
        [sg.Text("Enter the starting date: "), sg.CalendarButton("Choose Date", target="-CAL-START-", key="-CAL-START-", format="%d-%m-%Y")],
        [sg.Text("Enter the ending date: "), sg.CalendarButton("Choose Date", target="-CAL-END-", key="-CAL-END-", format="%d-%m-%Y")],
        [sg.Button("Submit"), sg.Button("Check Missions"), sg.Button("Cancel")],
        [
            sg.Column([
                [sg.Text("Available Employees:", font=("Helvetica", 12, "bold"))],
                [sg.Listbox(values=[], size=(30, 10), key="-AVAILABLE-")],
            ]),
            sg.Column([
                [sg.Text("Non-Available Employees:", font=("Helvetica", 12, "bold"))],
                [sg.Listbox(values=[], size=(30, 10), key="-NON_AVAILABLE-")],
            ]),
        ],
        [sg.Table(values=[], headings=["Employee", "Mission"], key="-MISSIONS-", auto_size_columns=True, num_rows=10, justification="left")]
    ]

    window = sg.Window("Employee Availability", layout)

    # Event loop to process GUI events
    while True:
        event, values = window.read()

        if event == "-CAL_START-":
            window["-CAL-START-"].update(values["-CAL_START-"].strftime("%d-%m-%Y"))

        elif event == "-CAL_END-":
            window["-CAL-END-"].update(values["-CAL_END-"].strftime("%d-%m-%Y"))

        elif event == "Submit":
            starting_date_input = values["-CAL-START-"]
            ending_date_input = values["-CAL-END-"]

            # Convert input dates to datetime objects
            starting_date = pd.to_datetime(starting_date_input, dayfirst=True)
            ending_date = pd.to_datetime(ending_date_input, dayfirst=True)

            # Filter the DataFrame based on the availability of employees
            available_employees = []
            non_available_employees = []  # New list for non-available employees
            for index, row in df.iterrows():
                employee_name = row['Employee Name']
                employee_data = row[starting_date:ending_date].values
                if not (employee_data == 8).any() and not pd.isnull(employee_data).any():
                    available_employees.append(employee_name)
                else:
                    non_available_employees.append(employee_name)  # Add non-available employee to the list

            # Update the ListBox in the GUI window
            window["-AVAILABLE-"].update(values=available_employees)
            window["-NON_AVAILABLE-"].update(values=non_available_employees)

        elif event == "Check Missions":
            missions_data = []
            for employee in non_available_employees:
                employee_row = df[df['Employee Name'] == employee]
                mission = employee_row['Mission '].values[0] if len(employee_row) > 0 else None
                missions_data.append([employee, mission])

            window["-MISSIONS-"].update(values=missions_data)

        elif event == "Cancel" or event == sg.WINDOW_CLOSED:
            break

    window.close()

# Insert & Modify function
def insert_modify(df):
    def data_entry(df):
        layout = [
            [sg.Text('Enter mission number:'), sg.Input(key='mission')],
            [sg.Text('Enter employee number:'), sg.Input(key='employee')],
            [sg.Text('Choose Start Date:'), sg.CalendarButton("Choose Date", target='-CAL-START-', key='-CAL-START-', format="%d-%m-%Y")],
            [sg.Text('Choose End Date:'), sg.CalendarButton("Choose Date", target='-CAL-END-', key='-CAL-END-', format="%d-%m-%Y")],
            [sg.Button('Submit'), sg.Button('Cancel')],
        ]

        window = sg.Window('Data Entry', layout)

        while True:
            event, values = window.read()

            if event == sg.WINDOW_CLOSED or event == 'Cancel':
                break

            if event == '-CAL-START-':
                window['-CAL-START-'].update(values['-CAL-START-'].strftime("%d-%m-%Y"))

            if event == '-CAL-END-':
                window['-CAL-END-'].update(values['-CAL-END-'].strftime("%d-%m-%Y"))

            if event == 'Submit':
                mission_number = values['mission']
                employee_number = values['employee'].encode('utf-8').decode('utf-8')

                mission = f"Mission {mission_number}"
                employee = f"Employé {employee_number}"

                start_date_input = values['-CAL-START-']
                end_date_input = values['-CAL-END-']

                start_date = pd.to_datetime(start_date_input, dayfirst=True)
                end_date = pd.to_datetime(end_date_input, dayfirst=True)

                date_ranges = [(mission, employee, start_date, end_date)]

                for (mission, employee, start_date, end_date) in date_ranges:
                    filter_condition = (df['Mission '] == mission) & (df['Employee Name'] == employee)

                    if filter_condition.any():
                        date_range = pd.date_range(start_date, end_date, freq='D').strftime('%-d/%-m').tolist()
                        columns_to_update = df.columns[df.columns.astype(str).isin(date_range)]

                        df.loc[filter_condition, columns_to_update] = 8
                    else:
                        new_row = [mission, employee] + [0] * (len(df.columns) - 2)
                        df.loc[df.index.max() + 1] = new_row

                        date_range = pd.date_range(start_date, end_date, freq='D').strftime('%-d/%-m').tolist()
                        columns_to_update = df.columns[df.columns.astype(str).isin(date_range)]

                        columns_to_zero = df.columns.difference(['Mission ', 'Employee Name'])
                        df.loc[df.index.max(), columns_to_zero] = 0
                        df.loc[df.index.max(), columns_to_update] = 8

                sg.popup('Data entry successful!', title='Success')

                window.close()

                return df

        window.close()

        return df

    save_option(data_entry(df))

# Delete function
def delete(df):
    def delete_row(df):
        mission_number = values['mission_number']
        employee_number = values['employee_number']

        mission = f"Mission {mission_number}"
        employee = f"Employé {employee_number}"

        # Delete the row where both mission and employee match the input values
        df = df.drop(df[(df['Mission '] == mission) & (df['Employee Name'] == employee)].index)
        sg.popup('Success!', 'Row has been deleted')
        return df

    # Create the GUI layout for delete function
    layout = [
        [sg.Text('Enter mission number: '), sg.Input(key='mission_number')],
        [sg.Text('Enter employee number: '), sg.Input(key='employee_number')],
        [sg.Button('Delete'),sg.Button("Cancel")]
    ]

    # Create the window for delete function
    window = sg.Window('Delete Row', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'Delete':
            # Call the delete_row function with the provided inputs
            df = delete_row(df)
        elif event == "Cancel" or event == sg.WINDOW_CLOSED:
            break

    window.close()

    save_option(df)
# Undo Mission function
def undo_mission(df):
    layout = [
        [sg.Text('Enter mission number: '), sg.Input(key='mission_number')],
        [sg.Text('Enter employee number: '), sg.Input(key='employee_number')],
        [sg.Text('Choose start date: '), sg.CalendarButton("Choose Date", target='-CAL-START-', key='-CAL-START-', format="%d-%m-%Y")],
        [sg.Text('Choose end date: '), sg.CalendarButton("Choose Date", target='-CAL-END-', key='-CAL-END-', format="%d-%m-%Y")],
        [sg.Button('Submit'), sg.Button('Cancel')]
    ]

    window = sg.Window('Undo Mission', layout)

    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == 'Cancel':
            break

        mission_number = values['mission_number']
        employee_number = values['employee_number']
        start_date_input = values['-CAL-START-']
        end_date_input = values['-CAL-END-']

        # Convert mission number and employee number to appropriate format
        mission = f"Mission {mission_number}"
        employee = f"Employé {employee_number}"

        # Convert start and end dates to appropriate format
        start_date = pd.to_datetime(start_date_input, format="%d-%m-%Y")
        end_date = pd.to_datetime(end_date_input, format="%d-%m-%Y")

        # Check if the mission and employee combination exists in the DataFrame
        mask = (df['Mission '] == mission) & (df['Employee Name'] == employee)
        if not any(mask):
            sg.popup("Mission and employee combination not found in the DataFrame.")
        else:
            # Update the values in the DataFrame
            df.loc[mask, start_date:end_date] = 0
            sg.popup('Success!', 'Mission data updated.')

        save_option(df)

    window.close()

# Save option
def save_option(df):
    layout = [
        [sg.Text('Save Excel File')],
        [sg.Input(key='save_file'), sg.FileSaveAs(file_types=(("Excel Files", "*.xlsx"),))],
        [sg.Button('Export'), sg.Button('Save')],
        [sg.Button('Skip')]
    ]

    window = sg.Window('Save Option', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Skip':
            break
        elif event == 'Export':
            save_file_path = values['save_file']
            try:
                df.to_excel(save_file_path + '.xlsx', index=False)
                sg.popup('File saved successfully.')
            except Exception as e:
                sg.popup(f'Error: {e}')
        elif event == 'Save':
            global original_df
            try:
                original_df.to_excel(import_file_path, index=False)
                sg.popup('File saved successfully.')
            except Exception as e:
                sg.popup(f'Error: {e}')

    window.close()
# Undo Option
def undo_changes(df):
    global original_df
    df = original_df.copy()
    sg.popup('Undo successful!', title='Success')
    return df

# Run the program
if __name__ == '__main__':
    import_file()
