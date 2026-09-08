import logging
import smartsheet
import json
import collections


class Smartsheets:
    def __init__(self, client_id, sheet_id):
        # Initialize client
        self.client = smartsheet.Smartsheet(client_id)

        # Make sure we don't miss any error
        self.client.errors_as_exceptions(True)

        # Log all calls
        logging.basicConfig(filename='smartsheetcalls.log', level=logging.INFO)

        # Set the Sheet ID
        self.sheet_id = sheet_id

        # The API identifies columns by Id, but it's more convenient to refer to column names. Store a map here
        self.column_map = {}

        # The API identifies rows by Id, but it's more convenient to refer to worker no for rows. Store a map here
        self.row_map = {}

        # Set the key for the row mapping
        self.map_key = 'Worker No'

        self.sheet = smartsheet.models.sheet.Sheet

        column_keys = {
            'Worker No': 'worker_no',
            'Employee Name': 'name',
            'Title': 'title',
            'Reporting Office Location': 'city',
            'State': 'state',
            'Division': 'division',
            'Supervisor': 'manager',
            'Office Location Code': 'office_nbr',
            'Job Code': 'work_type',
            'Worker Status': 'status',
            'Hire Date': 'hire_date',
            'Separation Date': 'term_date',
            'Email Address': 'email',
            'Hardware - Company Cell Phone Number': 'phone',
            'Employee AD': 'ad'
        }
        self.column_keys = column_keys

        pv = {
            "072881": "Yes",
            "072967": "Yes",
            "072928": "Yes",
            "078847": "Yes",
            "806819": "Yes",
            "807058": "Yes",
            "072471": "Yes",
            "079350": "Yes",
            "072909": "Yes",
            "072496": "Yes",
            "092531": "Yes",
            "072445": "Yes",
            "072602": "Yes",
            "088665": "Yes",
            "072924": "Yes",
            "818477": "Yes",
            "862042": "Yes",
            "072878": "Yes",
            "824621": "Yes",
            "072484": "Yes",
            "072517": "Yes",
            "817296": "Yes",
            "072569": "Yes",
            "078465": "Yes",
            "072576": "Yes",
            "072978": "Yes",
            "072593": "Yes",
            "072619": "Yes",
            "073008": "Yes"
        }
        self.pv = pv

    # Helper function to find cell in a row
    def get_cell_by_column_name(self, row, column_name):
        column_id = self.column_map[column_name]
        return row.get_column(column_id)

    # Helper function to find the index of a cell in a row
    def get_cell_index(self, row, new_cell):
        for x in range(len(row.cells)):
            if row.cells[x].column_id == new_cell.column_id:
                return x
        return None

    # Function to set the column and row maps
    def set_maps(self):
        # Build column map for later reference - translates column names to column id
        for column in self.sheet.columns:
            self.column_map[column.title] = column.id

        # Build rows map for later reference - translates worker no to row id
        for row in self.sheet.rows:
            worker_no_cell = self.get_cell_by_column_name(
                row, self.map_key)
            print(worker_no_cell)
            if worker_no_cell == None:
                continue
            worker_no_value = worker_no_cell.display_value
            self.row_map[worker_no_value] = row.id

        # print(self.row_map)
        return

    # Loads the smartsheet data
    def load_sheet(self):
        # Load entire sheet
        self.sheet = self.client.Sheets.get_sheet(self.sheet_id)

        print("Loaded " + str(len(self.sheet.rows)) +
              " rows from sheet: " + self.sheet.name)

        self.set_maps()

    # Function to set the sheet ID after initilization
    def set_sheet_id(self, id):
        self.sheet_id = id
        return

    def get_smartsheet_data(self, file_name):
        # Load the data from the sheet
        self.load_sheet()

        values = {}

        for row in self.sheet.rows:
            new_row = {
                "worker_no": "",
                "name": "",
                "title": "",
                "city": "",
                "state": "",
                "division_manager": "",
                "office_nbr": "",
                "division": "",
                "manager": "",
                "supervisor": "",
                "status": "",
                "work_type": "",
                "pvi": "",
                "hire_date": "",
                "term_date": "",
                "email": "",
                "phone": "",
                "ad": ""
            }

            request_type = self.get_cell_by_column_name(row, 'Request Type')
            if request_type == None:
                print("No request type????")
                return

            if request_type.display_value != 'New Hire Provisioning':
                # print('ignoring '+request_type.display_value)
                continue

            row_key = ""
            for key in self.column_keys:
                map_key = self.column_keys[key]

                cell = self.get_cell_by_column_name(row, key)
                if cell == None:
                    continue

                if key == self.map_key:
                    row_key = str(cell.display_value)

                if cell.display_value == None and cell.value == None:
                    new_row[map_key] = ""
                elif cell.display_value == None:
                    new_row[map_key] = cell.value
                else:
                    new_row[map_key] = cell.display_value

            # Check if the employee drives a personal vehicle or not
            try:
                p = self.pv[row_key]
                new_row["pvi"] = p
            except:
                new_row["pvi"] = "No"

            # Set the division/manager field
            div = new_row["division"]
            new_row["division"] = "Division " + div
            mgr = new_row["manager"].title()
            new_row["manager"] = mgr
            new_row["division_manager"] = div + " " + mgr

            # Add the row to our values map
            values[row_key] = new_row
            print(new_row)

        sorted_values = collections.OrderedDict(sorted(values.items()))
        data = {
            "data": sorted_values
        }

        json_object = json.dumps(data, indent=4)
        with open(file_name, 'w') as outfile:
            outfile.write(json_object)
        return values
