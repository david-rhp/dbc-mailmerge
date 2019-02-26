import pandas as pd
from tkinter import filedialog
# Keys are the column headers in the data source, keys and data source have to match or a key error will be raised
# values are the standardized way the respective data source is represented
FIELD_MAP_CLIENTS = {"client_id"}
FIELD_MAP_PROJECT = {"project_id": "project_id", "project_name": "project_name", "date_issuance": "date_issuance",
                     "date_maturity": "date_maturity", "coupon_rate": "coupon_rate",
                     "commercial_register_number": "commercial_register_number", "issue_volume_min": "issue_volume_min",
                     "issue_volume_max": "issue_volume_max", "collateral_string": "collateral_string"}


class MailProject:
    def __init__(self, project_id, project_name, date_issuance, date_maturity, coupon_rate, commercial_register_number,
                 issue_volume_min, issue_volume_max, collateral_string, client_list=None, ):
        self.project_id = project_id
        self.project_name = project_name
        self.date_issuance = date_issuance
        self.date_maturity = date_maturity
        self.coupon_rate = coupon_rate
        self.commercial_register_number = commercial_register_number
        self.issue_volume_min = issue_volume_min
        self.issue_volume_max = issue_volume_max
        self.collateral_string = collateral_string

        if client_list is None:
            self.client_list = []

    def __repr__(self):
        # Make sure that pd.Timestamp object gets created when using this string
        return (f"MailProject({self.project_id}, '{self.project_name}', pd.Timestamp('{self.date_issuance}'), "
                f"pd.Timestamp('{self.date_maturity}'), {self.coupon_rate}, '{self.commercial_register_number}',"
                f"{self.issue_volume_min}, {self.issue_volume_max}, '{self.collateral_string}')")

    def __eq__(self, other):
        # Assumption: two projects are the same if their attributes are the same.
        return vars(self) == vars(other)

    @classmethod
    def from_excel(cls, project_data_path, project_data_sheet_name, field_map):
        project_data = MailProject._parse_excel(project_data_path, project_data_sheet_name, field_map.keys())

        # Extract one or more projects from the data source and instantiate one or more instances of MailProject
        projects = []
        for _, project in project_data.iterrows():
            project = project.to_dict()

            # Use field map to translate excel column headers (keys) to expected values
            for key in project.copy():
                if field_map[key] not in project:
                    project[field_map[key]] = project[key]
                    del project[key]

            projects.append(cls(**project))

        if len(project_data) > 1:
            return projects
        else:
            return projects[0]

    @staticmethod
    def _parse_excel(filepath, sheet_name, field_list):
        # Extract only relevant fields: all fields in field_list
        df = pd.read_excel(filepath, sheet_name)[field_list]
        return df

    def create_clients(self):
        # TODO update this, doesn't work currently because class definitions of MailProject and Client have changed
        for _, row in self.client_data.iterrows():
            # Convert each row to a dict, each dict representing one client.
            # Use this dict as argument for instantiating Client instances.
            self.client_list.append(Client(*row.to_dict()))


class Client:
    def __init__(self, client_id, first_name, last_name,
                 address_mailing_street, address_mailing_zip, address_mailing_city,
                 address_notify_street, address_notify_zip, address_notify_city,
                 amount, subscription_am_authorized, mailing_as_email, depot_no, depot_bic):
        # Core data - client specific
        self.client_id = client_id
        self.first_name = first_name
        self.last_name = last_name

        # Addresses for mailing - client specific
        self.address_mailing_street = address_mailing_street
        self.address_mailing_zip = address_mailing_zip
        self.address_mailing_city = address_mailing_city
        self.address_notify_street = address_notify_street
        self.address_notify_zip = address_notify_zip
        self.address_notify_city = address_notify_city

        # Data pertaining to an individual bond subscription
        self.amount = amount
        self.subscription_am_authorized = subscription_am_authorized
        self.mailing_as_email = mailing_as_email
        self.depot_no = depot_no
        self.depot_bic = depot_bic

    def __repr__(self):
        return (f"Client({self.client_id}, '{self.first_name}', '{self.last_name}',"
                f"'{self.address_mailing}', '{self.address_notify}')")

    def __str__(self):
        return f"Client ID ({self.client_id}):{self.first_name}, {self.last_name}"


if __name__ == "__main__":

    #project = MailProject()
    #print(project.client_data)

    c_data = {"first_name": "Hans", "last_name": "Schmidt"}

    # client = Client(c_data)
    # print(client.__dict__.keys())
    # print(client.first_name)


    path = filedialog.askopenfilename()
    project = MailProject("151", "Some Project Name")
    print(project)
    print(vars(project))

    project2 = MailProject.from_excel(path, "project_data_single", FIELD_MAP_PROJECT)
    print(project2)



    #project.create_clients()
    #print(project.client_list)
    #print(dir(project.client_list[0]))