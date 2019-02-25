import pandas as pd
from tkinter import filedialog
# Keys are the column headers in the data source,
# values are the standardized way the respective data source is represented
FIELD_MAP_CLIENTS = {"client_id"}
FIELD_MAP_PROJECT = {"project_id": "project_id", "project_name": "project_name"}


class MailProject:
    def __init__(self, project_id, project_name, client_list=None):

        self.project_id = project_id
        self.project_name = project_name
        if client_list is None:
            self.client_list = []

    @classmethod
    def from_excel(cls, project_data_path, project_data_sheet_name, field_map):
        project_data = MailProject._parse_excel(project_data_path, project_data_sheet_name, field_map.keys())

        # Extract one or more projects from the data source and instantiate one or more instances of MailProject
        projects = []
        for _, project in project_data.iterrows():
            project = project.to_dict()
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
        # TODO update this, doesn' work currently because class definitions of MailProject and Client have changed
        for _, row in self.client_data.iterrows():
            # Convert each row to a dict, each dict representing one client.
            # Use this dict as argument for instantiating Client instances.
            self.client_list.append(Client(*row.to_dict()))

    def __repr__(self):
        return f"MailProject({self.project_id}, '{self.project_name}')"


class Client:
    def __init__(self, client_id, first_name, last_name, address_mailing, address_notify):
        self.client_id = client_id
        self.first_name = first_name
        self.last_name = last_name
        self.address_mailing = address_mailing
        self.address_notify = address_notify

    def __repr__(self):
        return (f"Client({self.client_id}, '{self.first_name}', '{self.last_name}',"
                f"'{self.address_mailing}', '{self.address_notify}')")

    def __str__(self):
        return f"Client ID ({self.client_id}):{self.first_name}, {self.last_name}"


class DbcClient(Client):
    def __init__(self, client_id, first_name, last_name, address_mailing, address_notify,
                 amount, subscription_am_authorized, mailing_as_email, depot):

        super().__init__(client_id, first_name, last_name, address_mailing, address_notify)
        self.amount = amount
        self.subscription_am_authorized = subscription_am_authorized
        self.mailing_as_email = mailing_as_email
        self.depot = depot  # as tuple depot[0]


#project = MailProject()
#print(project.client_data)

c_data = {"first_name": "Hans", "last_name": "Schmidt"}

# client = Client(c_data)
# print(client.__dict__.keys())
# print(client.first_name)


path = filedialog.askopenfilename()
project = MailProject("151", "Some Project Name")
print(project)

project2 = MailProject.from_excel(path, "project_data", FIELD_MAP_PROJECT)
print(project2)



#project.create_clients()
#print(project.client_list)
#print(dir(project.client_list[0]))