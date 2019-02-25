import pandas as pd
from tkinter import filedialog


class MailProject:
    def __init__(self, client_data_path, client_data_sheet_name, client_data_fields,
                 project_data_path, project_data_sheet_name, project_data_fields):

        self.client_data = self.parse_exel(client_data_path, client_data_sheet_name, client_data_fields)
        self.project_data = self.parse_exel(project_data_path, project_data_sheet_name, project_data_fields)
        self.client_list = []

    @staticmethod
    def parse_exel(filepath, sheet_name, field_list):
        df = pd.read_excel(filepath, sheet_name)[field_list]
        return df

    def create_clients(self):
        pass


class DbcMailProject(MailProject):
    def __init__(self, client_data_path, client_data_sheet_name, client_data_fields,
                 project_data_path, project_data_sheet_name, project_data_fields):

        super().__init__(client_data_path, client_data_sheet_name, client_data_fields,
                         project_data_path, project_data_sheet_name, project_data_fields)


class Client:
    def __init__(self, attribute_dict):
        for attribute_name, attribute_value in attribute_dict.items():
            setattr(self, attribute_name, attribute_value)





#project = MailProject()
#print(project.client_data)

c_data = {"first_name": "Hans", "last_name": "Schmidt"}

client = Client(c_data)
print(client.__dict__.keys())
print(client.first_name)


path = filedialog.askopenfilename()
project = MailProject(path, 'Zeichnerliste', ['NAME_TITEL_VNAME', "NAME_NNAME"], path, 'Zeichnerliste', ['something'])
print(project.client_data)
