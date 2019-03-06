import pandas as pd
from tkinter import filedialog
from .utility import mailmerge_factory


class MailProject:
    def __init__(self, project_id, project_name, date_issuance, date_maturity, coupon_rate, commercial_register_number,
                 issue_volume_min, issue_volume_max, collateral_string, clients=None):
        # Core data
        self.project_id = project_id
        self.project_name = project_name

        self.date_issuance = date_issuance
        self.date_maturity = date_maturity
        self.coupon_rate = coupon_rate
        self.commercial_register_number = commercial_register_number
        self.issue_volume_min = issue_volume_min
        self.issue_volume_max = issue_volume_max
        self.collateral_string = collateral_string

        if clients is None:
            self.clients = []

    def __repr__(self):
        # Make sure that pd.Timestamp object gets created when using this string
        return (f"MailProject({self.project_id},"
                f"'{self.project_name}',"
                f"pd.Timestamp('{self.date_issuance}'), "
                f"pd.Timestamp('{self.date_maturity}'), "
                f"{self.coupon_rate},"
                f"'{self.commercial_register_number}',"
                f"{self.issue_volume_min},"
                f"{self.issue_volume_max},"
                f"'{self.collateral_string}')")

    def __str__(self):
        return (f"Project ID ({self.project_id}): "
                f"{self.project_name}, "
                f"issuance {self.date_issuance.day}"
                f".{self.date_issuance.month}"
                f".{self.date_issuance.year}, "
                f"maturity {self.date_maturity.day}"
                f".{self.date_maturity.month}"
                f".{self.date_maturity.year}")

    def __eq__(self, other):
        # Assumption: two projects are the same if their attributes are the same.
        return vars(self) == vars(other)

    @classmethod
    def from_excel(cls, project_data_path, project_data_sheet_name, project_field_map):
        return mailmerge_factory(cls, project_data_path, project_data_sheet_name, project_field_map)

    def create_clients(self, client_data_path, client_data_sheet_name, client_field_map):
        # obtain DataFrame with only the columns of field_maps.keys()
        clients = mailmerge_factory(Client, client_data_path, client_data_sheet_name, client_field_map.keys())

        if self.clients:
            # prevent override of the clients stored in the MailProject instance.
            raise ValueError("At least one client has already been added to this project.")
        else:
            self.clients.extend(clients)


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
        # Watch out for data types. Strings are enclosed by '' (e.g., first_name), while numerics are not
        # (e.g., client_id)
        return (f"Client({self.client_id}, "
                f"'{self.first_name}', "
                f"'{self.last_name}',"
                f"'{self.address_mailing_street}', "
                f"'{self.address_mailing_zip}', "
                f"'{self.address_mailing_city}', "
                f"'{self.address_notify_street}', "
                f"'{self.address_notify_zip}', "
                f"'{self.address_notify_city}', "
                f"{self.amount}, "
                f"{self.subscription_am_authorized}, "
                f"{self.mailing_as_email}, "
                f"'{self.depot_no}', "
                f"'{self.depot_bic}')")

    def __str__(self):
        return f"Client ID ({self.client_id}):{self.first_name}, {self.last_name}"

    def __eq__(self, other):
        # Assumption: two projects are the same if their attributes are the same.
        return vars(self) == vars(other)


if __name__ == "__main__":


    path = filedialog.askopenfilename()
    project = MailProject("151", "Some Project Name")
    print(project)
    print(vars(project))


    #project.create_clients()
    #print(project.client_list)
    #print(dir(project.client_list[0]))