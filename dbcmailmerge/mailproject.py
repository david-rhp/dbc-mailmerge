import pandas as pd
from tkinter import filedialog
from .utility import mailmerge_factory, translate_dict, create_folder_hierarchy
from pathlib import Path
from mailmerge import MailMerge
from dbcmailmerge.fieldmap import FIELD_MAP_CLIENTS, FIELD_MAP_PROJECT


class MailProject:
    TOP_LEVEL_DIR = "client_correspondence"  # name of directory where the created documents should be stored
    DOCUMENT_TYPES = ["offer_documents", "appropriateness_test"]
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
                f"'{self.date_issuance}', "
                f"'{self.date_maturity}', "
                f"{self.coupon_rate}, "
                f"'{self.commercial_register_number}', "
                f"{self.issue_volume_min}, "
                f"{self.issue_volume_max}, "
                f"'{self.collateral_string}')")

    def __str__(self):
        return (f"Project ID ({self.project_id}): "
                f"{self.project_name}, "
                f"issuance {self.date_issuance}, "
                f"maturity {self.date_maturity}")

    def __eq__(self, other):
        # Assumption: two projects are the same if their attributes are the same.
        return vars(self) == vars(other)

    @classmethod
    def from_excel(cls, project_data_path, project_data_sheet_name, project_field_map):
        return mailmerge_factory(cls, project_data_path, project_data_sheet_name, project_field_map)

    def create_clients(self, client_data_path, client_data_sheet_name, client_field_map):
        # obtain DataFrame with only the columns of field_maps.keys()
        clients = mailmerge_factory(Client, client_data_path, client_data_sheet_name, client_field_map)

        if self.clients:
            # prevent override of the clients stored in the MailProject instance.
            raise ValueError("At least one client has already been added to this project.")
        else:
            self.clients.extend(clients)

    def select_clients(self, selection_criteria):
        # select only relevant clients
        selected_clients = []
        for client in self.clients:
            selected = True

            for criterion in selection_criteria.keys():
                if not selection_criteria[criterion](getattr(client, criterion)):
                    selected = False

            if selected:
                selected_clients.append(client)

        return selected_clients

    def create_client_documents(self, selected_clients):
        advisors = set()
        merge_records = []
        for client in selected_clients:
            # Apply formatting rules to title and street
            client_record = vars(client)

            # add client advisor for creation of sub_directories
            advisors.add(client_record["advisor"])

            # The MailMerge.merge method from docx-mailmerge strips the whitespace around each mergefield.
            # In the cases where a client has a title, the format is, for example, <title><first_name> <last_name>
            # This means that if there is no title, title should be replaced by an empty string, if there is, title
            # should have a trailing space in order to prevent title being 'together' with first_name, such as
            # Dr.Jane Doe => Dr. Jane Doe
            # Therefore, format the first_name field and salutation field (where the same as above occurs).
            if client_record["title"]:
                client_record["first_name"] = client_record["title"] + ' ' + client_record["first_name"]
                client_record["salutation"] += ' ' + client_record["title"]
                client_record["title"] = ''

            client_record["address_mailing_street"] += '\n'
            client_record["amount"] = format(client_record["amount"], ",.2f")

            # translate client attributes to match excel version, thus, matching the word mergefield placeholders
            client_record = translate_dict(client_record, FIELD_MAP_CLIENTS, reverse=True)
            client_record = {key: str(value) for key, value in client_record.items()}  # cast to str for MailMerge

            merge_records.append(client_record)





        # translate project data so that the fields (keys) match the names in the word template
        project_data = vars(self)
        del project_data["clients"]  # not part of the fields in the template
        project_data = translate_dict(project_data, FIELD_MAP_PROJECT, reverse=True)
        project_data = {key: str(value) for key, value in project_data.items()}  # convert value to str for MailMerge

        template_path = "../data/templates/cover_letter.docx"
        hierarchy_root = Path('../tests/')
        doc_type = "offer_documents"

        # create folder hierarchy for the storage of the created documents
        sub_directories = [list(advisors), type(self).DOCUMENT_TYPES]
        create_folder_hierarchy(hierarchy_root, type(self).TOP_LEVEL_DIR, sub_directories)

        reversed_field_map_clients = {value: key for key, value in FIELD_MAP_CLIENTS.items()}

        for client_record in merge_records:
            # add project data
            client_record.update(project_data)

            # copy word template and replace placeholders with client instance data and project data
            with MailMerge(template_path) as document:
                document.merge(**client_record)

                out_path = (hierarchy_root
                            / type(self).TOP_LEVEL_DIR
                            / client_record[reversed_field_map_clients["advisor"]]
                            / doc_type)

                filename = (client_record[reversed_field_map_clients["last_name"]]
                            + '_'
                            + client_record[reversed_field_map_clients["first_name"]]
                            + '_'
                            + client_record[reversed_field_map_clients["client_id"]])

                # save document in folder hierarchy
                document.write(out_path / (filename + ".docx"))


class Client:
    def __init__(self, client_id, advisor, title, first_name, last_name, salutation_address_field, salutation,
                 address_mailing_street, address_mailing_zip, address_mailing_city,
                 address_notify_street, address_notify_zip, address_notify_city,
                 amount, subscription_am_authorized, mailing_as_email, depot_no, depot_bic):
        # Core data - client specific
        self.client_id = client_id

        self.advisor = advisor
        self.title = title
        self.first_name = first_name
        self.last_name = last_name
        self.salutation_address_field = salutation_address_field
        self.salutation = salutation

        # Addresses for mailing - client specific
        self.address_mailing_street = address_mailing_street
        self.address_mailing_zip = str(address_mailing_zip)
        self.address_mailing_city = address_mailing_city
        self.address_notify_street = address_notify_street
        self.address_notify_zip = str(address_notify_zip)
        self.address_notify_city = address_notify_city

        # Data pertaining to an individual bond subscription
        # always an int or empty >= 0 in 5k increments, excel sometimes returns floats
        self.amount = int(amount) if amount else amount
        self.subscription_am_authorized = subscription_am_authorized
        self.mailing_as_email = mailing_as_email
        self.depot_no = str(depot_no)
        self.depot_bic = str(depot_bic)

    def __repr__(self):
        # Watch out for data types. Strings are enclosed by '' (e.g., first_name), while numerics are not
        # (e.g., client_id)
        return (f"Client({self.client_id}, "
                f"'{self.advisor}', "
                f"'{self.title}', "
                f"'{self.first_name}', "
                f"'{self.last_name}', "
                f"'{self.salutation_address_field}', "
                f"'{self.salutation}', "
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
        return f"Client ID ({self.client_id}): {self.first_name}, {self.last_name}"

    def __eq__(self, other):
        # Assumption: two projects are the same if their attributes are the same.
        return vars(self) == vars(other)

    def create_doc_from_word_template(self, template_path):
        advisors = set()
        merge_records = []
        for client in selected_clients:
            # Apply formatting rules to title and street
            client_record = vars(client)

            # add client advisor for creation of sub_directories
            advisors.add(client_record["advisor"])

            # The MailMerge.merge method from docx-mailmerge strips the whitespace around each mergefield.
            # In the cases where a client has a title, the format is, for example, <title><first_name> <last_name>
            # This means that if there is no title, title should be replaced by an empty string, if there is, title
            # should have a trailing space in order to prevent title being 'together' with first_name, such as
            # Dr.Jane Doe => Dr. Jane Doe
            # Therefore, format the first_name field and salutation field (where the same as above occurs).
            if client_record["title"]:
                client_record["first_name"] = client_record["title"] + ' ' + client_record["first_name"]
                client_record["salutation"] += ' ' + client_record["title"]
                client_record["title"] = ''

            client_record["address_mailing_street"] += '\n'
            client_record["amount"] = format(client_record["amount"], ",.2f")

            # translate client attributes to match excel version, thus, matching the word mergefield placeholders
            client_record = translate_dict(client_record, FIELD_MAP_CLIENTS, reverse=True)
            client_record = {key: str(value) for key, value in client_record.items()}  # cast to str for MailMerge

            merge_records.append(client_record)

        return merge_records, advisors

if __name__ == "__main__":


    path = filedialog.askopenfilename()
    project = MailProject("151", "Some Project Name")
    print(project)
    print(vars(project))


    #project.create_clients()
    #print(project.client_list)
    #print(dir(project.client_list[0]))