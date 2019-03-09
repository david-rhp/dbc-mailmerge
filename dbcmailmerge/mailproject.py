"""
Contains the main classes and business logic for creating a MailProject and corresponding Clients.

In its current state, the business logic is tied into the classes and should be factored out in a future release
to make the classes more maintainable and extendable.
"""
# TODO factor out business logic of classes

import os
from mailmerge import MailMerge
from PyPDF2 import PdfFileMerger
from dbcmailmerge.config import (FIELD_MAP_CLIENTS, FIELD_MAP_CLIENTS_REVERSED, FIELD_MAP_PROJECT,
                                 TEMPLATES, INCLUDE_STANDARDS)
from .utility import mailmerge_factory, translate_dict, create_folder_hierarchy, parse_excel
from .docx2pdfconverter import convert_to


class MailProject:
    TOP_LEVEL_DIR = "client_correspondence"  # name of directory where the created documents should be stored
    DOCUMENT_TYPES = ["offer_documents", "appropriateness_test"]

    def __init__(self, project_id, project_name, date_issuance, date_maturity, coupon_rate, commercial_register_number,
                 issue_volume_min, issue_volume_max, collateral_string, client_records=None):
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

        if client_records is None:
            self.client_records = []

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

    def create_client_records(self, client_data_path, client_data_sheet_name, client_field_map):
        # obtain DataFrame with only the columns of field_maps.keys()
        df = parse_excel(client_data_path, client_data_sheet_name, client_field_map)

        client_records = []
        for _, record in df.iterrows():
            # Convert pandas series to dict for translation of excel column names to the version used internally
            record = record.to_dict()
            record = translate_dict(record, client_field_map)

            client_records.append(record)

        if self.client_records:
            # prevent override of the client_records stored in the MailProject instance.
            raise ValueError("At least one client has already been added to this project.")
        else:
            self.client_records = client_records

        self.cast_client_records(True)

    def cast_client_records(self, silent=False):
        CONVERSION_MAP = {"client_id": str, "amount": int, "depot_no": str, "depot_bic": str, "address_mailing_zip": str,
                          "address_notify_zip": str}

        for record in self.client_records:
            for key in CONVERSION_MAP:
                conversion_function = CONVERSION_MAP[key]
                try:
                    record[key] = conversion_function(record[key])
                except ValueError as err:
                    # Conversion failed
                    if not silent:
                        raise ValueError(err,
                                         "The conversion didn't work. Probably due to the value in the data source"
                                         "having an incompatible type with the conversion function in the "
                                         "CONVERSION MAP")
                    else:
                        # don't change the value of the current record.
                        pass

    def select_clients(self, selection_criteria):
        # select only relevant client_records
        selected_clients = []
        for client in self.client_records:
            selected = True

            for criterion in selection_criteria.keys():
                if not selection_criteria[criterion](client[criterion]):
                    selected = False

            if selected:
                selected_clients.append(client)

        return selected_clients

    def create_project_record(self):
        # translate project data so that the fields (keys) match the names in the word template
        project_record = vars(self)
        del project_record["client_records"]  # not part of the fields in the template

        # convert coupon rate from decimal to percentage and use German comma
        project_record["coupon_rate"] = format(float(project_record["coupon_rate"]) * 100, ".2f").replace('.', ',')
        project_record = translate_dict(project_record, FIELD_MAP_PROJECT, reverse=True)
        project_record = {key: str(value) for key, value in project_record.items()}  # cast to str for MailMerge

        return project_record

    def create_client_documents(self, selected_clients, hierarchy_root, standard_pdfs):
        project_record = self.create_project_record()

        advisors = set()
        merge_records = []
        for client_record in selected_clients:
            # add client advisor for creation of sub_directories
            advisors.add(client_record["advisor"])

            # translate client to match placeholders in word
            client_record = translate_dict(client_record, FIELD_MAP_CLIENTS, reverse=True)
            # add project data
            client_record.update(project_record)

            merge_records.append(client_record)

        # create folder hierarchy for the storage of the created documents
        sub_directories = [list(advisors), type(self).DOCUMENT_TYPES]
        create_folder_hierarchy(hierarchy_root, type(self).TOP_LEVEL_DIR, sub_directories)

        for client_record in merge_records:
            self.create_client_document(client_record, standard_pdfs, hierarchy_root)

    def create_client_document(self, client_record, standard_pdfs, hierarchy_root):
        # copy word template and replace placeholders with client instance data and project data
        for doc_type in TEMPLATES.keys():
            created_documents_paths = []
            for template_path in TEMPLATES[doc_type]:
                with MailMerge(template_path) as document:
                    document.merge(**client_record)

                    out_path = (hierarchy_root
                                / type(self).TOP_LEVEL_DIR
                                / client_record[FIELD_MAP_CLIENTS_REVERSED["advisor"]]
                                / doc_type)

                    # template_name = template_path.split('/')[-1].replace(".docx", '')
                    template_name = template_path.parts[-1].replace(".docx", '')

                    filename = ("Nr._"
                                + str(self.project_id)
                                + '_'
                                + client_record[FIELD_MAP_CLIENTS_REVERSED["last_name"]]
                                + '_'
                                + client_record[FIELD_MAP_CLIENTS_REVERSED["first_name"]]
                                + '_'
                                + client_record[FIELD_MAP_CLIENTS_REVERSED["client_id"]]).replace(' ', '_')

                    out_path_full = out_path / (filename + '_' + template_name + ".docx")

                    # TODO Bottleneck here, file is written, read, converted, saved, deleted. Conversion takes long
                    # save document in folder hierarchy
                    document.write(out_path_full)
                    convert_to(out_path, out_path_full)
                    os.remove(out_path_full)

                    created_documents_paths.append(out_path_full.with_suffix('.pdf'))  # replace docx with pdf

            self.merge_pdfs_and_remove(created_documents_paths, standard_pdfs, out_path, filename,
                                       INCLUDE_STANDARDS[doc_type])

    @staticmethod
    def merge_pdfs_and_remove(customized_documents, standard_pdfs, out_path, filename, include_standards=False):
        merger = PdfFileMerger()

        all_documents = customized_documents.copy()

        if include_standards:
            all_documents.extend(standard_pdfs)

        for document in all_documents:
            # The reference to the file descriptor is reassigned in each loop, but a reference to each
            # descriptor is kept in the merger object. The descriptors are closed at the end when
            # merger.close() is called
            # See: https://pythonhosted.org/PyPDF2/PdfFileMerger.html
            in_pdf = open(document, "rb")
            merger.append(in_pdf)

        with open(out_path / (filename + ".pdf"), "wb") as out_pdf:
            merger.write(out_pdf)
            merger.close()  # closes all file descriptors (input and output)

        # delete old in_pdfs
        for document in customized_documents:
            os.remove(document)
