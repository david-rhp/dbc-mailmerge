"""
Author: David Meyer

Description
-----------
Contains the main classes and business logic for creating a MailProject and corresponding clients.

In its current state, the business logic is tied into the classes and should be factored out in a future release
to make the classes more maintainable and extendable.
"""
import os
from mailmerge import MailMerge
from PyPDF2 import PdfFileMerger
from dbcmailmerge.config import (FIELD_MAP_CLIENTS, FIELD_MAP_CLIENTS_REVERSED, FIELD_MAP_PROJECT,
                                 TEMPLATES, INCLUDE_STANDARDS, CONVERSION_MAP)
from dbcmailmerge.utility import translate_dict, create_folder_hierarchy, parse_excel
from dbcmailmerge.docx2pdfconverter import convert_to


class MailProject:
    """
    Models a mail project. For each mail project, a number of documents (customized and standardized) have to be
    created for the company's clients. Not all clients receive documents per project.

    An instance of this class should be created by using the factory method `MailProject.from_excel`.

    This class can be used to create a mail project, load in the clients, filter the clients, and create the customized
    and standardized documents.

    Attributes
    ----------
    project_id : int
        The id to distinctly identify a project.
    project_name : str
        The name of the project. Usually the issuing entity of a bond issue.
    date_issuance : str
        The date of the bond issuance. Stored as a string in the dd.mm.yyyy format, e.g., 31.12.2019 (format used in
        Germany).
    date_maturity : str
        The maturity date of the bond. Stored as a string in the dd.mm.yyyy format, e.g., 31.12.2019 (format used in
        Germany).
    coupon_rate : float
        Coupon rate of the bond. Has to be a decimal, i.e., 12% stored as 0.12. Impacts output format for creating
        the documents using the word templates.
    commercial_register_number : str
        The commercial register number used to distinctly identify companies in Germany (Handelsregisternummer).
    issue_volume_min : int
        Minimum issuance: Amount required for the bond to be issued. Can be any int in 5k increments >= 1 million.
        Smaller values are not feasible from a business perspective.
    issue_volume_max : int
        Maximum issuance: Upper limit of the bond issuance. Can be any int in 5k increments >= issue_volume_min.
    collateral_string : str
        An enumeration of the collateral for the bond creditors.
    client_records : list of dict, optional
        A list of dictionaries. Each dictionary represents a client record. A record contains information pertaining
        to a client, e.g., the id, the address, the subscription amount etc. (default: None).
    """
    # TODO factor out business logic of classes
    TOP_LEVEL_DIR = "client_correspondence"  # name of directory where the created documents should be stored
    AMOUNT_EMPTY_PLACEHOLDER = '_' * 20

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
        """
        Factory method to create 1 instance of cls per record of the data source.

        This function takes an excel file, processes the records, and the selects only the relevant columns, that are
        also in the field_map. Since the names of the excel columns might change, but the program level attribute names
        will stay the same, the map is also used to change the excel column names to the version used internally
        in the program.

        Parameters
        ----------
        cls : class
            Class for instantiating objects per record.
        project_data_path : pathlib.Path or pathlike str
            Filepath to the data source, has to be `.xlsx`.
        project_data_sheet_name : str
            The sheet name, in which the records for the object instantiation are stored.
        project_field_map : dict
            A dictionary containing the mapping of excel_column_name to class_attribute_name.
            Class attribute names are used consistently throughout the project, however, excel column names might
            change more often. When a change occurs, only the field map in the config file has to be updated.

        Returns
        -------
        instances : list of instances of cls or instance of cls
            A list containing the created instances, if multiple records are in the data source, otherwise one instance.
        """
        # obtain DataFrame with only the columns of field_maps.keys()
        data = parse_excel(project_data_path, project_data_sheet_name, project_field_map.keys())

        # Extract one or more records from the data source and instantiate one instance of cls per record
        instances = []
        for _, record in data.iterrows():
            # Convert pandas series to dict for translation of excel column names to the version used internally
            record = record.to_dict()
            record = translate_dict(record, project_field_map)

            instances.append(cls(**record))

        if len(data) > 1:
            return instances
        else:
            return instances[0]

    def create_client_records(self, client_data_path, client_data_sheet_name, client_field_map):
        """
        Method to load records into the class instance's client_records attribute based on a provided excel file.

        Reads an excel file and processes its rows. All processed records are stored as a list of dicts in the
        instance's client_records attribute.

        Parameters
        ----------
        client_data_path : pathlib.Path or pathlike str
            Filepath to the data source, has to be `.xlsx`.
        client_data_sheet_name : str
            The sheet name, in which the records for the object instantiation are stored.
        client_field_map : dict
            A dictionary containing the mapping of excel_column_name to class_attribute_name.
            Class attribute names are used consistently throughout the project, however, excel column names might
            change more often. When a change occurs, only the field map in the config file has to be updated.

        Returns
        -------
        None

        Raises
        ------
        ValueError
            If the MailProject instance is not empty before calling this method.
        """
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

        self.__cast_client_records(True)

    def __cast_client_records(self, silent=False):
        """
        Converts the instance attributes that are found in the CONVERSION MAP using the functions in the CONVERISON MAP.

        Mutates the dictionaries stored in client_records.

        Parameters
        ----------
        silent : bool, optional
            Indicates if an error in the conversion process should be silenced (default: False).

        Returns
        -------
        None

        Raises
        ------
        ValueError
            If an error occurs during the conversion process and silent=False
            For example, a function in CONVERSION_MAP tried casting an incompatible value.
        """
        for record in self.client_records:
            for key in CONVERSION_MAP:
                conversion_function = CONVERSION_MAP[key]

                # Try the conversion
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
                        # Value was not converted.
                        pass

    def select_clients(self, selection_criteria):
        """
        Selects client records from the instance attribute using a selection function.

        Parameters
        ----------
        selection_criteria : dict of functions
            The key represents the attribute on which the corresponding function should be applied.
            The function (value of selection_criteria dict) needs to take a single input and return True, if the client
            should be included, and False if the client should be excluded.
        Returns
        -------
        selected_clients : list of dicts
            A list containing the client_records (dicts) that evaluate to True for the function in selection_criteria.
        """
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

    def __create_project_record(self):
        """
        Creates a project record that can be used for populating the placeholders in a word template.

        This method also applies some formatting to the project data in the record, so that it has the intended
        format for the MailMerge.

        Returns
        -------
        project_record : dict
            The formatted and translated project record. Translated means, that its keys have the matching names for
            the placeholders in the word templates.
        """
        project_record = vars(self)
        del project_record["client_records"]  # not part of the fields in the template

        # convert coupon rate from decimal to percentage and use German comma
        project_record["coupon_rate"] = format(float(project_record["coupon_rate"]) * 100, ".2f").replace('.', ',')

        # translate project data so that the fields (keys) match the names in the word template
        project_record = translate_dict(project_record, FIELD_MAP_PROJECT, reverse=True)

        # cast to str for MailMerge
        project_record = {key: str(value) for key, value in project_record.items()}

        return project_record

    def create_client_documents(self, selected_clients, hierarchy_root, standard_pdfs):
        """
        Creates the customized docs, includes the standard pdfs where appropriate and saves the merged file as 1 PDF.

        The pdfs are saved in the folder structure required by the business need.

        Parameters
        ----------
        selected_clients : list of dicts
            A list containing the client_records (dicts) that evaluate to True for the function in selection_criteria.
        hierarchy_root : pathlib.Path
            The location in which the TOP_LEVEL_DIR should be created, which in turn will be used
            to store all created documents. The files will be saved first by advisor, and within advisor by doc type
            (see TEMPLATES.keys()).
        standard_pdfs : list of pathlib.Path or pathlike str
            File paths to the pdfs that should be included in the mail merge.

        Returns
        -------
        None
        """
        project_record = self.__create_project_record()

        advisors = set()
        merge_records = []
        for client_record in selected_clients:
            # add client advisor for creation of sub_directories
            advisors.add(client_record["advisor"])

            # Apply formatting to client record
            client_record = self.__format_client_records(client_record)

            # translate client to match placeholders in word
            client_record = translate_dict(client_record, FIELD_MAP_CLIENTS, reverse=True)

            # add project data
            client_record.update(project_record)

            merge_records.append(client_record)

        # create folder hierarchy for the storage of the created documents
        sub_directories = [list(advisors), INCLUDE_STANDARDS.keys()]
        create_folder_hierarchy(hierarchy_root, type(self).TOP_LEVEL_DIR, sub_directories)

        for client_record in merge_records:
            self.__create_client_document(client_record, standard_pdfs, hierarchy_root)

    def __create_client_document(self, client_record, standard_pdfs, hierarchy_root):
        """

        Parameters
        ----------
        client_record : dict
            The dict represents a client record. A record contains information pertaining to a client,
            e.g., the id, the address, the subscription amount etc.
        standard_pdfs : list of pathlib.Path or pathlike str
                File paths to the pdfs that should be included in the mail merge.
        hierarchy_root : pathlib.Path
            The location in which the TOP_LEVEL_DIR should be created, which in turn will be used
            to store all created documents. The files will be saved first by advisor, and within advisor by doc type
            (see TEMPLATES.keys()).

        Returns
        -------
        None
        """
        for doc_type in TEMPLATES.keys():
            created_documents_paths = []
            for template_path in TEMPLATES[doc_type]:
                with MailMerge(template_path) as document:
                    # copy word template and replace placeholders with client instance data and project data
                    document.merge(**client_record)

                    # Create path to location where the file should be saved
                    out_path = (hierarchy_root
                                / type(self).TOP_LEVEL_DIR
                                / client_record[FIELD_MAP_CLIENTS_REVERSED["advisor"]]
                                / doc_type)

                    # Name used for saving the file in order to be able to distinguish documents that were
                    # created based on different templates.
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

                    # TODO Bottleneck here, file is written, read, converted, saved, deleted. Conversion takes long.

                    # save document in folder hierarchy as docx
                    document.write(out_path_full)

                    # convert docx to pdf
                    convert_to(out_path, out_path_full)

                    # delete docx because it is not required for the final output
                    os.remove(out_path_full)

                    # collect recently created pdf file path so that they can be removed later on when all pdfs
                    # per client are merged
                    created_documents_paths.append(out_path_full.with_suffix('.pdf'))  # replace docx with pdf

            self.__merge_pdfs_and_remove(created_documents_paths, standard_pdfs, out_path, filename,
                                         INCLUDE_STANDARDS[doc_type])

    def __format_client_records(self, client_record):
        """
        Format the client record based on a pre-determined business need.

        Parameters
        ----------
        client_record : dict
            The to be formatted client record. It is not mutated.
        Returns
        -------
        client_record : dict
            The formatted client record.
        """
        client_record = client_record.copy()

        # Apply formatting rules

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

        client_record["address_mailing_street"] += '\n'  # create space between street and zip city combination

        if client_record["amount"]:
            client_record["amount"] = format(client_record["amount"], ",.2f")
        else:
            # if no amount entered in excel, use _ als placeholder for the customer to enter in handwritten form
            client_record["amount"] = type(self).AMOUNT_EMPTY_PLACEHOLDER

        client_record = {key: str(value) for key, value in client_record.items()}  # cast to str for MailMerge
        return client_record

    @staticmethod
    def __merge_pdfs_and_remove(customized_documents_paths, standard_pdfs, out_path, filename, include_standards=False):
        """
        Merges the pdfs at the provided filepaths together into one file, and removes the unmerged versions.

        Parameters
        ----------
        customized_documents_paths : list
            Contains the filepaths to the recently created pdfs. Each element represents one pdf. All elements of the
            list correspond to one specific client. The list is not mutated.
        standard_pdfs : list of pathlib.Path or pathlike str
                File paths to the pdfs that should be included in the mail merge.
        out_path : pathlib.Path
            Path to the folder where the merged pdfs should be saved.
        filename : str
            Name for the to be saved pdf file
        include_standards : bool, optional
            Indicates if the standard_pdfs should be included at the end of the customized and merged pdfs.
            Refer to the docs in config.py for an example.
        Returns
        -------
        None
        """
        merger = PdfFileMerger()

        all_documents = customized_documents_paths.copy()

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
        for document in customized_documents_paths:
            os.remove(document)
