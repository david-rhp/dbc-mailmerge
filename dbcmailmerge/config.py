"""
Contains the constants for configuring the application.

Constants
---------
TEMPLATES : dict
    Contains doc_type, [template_paths] pairs.
    Add/remove templates that should be used for creating customized documents per clients. Currently, each template
    in TEMPLATES will be used for creating the documents. Each template needs to have a corresponding doc_type (key),
    which will be used to create folders for each doc_type. The documents created on the basis of a template,
    will be saved in their corresponding folder
    (e.g. files created using the offer_documents template will be stored in the offer_documents folder)

INCLUDE_STANDARDS : dict
    Determines if standardized pdfs should be appended at the end of a specific doc_type (key). Standardized
    pdfs are files, which contain no customized data and a generic for each client. This means, that each customized,
    document based on a template may or may not receive these standardized docs(dependant on the doc_type, specified
    in this dict).

    For example, these could be general terms and conditions, that don't change from client to client and apply
    to each client the same. However, this might not be relevant for each doc_type.

    In this example, offer_documents are sent to the client, for which the client has to receive
    standardized documents. The appropriateness_test is an internal form, that contains client specific that, but that
    is for internal use only, thus, it does not need standardized documents, such as general terms and conditions,
    since they are available for internal use anyways.

FIELD_MAP_CLIENTS, FIELD_MAP_PROJECT : dict
    Contains the translation from excel column names to names that are used internally in this project. This is
    required because external files, such as the data source (excel column names) or the word templates (placeholders)
    are changed from time to time and don't adhere to a standard.

    If the name in the excel file or the word template changes, simply update the field map here. No other, changes
    have to be made.

    However, this does not include removing field_names. The current attributes are the minimally required information
    needed for each mailing project in general.

FIELD_MAP_CLIENTS_REVERSED : dict
    The reversal of FIELD_MAP_CLIENTS, created once, dynamically at runtime. This is used when the original names of
    the excel column names / word template placeholder names should be used.

    For example, when writing a created document to disk, the names of the client records (keys) were translated back
    to its original form, so that they can be used to populate the word templates' placeholders. However, during the
    writing process, this specific value might be relevant. (e.g., when the storage location of the created document
    should be determined, it used the advisor name, documents are aggregated per client. In order to get the
    advisor name of that specific client record, it's value for the advisor key has to accessed,
    but the key is in its original, untranslated format.

"""

import os
from pathlib import Path

# Document Templates
####################

base_path = Path(os.path.dirname(os.path.realpath(__file__))).parents[0]  # 1 level up relative to this module

TEMPLATES = {"offer_documents": [base_path / "data/templates/cover_letter.docx",
                                 base_path / "data/templates/subscription_agreement.docx"],

             "appropriateness_test": [base_path / "data/templates/appropriateness_test.docx"]}

INCLUDE_STANDARDS = {"offer_documents": True, "appropriateness_test": False}


# Field Maps
############

# Keys are the column headers in the data source, keys and data source have to match or a key error will be raised
# values are the standardized way the respective data source is represented
FIELD_MAP_CLIENTS = {"db_id": "client_id",
                     "betreuer": "advisor",
                     "titel": "title",
                     "vorname": "first_name",
                     "nachname": "last_name",
                     "anrede_adressfeld": "salutation_address_field",
                     "anrede": "salutation",

                     "post_str": "address_mailing_street",
                     "post_plz": "address_mailing_zip",
                     "post_ort": "address_mailing_city",
                     "melde_str": "address_notify_street",
                     "melde_plz": "address_notify_zip",
                     "melde_ort": "address_notify_city",

                     "zeichnungssumme": "amount",
                     "in_vv": "subscription_am_authorized",
                     "medium_email": "mailing_as_email",
                     "depot_nummer": "depot_no",
                     "depot_bic": "depot_bic"}

FIELD_MAP_CLIENTS_REVERSED = {value: key for key, value in FIELD_MAP_CLIENTS.items()}

FIELD_MAP_PROJECT = {"projektnummer": "project_id",
                     "projektname": "project_name",

                     "datum_emission": "date_issuance",
                     "datum_fälligkeit": "date_maturity",

                     "zinssatz": "coupon_rate",
                     "handelsregisternummer": "commercial_register_number",
                     "emissionsvolumen_min": "issue_volume_min",
                     "emissionsvolumen_max": "issue_volume_max",
                     "sicherheiten": "collateral_string"}
