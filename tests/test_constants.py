"""
Author: David Meyer

Description
-----------
Contains the constants used in the test suite. Changes to project and client dummy data need to also be made in the
test_data_source.xlsx file.
"""
from pathlib import Path


# Test Documents and Paths
##########################

HIERARCHY_ROOT = Path('../tests/')
STANDARD_PDFS = ["../data/tests/pib.pdf", "../data/tests/factsheet.pdf"]


# Project and Client Dummy Data
###############################

TEST_DATA_SOURCE_PATH = Path("../data/tests/test_data_source.xlsx")

TEST_PROJECT_SINGLE_1 = {"project_id": 141,
                         "project_name": "Certainly a Project GmbH & Co. KG",

                         "date_issuance": "30.06.2019",
                         "date_maturity": "30.06.2022",

                         "coupon_rate": 0.12,
                         "commercial_register_number": "HRA 12345 B",
                         "issue_volume_min": 2000000,
                         "issue_volume_max": 3000000,
                         "collateral_string": "Land Charge and Letter of Comfort"}

TEST_PROJECT_SINGLE_2 = {"project_id": 178,
                         "project_name": "Another Project GmbH",

                         "date_issuance": "31.12.2019",
                         "date_maturity": "31.12.2022",

                         "coupon_rate": 0.11,
                         "commercial_register_number": "HRB 04321 A",
                         "issue_volume_min": 4000000,
                         "issue_volume_max": 5000000,
                         "collateral_string": "Letter of Comfort"}

TEST_PROJECT_MULTIPLE = [TEST_PROJECT_SINGLE_1, TEST_PROJECT_SINGLE_2]


TEST_CLIENT_1 = {"client_id": 1,
                 "title": "",
                 "advisor": "Betreuer 1",
                 "first_name": "John1",
                 "last_name": "Doe1",
                 "salutation_address_field": "Herrn",
                 "salutation": "er Herr",

                 "address_mailing_street": "Client 1 Str. 1",
                 "address_mailing_zip": "80001",
                 "address_mailing_city": "Munich",
                 "address_notify_street": "Client 1 Str. 1",
                 "address_notify_zip":	"80001",
                 "address_notify_city":	"Munich",

                 "amount":	50000,
                 "subscription_am_authorized": 1,
                 "mailing_as_email": 1,
                 "depot_no": "123456789",
                 "depot_bic": "SOMEALPHANUMERICSTRING"}

TEST_CLIENT_2 = {"client_id": 4,
                 "advisor": "Betreuer 2",
                 "title": "Dr.",
                 "first_name": "Jane4",
                 "last_name": "Doe4",
                 "salutation_address_field": "Frau",
                 "salutation": "e Frau",

                 "address_mailing_street": "Client 4 Str. 4",
                 "address_mailing_zip": "80004",
                 "address_mailing_city": "São Paulo",
                 "address_notify_street": "Different Client 4 Str. 4",
                 "address_notify_zip":	"90004",
                 "address_notify_city":	"Barcelona",

                 "amount": 25000,
                 "subscription_am_authorized": 0,
                 "mailing_as_email": 0,
                 "depot_no": "0048522358",
                 "depot_bic": "00BICNUM120X0"}

TEST_CLIENT_MULTIPLE = [TEST_CLIENT_1, TEST_CLIENT_2]

