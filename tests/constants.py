import pandas as pd
from pathlib import Path

TEST_DATA_SOURCE_PATH = Path("../data/test/test_data_source.xlsx")

TEST_PROJECT_SINGLE_1 = {"project_id": 141,
                         "project_name": "Certainly a Project GmbH & Co. KG",

                         "date_issuance": pd.Timestamp(2019, 6, 30),
                         "date_maturity": pd.Timestamp(2022, 6, 30),

                         "coupon_rate": 0.12,
                         "commercial_register_number": "HRA 12345 B",
                         "issue_volume_min": 2000000,
                         "issue_volume_max": 3000000,
                         "collateral_string": "Land Charge and Letter of Comfort"}

TEST_PROJECT_SINGLE_2 = {"project_id": 178,
                         "project_name": "Another Project GmbH",

                         "date_issuance": pd.Timestamp(2019, 12, 31),
                         "date_maturity": pd.Timestamp(2022, 12, 31),

                         "coupon_rate": 0.11,
                         "commercial_register_number": "HRB 04321 A",
                         "issue_volume_min": 4000000,
                         "issue_volume_max": 5000000,
                         "collateral_string": "Letter of Comfort"}

TEST_PROJECT_MULTIPLE = [TEST_PROJECT_SINGLE_1, TEST_PROJECT_SINGLE_2]


TEST_CLIENT_1 = {"client_id": 1,
                 "first_name": "John1",
                 "last_name": "Doe1",

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

TEST_CLIENT_2 = {}