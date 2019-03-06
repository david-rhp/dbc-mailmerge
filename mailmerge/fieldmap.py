# Keys are the column headers in the data source, keys and data source have to match or a key error will be raised
# values are the standardized way the respective data source is represented

FIELD_MAP_CLIENTS = {"client_id": "client_id",
                     "first_name": "first_name",
                     "last_name": "last_name",

                     "address_mailing_street": "address_mailing_street",
                     "address_mailing_zip": "address_mailing_zip",
                     "address_mailing_city": "address_mailing_city",
                     "address_notify_street": "address_notify_street",
                     "address_notify_zip": "address_notify_zip",
                     "address_notify_city": "address_notify_city",

                     "amount": "amount",
                     "subscription_am_authorized": "subscription_am_authorized",
                     "mailing_as_email": "mailing_as_email",
                     "depot_no": "depot_no",
                     "depot_bic": "depot_bic"}

FIELD_MAP_PROJECT = {"projektnummer": "project_id",
                     "projektname": "project_name",

                     "datum_emission": "date_issuance",
                     "datum_fälligkeit": "date_maturity",

                     "zinssatz": "coupon_rate",
                     "handelsregisternummer": "commercial_register_number",
                     "emissionsvolumen_min": "issue_volume_min",
                     "emissionsvolumen_max": "issue_volume_max",
                     "sicherheiten": "collateral_string"}
