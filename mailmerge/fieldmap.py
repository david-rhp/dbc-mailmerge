# Keys are the column headers in the data source, keys and data source have to match or a key error will be raised
# values are the standardized way the respective data source is represented

FIELD_MAP_CLIENTS = {"db_id": "client_id",
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

FIELD_MAP_PROJECT = {"projektnummer": "project_id",
                     "projektname": "project_name",

                     "datum_emission": "date_issuance",
                     "datum_fälligkeit": "date_maturity",

                     "zinssatz": "coupon_rate",
                     "handelsregisternummer": "commercial_register_number",
                     "emissionsvolumen_min": "issue_volume_min",
                     "emissionsvolumen_max": "issue_volume_max",
                     "sicherheiten": "collateral_string"}
