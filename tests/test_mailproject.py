"""
Author: David Meyer

Description
-----------
Contains the test suite for the MailProject class in mailproject.py
"""
from dbcmailmerge.mailproject import MailProject
from tests.test_constants import (HIERARCHY_ROOT, STANDARD_PDFS, TEST_DATA_SOURCE_PATH,
                                  TEST_PROJECT_SINGLE_1, TEST_PROJECT_SINGLE_2, TEST_PROJECT_MULTIPLE,
                                  TEST_CLIENT_MULTIPLE)
from dbcmailmerge.config import FIELD_MAP_CLIENTS, FIELD_MAP_PROJECT


class TestMailProject:
    def test_eq_operator(self):
        # All other tests are dependent on the equality operation defined in the class
        project1 = MailProject(**TEST_PROJECT_SINGLE_1)
        project2 = MailProject(**TEST_PROJECT_SINGLE_1)

        # Test for the same projects
        assert project1 == project2

        # Test for different projects
        project3 = MailProject(**TEST_PROJECT_SINGLE_2)
        assert project1 != project3

    def test_repr(self):
        project1 = MailProject(**TEST_PROJECT_SINGLE_1)
        project2 = eval(repr(project1))

        project1_attributes = vars(project1)
        project2_attributes = vars(project2)

        # Test if each attribute matches
        for key in project1_attributes:
            assert project1_attributes[key] == project2_attributes[key]

        # Test using __eq__ of the class
        assert project1 == project2

    def test_str(self):
        project = MailProject(**TEST_PROJECT_SINGLE_1)

        expected = (f"Project ID ({TEST_PROJECT_SINGLE_1['project_id']}): {TEST_PROJECT_SINGLE_1['project_name']}, "
                    f"issuance {TEST_PROJECT_SINGLE_1['date_issuance']}, "
                    f"maturity {TEST_PROJECT_SINGLE_1['date_maturity']}")

        assert str(project) == expected

    def test_direct_instantiation_project_1(self):
        project = MailProject(**TEST_PROJECT_SINGLE_1)

        # Test that all attributes have been set properly.
        for key in TEST_PROJECT_SINGLE_1:
            assert getattr(project, key) == TEST_PROJECT_SINGLE_1[key]

    def test_direct_instantiation_project_2(self):
        project = MailProject(**TEST_PROJECT_SINGLE_2)

        # Test that all attributes have been set properly.
        for key in TEST_PROJECT_SINGLE_2:
            assert getattr(project, key) == TEST_PROJECT_SINGLE_2[key]

    def test_from_excel_single_project_1(self):
        expected = MailProject(**TEST_PROJECT_SINGLE_1)
        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_single_1", FIELD_MAP_PROJECT)

        assert result == expected

    def test_from_excel_single_project_2(self):
        expected = MailProject(**TEST_PROJECT_SINGLE_2)
        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_single_2", FIELD_MAP_PROJECT)

        assert result == expected

    def test_from_excel_multiple_projects(self):
        expected = []
        for project in TEST_PROJECT_MULTIPLE:
            expected.append(MailProject(**project))

        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_multiple", FIELD_MAP_PROJECT)

        assert result == expected

    def test_create_clients(self):
        # set up project
        project = MailProject(**TEST_PROJECT_SINGLE_1)
        project.create_client_records(TEST_DATA_SOURCE_PATH, "client_data", FIELD_MAP_CLIENTS)

        # The first and the last client record of the data source match the projects in TEST_CLIENT_MULTIPLE
        result = [project.client_records[0], project.client_records[-1]]

        assert result == TEST_CLIENT_MULTIPLE

    def test_select_clients(self):
        # set up project
        project = MailProject(**TEST_PROJECT_SINGLE_1)
        project.create_client_records(TEST_DATA_SOURCE_PATH, "client_data", FIELD_MAP_CLIENTS)

        # documents for client_records that have an amount >= 0 will be contacted. This means that if there is a numeric
        # value entered in the corresponding cell in Excel, the documents will be created. If not, parse excel
        # reads the excel file, sets NaN for the missing value, and replaces NaN by an empty string.
        # As a result, records with no entered amount, are converted to empty strings and are filtered out here.
        selection_criteria = {"amount": lambda x: isinstance(x, (int, float))}
        result = project.select_clients(selection_criteria)

        # matches tests data source, id 3 excluded because no amount entered in excel
        expected_client_ids = [1, 2, 4]

        # Check if the correct amount of client_records has been returned
        assert len(result) == len(expected_client_ids)

        # Check if the correct client_records have been selected and returned
        for client in result:
            assert client["client_id"] in expected_client_ids

    def test_create_client_documents_with_filter(self, with_filter=True):
        # set up project
        project = MailProject(**TEST_PROJECT_SINGLE_1)
        project.create_client_records(TEST_DATA_SOURCE_PATH, "client_data", FIELD_MAP_CLIENTS)

        if with_filter:
            selection_criteria = {"amount": lambda x: bool(x)}
        else:
            selection_criteria = {}

        selected_clients = project.select_clients(selection_criteria)

        project.create_client_documents(selected_clients, HIERARCHY_ROOT, STANDARD_PDFS)

        pass

    def test_create_client_documents_without_filter(self):
        self.test_create_client_documents_with_filter(with_filter=False)
        pass
