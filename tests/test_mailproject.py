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

# TODO refactor tests to use 1 or 2 setup functions instead of duplicating setup code


class TestMailProject:
    def test_eq_operator(self):
        """
        Tests if testing for equality works as expected.

        Many of the other tests are dependant on this method working. If this test fails along with others,
        debug here first.
        """
        # All other tests are dependent on the equality operation defined in the class
        project1 = MailProject(**TEST_PROJECT_SINGLE_1)
        project2 = MailProject(**TEST_PROJECT_SINGLE_1)

        # Test for the same projects
        assert project1 == project2

        # Test for different projects
        project3 = MailProject(**TEST_PROJECT_SINGLE_2)
        assert project1 != project3

    def test_repr(self):
        """
        Tests the __repr__ method of the class by checking if its output can be used to construct an instance
        with identical class attributes.
        """
        project1 = MailProject(**TEST_PROJECT_SINGLE_1)
        project2 = eval(repr(project1))

        assert project1 == project2

    def test_str(self):
        project = MailProject(**TEST_PROJECT_SINGLE_1)

        expected = (f"Project ID ({TEST_PROJECT_SINGLE_1['project_id']}): {TEST_PROJECT_SINGLE_1['project_name']}, "
                    f"issuance {TEST_PROJECT_SINGLE_1['date_issuance']}, "
                    f"maturity {TEST_PROJECT_SINGLE_1['date_maturity']}")

        assert str(project) == expected

    def test_from_excel_single_project_1(self):
        """
        Tests if an instance created by the from_excel factory method is the same as one created by its constructor.
        The test assumes that the data in TEST_PROJECT_SINGLE matches the one in the excel file.
        """
        expected = MailProject(**TEST_PROJECT_SINGLE_1)
        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_single_1", FIELD_MAP_PROJECT)

        assert result == expected

    def test_from_excel_single_project_2(self):
        """
        Tests if an instance created by the from_excel factory method is the same as one created by its constructor.
        The test assumes that the data in TEST_PROJECT_SINGLE matches the one in the excel file.
        """
        expected = MailProject(**TEST_PROJECT_SINGLE_2)
        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_single_2", FIELD_MAP_PROJECT)

        assert result == expected

    def test_from_excel_multiple_projects(self):
        """Tests if multiple instances can be created from one call to the factory method."""
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

        # Business Requirement:
        # documents for client_records that have an amount >= 0 will be contacted. This means that if there is a numeric
        # value entered in the corresponding cell in Excel, the documents will be created. If not, parse_excel
        # reads the excel file, sets NaN for the missing value, and replaces NaN by an empty string.
        # As a result, records with no entered amount, are converted to empty strings and are filtered out by the
        # following lambda.
        selection_criteria = {"amount": lambda x: isinstance(x, (int, float))}
        result = project.select_clients(selection_criteria)

        # matches tests data source, id 3 excluded because no amount entered in test_data_source.xlsx
        expected_client_ids = [1, 2, 4]

        # Check if the correct amount of client_records has been returned
        assert len(result) == len(expected_client_ids)

        # Check if the correct client_records have been selected and returned
        for client in result:
            assert client["client_id"] in expected_client_ids

    def test_create_client_documents_with_filter_unix(self, with_filter=True):
        """
        Tests if the function was able to create the documents. It does NOT check if the content has been created
        correctly. Please check the output manually (visually).

        The test will always pass if no exception occurred and the files could be created.

        Run this function individually. Delete the client_correspondence folder manually in the ../tests/ directory
        before each run. Running this test in conjunction with `test_create_client_documents_without_filter_unix` will
        lead to inconsistent results.
        """
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

    def test_create_client_documents_without_filter_unix(self):
        """
        Tests if the function was able to create the documents. It does NOT check if the content has been created
        correctly. Please check the output manually (visually).

        The test will always pass if no exception occurred and the files could be created.

        Run this function individually. Delete the client_correspondence folder manually in the ../tests/ directory
        before each run. Running this test in conjunction with `test_create_client_documents_with_filter_unix` will
        lead to inconsistent results.
        """
        self.test_create_client_documents_with_filter_unix(with_filter=False)
        pass
