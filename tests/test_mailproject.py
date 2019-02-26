from pathlib import Path
import pandas as pd
from mailmerge.mailproject import MailProject, Client, FIELD_MAP_PROJECT

TEST_PROJECT_SINGLE_1 = {"project_id": 141, "project_name": "Certainly a Project GmbH & Co. KG",
                        "date_issuance": pd.Timestamp(2019, 6, 30), "date_maturity": pd.Timestamp(2022, 6, 30),
                        "coupon_rate": 0.12, "commercial_register_number": "HRA 12345 B", "issue_volume_min": 2000000,
                        "issue_volume_max": 3000000, "collateral_string": "Land Charge and Letter of Comfort"}

TEST_PROJECT_SINGLE_2 = {"project_id": 178, "project_name": "Another Project GmbH",
                       "date_issuance": pd.Timestamp(2019, 12, 31), "date_maturity": pd.Timestamp(2022, 12, 31),
                       "coupon_rate": 0.11, "commercial_register_number": "HRB 04321 A", "issue_volume_min": 4000000,
                       "issue_volume_max": 5000000, "collateral_string": "Letter of Comfort"}

TEST_PROJECT_MULTIPLE = [TEST_PROJECT_SINGLE_1, TEST_PROJECT_SINGLE_2]
TEST_DATA_SOURCE_PATH = Path("../data/test/test_data_source_base.xlsx")


class TestMailProject:
    def test_simple_instantiation_project_1(self):
        project = MailProject(**TEST_PROJECT_SINGLE_1)

        # Test that all attributes have been set properly.
        for key in TEST_PROJECT_SINGLE_1:
            assert getattr(project, key) == TEST_PROJECT_SINGLE_1[key]

    def test_simple_instantiation_project_2(self):
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

    def test_eq_operator(self):
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

