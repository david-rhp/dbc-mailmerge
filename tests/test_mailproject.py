from pathlib import Path

from mailmerge.mailproject import MailProject, Client, FIELD_MAP_PROJECT

TEST_PROJECT_SINGLE = {"project_id": 141, "project_name": "Certainly a Project GmbH & Co. KG"}
TEST_PROJECT_MULTIPLE = [{"project_id": 141, "project_name": "Certainly a Project GmbH & Co. KG"},
                         {"project_id": 178, "project_name": "Another Project GmbH"}]
TEST_DATA_SOURCE_PATH = Path("../data/test/test_data_source_base.xlsx")


class TestMailProject:
    def test_simple_instantiation(self):
        project = MailProject(**TEST_PROJECT_SINGLE)

        assert project.project_id == TEST_PROJECT_SINGLE["project_id"]
        assert project.project_name == TEST_PROJECT_SINGLE["project_name"]

    def test_from_excel_single_project(self):
        expected = MailProject(**TEST_PROJECT_SINGLE)
        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_single", FIELD_MAP_PROJECT)

        assert result == expected

    def test_from_excel_multiple_projects(self):
        expected = []
        for project in TEST_PROJECT_MULTIPLE:
            expected.append(MailProject(**project))

        result = MailProject.from_excel(TEST_DATA_SOURCE_PATH, "project_data_multiple", FIELD_MAP_PROJECT)

        assert result == expected

    def test_eq_operator(self):
        project1 = MailProject(**TEST_PROJECT_SINGLE)
        project2 = MailProject(**TEST_PROJECT_SINGLE)

        assert project1 == project2

        # Test for 1 attribute different
        project3 = MailProject(1, TEST_PROJECT_SINGLE["project_name"])
        assert project1 != project3

        # Test both attributes different
        project4 = MailProject(2, "Name")
        assert project1 != project4

    def test_repr(self):
        project1 = MailProject(**TEST_PROJECT_SINGLE)
        project2 = eval(repr(project1))

        assert project1 == project2
