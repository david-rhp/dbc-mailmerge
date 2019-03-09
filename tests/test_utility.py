from pathlib import Path

from dbcmailmerge.mailproject import MailProject, Client
from dbcmailmerge.utility import path_creator, create_folder_hierarchy, prompt_filepath, mailmerge_factory
from tests.test_constants import TEST_DATA_SOURCE_PATH, TEST_PROJECT_SINGLE_1, TEST_CLIENT_1, TEST_CLIENT_2
from dbcmailmerge.constants import FIELD_MAP_CLIENTS, FIELD_MAP_PROJECT


def test_path_creator():
    test_input = [["advisor_1", "advisor_2"], ["offer_documents", "appropriateness_test"]]
    expected = [Path().joinpath("advisor_1", "offer_documents"), Path().joinpath("advisor_1", "appropriateness_test"),
                Path().joinpath("advisor_2", "offer_documents"), Path().joinpath("advisor_2", "appropriateness_test")]

    paths = path_creator(test_input)
    assert paths == expected


def test_create_folder_hierarchy(tmp_path):
    top_level = "client_correspondence"

    base_path = tmp_path / top_level
    expected_paths = [base_path.joinpath("advisor_1", "offer_documents"),
                      base_path.joinpath("advisor_1", "appropriateness_test"),
                      base_path.joinpath("advisor_2", "offer_documents"),
                      base_path.joinpath("advisor_2", "appropriateness_test")]

    create_folder_hierarchy(tmp_path, top_level, [["advisor_1", "advisor_2"],
                                                  ["offer_documents", "appropriateness_test"]])

    # check if paths have been created by create_folder_hierarchy
    for path in expected_paths:
        assert path.exists()


def test_prompt_filepath(tmp_path, mocker):
    expected = Path(tmp_path)

    # Patch askdirectory method with lambda function
    # instead of obtaining the path through tkinter's user prompt, return tmp_path
    mocker.patch("tkinter.filedialog.askdirectory", lambda: tmp_path)
    result = prompt_filepath()
    assert result == expected


def test_mailmerge_factory():
    # Test with instantiation of 1 project
    result = mailmerge_factory(MailProject, TEST_DATA_SOURCE_PATH, "project_data_single_1", FIELD_MAP_PROJECT)
    expected = MailProject(**TEST_PROJECT_SINGLE_1)

    assert result == expected

    # Test with instantiation of 1 client
    clients = mailmerge_factory(Client, TEST_DATA_SOURCE_PATH, "client_data", FIELD_MAP_CLIENTS)
    result = clients[0]  # first client matches data in test_client_1
    expected = Client(**TEST_CLIENT_1)

    assert result == expected

    # Test with instantiation of 2 clients
    result = [clients[0], clients[-1]]
    expected = [expected, Client(**TEST_CLIENT_2)]

    assert result == expected
