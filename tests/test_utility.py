from pathlib import Path
from mailmerge.utility import path_creator, create_folder_hierarchy


def test_path_creator():
    test_input = [["advisor_1", "advisor_2"], ["ZU", "AP"]]
    expected = [Path().joinpath("advisor_1", "ZU"), Path().joinpath("advisor_1", "AP"),
                Path().joinpath("advisor_2", "ZU"), Path().joinpath("advisor_2", "AP")]

    paths = path_creator(test_input)
    assert paths == expected


def test_create_folder_hierarchy(tmp_path):
    top_level = "client_correspondence"

    base_path = tmp_path / top_level
    expected_paths = [base_path.joinpath("advisor_1", "ZU"), base_path.joinpath("advisor_1", "AP"),
                      base_path.joinpath("advisor_2", "ZU"), base_path.joinpath("advisor_2", "AP")]

    create_folder_hierarchy(tmp_path, top_level, [["advisor_1", "advisor_2"], ["ZU", "AP"]])

    # check if paths have been created by create_folder_hierarchy
    for path in expected_paths:
        assert path.exists()
