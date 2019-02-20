from pathlib import Path
from mailmerge.utility import path_creator


def test_path_creator():
    test_input = [["advisor_1", "advisor_2"], ["ZU", "AP"]]
    expected = [Path().joinpath("advisor_1", "ZU"), Path().joinpath("advisor_1", "AP"),
                Path().joinpath("advisor_2", "ZU"), Path().joinpath("advisor_2", "AP")]

    paths = path_creator(test_input)
    assert paths == expected
