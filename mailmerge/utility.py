"""
Contains various helper functions.
"""
from pathlib import Path
from tkinter import filedialog, Tk
from itertools import product


def create_folder_hierarchy(top_level_dir, sub_directories):
    # prevent second window pop up
    root = Tk()
    root.withdraw()

    # receive 'root' directory path in which the folder hierarchy should be created,
    hierarchy_root = Path(filedialog.askdirectory())

    # construct absolute paths
    paths = []
    for path in path_creator(sub_directories):
        paths.append(hierarchy_root / top_level / path)

    for path in paths:
        path.mkdir(parents=True)

    # create directories, abort if already existing?
    # return absolute paths of the created folders
    # return None, if error?
    # reverse created folders if error encountered


def path_creator(directories):
    """
    Takes a 2d list, each row representing a directory level and returns the product of the of these directories
    as individual Path objects.

    Parameters
    ----------
    directories: list
        A 2d list, each row represents one directory level. Row 0 == top or `root` level directory.
        Row 1 would then be a list containing the subdirectories for ALL directories in row 0.
        This means, each directory in row 0 has all the subdirectories of row 1.

    Returns
    -------
    paths: list
        A list of individual pathlib.Path objects, each object representing one distinct path.

    Examples
    --------
    [[dir_1, dir_2], [sub_1, sub_2]]

    [Path(dir_1/sub_1), Path(dir_1/sub_2), Path(dir_2/sub_1), Path(dir_2/sub_2)]
    """
    parts = product(*directories)
    paths = []
    for part in parts:
        path = Path().joinpath(*part)
        paths.append(path)

    return paths

#print(path_creator([["advisor_1", "advisor_2"], ["ZU", "AP"]]))

top_level = "client_correspondence"

create_folder_hierarchy(top_level, [["advisor_1", "advisor_2"], ["ZU", "AP"]])