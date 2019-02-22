"""
Contains various helper functions.
"""
from pathlib import Path
from tkinter import filedialog, Tk
from itertools import product


def prompt_filepath():
    """
    Prompts the user to select a directory and returns its absolute OS-specific filepath as a pathlib.Path object.

    Returns
    -------
    hierarchy_root: pathlib.Path
        Contains an absolute filepath.
    """
    # prevent second window pop up when prompting in askdirectory
    root = Tk()
    root.withdraw()

    # receive 'root' directory path in which the folder hierarchy should be created,
    hierarchy_root = Path(filedialog.askdirectory())

    return hierarchy_root
    # TODO create test


def create_folder_hierarchy(hierarchy_root, top_level_dir, sub_directories):
    """
    Creates directory tree in a given directory.

    Parameters
    ----------
    top_level_dir: str
        Specifies the name of the top level directory, in which the `sub_directories` should be created. This needs
        to be a valid directory name for the respective OS.

    sub_directories: list
        A 2d list, each row represents one directory level. Row 0 == top or `root` level directory.
        Row 1 would then be a list containing the sub directories for ALL directories in row 0.
        This means, each directory in row 0 has all the subdirectories of row 1. The elements of the lists need
        to be strings, specifying valid directory names for the respective OS.

    Returns
    -------
    None
    """
    # construct absolute paths
    paths = []
    for path in path_creator(sub_directories):
        paths.append(hierarchy_root / top_level_dir / path)

    for path in paths:
        path.mkdir(parents=True, exist_ok=True)

    # TODO reverse created folders if error encountered?
    # TODO return value?


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
