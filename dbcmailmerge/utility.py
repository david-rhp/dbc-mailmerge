"""
Contains various helper functions.
"""
from pathlib import Path
from tkinter import filedialog, Tk
from itertools import product
import pandas as pd
import numpy as np


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


def parse_excel(filepath, sheet_name, field_list):
    # TODO add docstring
    # TODO add tests
    # Extract only relevant fields: all fields in field_list
    df = pd.read_excel(filepath, sheet_name)[field_list]
    df.fillna('', inplace=True)  # fill NaN with empty string so comparisons for the entire instance works

    # format datetime
    df_dates = df.select_dtypes([np.datetime64])
    for column in df_dates:
        df[column] = df_dates[column].dt.strftime("%d.%m.%Y")  # example: 31.12.2019

    return df


def mailmerge_factory(cls, data_path, data_sheet_name, field_map):
    """
    Factory function to create 1 instance of cls per record of the data source.

    This function takes an excel file, processes the records, and the selects only the relevnt columns, that are also
    in the field_map. Since the names of the excel columns might change, but the program level attribute names will stay
    the same, the map is also used to change the excel column names to the version used internally in the program.

    Parameters
    ----------
    cls : class
        Class for instantiating objects per record, either of type MailProject or Client.

    data_path : pathlib.Path or pathlike str
        Filepath to the data source, has to be `.xlsx`.
    data_sheet_name : str
        The sheet name, in which the records for the object instantiation are stored.
    field_map : dict
        A dictionary containing the mapping of excel_column_name to class_attribute_name (key: value). Class attribute
        names are used consistently throughout the project, however, excel column names might change more often.
        When a change occurs, only the field map has to be updated.

    Returns
    -------
    instances : list of instances of cls or instance of cls
        A list containing the created instances, if multiple records are in the data source, otherwise one instance.
    """
    # obtain DataFrame with only the columns of field_maps.keys()
    data = parse_excel(data_path, data_sheet_name, field_map.keys())

    # Extract one or more records from the data source and instantiate one instance of cls per record
    instances = []
    for _, record in data.iterrows():
        # Convert pandas series to dict for translation of excel column names to the version used internally
        record = record.to_dict()
        record = translate_dict(record, field_map)

        instances.append(cls(**record))

    if len(data) > 1:
        return instances
    else:
        return instances[0]


def translate_dict(in_dict, field_map, reverse=False):
    """
    Translates the keys of a dictionary to the values in field map.

    Can be used to translate excel column names to the names used internally in the project. Additionally, if the
    processing of the data has concluded, it can be used to `reverse` (kwarg) the translation, e.g., when the data
    using the internal names is used to populate the word template placeholders (the placeholders in word match the
    excel column names).

    Parameters
    ----------
    in_dict : dict
        Dictionary of which the keys should be translated to the values in field_map. It is not mutated.
    field_map : dict
        Dictionary of which contains the mapping of to be translated names (keys) to translated result (values).
        Values need to be hashable or the reversal process won't work, as the key, value pairs in field_map are swapped,
        i.e., values become keys of field_map.
    reverse : bool, optional
        Indicates if the field_map key: value pairs should be swapped (default: False). This would need to be done,
        if the in_dict has already been translated once, and the translation should be reversed. Otherwise the current
        implementation would throw an exception.

    Returns
    -------
    new_dict : dict
        A translated COPY of in_dict.

    """
    # TODO Automatically recognize when field_map needs to be recognized.
    new_dict = in_dict.copy()

    # use keys or values to translate

    if reverse:
        # reverse field_map for translation purposes
        field_map = {value: key for key, value in field_map.items()}

    for key in new_dict.copy():
        if field_map[key] not in new_dict:
            new_dict[field_map[key]] = new_dict[key]
            del new_dict[key]

    return new_dict
