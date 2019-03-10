"""
Author: David Meyer

Description
-----------
Main file to run the application. Also contains helper functions used for creating the CLI.

Prompts the user to select a directory where the documents should be saved, asks
for the data source (excel), and asks for standard pdfs (generic) that should be appended at the end of each created
document (from word template) per client.
"""
import sys
import tkinter as tk
from tkinter import messagebox, filedialog
from pathlib import Path
from xlrd import XLRDError

from dbcmailmerge.config import FIELD_MAP_CLIENTS, FIELD_MAP_PROJECT
from dbcmailmerge.mailproject import MailProject

ABORT_KEYWORDS = ('q', "quit")
# First element of the tuple is an explanation,the second a key to a filter
FILTER_MENU = {0: ("all filters selected",), 1: ("filter by amount >= 0", "amount")}


def prompt_str(prompt="Please enter a value or type `q` or `quit` to abort"):
    """
    Prompts the user to enter a non-empty string and returns the result. Re-prompts if no value entered.

    Parameters
    ----------
    prompt : str, optional
        A message to prompt the user for input (default: generic str input request). No trailing colon or whitespace
        required.

    Returns
    -------
    value : str
        The non-empty string the user was prompted for.

    Raises
    ------
    SystemExit
        When the users enters one of the ABORT_KEYWORDS.
    """
    while True:
        value = input(prompt.strip() + ": ").strip()

        if value.lower() in ABORT_KEYWORDS:
            sys.exit(1)

        if not value:
            print("You didn't provide any input. Please try again.\n")
        else:
            return value


def prompt_data_source(data_kind, prompt_source_file=True, prompt_sheet_name=True):
    """
    Asks the user to select a data source using the OS-specific explorer and/or to enter a sheet_name in the console.

    Parameters
    ----------
    data_kind : str
        Name for which data source the user should be prompted.
    prompt_source_file : bool, optional
        Determines if the user should be asked for a data source file using the explorer (default: True).
    prompt_sheet_name : bool, optional
        Determines if the user should be asked for a sheet_name (excel) using the console (default: True).
    Returns
    -------
    data_source, sheet_name : tuple of str or Tuple of None
        For each position in the tuple, returns None if the user was not prompted for the corresponding category
        (prompt_source_file, prompt_sheet_name). Otherwise, returns the filepath to the datasource and its sheet_name.

    Warnings
    --------
    This function does NOT check if the sheet_name actually exists in the data source and raises no related error.
    This can be remedied by nesting this function in a loop and at try, except block and catch the corresponding error
    when the sheet_name has been received and used for accessing an excel file.
    """
    data_source = None
    if prompt_source_file:
        messagebox.showinfo("Select Data Source",
                            "Select the data source for the project and the respective client_records.")
        data_source = filedialog.askopenfilename()
        root.update()

    sheet_name = None
    if prompt_sheet_name:
        sheet_name = prompt_str(f"Please provide the sheet_name of the {data_kind} data source")

    return data_source, sheet_name


def prompt_int(prompt="Please enter one integer"):
    """
    Prompts a user to enter one integer and returns the single int.

    Parameters
    ----------
    prompt: str, optional
        The text displayed when prompting the user for input (default: generic int input request).

    Returns
    -------
    result: int
        The integer provided by the user.

    Warnings
    --------
    The user can terminate the program by entering one of the abort keywords.

    Raises
    ------
    SystemExit
        If the user enters one of the ABORT_KEYWORDS -> catch to prevent premature termination of the program.
    """
    while True:
        user_input = input(prompt + ': ').strip()

        if user_input.lower() in ABORT_KEYWORDS:
            sys.exit(0)

        try:
            result = int(user_input)
        except ValueError:
            print("Please enter only one integer, e.g., `5` or `quit` to abort.\n---\n")
        else:
            return result


def print_dictionary_menu(dictionary):
    """
    Prints a dictionary displaying its keys as options in order to represent a menu.

    Parameters
    ----------
    dictionary : dict
        Dictionary, where each key maps to a list/tuple. The first element of the value is an explanation regarding
        the second element.

    Returns
    -------
    None
    """
    for key in dictionary:
        print(f"Option {key}: {dictionary[key][0]}")


def select_filter():
    """
    Prompts the user to select 1 or multiple filter functions. Currently, only one filter is implemented.

    Returns
    -------
    selected_filters : dict
        Contains the filter functions that can be used in MailProject.select_clients.
    """
    # TODO add multiple filters
    selection = None
    filters = {"amount": lambda x: bool(x)}  # evaluates to False for cells that had no value in the data source

    selected_filters = {}
    while selection != 0 and not selected_filters:

        print()
        print_dictionary_menu(FILTER_MENU)
        selection = prompt_int("Please select an option from the menu")

        if selection not in FILTER_MENU:
            print(f"This option does not exist, please select one of the following {FILTER_MENU.keys()}\n")

        # add filter
        elif selection != 0:
            key = FILTER_MENU[selection][1]  # choose second element of tuple
            selected_filters[key] = filters[key]

    return selected_filters


def prompt_files():
    """
    Prompts the user to select one or more files using the OS-specific GUI explorer.

    Returns
    -------
    selected_files : list
        Contains the file paths of the selected files.
    """
    messagebox.showinfo("Select Files", "Please select the standardized documents that you would like to include.")

    selected_files = []
    while True:
        file = filedialog.askopenfile()
        root.update()

        selected_files.append(file.name)  # .name = append filepath instead of _io.TextIOWrapper

        ask_for_next_file = messagebox.askyesno("Select File", "Do you want to select another file?")
        root.update()

        if not ask_for_next_file:
            return selected_files


def create_project_and_clients(data_source=None, data_kind=("project", "client"), project_object=None, counter=0):
    """
    Prompts the user to provide select an excel file, provide a sheet name, and creates an instance of a project
    or calls the project's create_client method based on the data in that sheet.

    This function is designed to run twice, first invoked by the caller, second invoked by itself. The first invocation
    creates a MailProject instance, the second invocation calls its create_clients method,
    thus, loading client instances into the project. For each call, a different excel sheet is processed.

    Parameters
    ----------
    data_source : pathlib.Path or pathlike str or None, optional
        Filepath to the data source, has to be `.xlsx` (default: None). If no data source is provided, the user
        will be prompted to select one.
    data_kind : tuple, optional
        For which data categories the user should provide the excel sheet names for (default: project, client).
    project_object : MailProject or None, optional
        The project_object for which the client_records should be created. None is used when no project instance has
        been instantiated.
    counter : int, optional
        Counts how often the function has called itself. If the function has invoked itself once (counter=1), it has
        run 2 times in total (once by the original caller, once by itself) and can return to the original caller.

    Returns
    -------
    project_object :
        The created MailProject instance with instantiated client_records.
    """
    while True:
        if not data_source:
            data_source, data_sheet_name = prompt_data_source(data_kind[counter], prompt_source_file=True)
        else:
            _, data_sheet_name = prompt_data_source(data_kind[counter], prompt_source_file=False)

        try:
            if project_object:
                project_object.create_client_records(data_source, data_sheet_name, FIELD_MAP_CLIENTS)
            else:
                project_object = MailProject.from_excel(data_source, data_sheet_name, FIELD_MAP_PROJECT)
        except XLRDError:
            print(f"Couldn't find the sheet name you specified for {data_kind[counter]} data, please try again.\n")
        else:
            break

    if counter == 0:
        counter += 1
        project_object = create_project_and_clients(data_source, data_kind, project_object, counter)

    return project_object


if __name__ == "__main__":
    # hide root window
    root = tk.Tk()
    root.withdraw()

    # Obtain location from the user for saving the documents
    # The folder structure is saved as a constant, and is determined by the business need
    messagebox.showinfo("Select Directory", "Select the directory where you want to save the created documents.")
    hierarchy_root = Path(filedialog.askdirectory())
    root.update()

    # Prompt for data source, create project, and load clients
    project = create_project_and_clients()

    # Prompt the user to select a filter for selecting only records from the data source, that are relevant
    selection_criteria = select_filter()
    selected_clients = project.select_clients(selection_criteria)

    # Ask the user to select the standard pdfs that should be appended at the end of the created document per client.
    standard_pdfs = prompt_files()

    summary_msg = (f"You have selected the project:\n"
                   f"{project.project_id}: {project.project_name}.\n\n"
                   f"You are about to create documents for\n"
                   f"{len(selected_clients)} clients\n\n"
                   f"Do you wish to continue?")

    start_mailmerge = messagebox.askyesno("Start Mailmerge", summary_msg)

    if start_mailmerge:
        # Create documents and save them at the desired location (hierarchy_root)
        project.create_client_documents(selected_clients, hierarchy_root, standard_pdfs)
    else:
        sys.exit(0)
