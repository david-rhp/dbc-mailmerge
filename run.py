from tkinter import filedialog
from dbcmailmerge.mailproject import MailProject
from dbcmailmerge.constants import FIELD_MAP_PROJECT, FIELD_MAP_CLIENTS
import sys
from pathlib import Path
from xlrd import XLRDError

import tkinter as tk
from tkinter import messagebox


ABORT_KEYWORDS = ('q', "quit")
# First element of the tuple is an explanation,the second a key to a filter
FILTER_MENU = {0: ("all filters selected",), 1: ("filter by amount >= 0", "amount")}


def prompt_str(prompt="Please enter a value or type `q` or `quit` to abort"):
    while True:
        value = input(prompt + ": ")

        if value.lower() in ABORT_KEYWORDS:
            sys.exit(1)

        if not value:
            print("You didn't provide any input. Please try again.\n")
        else:
            return value


def prompt_data_source(data_kind, prompt_source_file=True, prompt_sheet_name=True):
    data_source = None
    if prompt_source_file:
        messagebox.showinfo("Select Data Source", "Select the data source for the project and the respective clients.")
        data_source = filedialog.askopenfilename()
        root.update()

    sheet_name = None
    if prompt_sheet_name:
        sheet_name = prompt_str(f"Please provide the sheet_name of the {data_kind} data source")

    return data_source, sheet_name


def prompt_boolean(prompt="Please enter 0 or 1 to indicate True or False respectively"):
    """
    Prompts a user to enter boolean represented by 0 or 1 and returns the result.

    Parameters
    ----------
    prompt: str, optional
        The text displayed when prompting the user for input (default: generic boolean input request).

    Returns
    -------
    bool
        True if the user entered 1, False if 0. Any other input value: the user gets re-prompted or the program
        is terminated.

    Warnings
    --------
    The user can terminate the program by entering one of the ABORT_KEYWORDS.
    """
    while True:
        user_input = input(prompt + ": ").strip()

        if user_input.lower() in ABORT_KEYWORDS:
            print("The user has terminated the program.")
            sys.exit(0)
        elif len(user_input) != 1:
            print("You must enter exactly 1 value, either `1` or `0` or `quit` to terminate the program")
            continue
        else:
            if user_input in ('0', '1'):
                user_input = bool(int(user_input))
                return user_input


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


def print_dictionary(dictionary):
    for key in dictionary:
        print(f"Option {key}: {dictionary[key][0]}")


def select_filter():
    selection = None

    filters = {"amount": lambda x: bool(x)}

    selected_filters = {}
    while selection != 0 and not selected_filters:

        print()
        print_dictionary(FILTER_MENU)
        selection = prompt_int("Please select an option from the menu")

        if selection not in FILTER_MENU:
            print(f"This option does not exist, please select one of the following {FILTER_MENU.keys()}\n")

        # add filter
        elif selection != 0:
            key = FILTER_MENU[selection][1]  # choose second element of tuple
            selected_filters[key] = filters[key]

    return selected_filters


def prompt_files():
    messagebox.showinfo("Select Files", "Please select the standardized documents that you would like to include.")

    selected_files = []
    while True:
        file = filedialog.askopenfile()
        root.update()

        selected_files.append(file.name)  # append filename instead of io.TextWrapper

        result = messagebox.askyesno("Select File", "Do you want to select another file?")
        root.update()

        if not result:
            return selected_files


def create_project_or_clients(data_source, data_kind, project_object=None):
    while True:
        _, data_sheet_name = prompt_data_source(data_kind, prompt_source_file=False)
        try:
            if project_object:
                project_object.create_clients(data_source, data_sheet_name, FIELD_MAP_CLIENTS)
            else:
                project_object = MailProject.from_excel(data_source, data_sheet_name, FIELD_MAP_PROJECT)
        except XLRDError:
            print(f"Couldn't find the sheet name you specified for {data_kind} data, please try again.\n")
        else:
            break

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

    # Get data source, currently 1 excel sheet per project including both the client and the project data
    data_source, _ = prompt_data_source("client", prompt_sheet_name=False)

    project = create_project_or_clients(data_source, "project")
    project = create_project_or_clients(data_source, "client", project)

    selection_criteria = select_filter()
    selected_clients = project.select_clients(selection_criteria)

    standard_pdfs = prompt_files()
    print(standard_pdfs)
    # Create documents and merge
    project.create_client_documents(selected_clients, hierarchy_root, standard_pdfs)
