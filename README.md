# dbc-mailmerge

A project to create one single PDF per recipient by customizing different word templates with each respective recipient's data, converting the files to PDF, and merging them.


## Prerequisites

In its current iteration, this project requires LibreOffice to be installed on your machine in order to convert the created docx files to PDF.

[Download LibreOffice](https://www.libreoffice.org/download/download/)

*Windows users*
Ensure that the `LibreOffice\program` folder is in your PATH. The converter will look for `soffice` on Windows and requires the preceding path to be in the user's PATH.

For example, the qualified path for LibreOffice is  something like this (by default): 

```
C:\Program Files\LibreOffice\program
```

which contains the `soffice` executable.


## Dependencies

Please install all the dependencies from the [requirements.txt](./requirements.txt) in the root directory of the project.

```
pip install -r requirements.txt
```

## Usage
To run this program, simply execute [run.py](run.py). The user will be prompted to select the appropriate files using tkinter message boxes, file dialogues, and the console.

## Testing

### General Instructions

The project uses pytest. Please note that the created documents are not automatically tested at the moment (the correct formatting, structure, etc.). This means the final output has to be verified visually at the moment. Please check the function's docstrings to verify if that's the case. The documents are stored in [./data/tests/client_correspondence](./data/tests/client_correspondence). If `client_correspondence` is not available, please run the test suite. It will automatically create the directory structure and save the files that have been created and which need manual checking. 

New test runs will overwrite existing files, but only if the respective file is created again, i.e., the suite only creates new directories, if they do not exist yet, and overwrites old files with new files. It does not delete the directory in its entirety beforehand. For a clean result, delete the folder for each run. 

**Side Note:** To test if the filter properly excludes particular clients when creating the documents, run the `test_create_client_documents_with_filter` and the `test_create_client_documents_without_filter` in [test_mailproject.py](./tests/test_mailproject.py) separately from one another, as they produce different results. If you run both at once, the client_correspondence will also have documents that wouldn't have been created by the test function using a filter. Furthermore, both pass even if the output documents' contents are wrong (hence the requirement for the visual checking). They only fail if an exception is raised at runtime. **Do not assume proper output. Check the content of the created files manually.**


### Data

Simple dummy data is found in [test_constants.py](./tests/test_constants.py). The tests also use dummy data from an excel sheet, see [test_data_source.xlsx](./data/tests/test_data_source.xlsx). A file similar to this would be used when using this project in production. Each row, represents data pertaining to a particular client. Do not change the sheet_names as they are statically stored in the test suite.

Please note, that it is assumed that the data of the excel file is clean, i.e., not a lot of error checking is performed. This is because the validity of the file is checked beforehand in our company. Also, the data of the file isn't used for any computations and is mostly just formatted to populate a placeholder in a word template.

**Exception**: The data format used for the filtering in the `MailProject.select_clients` method is relevant. Currently, there is only one attribute for which a filter can be applied (amount). See the `select_filter` function in [run.py](run.py). 


## Configuration

This project is designed to fit a specific business need and has the appropriate business logic in the main class. However, it is not too difficult to adjust it for a different need. Please update the test suite to match your adjustments.


### General adjustments
Most options can be adjusted in [config.py](./dbcmailmerge/config.py). Please have a look at the module's docstring to see further explanations.


### Formatting
If you want to change the formatting of the data that should populate the placeholders in the word templates, update the `MailProject.__format_client_records` method.

If you want to change the filename of the output PDF make adjustments in the `MailProject.__create_client_document` method.

If you want to change the formatting of the MailProject data, which is also used for populating word templates, make adjustments to the `MailProject.__create_project_record` method.

For making changes to the current UI, see [run.py](run.py).

## Known issues

The current way of converting the docx files to PDF in the `MailProject.__create_client_document` method is slow. However, this is not a critical issue for our business, because the to be processed client records never exceeds 150 clients (for legal reasons).

In the future, possible solutions include using a different conversion method, such as using the [pywin32](https://pypi.org/project/pywin32/) to directly use the Windows API for the conversion process (caveat, Windows only). The conversion works, but the speed difference would have to be measured.

Alternatively, the conversion process could be done on a server, as described [here](https://michalzalecki.com/converting-docx-to-pdf-using-python/). 

Another option might be using some form of concurrency.


## Acknowledgments
* [Michal Zalecki](https://michalzalecki.com/converting-docx-to-pdf-using-python/) who published code for converting docx files to PDF. See: [docx2pdfconverter.py](./dbcmailmerge/docx2pdfconverter.py)

