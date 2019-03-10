# dbc-mailmerge

A project to create one single PDF per recipient by customizing different word templates with each respective recipient's data, converting the files to PDF, and merging them.

## Prerequisites

In its current iteration, this project requires LibreOffice to be installed on your machine in order to convert the created docx files to PDF.

[Download LibreOffice](https://www.libreoffice.org/download/download/)

## Dependencies

Please install all the dependencies from the [requirements.txt](./requirements.txt) in the root directory of the project.

```
pip install -r requirements.txt
```

## Testing

The project uses pytest. Please note that the created documents are not automatically tested at the moment (the correct formatting, structure, etc.). This means the final output has to be verified visually at the moment. Please check the function's docstrings to verify if that's the case. 

Simple dummy data is found in [test_constants.py](./tests/test_constants.py). The tests also use dummy data from an excel sheet, see [test_data_source.xlsx](./data/tests/test_data_source.xlsx). A file similar to this would be used when using this project in production. Each row, represents data pertaining to a particular client. Do not change the sheet_names as they are statically stored in the test suite.

Please note, that it is assumed that the data of the excel file is clean, i.e., not a lot of error checking is performed. This is because the validity of the file is checked beforehand in our company. Also, the data of the file isn't used for any computations and is mostly just formatted to populate a placeholder in a word template.

**Exception**: The data format used for the filtering in the `MailProject.select_clients` method is relevant. Currently, there is only one attribute for which a filter can be applied (amount). See the `select_filter` function in [run.py](run.py). 

## Configuration

This project is designed to fit a specific business need and has the appropriate business logic in the main class. However, it is not too difficult to adjust it for a different need.

### General adjustments
Most options can be adjusted in [config.py](./dbcmailmerge/config.py). Please have a look at the module's docstring to see further explanations.

### Formatting
If you want to change the formatting of the data that should populate the placeholders in the word templates, please adjust the `MailProject.__format_client_records` method.

If you want to change the filename of the output PDF make adjustments in the `MailProject.__create_client_document` method.

If you want to change the formatting of the MailProject data, which is also used for populating word templates, please make adjustments to the `MailProject.__create_project_record` method.

## Acknowledgments
* [Michal Zalecki](https://michalzalecki.com/converting-docx-to-pdf-using-python/) who published code for converting docx files to PDF. See: [docx2pdfconverter.py](./dbcmailmerge/docx2pdfconverter.py)

