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

The project uses pytest. Please note that the created documents are not automatically tested at the moment (the correct formatting, structure, etc.). 

## Configurations

This project is designed to fit a specific business need and has the appropriate business logic in the main class. However, it is not too difficult to adjust it for a different need.

### General adjustments
Most options can be adjusted in [config.py](./dbcmailmerge/config.py). Please have a look at the module's docstring to see further explanations.

### Formatting
If you want to change the formatting of the data that should populate the placeholders in the word templates, please adjust the `MailProject.__format_client_records` method.

If you want to change the filename of the output PDF make adjustments in the `MailProject.__create_client_document` method.

## Acknowledgments
* [Michal Zalecki](https://michalzalecki.com/converting-docx-to-pdf-using-python/) who published code for converting docx file to PDF. See: [docx2pdfconverter.py](./dbcmailmerge/docx2pdfconverter.py)

