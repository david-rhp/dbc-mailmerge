"""
Author: James Mills
Thank you for publishing this code.

Description
-----------
Converts a doc file to PDF using Windows' COM objects

References
----------
Source of this file
    https://stackoverflow.com/questions/30481136/pywin32-save-docx-as-pdf

"""
import comtypes.client

wdFormatPDF = 17  # code for PDF format


def covx_to_pdf(infile, outfile):
    """Convert a Word .docx to PDF"""

    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(infile)
    doc.SaveAs(outfile, FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
