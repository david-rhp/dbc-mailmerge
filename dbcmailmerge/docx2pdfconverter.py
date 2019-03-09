"""
Author: Michal Zaleki
Thank you for publishing this code.

Description
-----------
Provides a wrapper to use the LibreOffice API, with which LibreOffice-compatible file types can be converted.

References
----------
Source of this file
    https://michalzalecki.com/converting-docx-to-pdf-using-python/
How to use LibreOffice in a CLI
    https://help.libreoffice.org/Common/Starting_the_Software_With_Parameters
"""
import sys
import subprocess
import re


def convert_to(folder, source, timeout=None):
    args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]

    process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
    filename = re.search('-> (.*?) using filter', process.stdout.decode())

    if filename is None:
        raise LibreOfficeError(process.stdout.decode())
    else:
        return filename.group(1)


def libreoffice_exec():
    if sys.platform == 'darwin':
        return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    return 'libreoffice'


class LibreOfficeError(Exception):
    def __init__(self, output):
        self.output = output


if __name__ == '__main__':
    print('Converted to ' + convert_to(sys.argv[1], sys.argv[2]))
