==================
'pydict2xlsx' task
==================

'pydict2xlsx' task converts a Python dictionary to Microsoft Excel file
or CSV file.

Installation
------------

Before installing 'pydict2xlsx' task, please make sure that 'pyloco' is installed.
Run the following command if you need to install 'pyloco'.

>>> pip install pyloco

Or, if 'pyloco' is already installed, upgrade 'pyloco' with the following command

>>> pip install -U pyloco

To install 'pydict2xlsx' task, run the following 'pyloco' command.

>>> pyloco install pydict2xlsx

Command-line syntax
-------------------

usage: pyloco pydict2xlsx [-h] [-t type] [-o OUTPUT] [--general-arguments]
                          data

converts Python dictionary to Microsoft Excel file

positional arguments:
  data                  input Python dictionary

optional arguments:
  -h, --help            show this help message and exit
  -t type, --type type  output file format (default='xlsx')
  -o OUTPUT, --output OUTPUT
                        generate output file
  --general-arguments   Task-common arguments. Use --verbose to see a list of
                        general arguments


Example(s)
----------

Current version of the task assumes that 'docx2text' pyloco task is used as
a previous task as shown below.

First, make sure that 'docx2text' task is installed by running the following
command.

>>> pyloco docx2text -h

You will see a help message of 'docx2text' on screen. If failed, please use
following command to install.

>>> pyloco install docx2text

Follwoing command reads my.docx MS World file and convert tables in the file
to worksheets of MS Excel file.

>>> pyloco docx2text my.docx -- pydict2xlsx -t xlsx
tables.xlsx

Follwoing command reads my.docx MS World file and convert tables in the file
to CSV format text file. 

>>> pyloco docx2text my.docx -- pydict2xlsx -t csv
tables.csv
