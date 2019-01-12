#!/usr/bin/env python3
## PYTHON_ARGCOMPLETE_OK

"""Create a new libreoffice document and open it

EXAMPLES:

Create a new writer document with the same name as the template in
the current dir and open it:

  newlibreoffice.py blank_writer.odt

Create a new calc document with the name "My fancy spreadsheet.ods" in
the current and open it

  newlibreoffice.py calc_blank.ods "My fancy spreadsheet"

Like last example, but to not open the file after creating it:

  newlibreoffice.py calc_blank.ods "My fancy spreadsheet" -d

DETAILS

NOTE: This script will select from templates in your templates
folder. The name of the template should start with the program name
i.e. writer_... calc_... or impress_... to make it easier to
autocomplete

NOTE AUTOCOMPLETION: This script uses argcomplete to make it possible
to autocomplete the template names. It will need to be installed and
activated, see here: https://argcomplete.readthedocs.io/en/latest/.

Return codes:
 0 Success
 1 Missing xdg-user-dir command

"""

import os
from os import path
import sys
import shutil
import argparse
import subprocess
import argcomplete

# Template dir
try:
    template_dir_bytes = subprocess.check_output(
        "xdg-user-dir TEMPLATES",
        shell=True,
    )
    TEMPLATE_DIR = template_dir_bytes.decode(sys.getdefaultencoding()).strip()
except subprocess.CalledProcessError:
    print("This script requires the command xdg-user-dir")
    raise SystemExit(1)

#TEMPLATE_DIR = path.join(path.expanduser('~'), "Skabeloner")
EXTENSIONS = ('.odt', '.ods', '.odp')

# Find templates
templates = []
for filename in os.listdir(TEMPLATE_DIR):
    filename_extension = os.path.splitext(filename)[1]
    if filename_extension in EXTENSIONS:
        templates.append(filename)

# Create parser
parser = argparse.ArgumentParser(
    description=__doc__,
    formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument('template', choices=templates,)
parser.add_argument('name', default=None, nargs='?',
                    help='sum the integers (default: find the max)')
parser.add_argument('-d', '--do-not-open', action="store_true", default=False,
                    help='Don\'t open the file after creating it')
argcomplete.autocomplete(parser)
args = parser.parse_args()

# Determine destination name
if args.name:
    destination_name = args.name
    _, extension = path.splitext(args.template)
    if not destination_name.lower().endswith(extension):
        destination_name += extension
else:
    destination_name = args.template

# Form paths
source_path = path.join(TEMPLATE_DIR, args.template)
destination_path = path.join(os.getcwd(), destination_name)

# Create the file
if args.do_not_open:
    print("Creating file\n{}".format(destination_path))
else:
    print("Creating and opening file\n{}".format(destination_path))
shutil.copy(source_path, destination_path)

# Unless specifically requested, open the file on completion
if not args.do_not_open:
    subprocess.Popen(['libreoffice', destination_name])
