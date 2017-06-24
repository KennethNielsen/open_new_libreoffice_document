#!/usr/bin/env python3

"""Create a new libreoffice document and open it

EXAMPLES:

Creates a new writer document with the same name as the template in
the current dir and open it:

  newlibreoffice.py writer

Create a new calc document with the name "My fancy spreadsheet.ods" in
the current and open it

  newlibreoffice.py calc "My fancy spreadsheet"

Like last example, but to not open the file after creating it:

  newlibreoffice.py calc "My fancy spreadsheet" -d

DETAILS

NOTE: In order for the script to work, your templates folder must
contain a template for each of the document types that you wish to
open (writer, calc and impress) with that word (e.g. "writer") in the
file name (e.g. "New writer document.odt").

Return codes:
 0 Success
 1 Missing xdg-user-dir command
 2 No template found for requested type

"""

import os
from os import path
import sys
import shutil
import argparse
import subprocess


# Template dir
try:
    template_dir_bytes = subprocess.check_output(
        "xdg-user-dir TEMPLATES",
        shell=True,
    )
    TEMPLATE_DIR = template_dir_bytes.decode(sys.getdefaultencoding())
except subprocess.CalledProcessError:
    print("This script requires the command xdg-user-dir")
    raise SystemExit(1)

TEMPLATE_DIR = path.join(path.expanduser('~'), "Skabeloner")

parser = argparse.ArgumentParser(
    description=__doc__,
    formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument('type', choices=['writer', 'calc', 'impress'],
                    default='writer',
                    help='The type of document to create')
parser.add_argument('name', default=None, nargs='?',
                    help='sum the integers (default: find the max)')
parser.add_argument('-d', '--do-not-open', action="store_true", default=False,
                    help='Don\'t open the file after creating it')

args = parser.parse_args()

# Determine source name
for source_name in os.listdir(TEMPLATE_DIR):
    if args.type in source_name:
        break
else:
    print("No template found for type {} in folder {}"\
          .format(args.type, TEMPLATE_DIR))
    raise SystemExit(2)

# Determine destination name
if args.name is None:
    destination_name = source_name
else:
    destination_name = args.name
    _, extension = path.splitext(source_name)
    if not destination_name.lower().endswith(extension):
        destination_name += extension

# Form paths
source_path = path.join(TEMPLATE_DIR, source_name)
destination_path = path.join(os.getcwd(), destination_name)

# Create the file
print("Creating and opening file\n{}".format(destination_path))
shutil.copy(source_path, destination_path)

# Unless specifically requested, open the file on completion
if not args.do_not_open:
    subprocess.Popen(['libreoffice', '--' + args.type, destination_name])
