#!/usr/bin/python3

# Author : zhanjunkai
# Date   : 2019/09/08

from klib import ktp
import re
import openpyxl

# =============================================================
# Function    : replace_in_file(file_name, pattern, new_string)
# Description : Replace the string which are matched by the
#               pattern with the new string in file
# =============================================================
def replace_in_file(file_name, pattern, new_string):
    ktp.check_file(file_name)
    file_in     = open(file_name, 'r')
    content     = file_in.read()
    content     = re.sub(pattern, new_string, content)
    file_out    = open(file_name, 'w')
    file_out.write(content)
    file_in.close()
    file_out.close()

# =============================================================
# Function    : delete_in_file(file_name, pattern)
# Description : Delete the string which are matched by the
#               pattern with the new string in file
# =============================================================
def delete_in_file(file_name, pattern):
    replace_in_file(file_name, pattern, '')


