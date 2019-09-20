#!/usr/bin/python3

# Author : zhanjunkai
# Date   : 2019/09/08

from klib import ktp
import re
import openpyxl
import smtplib
import os
import shutil

from email.mime.text import MIMEText

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

# =================================================================================
# Function    : send_email(text, dst_mailbox, mailserver, port, src_mailbox, passwd)
# Description : sent text from src_mailbox on mailserver.port to 
#               dst_mailbox
# =================================================================================
def send_email(subject, ext, dst_mailbox, src_mailbox, passwd, mailserver='smtp.163.com', port=25):
    smtpobj = smtplib.SMTP(mailserver, port)
    smtpobj.login(src_mailbox, passwd)
    msg = MIMEText(text, 'plain', 'utf-8')
    msg['From'] = src_mailbox
    msg['To'] = dst_mailbox
    msg['Subject'] =  subject
    status = smtpobj.sendmail(src_mailbox, [dst_mailbox], msg.as_string())
    if status != {}:
        print('Error occured: From %s to %s, Status is %s' % (src_mailbox, dst_mailbox, status))
    else:
        print('The email is sent successfully from %s to %s' % (src_mailbox, dst_mailbox))
    smtpobj.quit()

# ========================================================
# Function    : organize_file_into_folder(file_type, path)
# Description : organize the file in file_type in path
#               into the folder or subfolder with the 
#               same name
# ======================================================== 
def organize_file_into_folder(file_type, path):
    os.chdir(path)
    files = []
    for obj in os.listdir():
        if obj.endswith(file_type):
            files.append(obj.split('.')[0])
    print("The files organized in %s are: " % path)
    print(files)
    for folder_name, subfolders, filenames in os.walk(path):
        current_folder = folder_name.split(os.path.sep)[-1]
        if current_folder in files:
            file_name = current_folder + "." + file_type
            shutil.move(file_name, folder_name)
            print("I have moved the file: %s into folder: %s" % (file_name, folder_name))


