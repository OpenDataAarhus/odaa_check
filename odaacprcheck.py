from fileinput import filename

__author__ = 'arne'

import os.path
import fnmatch
import time
from datetime import timedelta
import datetime
import re
import zipfile
import layout_scanner
import xlrd
from ODSReader import *
from pprint import pprint
import sys
import rtf2xml
import smtplib
import csv
import zipfile
import subprocess


g_u_line        = '\n--------------------------------------------------------------------------------------------------------\n\n'
g_end_u_line    = '\n\n--------------------------------------------------------------------------------------------------------'
g_log_file_name = 'odaachecklog'

g_debug = False
g_error_template = "Error ==> Exception %s"
g_temp_csv_file = "/tmp/odaa_xlsx2csv.csv"

g_not_readable = '.dwg, .mdb, .zip,.wk3'
g_sender = 'odaa_check@odaa.aakb.dk'


#odaa_file_path = '/home/arne/newckan/data/files/'
odaa_file_path = '/home/ckan/ckan2/files/pairtree_root/'

#odaa_file_path= '/home/arne/Odaa_Cpr_testData/spec_test/'

def send_mail(subject, _body):
    
    to = ['xxx@xxx.xx','xxx@xxx.xx']

    body = _body.replace('\n', '\r\n')
    #print "sendeing mail"

    # Prepare actual message
    message = """\
From: %s
To: %s
Subject: %s
%s
""" % (g_sender, ", ".join(to), subject, body)
    try:
        smtpObj = smtplib.SMTP('localhost')
        smtpObj.sendmail(g_sender, to, message)
        print "Successfully sent email"
    except smtplib.SMTPException:
        print "Error: unable to send email"
        sys.exit(1)



def fix(data):
    if isinstance(data, unicode):
        return data.encode('utf-8')
    elif isinstance(data, dict):
        data = dict((fix(k), fix(data[k])) for k in data)
    elif isinstance(data, list):
        for i in xrange(0, len (data)):
            data[i] = fix(data[i])
    return data



#
def listFiles(root, patterns='*', recurse=1, return_folders=0):

    #Begin check existing of log file, if not found the is the first time we are checking go looong time back
    if (os.path.isfile(g_log_file_name)) and (not g_debug):
        yesterday = datetime.datetime.now() - timedelta(1)
    else:
        yesterday = datetime.datetime.now() - timedelta(10000)

    #print yesterday
    #print yesterday.toordinal()
    yesterday_epoc = time.mktime(yesterday.timetuple())
    # Expand patterns from semicolon-separated string to list
    pattern_list = patterns.split(';')
    # Collect input and output arguments into one bunch

    class Bunch:
        def __init__(self, **kwds): self.__dict__.update(kwds)

    arg = Bunch(recurse=recurse, pattern_list=pattern_list, return_folders=return_folders, results=[])

    def visit(arg, dirname, files):
        # Append to arg.results all relevant files (and perhaps folders)
        for name in files:
            #Checking:" + newfile
        #    print name

            fullname = os.path.normpath(os.path.join(dirname, name))
            if arg.return_folders or os.path.isfile(fullname):
                for pattern in arg.pattern_list:
                    if fnmatch.fnmatch(name, pattern):

                        if os.path.getmtime(fullname) - yesterday_epoc > 0:
                            # arg.results.append(fullname)
                            arg.results.append(fullname)
                            #arg.results.append( "NEW: %s modified: %s" % (fullname, time.ctime(os.path.getmtime(fullname))))
                       # else:
                       #     arg.results.append(
                       #         "OLD: %s modifyed: %s" % (fullname, time.ctime(os.path.getmtime(fullname))))
                        break

                        # Block recursion if recursion was disallowed
        if not arg.recurse: files[:] = []

    os.path.walk(root, visit, arg)

    return arg.results




def process_pdf(path):
    """Extract text from PDF file using PDFMiner with whitespace inatact."""
    str = ""
    try:
        pages = layout_scanner.get_pages(path)
        i = 0
        l = len(pages)
        while i < l:
            str += pages[i]
            i += 1
    except Exception, e:
        return g_error_template % e, ""

    return "", str


def process_docx(newfile):

    with open(newfile) as f:
        unzip = zipfile.ZipFile(f)
        xml_content = unzip.read('word/document.xml')

    return "", xml_content


def process_doc(newfile):
    proc = subprocess.Popen(['/usr/bin/antiword', newfile], stdout=subprocess.PIPE)

 #    error handling missing
    contents = ""
    for line in iter(proc.stdout.readline, ''):
        contents += line.rstrip()

    return "", contents


def process(newfile):
    print "in process: %s " % filename

def process_xls(newfile):

    print "now ind process_xls"

    print "start process_xls file %s" % newfile

    try:
        book = xlrd.open_workbook(newfile)
        print "open book"
        contents = ""
        sheets = book.sheet_names()
    except Exception, e:
        print " except  " + repr(e)
        return g_error_template % e, ""


    print "trace sheet name"
    pprint(sheets)
    try:
        for sheet_name in sheets:
            worksheet = book.sheet_by_name(sheet_name)
            print " handle sheet " + sheet_name
            for row_index in xrange(worksheet.nrows):
                for coll_index in xrange ( worksheet.ncols):
                    #contents += worksheet.cell_value(row_index, coll_index).encode('utf-8').strip() + " "
                    #print "type : %d" % worksheet.cell_type(row_index, coll_index)
                    #print " val :" + str( worksheet.cell_value(row_index, coll_index) )

                    ctype = worksheet.cell_type(row_index, coll_index)
                    if ctype == xlrd.XL_CELL_EMPTY or ctype == xlrd.XL_CELL_BLANK:
                        continue

                    if ctype == xlrd.XL_CELL_TEXT:
                        contents += worksheet.cell_value(row_index, coll_index).encode('utf-8') + " "
                    else:
                        contents += str(worksheet.cell_value(row_index, coll_index)  ) + " "
    except Exception, e:
        print " except  " + repr(e)
        return g_error_template % e, ""


    return "", contents



def process_xlsx(newfile):

    error = ""
    print "start process_xlsx file %s " % newfile
    try:

        try:
            wb = xlrd.open_workbook(newfile,  ragged_rows=True)
           #wb = xlrd.open_workbook(newfile)

        except Exception, e:
            print "===> 01. Execp: %s " % e
            return g_error_template % e, ""

        print wb.biff_version, wb.codepage, wb.encoding
        #print "after xlrd.open_workbook "

        sheet_name = wb._sheet_names
       # print ','.join(sheet_name)
        print "ready to go"
        with open(g_temp_csv_file, 'w') as csv_file:

            wr = csv.writer(csv_file, quoting=0 )

            #print "---- runningh all the rows ----"
            for name in sheet_name:
                #print "process %s" % name
                sh = wb.sheet_by_name(name)

                #print "Before row process process_xlsx"
                row = " "
                try:
                    for rownum in xrange(sh.nrows):

                       # rw = sh.row_values(rownum)
                        row = fix(sh.row_values(rownum))
                        # row = ''.join(unicode(str(e), errors='ignore') for e in rw)
                        #wr.writerow(sh.row_values(rownum).unicodeData.encode('ascii', 'ignore'))
                        wr.writerow(row)
                except Exception, e:
                    #row = ''.join(row).encode(encoding='utf-8', error='ignore')
                    print "except 002"
                    print "Execp 002: %s  %s" % (e, row)


            csv_file.close()

        with open(g_temp_csv_file, 'r') as tfile:
            contents = tfile.read()
            tfile.close()
        return " ", contents.strip('\n')

    except Exception, e:
        print "Execp 2: %s " % e
        return g_error_template % e, ""



"""
process OpenDocument Spreadsheet
"""
def process_ods(newfile, reg=None):

    doc = ODSReader(newfile)

    #print doc.SHEETS
    #print ','.join(doc.SHEEexceptTS)
    #print doc.SHEETS.keys()

    contents = ""
    if reg <> None:
        for key in doc.SHEETS.keys():
            thisrow =  ' '.join(' '.join(u''.join(el) for el in list) for list in doc.SHEETS[key])
            contents += u''.join(reg.findall(thisrow) )
    else:
        for key in doc.SHEETS.keys():
            #print doc.SHEETS[key]
            #contents += u' '.join(" ".join(map(str, l)).decode('ascii', 'ignore') for l in doc.SHEETS[key])
            contents += '\n'.join('\t'.join(u''.join(el) for el in list) for list in doc.SHEETS[key])
            #print contents
            #l =  doc.SHEETS[key]
            #print type(l)
    return "", contents

##
## process OpenDocument writer
##

def process_odt(newfile):

    return " ", odf.opendocument.load(newfile).xml()




def process_zip(newfile):

    return " ", ""

"""
    with open(newfile, 'rb') as fh:
        z = zipfile.ZipFile(fh)
        for name in z.namelist():
            outP = '/tmp/'
            z.extract(name, outP)
    fh.close()


    with zipfile.ZipFile(newfile, 'r') as myzip:
        contents = myzip.filelist
        pprint( contents)

        count_member = len(myzip.namelist())
        if count_member > 1:
            pass
        myzip.extractall(g_temp_exp_zip)
    return " ", " "

"""


import rtf2xml.ParseRtf as RtfParse
def process_rtf(newfile):

    try:
        parse_obj =rtf2xml.ParseRtf.ParseRtf(
            in_file = newfile,

            # these are optional

            # determine the output file
            out_file = '/tmp/out.xml',

            # Form lists from RTF. Default is 1.
            form_lists = 1,

            # Convert headings to sections. Default is 0.
            headings_to_sections = 1,

            # Group paragraphs with the same style name. Default is 1.
            group_styles = 1,

            # Group borders. Default is 1.
            group_borders = 1,

            # Write or do not write paragraphs. Default is 0.
            empty_paragraphs = 0,
        )

        parse_obj.parse_rtf()

        with open('/tmp/out.xml') as wb:
            contents = wb.read()
        return " ", contents

    except rtf2xml.ParseRtf.InvalidRtfException, msg:
        sys.stderr.write(str(msg))
        return g_error_template % msg, ""

    except rtf2xml.ParseRtf.RtfInvalidCodeException, msg:
        sys.stderr.write(str(msg))
        return g_error_template % msg, ""


def handle_newfile(filelist):

    """

    :param filelist:
    """
    r = '\d{5,6}[- ]\d{4} '
    reg = re.compile(r)


    abstract = ""
    count_cpr = 0


    #print "start handle_newfile"
    logfile = open(g_log_file_name, 'a')
    logfile.write("\n\n -- Start cpr check" + datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))
    logfile.write(g_u_line)
    file_ext = ""
    fileNo = 0
    error_in_read = ""
    try:

        logfile.write("\nFiler to check : %d" % len(filelist))

        for newfile in filelist:
            fileNo += 1
            loglines = ""
            fileName, fileExtension = os.path.splitext(newfile)
            file_size = os.path.getsize(newfile)
            filetext = " "

            logfile.write("\nChecking: %s size: %d" % (newfile, len(filetext)))
            print " file: %s ext %s" % (fileName, fileExtension)

            #collecting ext, just fore info in logfile
            if fileExtension not in file_ext:
                file_ext += ' ' + fileExtension

            # there are some types we know that we can't extract
            if fileExtension in g_not_readable:
                print "chekc global file ext ext %s  file %s " % (fileExtension, newfile)
                logfile.write("\n\n ==> File type %s kan ikke auto kontrollers\n" % (newfile))
                abstract += "\n ==> File type %s kan ikke auto kontrollers" % (newfile)
                continue

            if fileExtension == '.pdf':
                error_in_read, filetext = process_pdf(newfile)

            elif fileExtension == '.docx':
                #Check m$ docc
                error_in_read, filetext  = process_docx(newfile)

            elif fileExtension == '.doc':
                #Check m$ doc
                error_in_read, filetext  = process_doc(newfile)

            elif fileExtension == '.odt':
                #open office writer
                error_in_read, filetext  = process_odt(newfile)

            elif fileExtension == '.xlsx':
                 #Check m$ excel
                error_in_read, filetext  = process_xlsx(newfile)

            elif fileExtension == '.rtf':
                #continue
                error_in_read, filetext  = process_rtf(newfile)
                #print filetext
#
            elif fileExtension == '.xls':
                #Check m$ old excel
                error_in_read, filetext = process_xls(newfile)

            elif fileExtension == '.ods':
                #open office calc
                if file_size > 4000000:
                    print "File %s is too large (%d byte) to auto control\n" % (newfile, file_size)
                    logfile.write("\n\n ==> NB NB File %s is too large (%d byte) to auto control\n" % (newfile, file_size))
                    abstract += "\n ==>  File %s is too large (%d byte) to auto control" % (newfile, file_size)
                    continue
                else:
                    error_in_read, filetext  = process_ods(newfile)

            elif fileExtension == '.zip':
                #not in use cfr. g_not_readable
                filetext = process_zip(newfile)
                print filetext

            else:  # all other filetypes and default handling
                inf = open(newfile, "r")
                filetext = inf.read()
                inf.close()

            # if content in "error_in_read" we have an error
            if len(error_in_read) > 5:
                logfile.write("\nNB NB == > Can not read : %s size: %d" % (newfile, len(filetext)) )
                abstract += "\n == > Can not read : %s size: %d %s " % (newfile, len(filetext), error_in_read)
                error_in_read = ""
                continue

            # now we can check for cpr
            fundet = reg.findall(filetext)

           #fundet = reg.findall(filetext, pos=0, endpos=-1)
            #filetext = "" # try to release memory

            print " Found  %s" % fundet

            for f in fundet:
                checkity = f[:6].strip(' ').strip('-')
                try:
                    datetime.datetime.strptime(checkity, '%d%m%y')
                    #print "valid date %s possibly cpr: %s" % (checkity, f)
                    loglines += "\n   ==> NB NB possibly cpr: %s filename %s \n" % (f, newfile)
                    count_cpr += 1
                    abstract += "\n ==>  %s found in filename %s" % (f, newfile)

                except ValueError, e:
                    pass

            logfile.write("\n*********** ######################## ***************")

            logfile.write("\n done handling no: %d file: %s " % (fileNo, newfile) )

            print "done handling %s No : %d " % (newfile, fileNo)

            if loglines != "":
                logfile.write('\n')
                logfile.write(loglines)



            logfile.flush()

        logfile.write('\n\nprocessed file types:' + file_ext)

        logfile.write(g_end_u_line)
        logfile.write("\n -- End cpr check " + datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))

        logfile.write("\n\n.......  summary ........")
        logfile.write("\n\nFound %d possibly cprno. or file not able to read" % count_cpr)
        logfile.write(abstract)
        logfile.write("\n********** end summary **************")
        logfile.flush()

    except Exception, e:
       print repr(e)
       print " e ==" + repr(e)
       logfile.close()
       sys.exit(1)

    logfile.close()

    if not g_debug:
        if len(abstract) < 1:
            send_mail("Nothing to do", abstract)
        else:
            send_mail("NB CHECK " ,  abstract)


def main():

    thefiles = listFiles(odaa_file_path, "*", recurse=1, return_folders=0)

    #print '\n'.join(thefiles)

    handle_newfile(thefiles)

    #pprint ( thefiles )


if __name__ == '__main__':
    main()

