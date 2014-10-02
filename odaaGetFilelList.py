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



g_u_line = '\n--------------------------------------------------------------------------------------------------------\n\n'
g_end_u_line = '\n\n--------------------------------------------------------------------------------------------------------'
g_log_file_name = 'odaachecklog'



#odaa_file_path = '/home/arne/newckan/data/files/'
odaa_file_path = '/home/ckan/ckan2/files/pairtree_root/'

#odaa_file_path= '/home/arne/Odaa_Cpr_testData/spec_test/'


#
def listFiles(root, patterns='*', recurse=1, return_folders=0):

    #Begin check existing of log file, if not found the is the first time we are checking go looong time back
    if os.path.isfile(g_log_file_name):
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
            i = i + 1

    except:
        print "Error load pdf"
        sys.exit(1)

    return str


def process_docx(filename):

    with open(filename) as f:
        unzip = zipfile.ZipFile(f)
        xml_content = unzip.read('word/document.xml')

    return xml_content



def getDataFromFile(fileName):
    with xlrd.open_workbook(fileName) as wb:
        # we are using the first sheet here
        worksheet = wb.sheet_by_index(0)
        # getting number or rows and setting current row as 0 -e.g first
        num_rows, curr_row = worksheet.nrows - 1, 0
        # retrieving keys values(first row values)
        keyValues = [x.value for x in worksheet.row(0)]
        # building dict
        data = dict((x, []) for x in keyValues)
        # iterating through all rows and fulfilling our dictionary
        while curr_row < num_rows:
            curr_row += 1
            for idx, val in enumerate(worksheet.row(curr_row)):
                print " cell type  %d " % val.ctype
                pprint(val)
                if val.ctype <> 2 and val.ctype <> 3:
                    if val.value.strip():
                        data[keyValues[idx]].append(val.value)
                else:
                        data[keyValues[idx]].append(str(val.value))

        return data

def process_xls(filename):

    #return getDataFromFile(filename)


    book = xlrd.open_workbook(filename)

    contents = ""
    sheets = book.sheet_names()

    for sheet_name in sheets:
        worksheet = book.sheet_by_name(sheet_name)
        #worksheet.ragged_rows
        for row_index in xrange(worksheet.nrows):
            for coll_index in xrange ( worksheet.ncols):
                #contents += worksheet.cell_value(row_index, coll_index).encode('utf-8').strip() + " "
                #print "type : %d" % worksheet.cell_type(row_index, coll_index)
                #print " val :" + str( worksheet.cell_value(row_index, coll_index) )

                ctype = worksheet.cell_type(row_index, coll_index)

                if  ctype <> 2 and ctype <> 3:
                    contents += worksheet.cell_value(row_index, coll_index).encode('utf-8') + " "
                else:
                    contents += str(worksheet.cell_value(row_index, coll_index)  ) + " "


    return contents

"""
process OpenDocument Spreadsheet
"""
def process_ods(newfile, reg=None):

    doc = ODSReader(newfile)

    #print doc.SHEETS
    #print ','.join(doc.SHEEexceptTS)
    #print doc.SHEETS.keys()
    return newfile
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
    return contents

"""
process OpenDocument writer
"""
def process_odt(newfile):

    #doc = odf.opendocument.load(newfile)
    return odf.opendocument.load(newfile).xml()


import zipfile

def process_zip(newfile):

    with zipfile.ZipFile(newfile, 'r') as myzip:
        contents = myzip.read('*')
        count_member = len(myzip.namelist())
        if count_member > 1:
            pass
        myzip.extractall('/tmp/odaa_check/zip')
    return contents



    zip = zipfile.ZipFile.open(newfile)
    zip.read()
    contents =  zip.read()
    zip.close()

    print contents

    return contents

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
        return contents

    except rtf2xml.ParseRtf.InvalidRtfException, msg:
        sys.stderr.write(str(msg))
        sys.exit(1)
    except rtf2xml.ParseRtf.RtfInvalidCodeException, msg:
        sys.stderr.write(str(msg))
        sys.exit(1)






def handle_newfile(filelist):

    """

    :param filelist:
    """
    r = '\d{5,6}[- ]\d{4} '
    reg = re.compile(r)
    fundet = []
    abstract = ""
    g_not_readable = '.dwg, .zip, .doc, .mdb'

    print "start handle_newfile"
    logfile = open(g_log_file_name, 'a')
    logfile.write("\n\n -- Start cpr check" + datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))
    logfile.write(g_u_line)
    file_ext = ""
    try:

        for newfile in filelist:
            #logfile.write("\nChecking:" + newfile)
            loglines = ""
            fileName, fileExtension = os.path.splitext(newfile)
            file_size = os.path.getsize(newfile)

            filetext = " "

            print " file size %d" % file_size

            if fileExtension not in file_ext:
                file_ext += ' ' + fileExtension

            print "checker %s" %newfile

            if fileExtension in g_not_readable:
                print "chekc global file ext ext %s  file %s " % (fileExtension, newfile)
                logfile.write("\n\n ==> File type %s kan ikke auto kontrollers\n" % (newfile))
                abstract += "\n ==> File type %s kan ikke auto kontrollers" % (newfile)
                continue

            if fileExtension == '.pdf':
                #continue
                filetext = process_pdf(newfile)

            elif fileExtension == '.docx':
                filetext  = process_docx(newfile)
                #print filetext

            elif fileExtension == '.odt':
                print "Check odt"
                filetext  = process_odt(newfile)

            elif fileExtension == '.xlsx':
                #continue
                filetext  = process_xls(newfile)
                #print filetext

            elif fileExtension == '.rtf':
                #continue
                filetext  = process_rtf(newfile)
                #print filetext
#
#            elif fileExtension == '.zip':
#
#                #filetext  = process_zip(newfile)
#                #print filetext
#                logfile.write("\n\n ==> NB NB Zip-file %s  kan ikke  auto kontrollers\n" % (newfile))#
#
#                continue

            elif fileExtension == '.ods':

                if file_size > 4000000:
                    print "NB NB File %s er for stor (%d) til auto kontrol" % (newfile, file_size)
                    logfile.write("\n\n ==> NB NB File %s er for stor (%d) til auto kontrol\n" % (newfile, file_size))
                    abstract += "\n ==> File %s er for stor (%d) til auto kontrol" % (newfile, file_size)
                    continue
                filetext  = process_ods(newfile)

            else:  # all other filetypes and default handling
                inf = open(newfile, "r")
                filetext = inf.read()
                inf.close()

            logfile.write("\nChecking: %s size: %d" % (newfile, len(filetext)) )

            fundet = reg.findall(filetext)

            filetext = "" # try to release memory

            print " Found  %s" % fundet

            for f in fundet:
                checkity = f[:6].strip(' ').strip('-')
                try:
                    datetime.datetime.strptime(checkity, '%d%m%y')
                    print "valid date %s possibly cpr: %s" % (checkity, f)
                    loglines += "\n   ==> NB NB possibly cpr: %s filename %s \n" % (f, newfile)
                    abstract += "\n==>  %s found in filename %s" % (f, newfile)

                except ValueError:
                    print "not a date  %s" % checkity
            if loglines <> "":
                logfile.write('\n')
                logfile.write(loglines)

        logfile.write('\n\nprocessed file types:' + file_ext)
        logfile.write(g_end_u_line)
        logfile.write("\n -- End cpr check " + datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))

        logfile.write("\n\n.......  summary ........")
        logfile.write(abstract)

    except:
        logfile.close()
        sys.exit(1)

    logfile.close()


def main():

    thefiles = listFiles(odaa_file_path, "*", recurse=1, return_folders=0)

    #print '\n'.join(thefiles)

    handle_newfile(thefiles)

    #pprint ( thefiles )


if __name__ == '__main__':
    main()

