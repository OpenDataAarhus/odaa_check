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

#odaa_file_path = '/home/arne/newckan/data/files/'
odaa_file_path = '/home/arne/Backup/Hulk_backup/20012014/files/'

#odaa_file_path= '/home/arne/Odaa_Cpr_testData/spec_test/'


#
def listFiles(root, patterns='*', recurse=1, return_folders=0):
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
    return contents


def process_odt(newfile):

    doc = odf.opendocument.load(newfile)
    return doc.xml()


def handle_newfile(filelist):

    """

    :param filelist:
    """
    r = '\d{5,6}[- ]\d{4}'
    reg = re.compile(r)
    fundet = []

    logfile = open('odaachecklog', 'a')
    logfile.write("\n\n -- Start cpr check" + datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))
    logfile.write('\n__________________________________________________________________')
    file_ext = ""
    try:

        for newfile in filelist:
            logfile.write("\nChecking:" + newfile)
            loglines = ""
            fileName, fileExtension = os.path.splitext(newfile)
            if fileExtension not in file_ext:
                file_ext += ' ' + fileExtension

            print "checker %s" %newfile

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

            elif fileExtension == '.ods':
                filetext  = process_ods(newfile,reg)
                print "After ods"
                print filetext

                continue

            else:  # all other filetypes and default handling
                inf = open(newfile, "r")
                filetext = inf.read()
                inf.close()
            logfile.write("\nExctacted textsize  %d" % len(filetext))

            fundet = reg.findall(filetext)

            filetext = "" # try to release memory

            print " Found  %s" % fundet

            for f in fundet:
                checkity = f[:6].strip(' ').strip('-')
                try:
                    datetime.datetime.strptime(checkity, '%d%m%y')
                    print "valid date %s possibly cpr: %s" % (checkity, f)
                    loglines += "ok date %s possibly cpr: %s" % (checkity, f) + '\n'
                except ValueError:
                    print "not a date  %s" % checkity
            if loglines <> "":

                logfile.write('\n')
                logfile.write(loglines)

        logfile.write('\nprocessed file types:' + file_ext)
        logfile.write("\n\n -- End cpr check " +  datetime.datetime.fromtimestamp(time.time()).strftime('%Y-%m-%d %H:%M:%S'))
        logfile.write('\n__________________________________________________________________')

    except:
        logfile.close()
        raise

    logfile.close()



def main():

    thefiles = listFiles(odaa_file_path, "*", recurse=1, return_folders=0)

    print '\n'.join(thefiles)

    handle_newfile(thefiles)

    #pprint ( thefiles )


if __name__ == '__main__':
    main()

