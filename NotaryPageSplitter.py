#! /usr/bin/env python3
#
# Python program to create a pdf file splitter for notary use.
# This program will determine the size of the pdf pages and 
# produce files of each size type.  A report is then generated to
# help with putting the pages back together in the correct order.
#
# Options include
#     -e for encrypting files, a new file with *_enc.pdf is created
#     -d for decrypting files, a new file with *_dec.pdf is created
#     -s number for splitting the pdf file into multiple files
#     -m merges files into one timestamped file
#
#     -G Debug option - gives more info on each page
#
# Note: if you give the number of section parts the same size as the number
# of pages then each page will be split out into a single file.
#
# Frank Anderson: November 6, 2020
# Nov 16,2020 Added strict=False, warnings on merged file errors
# Nov 16,2020 Added pdf file decryption
# Nov 17,2020 Added pdf splitter function
# Dec 25,2020 Added pdf merge function
#   note: excluding and sorting are not implemented because I'm assuming
#         the user knows what they want to do.  There are often duplicate
#         pages in a notary packet.
# Dec 26,2020 Added pdf file encryption and decrytion option
#         Rework password entry so the text is hidden
#

import glob
import os
import sys
import getpass
from PyPDF2 import PdfFileReader, PdfFileWriter
from datetime import datetime

gDebug = False

def mergePdfFiles( files ):
    print( "Merging files ..." )
    totalNumPages = 0
    pdfW = PdfFileWriter()
    for fn in files:
        (path, filename) = os.path.split( fn )
        if path == "":
            path = "."
        (fname, extension) = os.path.splitext(filename)
        currentPath = os.getcwd()
        os.chdir( path )

        try:
            pdfFP = PdfFileReader( open( filename, 'rb' ), strict = False)
        except:
            print("Something is wrong with file: ", filename )
            continue

        if pdfFP.isEncrypted:
            print( "<<<<<<<<<<<<< WARNING >>>>>>>>>>>>>>>" )
            print( fn, "is encrypted with a password")
            result = 0
            for ctr in range(3):
                password = getpass.getpass( "Enter a decrytion password: " )
                result = pdfFP.decrypt( password )
                if result == 0:
                    print( "Wrong password, try again" )
                    continue
                else:
                    break
            if result == 0:
                print( "Incorrect password failed 3 times, exiting" )
                sys.exit(0)

        numPages = pdfFP.getNumPages()
        for pg in range( numPages ):
            pdfW.addPage( pdfFP.getPage(pg) )

        msg = "%s: %d" % (fn, numPages )
        print( msg )
        os.chdir( currentPath )
        totalNumPages += numPages

    ts = datetime.now()
    pdfFile = "%d%d%d-%d%d%d_mergedFile.pdf" % (ts.year, ts.month, ts.day, ts.hour, ts.minute, ts.second)
    with open( pdfFile, 'wb' ) as file1:
        pdfW.write( file1 )
    file1.close()
    msg = "Created %s: %d pages" % ( pdfFile, totalNumPages )
    print( msg )


def pdfDivide( fullFileName, parts ):
    print( "dividing file: ", fullFileName, " into ", parts, " parts" )
    if parts <= 0:
        return

    (path, filename) = os.path.split( fullFileName )
    if path == "":
        path = "."
    (fname, extension) = os.path.splitext(filename)
    currentPath = os.getcwd()
    os.chdir( path )

    try:
        pdfFP = PdfFileReader( open( filename, 'rb' ), strict = False)
    except:
        print("Something is wrong with file: ", filename )
        return


    if pdfFP.isEncrypted:
        print( "<<<<<<<<<<<<< WARNING >>>>>>>>>>>>>>>" )
        print( fn, "is encrypted with a password")
        result = 0
        for ctr in range(3):
            password = getpass.getpass( "Enter a decrytion password: " )
            result = pdfFP.decrypt( password )
            if result == 0:
                print( "Wrong password, try again" )
                continue
            else:
                break
        if result == 0:
            print( "Incorrect password failed 3 times, exiting" )
            sys.exit(0)

    pdfW = PdfFileWriter()

    numPages = pdfFP.getNumPages()
    numPagesPerPart = int( numPages/parts )
    lastPart = numPages - numPagesPerPart*(parts-1)

    message = "PDF FILE: %s\n" % fullFileName
    message += "Number of Pages: %d\n" % numPages
    message += "Splitting PDF into %d parts\n" % parts
    page_count = 0;

    for section in range(parts-1):
        for page in range(numPagesPerPart):
            pdfW.addPage( pdfFP.getPage(page_count) )
            page_count += 1

        pdfFile = "%s_part%d.pdf" % (fname, section+1)
        with open( pdfFile, 'wb' ) as file1:
            message += "creating %s numPages: %d\n" % (pdfFile, numPagesPerPart)
            pdfW.write( file1 )
        file1.close()
        pdfW = PdfFileWriter()

    pdfW = PdfFileWriter()
    for page in range(lastPart):
        pdfW.addPage( pdfFP.getPage(page_count) )
        page_count += 1

    pdfFile = "%s_part%d.pdf" % (fname, parts)
    with open( pdfFile, 'wb' ) as file1:
        message += "creating %s numPages: %d\n" % (pdfFile, lastPart)
        pdfW.write( file1 )
    file1.close()

    print( message )
    os.chdir( currentPath )

def pdfSplitter( fullFileName ):
    global gDebug

    (path, filename) = os.path.split( fullFileName )
    if path == "":
        path = "."
    (fname, extension) = os.path.splitext(filename)
    currentPath = os.getcwd()
    os.chdir( path )

    try:
        pdfFP = PdfFileReader( open( filename, 'rb' ), strict = False)
    except:
        print("Something is wrong with file: ", filename )
        return

    if pdfFP.isEncrypted:
        print( "<<<<<<<<<<<<< WARNING >>>>>>>>>>>>>>>" )
        print( fn, "is encrypted with a password")
        result = 0
        for ctr in range(3):
            password = getpass.getpass( "Enter a decrytion password: " )
            result = pdfFP.decrypt( password )
            if result == 0:
                print( "Wrong password, try again" )
                continue
            else:
                break
        if result == 0:
            print( "Incorrect password failed 3 times, exiting" )
            sys.exit(0)

    pdfLetter = PdfFileWriter()
    pdfLegal = PdfFileWriter()
    pdfTabloid = PdfFileWriter()
    pdfUnknown = PdfFileWriter()

    pageNumber = 1
    oldSize = "nosize"
    letterFound = False
    legalFound = False
    tabloidFound = False
    unknownFound = False
    numType = 0
    numLetter = 0
    numLegal = 0
    numTabloid = 0
    numUnknown = 0;

    numPages = pdfFP.getNumPages()
    docInfo = pdfFP.getDocumentInfo()
    message = "Splitting file: %s\n" % fullFileName
    message += "Title: %s\n" % docInfo.title
    message += "Subject: %s\n" % docInfo.subject
    message += "Author: %s\n" % docInfo.author
    message += "Producer: %s\n" % docInfo.producer
    message += "Creator: %s\n" % docInfo.creator
    message += "Number of Pages: %d\n" % numPages
    if pdfFP.isEncrypted:
        message += "Encryption: Password Required\n"
    else:
        message += "Encryption: No Password Reqiured\n"
    message += "\n"

    tol=1.0

    for page in range( numPages):
        ro = pdfFP.getPage(page).mediaBox
        (x,y) = ro.upperRight
        if gDebug:
            print( x, "x", y )

        try:
            fx = float(x)
        except:
            # default to 8.5 inches
            print( "Warning: float exception. ", y, " is not converting to float, using 8.5 inches" )
            fx = 72.0*8.5
        try:
            fy = float(y)
        except:
            # default to 11.0 inches
            print( "Warning: float exception. ", y, " is not converting to float, using 11 inches" )
            fy = 72.0*11.0

        width = fx/72.0
        height = fy/72.0


        size = "letter"
        if  width >= 8.5-tol and width <= 8.5+tol and \
            height >= 11.0-tol and height <= 11.0+tol:
                size = "letter"
                letterFound = True
        elif  width >= 8.5-tol and width <= 8.5+tol and \
            height >= 14.0-tol and height <= 14.0+tol:
                size = "legal"
                legalFound = True
        elif  width >= 11.0-tol and width <= 11.0+tol and\
            height >= 17.0-tol and height <= 17.0+tol:
                size = "tabloid"
                tabloidFound = True

        elif height >= 8.5-tol and height <= 8.5+tol and\
            width >= 11.0-tol and width <= 11.0+tol:
                size = "letterRot"
                letterFound = True
        elif height >= 8.5-tol and height <= 8.5+tol and\
            width >= 14.0-tol and width <= 14.0+tol:
                size = "legalRot"
                legalFound = True

        elif width >= 17.0-tol and width <= 17.0+tol and\
            height >= 11.0-tol and height <= 11.0+tol:
                size = "tabloidRot"
                tabloidFound = True
        else:
            size = "unknown"
            unknownFound = True
            message += "width: %f, height: %f\n" % (width,height)

        if gDebug:
            msg = "Page\t%d:\t%.1fx%.1f\t%s" % (page+1, width, height, size)
            print( msg )

        if ( oldSize != size ):
            if numType > 0:
                message += "(qty: %d)\n" % numType
                if size == "letter":
                    message += "\n"
                else if size == "letterRot":
                    message += "\n"
            message += "%s: " % size
            numType = 0

        message += "%d " % pageNumber

        if size == "letter":
            pdfLetter.addPage( pdfFP.getPage(page) )
            numLetter += 1
        elif size == "letterRot":
            pdfLetter.addPage( pdfFP.getPage(page).rotateClockwise(90))
            numLetter += 1
        elif size == "legal":
            pdfLegal.addPage( pdfFP.getPage(page) )
            numLegal += 1
        elif size == "legalRot":
            pdfLegal.addPage( pdfFP.getPage(page).rotateClockwise(90))
            numLegal += 1
        elif ( size == "tabloid" ):
            pdfTabloid.addPage( pdfFP.getPage(page) )
            numTabloid += 1
        elif ( size == "tabloidRot" ):
            pdfTabloid.addPage( pdfFP.getPage(page).rotateClockwise(90))
            numTabloid += 1
        else:
            pdfUnknown.addPage( pdfFP.getPage(page) )
            numUnknown += 1

        oldSize = size

        pageNumber += 1
        numType += 1

    if numType > 0:
        message += "(qty: %d)\n\n" % numType
    else:
        message += "\n"

    message += "numPages:   %d\n" % numPages
    message += "numLetter:  %d\n" % numLetter
    message += "numLegal:   %d\n" % numLegal
    message += "numTabloid: %d\n" % numTabloid
    message += "numUnknown: %d\n" % numUnknown
    message += "\n"

    if letterFound:
        pdfFile = "%s_letter.pdf" % (fname)
        with open( pdfFile, 'wb' ) as file1:
            message += "creating %s\n" % pdfFile
            pdfLetter.write( file1 )
        file1.close()
    if legalFound:
        pdfFile = "%s_legal.pdf" % (fname)
        with open( pdfFile, 'wb' ) as file2:
            message += "creating %s\n" % pdfFile
            pdfLegal.write( file2 )
        file2.close()
    if tabloidFound:
        pdfFile = "%s_tabloid.pdf" % (fname)
        with open( pdfFile, 'wb' ) as file3:
            message += "creating %s\n" % pdfFile
            pdfTabloid.write( file3 )
        file3.close()
    if unknownFound:
        pdfFile = "%s_unknown.pdf" % (fname)
        with open( pdfFile, 'wb' ) as file4:
            message += "creating %s\n" % pdfFile
            pdfTabloid.write( file4 )
        file4.close()


    logFileName = "%s_logfile.txt" % (fname)
    message += "creating %s" % logFileName
    logfile = open(logFileName, 'w')
    logfile.write( message )
    logfile.close()

    print( message )
    os.chdir( currentPath )

def encryptFile( fullFileName, password ):
    (path, filename) = os.path.split( fullFileName )
    if path == "":
        path = "."
    (fname, extension) = os.path.splitext(filename)
    currentPath = os.getcwd()
    os.chdir( path )

    try:
        pdfFP = PdfFileReader( open( filename, 'rb'), strict = False )
    except:
        print("Something is wrong with file: ", filename)
        return

    if pdfFP.isEncrypted:
        print("PDF File is already encrypted: ", filename )
        return
    elif len(password) <= 0:
        password = getpass.getpass( "Enter a password: " )

    pdfW = PdfFileWriter()
    numPages = pdfFP.getNumPages()
    for pageNum in range(numPages):
        page = pdfFP.getPage( pageNum )
        pdfW.addPage( page )

    # only the user password is added with a default 128 bit encryption
    pdfW.encrypt( password )
    newFileName = "%s_enc.pdf" % fname
    with open( newFileName, "wb" ) as file1:
        print( "Writing Encrypted File: ", newFileName )
        pdfW.write( file1 )

    file1.close()
    os.chdir( currentPath )

def decryptFile( fullFileName, password ):
    (path, filename) = os.path.split( fullFileName )
    if path == "":
        path = "."
    (fname, extension) = os.path.splitext(filename)
    currentPath = os.getcwd()
    os.chdir( path )

    try:
        pdfFP = PdfFileReader( open( filename, 'rb'), strict = False )
    except:
        print("Something is wrong with file: ", filename)
        return

    if pdfFP.isEncrypted == False:
        print("PDF File is not encrypted: ", filename )
        return
    else:
        if len(password) <= 0:
            print( "You must supply a password do decrypt the file." )
            password = getpass.getpass( "Enter a password: " )
        if pdfFP.decrypt( password ) == 0:
            print( "Wrong password, exiting" )
            sys.exit(0)

    pdfW = PdfFileWriter()
    numPages = pdfFP.getNumPages()
    for pageNum in range(numPages):
        page = pdfFP.getPage( pageNum )
        pdfW.addPage( page )

    newFileName = "%s_dec.pdf" % fname
    with open( newFileName, "wb" ) as file1:
        print( "Writing Decrypted File: ", newFileName )
        pdfW.write( file1 )

    file1.close()
    os.chdir( currentPath )

if __name__ == "__main__":
    if len(sys.argv) <= 1:
        print("Please add some file names" )
        sys.exit()

    files = []
    password = ""
    split_parts = 0
    mergeFiles = False
    splitPDFs = False
    encryptFiles = False
    decryptFiles = False

    for i, arg in enumerate( sys.argv ):
        if arg == "-h":
            print( sys.argv[0], "[-d|-e|-m|-s|-h] [-G] filenames" )
            helpStr = '''Options:
  -d \tDecrypt the file with a user password, New File *_dec.pdf
  -e \tEncrypt the file with a user password, New File *_enc.pdf
  -m \tmerge pdf files
  -s \tsplit pdf files into separate files to make them smaller
  -G \tDebug option will print page sizes
  -h \tReproduces this help info

Note: glob filenames are supported. E.G. *.pdf

This software aids with printing different pdf page sizes for use in building
up Notary packets for signaures.  This software will will split the pdf files
into legal, letter, and tabloid size pdf files.  The new files will have
filename_{size}.pdf names.  A log file will be created to help you merge the
printed pages back into a packet: filename_logfile.txt.  Encryption and
decryption are supported and you will be prompted for a password when needed.
When a landscape page is found the page will be rotated by 90 degrees.'''

            print( helpStr )
            sys.exit(0)

    for i, arg in enumerate( sys.argv ):
        if arg == "-s":
            split_parts = int(input("Enter number of ways to split doc: "))
            if split_parts <= 1:
                print( "number of parts must be greater than 1" )
                sys.exit(0)
            splitPDFs = True
            continue
        elif arg == "-m":
            mergeFiles = True
            continue
        elif arg == "-G":
            gDebug = True
            continue
        elif arg == "-e":
            password = getpass.getpass()
            encryptFiles = True
            continue
        elif arg == "-d":
            password = getpass.getpass()
            decryptFiles = True
            continue
        elif i > 0:
            files.append( arg )

    #print( "files: ", files )
    #print( "Password: ", password )
    allFiles = []

    for file in files:
        fileList = glob.glob( file )
        for fn in fileList:
            allFiles.append( fn )

    # The actual list is not "uniqued" or sorted because I'm
    # assuming the user knows what they want to do.
    newFiles = list( set( allFiles ))
    newFiles.sort()

    if len(allFiles) == 0:
        print("ERROR: No files were entered")
        sys.exit()

    if gDebug:
        print( allFiles )

    if splitPDFs:
        for fn in allFiles:
            pdfDivide( fn, split_parts )

    elif mergeFiles:
        mergePdfFiles( allFiles )
        if len( allFiles ) != len( newFiles ):
            print("Warning: duplicate files were merged")

    elif encryptFiles:
        for fn in allFiles:
            encryptFile( fn, password )

    elif decryptFiles:
        for fn in allFiles:
            decryptFile( fn, password )

    else:
        for fn in allFiles:
            pdfSplitter( fn )

