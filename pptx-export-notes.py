# coding=utf-8
#############################################################################
#############################################################################
###                                                                       ###
###        pptx notes exporter v1.0 copyright Eric Jang 2012              ###
###        ericjang2004@gmail.com                                         ###
###                                                                       ###
#############################################################################
#############################################################################

# !/usr/bin/env python
import argparse
import os
import glob
import shutil
import codecs
from zipfile import ZipFile
from xml.dom.minidom import parse

TMP_PATH = '/tmp/pptx-export-notes/'
SLIDE_DELIMITER = u'\n────────────────────────────────────────\n'


# main function
def run():
    parser = argparse.ArgumentParser(description='exports speaker notes from pptx files by parsing the XML')
    parser.add_argument('-v', action='version', version='%(prog)s 1.0')
    parser.add_argument('-p', metavar='<path/to/pptx/file>', help='path to the Powerpoint 2007+ file', action='store',
                        type=argparse.FileType('rb'), dest='pptxfile')
    # add more arguments here in future if you wish to expand
    args = parser.parse_args()
    # extract the pptx file as a zip archive
    # note: only extract from pptx files that you trust. they could potentially overwrite your important files.
    shutil.rmtree(TMP_PATH, ignore_errors=True)
    ZipFile(args.pptxfile).extractall(path=TMP_PATH, pwd=None)
    path = TMP_PATH + 'ppt/notesSlides/'

    notesDict = {}
    # open up the file that you wish to write to
    writepath = os.path.dirname(args.pptxfile.name) + '/' + os.path.basename(args.pptxfile.name).rsplit('.', 1)[
        0] + '_presenter_notes.txt'
    print(writepath)
    f = codecs.open(writepath, mode='w', encoding='utf-8')
    # f = open(writepath, 'w')

    for infile in glob.glob(os.path.join(path, '*.xml')):
        # parse each XML notes file from the notes folder.
        dom = parse(infile)
        noteslist = dom.getElementsByTagName('a:t')
        if not noteslist:
            continue

        if noteslist[-1].parentNode.getAttribute('type') == 'slidenum':
            # separate last element of noteslist for use as the slide marking.
            slideNumber = noteslist.pop()
            slideNumber = slideNumber.firstChild.nodeValue
        else:
            # fallback for pptx generated from LibreOffice/Google Slides
            slideNumber = os.path.basename(infile).rsplit('.', 1)[0][len('notesSlide'):]

        # start with this empty string to build the presenter note itself
        tempstring = '\n'.join(node.firstChild.nodeValue if node.firstChild else '' for node in noteslist)

        # store the tempstring in the dictionary under the slide number
        notesDict[slideNumber] = tempstring

    # print/write the dictionary to file in sorted order by key value.
    for x in [key for key in sorted(notesDict.keys(), key=int)]:
        f.write('Slide ' + str(x) + '\n')
        f.write(notesDict[str(x)])
        f.write(SLIDE_DELIMITER)

    f.close()
    print('file successfully written to' + '\'' + writepath + '\'')


if __name__ == "__main__":
    run()
