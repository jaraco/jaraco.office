#!python

import re
import inspect
from optparse import OptionParser
from win32com.client import Dispatch, constants

def doc2pdf(docfile, pdffile = None):
	word = Dispatch('Word.Application')
	if pdffile is None:
		pdffile = re.sub('\.doc(x)?$', '.pdf', docfile)
	doc = word.Documents.Open(docfile)
	wdFormatPDF = getattr(constants, 'wdFormatPDF', 17)
	doc.SaveAs(pdffile, wdFormatPDF)
	wdDoNotSaveChanges = getattr(constants, 'wdDoNotSaveChanges', 0)
	doc.Close(wdDoNotSaveChanges)
	word.Quit()

def handle_command_line():
	"%prog <word doc> [<pdf file>]"
	usage = inspect.getdoc(handle_command_line)
	parser = OptionParser(usage=usage)
	options, args = parser.parse_args()
	if not 1 <= len(args) <= 2:
		parser.error("Incorrect number of arguments")
	doc2pdf(*args)

if __name__ == '__main__': handle_command_line()