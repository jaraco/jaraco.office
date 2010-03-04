#!python

import re
import os
import inspect
import logging
from glob import glob
from optparse import OptionParser
from win32com.client import Dispatch, constants

log = logging.getLogger(__name__)

class Converter(object):
	def __init__(self):
		self.word = Dispatch('Word.Application')

	def convert(self, docfile, pdffile = None):
		if pdffile is None:
			base, ext = os.path.splitext(docfile)
			pdffile = base + '.pdf'
		if os.path.exists(pdffile):
			raise Exception("Target already exists: " + pdffile)
		log.info('converting {docfile} to {pdffile}'.format(**vars()))
		doc = self.word.Documents.Open(docfile)
		wdFormatPDF = getattr(constants, 'wdFormatPDF', 17)
		doc.SaveAs(pdffile, wdFormatPDF)
		wdDoNotSaveChanges = getattr(constants, 'wdDoNotSaveChanges', 0)
		doc.Close(wdDoNotSaveChanges)

	__call__ = convert

	def __del__(self):
		self.word.Quit()

def handle_multiple(docfile, pdffile=None):
	"""
	Handle docfile if it matches more than one file.
	"""
	files = glob(docfile)
	n_files = len(files)
	if n_files > 1:
		if pdffile is not None:
			raise Exception("Cannot specify output file with multiple sources")
	log.info("Processing {n_files} source files...".format(**vars()))
	converter = Converter()
	map(converter, files)

def handle_command_line():
	"%prog <word doc> [<pdf file>]"
	usage = inspect.getdoc(handle_command_line)
	parser = OptionParser(usage=usage)
	options, args = parser.parse_args()
	if not 1 <= len(args) <= 2:
		parser.error("Incorrect number of arguments")
	logging.basicConfig(level=logging.INFO)
	handle_multiple(*args)

if __name__ == '__main__': handle_command_line()