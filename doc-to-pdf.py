#!python

import re
import os
import tempfile
import inspect
import logging
import itertools
from glob import glob
from optparse import OptionParser
from win32com.client import Dispatch, constants

log = logging.getLogger(__name__)

class save_to_file():
	def __init__(self, content):
		self.content = content

	def __enter__(self):
		fd, self.filename = tempfile.mkstemp()
		file = os.fdopen(fd, 'wb')
		file.write(self.content)
		file.close()
		return self.filename

	def __exit__(self, type, value, traceback):
		os.remove(self.filename)

class Converter(object):
	def __init__(self):
		self.word = Dispatch('Word.Application')

	def convert(self, docfile_string):
		with save_to_file(docfile_string) as docfile:
			doc = self.word.Documents.Open(docfile)
			wdFormatPDF = getattr(constants, 'wdFormatPDF', 17)
			pdffile = docfile+'.pdf' # if I don't put a pdf extension on it, Word will
			res = doc.SaveAs(pdffile, wdFormatPDF)
			wdDoNotSaveChanges = getattr(constants, 'wdDoNotSaveChanges', 0)
			doc.Close(wdDoNotSaveChanges)
			content = open(pdffile, 'rb').read()
			os.remove(pdffile)
		return content

	__call__ = convert

	def __del__(self):
		self.word.Quit()

class ExtensionReplacer():
	"""
	>>> ExtensionReplacer('.pdf')('myfile.doc')
	'myfile.pdf'
	"""
	def __init__(self, new_ext):
		self.new_ext = new_ext

	def __call__(self, orig_name):
		return os.path.splitext(orig_name)[0] + self.new_ext

def save_to(content, filename):
	open(filename, 'wb').write(content)

def handle_multiple(docfile, pdffile=None):
	"""
	Handle docfile if it matches more than one file.
	"""
	doc_files = glob(docfile)
	n_files = len(doc_files)
	if n_files > 1:
		if pdffile is not None:
			raise Exception("Cannot specify output file with multiple sources")
	log.info("Processing {n_files} source files...".format(**vars()))
	converter = Converter()
	doc_content = itertools.imap(lambda f: open(f, 'rb').read(), doc_files)
	pdf_content = itertools.imap(converter, doc_content)
	pdf_files = itertools.imap(ExtensionReplacer('.pdf'), doc_files) if not pdffile else [pdffile]
	map(save_to, pdf_content, pdf_files)

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