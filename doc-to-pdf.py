#!python

import re
import os
import inspect
import logging
import itertools
from glob import glob
from optparse import OptionParser

from jaraco.windows.office import Converter

log = logging.getLogger(__name__)

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