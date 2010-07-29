import os
import optparse
from contextlib import contextmanager
from jaraco.util.filesystem import save_to_file, replace_extension

@contextmanager
def word_context(word, filename, close_flags):
	doc = word.Documents.Open(filename)
	yield doc
	doc.Close(close_flags)

class Converter(object):
	"""
	An object that will convert a Word-readable file to one of the Word-
	savable formats (defaults to PDF).
	
	Requires Microsoft Word 2007 or later.
	"""
	def __init__(self):
		from win32com.client import Dispatch
		import pythoncom
		import threading
		if threading.current_thread().getName() != 'MainThread':
			pythoncom.CoInitialize()
		self.word = Dispatch('Word.Application')

	def convert(self, docfile_string, target_format=None):
		"""
		Take a string (in memory) and return it as a string of the
		target format (also as a string in memory).
		"""
		from win32com.client import constants
		target_format = target_format or getattr(constants, 'wdFormatPDF', 17)

		with save_to_file(docfile_string) as docfile:
			# if I don't put a pdf extension on it, Word will
			pdffile = replace_extension('.pdf', docfile)
			dont_save = getattr(constants, 'wdDoNotSaveChanges', 0)
			with word_context(self.word, docfile, dont_save) as doc:
				res = doc.SaveAs(pdffile, target_format)
			content = open(pdffile, 'rb').read()
			os.remove(pdffile)
		return content

	__call__ = convert

	def __del__(self):
		self.word.Quit()

class ConvertServer(object):
	def default(self, filename):
		cherrypy.response.headers['Content-Type'] = 'application/pdf'
		return Converter().convert(cherrypy.request.body.fp.read())
	default.exposed = True

	@staticmethod
	def start_server():
		global cherrypy
		import cherrypy
		_, args = optparse.OptionParser().parse_args()
		if args: config, = args
		cherrypy.quickstart(ConvertServer(), config=config)