from setuptools import setup, find_packages
import sys, os

version = '1.0'

setup(
	name='jaraco.office',
	version=version,
	description="Utility library for working with MS Office documents",
	keywords='excel office word'.split(),
	author='Jason R. Coombs',
	author_email='jaraco@jaraco.com',
	url='http://pypi.python.org/pypi/jaraco.office',
	license='MIT',
	packages=find_packages(),
	namespace_packages = ['jaraco'],
	classifiers = [
		"Development Status :: 4 - Beta",
		"Intended Audience :: Developers",
		"Programming Language :: Python",
	],
	zip_safe=True,
	entry_points = dict(
		console_scripts = [
			'doc-to-pdf = jaraco.office.word:doc_to_pdf_cmd',
		],
	)
	)
