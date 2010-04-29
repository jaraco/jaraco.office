from setuptools import setup, find_packages
import sys, os

version = '1.0'

setup(
	name='jaraco.excel',
	version=version,
	description="Utility library for working with MS Excel documents",
	keywords='excel',
	author='Jason R. Coombs',
	author_email='jaraco@jaraco.com',
	url='http://pypi.python.org/pypi/jaraco.excel',
	license='MIT',
	packages=find_packages(),
	namespace_packages = ['jaraco'],
	classifiers = [
		"Development Status :: 4 - Beta",
		"Intended Audience :: Developers",
		"Programming Language :: Python",
	],
	zip_safe=True,
	)
