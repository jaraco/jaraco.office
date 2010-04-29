from setuptools import setup, find_packages
import sys, os

version = '1.0'

setup(
	name='jaraco.excel',
	version=version,
	description="Utility library for working with MS Excel documents",
	long_description="""\
""",
	keywords='excel',
	author='Jason R. Coombs',
	author_email='jaraco@jaraco.com',
	url='http://www.jaraco.com',
	license='MIT',
	packages=find_packages(),
	namespace_packages = ['jaraco']
	classifiers = [
		"Development Status :: 4 - Beta",
		"Intended Audience :: Developers",
		"Programming Language :: Python",
	],
	include_package_data=True,
	zip_safe=True,
	install_requires=[
	  # -*- Extra requirements: -*-
	],
	entry_points="""
	# -*- Entry points: -*-
	""",
	)
