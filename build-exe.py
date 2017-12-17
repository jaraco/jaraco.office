"""
Build script to create a doc-to-pdf convert server as a Windows executable.
"""

import os
import textwrap

# due to the funny way that email.mime does its imports, it doesn't get
#  picked up. CherryPy needs these to be present to not throw errors in
#  the autoreloader.
import email
mime_names = ['email.mime.' + name.lower() for name in email._MIMENAMES]

setup_params = dict(
	console=['server.py'],
	options=dict(
		py2exe=dict(
			packages=['pkg_resources'] + mime_names,
		),
	),
	script_args=('py2exe',),
)

if __name__ == '__main__':
	from setuptools import setup
	__import__('py2exe')
	code = """
		from jaraco.office import convert
		convert.ConvertServer.start_server()
		"""
	open('server.py', 'w').write(textwrap.dedent(code))
	setup(**setup_params)
	os.remove('server.py')
