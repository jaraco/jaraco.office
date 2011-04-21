import os
import textwrap
import email
mime_names = ['email.mime.'+ name.lower() for name in email._MIMENAMES]
setup_params = dict(
	console=['server.py'],
	options = dict(
		py2exe = dict(
			packages = ['pkg_resources'] + mime_names,
		),
	),
)

if __name__ == '__main__':
	from setuptools import setup
	import py2exe
	open('server.py', 'w').write(textwrap.dedent(
		"""
		from jaraco.office import convert
		convert.ConvertServer.start_server()
		"""))
	setup(**setup_params)
	os.remove('server.py')
