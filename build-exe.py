"""
Build script to create a doc-to-pdf convert server as a Windows executable.
"""

if __name__ == '__main__':
    import os
    import textwrap

    from setuptools import setup

    __import__('py2exe')
    code = """
        from jaraco.office import convert
        convert.ConvertServer.start_server()
        """
    open('server.py', 'w').write(textwrap.dedent(code))
    setup(
        console=['server.py'],
        options=dict(
            py2exe=dict(
                packages=['pkg_resources'],
            ),
        ),
        script_args=('py2exe',),  # type: ignore[arg-type] # python/typeshed#12595
    )
    os.remove('server.py')
