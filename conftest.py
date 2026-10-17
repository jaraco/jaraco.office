import importlib.util

collect_ignore = (
    [] if importlib.util.find_spec('win32com') else ['jaraco/office/grep.py']
)
