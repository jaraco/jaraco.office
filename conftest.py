import platform


collect_ignore = (
    [
        'jaraco/office/grep.py',
    ]
    if platform.system() != "Windows"
    else []
)
