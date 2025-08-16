a = Analysis(['DocCreatorv6_AI_Addition.py'],
             pathex=['c:\\Users\\Pete\\OneDrive\\4. Programs\\Git\\Word Template Generator\\Python-Word-Template-Generator\\Python-Word-Template-Generator'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[
                 'google.generativeai', # Add this line to exclude the AI package
                 'matplotlib',
                 'PyQt5',
                 'PySide2',
                 'sqlite3',
                 'unittest',
                 'pydoc_data',
                 'bz2',
                 'lzma',
                 'tkinter.test'
             ],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=None,
             noarchive=False)