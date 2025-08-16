# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.3 (tags/v3.13.3:6280bb5, Apr  8 2025, 14:47:33) [MSC v.1943 64 bit (AMD64)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Aug 15 20:38:25 2025
'Microsoft Word 16.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30d03f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{00020905-0000-0000-C000-000000000046}')
MajorVersion = 8
MinorVersion = 7
LibraryFlags = 8
LCID = 0x0

from win32com.client import CoClassBaseClass
import sys
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7.DocumentEvents')
DocumentEvents = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7.DocumentEvents'].DocumentEvents
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7.DocumentEvents2')
DocumentEvents2 = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7.DocumentEvents2'].DocumentEvents2
__import__('win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7._Document')
_Document = sys.modules['win32com.gen_py.00020905-0000-0000-C000-000000000046x0x8x7._Document']._Document
# This CoClass is known by the name 'Word.Document.8'
class Document(CoClassBaseClass): # A CoClass
	CLSID = IID('{00020906-0000-0000-C000-000000000046}')
	coclass_sources = [
		DocumentEvents,
		DocumentEvents2,
	]
	default_source = DocumentEvents2
	coclass_interfaces = [
		_Document,
	]
	default_interface = _Document

win32com.client.CLSIDToClass.RegisterCLSID( "{00020906-0000-0000-C000-000000000046}", Document )
