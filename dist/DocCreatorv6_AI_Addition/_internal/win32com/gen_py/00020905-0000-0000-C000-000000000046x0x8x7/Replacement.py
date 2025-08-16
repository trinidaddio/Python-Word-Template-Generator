# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.3 (tags/v3.13.3:6280bb5, Apr  8 2025, 14:47:33) [MSC v.1943 64 bit (AMD64)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Aug 15 20:38:27 2025
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

from win32com.client import DispatchBaseClass
class Replacement(DispatchBaseClass):
	CLSID = IID('{000209B1-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ClearFormatting(self):
		return self._oleobj_.InvokeTypes(20, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'Font' returns object of type 'Font'
		"Font": (10, 2, (13, 0), (), "Font", '{000209F5-0000-0000-C000-000000000046}'),
		# Method 'Frame' returns object of type 'Frame'
		"Frame": (18, 2, (9, 0), (), "Frame", '{0002092A-0000-0000-C000-000000000046}'),
		"Highlight": (17, 2, (3, 0), (), "Highlight", None),
		"LanguageID": (16, 2, (3, 0), (), "LanguageID", None),
		"LanguageIDFarEast": (19, 2, (3, 0), (), "LanguageIDFarEast", None),
		"NoProofing": (21, 2, (3, 0), (), "NoProofing", None),
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat'
		"ParagraphFormat": (11, 2, (13, 0), (), "ParagraphFormat", '{000209F4-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"Style": (12, 2, (12, 0), (), "Style", None),
		"Text": (15, 2, (8, 0), (), "Text", None),
	}
	_prop_map_put_ = {
		"Font": ((10, LCID, 4, 0),()),
		"Highlight": ((17, LCID, 4, 0),()),
		"LanguageID": ((16, LCID, 4, 0),()),
		"LanguageIDFarEast": ((19, LCID, 4, 0),()),
		"NoProofing": ((21, LCID, 4, 0),()),
		"ParagraphFormat": ((11, LCID, 4, 0),()),
		"Style": ((12, LCID, 4, 0),()),
		"Text": ((15, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000209B1-0000-0000-C000-000000000046}", Replacement )
# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.3 (tags/v3.13.3:6280bb5, Apr  8 2025, 14:47:33) [MSC v.1943 64 bit (AMD64)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Aug 15 20:38:27 2025
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

Replacement_vtables_dispatch_ = 1
Replacement_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 10, (10, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 10, (10, (), [ (13, 1, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 11, (11, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 11, (11, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 12, (12, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 12, (12, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 15, (15, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 15, (15, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 16, (16, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 16, (16, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Highlight' , 'prop' , ), 17, (17, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Highlight' , 'prop' , ), 17, (17, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'Frame' , 'prop' , ), 18, (18, (), [ (16393, 10, None, "IID('{0002092A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 19, (19, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 19, (19, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'ClearFormatting' , ), 20, (20, (), [ ], 1 , 1 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 21, (21, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 21, (21, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000209B1-0000-0000-C000-000000000046}", Replacement )
