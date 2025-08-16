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

from win32com.client import DispatchBaseClass
class Style(DispatchBaseClass):
	CLSID = IID('{0002092C-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def Delete(self):
		return self._oleobj_.InvokeTypes(100, LCID, 1, (24, 0), (),)

	def LinkToListTemplate(self, ListTemplate=defaultNamedNotOptArg, ListLevelNumber=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(101, LCID, 1, (24, 0), ((9, 1), (16396, 17)),ListTemplate
			, ListLevelNumber)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"AutomaticallyUpdate": (13, 2, (11, 0), (), "AutomaticallyUpdate", None),
		"BaseStyle": (1, 2, (12, 0), (), "BaseStyle", None),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (8, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"BuiltIn": (4, 2, (11, 0), (), "BuiltIn", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"Description": (2, 2, (8, 0), (), "Description", None),
		# Method 'Font' returns object of type 'Font'
		"Font": (10, 2, (13, 0), (), "Font", '{000209F5-0000-0000-C000-000000000046}'),
		# Method 'Frame' returns object of type 'Frame'
		"Frame": (11, 2, (9, 0), (), "Frame", '{0002092A-0000-0000-C000-000000000046}'),
		"Hidden": (17, 2, (11, 0), (), "Hidden", None),
		"InUse": (6, 2, (11, 0), (), "InUse", None),
		"LanguageID": (12, 2, (3, 0), (), "LanguageID", None),
		"LanguageIDFarEast": (16, 2, (3, 0), (), "LanguageIDFarEast", None),
		"LinkStyle": (104, 2, (12, 0), (), "LinkStyle", None),
		"Linked": (26, 2, (11, 0), (), "Linked", None),
		"ListLevelNumber": (15, 2, (3, 0), (), "ListLevelNumber", None),
		# Method 'ListTemplate' returns object of type 'ListTemplate'
		"ListTemplate": (14, 2, (9, 0), (), "ListTemplate", '{0002098F-0000-0000-C000-000000000046}'),
		"Locked": (22, 2, (11, 0), (), "Locked", None),
		"NameLocal": (0, 2, (8, 0), (), "NameLocal", None),
		"NextParagraphStyle": (5, 2, (12, 0), (), "NextParagraphStyle", None),
		"NoProofing": (18, 2, (3, 0), (), "NoProofing", None),
		"NoSpaceBetweenParagraphsOfSameStyle": (20, 2, (11, 0), (), "NoSpaceBetweenParagraphsOfSameStyle", None),
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat'
		"ParagraphFormat": (9, 2, (13, 0), (), "ParagraphFormat", '{000209F4-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"Priority": (23, 2, (3, 0), (), "Priority", None),
		"QuickStyle": (25, 2, (11, 0), (), "QuickStyle", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (7, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		# Method 'Table' returns object of type 'TableStyle'
		"Table": (21, 2, (9, 0), (), "Table", '{B7564E97-0519-4C68-B400-3803E7C63242}'),
		"Type": (3, 2, (3, 0), (), "Type", None),
		"UnhideWhenUsed": (24, 2, (11, 0), (), "UnhideWhenUsed", None),
		"Visibility": (19, 2, (11, 0), (), "Visibility", None),
	}
	_prop_map_put_ = {
		"AutomaticallyUpdate": ((13, LCID, 4, 0),()),
		"BaseStyle": ((1, LCID, 4, 0),()),
		"Borders": ((8, LCID, 4, 0),()),
		"Font": ((10, LCID, 4, 0),()),
		"Hidden": ((17, LCID, 4, 0),()),
		"LanguageID": ((12, LCID, 4, 0),()),
		"LanguageIDFarEast": ((16, LCID, 4, 0),()),
		"LinkStyle": ((104, LCID, 4, 0),()),
		"Locked": ((22, LCID, 4, 0),()),
		"NameLocal": ((0, LCID, 4, 0),()),
		"NextParagraphStyle": ((5, LCID, 4, 0),()),
		"NoProofing": ((18, LCID, 4, 0),()),
		"NoSpaceBetweenParagraphsOfSameStyle": ((20, LCID, 4, 0),()),
		"ParagraphFormat": ((9, LCID, 4, 0),()),
		"Priority": ((23, LCID, 4, 0),()),
		"QuickStyle": ((25, LCID, 4, 0),()),
		"UnhideWhenUsed": ((24, LCID, 4, 0),()),
		"Visibility": ((19, LCID, 4, 0),()),
	}
	# Default property for this class is 'NameLocal'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "NameLocal", None))
	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{0002092C-0000-0000-C000-000000000046}", Style )
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

Style_vtables_dispatch_ = 1
Style_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'NameLocal' , 'prop' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'NameLocal' , 'prop' , ), 0, (0, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'BaseStyle' , 'prop' , ), 1, (1, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'BaseStyle' , 'prop' , ), 1, (1, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Description' , 'prop' , ), 2, (2, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Type' , 'prop' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'BuiltIn' , 'prop' , ), 4, (4, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'NextParagraphStyle' , 'prop' , ), 5, (5, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'NextParagraphStyle' , 'prop' , ), 5, (5, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'InUse' , 'prop' , ), 6, (6, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 7, (7, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 8, (8, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 8, (8, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 9, (9, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 9, (9, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 10, (10, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 10, (10, (), [ (13, 1, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'Frame' , 'prop' , ), 11, (11, (), [ (16393, 10, None, "IID('{0002092A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 12, (12, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 12, (12, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'AutomaticallyUpdate' , 'prop' , ), 13, (13, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'AutomaticallyUpdate' , 'prop' , ), 13, (13, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'ListTemplate' , 'prop' , ), 14, (14, (), [ (16393, 10, None, "IID('{0002098F-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'ListLevelNumber' , 'prop' , ), 15, (15, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 16, (16, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 16, (16, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Hidden' , 'prop' , ), 17, (17, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 64 , )),
	(( 'Hidden' , 'prop' , ), 17, (17, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 64 , )),
	(( 'Delete' , ), 100, (100, (), [ ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'LinkToListTemplate' , 'ListTemplate' , 'ListLevelNumber' , ), 101, (101, (), [ (9, 1, None, "IID('{0002098F-0000-0000-C000-000000000046}')") , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 312 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 18, (18, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 18, (18, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'LinkStyle' , 'prop' , ), 104, (104, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'LinkStyle' , 'prop' , ), 104, (104, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Visibility' , 'prop' , ), 19, (19, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'Visibility' , 'prop' , ), 19, (19, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'NoSpaceBetweenParagraphsOfSameStyle' , 'prop' , ), 20, (20, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'NoSpaceBetweenParagraphsOfSameStyle' , 'prop' , ), 20, (20, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Table' , 'prop' , ), 21, (21, (), [ (16393, 10, None, "IID('{B7564E97-0519-4C68-B400-3803E7C63242}')") , ], 1 , 2 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'Locked' , 'prop' , ), 22, (22, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'Locked' , 'prop' , ), 22, (22, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'Priority' , 'prop' , ), 23, (23, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'Priority' , 'prop' , ), 23, (23, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'UnhideWhenUsed' , 'prop' , ), 24, (24, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'UnhideWhenUsed' , 'prop' , ), 24, (24, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'QuickStyle' , 'prop' , ), 25, (25, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'QuickStyle' , 'prop' , ), 25, (25, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'Linked' , 'prop' , ), 26, (26, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002092C-0000-0000-C000-000000000046}", Style )
