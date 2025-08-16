# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.3 (tags/v3.13.3:6280bb5, Apr  8 2025, 14:47:33) [MSC v.1943 64 bit (AMD64)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Aug 15 20:38:26 2025
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
class Section(DispatchBaseClass):
	CLSID = IID('{00020959-0000-0000-C000-000000000046}')
	coclass_clsid = None

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'Footers' returns object of type 'HeadersFooters'
		"Footers": (122, 2, (9, 0), (), "Footers", '{00020984-0000-0000-C000-000000000046}'),
		# Method 'Headers' returns object of type 'HeadersFooters'
		"Headers": (121, 2, (9, 0), (), "Headers", '{00020984-0000-0000-C000-000000000046}'),
		"Index": (124, 2, (3, 0), (), "Index", None),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (1101, 2, (9, 0), (), "PageSetup", '{00020971-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"ProtectedForForms": (123, 2, (11, 0), (), "ProtectedForForms", None),
		# Method 'Range' returns object of type 'Range'
		"Range": (0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'),
	}
	_prop_map_put_ = {
		"Borders": ((1100, LCID, 4, 0),()),
		"PageSetup": ((1101, LCID, 4, 0),()),
		"ProtectedForForms": ((123, LCID, 4, 0),()),
	}
	# Default property for this class is 'Range'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'))
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020959-0000-0000-C000-000000000046}", Section )
# -*- coding: utf-8 -*-
# Created by makepy.py version 0.5.01
# By python version 3.13.3 (tags/v3.13.3:6280bb5, Apr  8 2025, 14:47:33) [MSC v.1943 64 bit (AMD64)]
# From type library '{00020905-0000-0000-C000-000000000046}'
# On Fri Aug 15 20:38:26 2025
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

Section_vtables_dispatch_ = 1
Section_vtables_ = [
	(( 'Range' , 'prop' , ), 0, (0, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (9, 1, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Headers' , 'prop' , ), 121, (121, (), [ (16393, 10, None, "IID('{00020984-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Footers' , 'prop' , ), 122, (122, (), [ (16393, 10, None, "IID('{00020984-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'ProtectedForForms' , 'prop' , ), 123, (123, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'ProtectedForForms' , 'prop' , ), 123, (123, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Index' , 'prop' , ), 124, (124, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020959-0000-0000-C000-000000000046}", Section )
