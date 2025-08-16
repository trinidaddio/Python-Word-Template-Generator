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
class Sections(DispatchBaseClass):
	CLSID = IID('{0002095A-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Section
	def Add(self, Range=defaultNamedOptArg, Start=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(5, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Range
			, Start)
		if ret is not None:
			ret = Dispatch(ret, 'Add', '{00020959-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Section
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{00020959-0000-0000-C000-000000000046}')
		return ret

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Count": (2, 2, (3, 0), (), "Count", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'First' returns object of type 'Section'
		"First": (3, 2, (9, 0), (), "First", '{00020959-0000-0000-C000-000000000046}'),
		# Method 'Last' returns object of type 'Section'
		"Last": (4, 2, (9, 0), (), "Last", '{00020959-0000-0000-C000-000000000046}'),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (1101, 2, (9, 0), (), "PageSetup", '{00020971-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
	}
	_prop_map_put_ = {
		"PageSetup": ((1101, LCID, 4, 0),()),
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00020959-0000-0000-C000-000000000046}')
		return ret

	def __str__(self, *args):
		return str(self.__call__(*args))
	def __int__(self, *args):
		return int(self.__call__(*args))
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,2,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, '{00020959-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(2, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __bool__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002095A-0000-0000-C000-000000000046}", Sections )
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

Sections_vtables_dispatch_ = 1
Sections_vtables_ = [
	(( '_NewEnum' , 'prop' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 1024 , )),
	(( 'Count' , 'prop' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'First' , 'prop' , ), 3, (3, (), [ (16393, 10, None, "IID('{00020959-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Last' , 'prop' , ), 4, (4, (), [ (16393, 10, None, "IID('{00020959-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (9, 1, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'prop' , ), 0, (0, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{00020959-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Add' , 'Range' , 'Start' , 'prop' , ), 5, (5, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{00020959-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 136 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002095A-0000-0000-C000-000000000046}", Sections )
