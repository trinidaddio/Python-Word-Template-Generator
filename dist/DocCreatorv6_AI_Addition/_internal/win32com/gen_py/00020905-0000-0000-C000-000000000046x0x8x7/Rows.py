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
class Rows(DispatchBaseClass):
	CLSID = IID('{0002094C-0000-0000-C000-000000000046}')
	coclass_clsid = None

	# Result is of type Row
	def Add(self, BeforeRow=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(100, LCID, 1, (9, 0), ((16396, 17),),BeforeRow
			)
		if ret is not None:
			ret = Dispatch(ret, 'Add', '{00020950-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def ConvertToText(self, Separator=defaultNamedOptArg, NestedTables=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(210, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Separator
			, NestedTables)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToText', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def ConvertToTextOld(self, Separator=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(16, LCID, 1, (9, 0), ((16396, 17),),Separator
			)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTextOld', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def Delete(self):
		return self._oleobj_.InvokeTypes(200, LCID, 1, (24, 0), (),)

	def DistributeHeight(self):
		return self._oleobj_.InvokeTypes(206, LCID, 1, (24, 0), (),)

	# Result is of type Row
	def Item(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, 'Item', '{00020950-0000-0000-C000-000000000046}')
		return ret

	def Select(self):
		return self._oleobj_.InvokeTypes(199, LCID, 1, (24, 0), (),)

	def SetHeight(self, RowHeight=defaultNamedNotOptArg, HeightRule=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(203, LCID, 1, (24, 0), ((4, 1), (3, 1)),RowHeight
			, HeightRule)

	def SetLeftIndent(self, LeftIndent=defaultNamedNotOptArg, RulerStyle=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(202, LCID, 1, (24, 0), ((4, 1), (3, 1)),LeftIndent
			, RulerStyle)

	_prop_map_get_ = {
		"Alignment": (4, 2, (3, 0), (), "Alignment", None),
		"AllowBreakAcrossPages": (3, 2, (3, 0), (), "AllowBreakAcrossPages", None),
		"AllowOverlap": (22, 2, (3, 0), (), "AllowOverlap", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"Count": (2, 2, (3, 0), (), "Count", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DistanceBottom": (14, 2, (4, 0), (), "DistanceBottom", None),
		"DistanceLeft": (20, 2, (4, 0), (), "DistanceLeft", None),
		"DistanceRight": (21, 2, (4, 0), (), "DistanceRight", None),
		"DistanceTop": (13, 2, (4, 0), (), "DistanceTop", None),
		# Method 'First' returns object of type 'Row'
		"First": (10, 2, (9, 0), (), "First", '{00020950-0000-0000-C000-000000000046}'),
		"HeadingFormat": (5, 2, (3, 0), (), "HeadingFormat", None),
		"Height": (7, 2, (4, 0), (), "Height", None),
		"HeightRule": (8, 2, (3, 0), (), "HeightRule", None),
		"HorizontalPosition": (15, 2, (4, 0), (), "HorizontalPosition", None),
		# Method 'Last' returns object of type 'Row'
		"Last": (11, 2, (9, 0), (), "Last", '{00020950-0000-0000-C000-000000000046}'),
		"LeftIndent": (9, 2, (4, 0), (), "LeftIndent", None),
		"NestingLevel": (103, 2, (3, 0), (), "NestingLevel", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"RelativeHorizontalPosition": (18, 2, (3, 0), (), "RelativeHorizontalPosition", None),
		"RelativeVerticalPosition": (19, 2, (3, 0), (), "RelativeVerticalPosition", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (102, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		"SpaceBetweenColumns": (6, 2, (4, 0), (), "SpaceBetweenColumns", None),
		"TableDirection": (104, 2, (3, 0), (), "TableDirection", None),
		"VerticalPosition": (17, 2, (4, 0), (), "VerticalPosition", None),
		"WrapAroundText": (12, 2, (3, 0), (), "WrapAroundText", None),
	}
	_prop_map_put_ = {
		"Alignment": ((4, LCID, 4, 0),()),
		"AllowBreakAcrossPages": ((3, LCID, 4, 0),()),
		"AllowOverlap": ((22, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"DistanceBottom": ((14, LCID, 4, 0),()),
		"DistanceLeft": ((20, LCID, 4, 0),()),
		"DistanceRight": ((21, LCID, 4, 0),()),
		"DistanceTop": ((13, LCID, 4, 0),()),
		"HeadingFormat": ((5, LCID, 4, 0),()),
		"Height": ((7, LCID, 4, 0),()),
		"HeightRule": ((8, LCID, 4, 0),()),
		"HorizontalPosition": ((15, LCID, 4, 0),()),
		"LeftIndent": ((9, LCID, 4, 0),()),
		"RelativeHorizontalPosition": ((18, LCID, 4, 0),()),
		"RelativeVerticalPosition": ((19, LCID, 4, 0),()),
		"SpaceBetweenColumns": ((6, LCID, 4, 0),()),
		"TableDirection": ((104, LCID, 4, 0),()),
		"VerticalPosition": ((17, LCID, 4, 0),()),
		"WrapAroundText": ((12, LCID, 4, 0),()),
	}
	# Default method for this class is 'Item'
	def __call__(self, Index=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(0, LCID, 1, (9, 0), ((3, 1),),Index
			)
		if ret is not None:
			ret = Dispatch(ret, '__call__', '{00020950-0000-0000-C000-000000000046}')
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
		return win32com.client.util.Iterator(ob, '{00020950-0000-0000-C000-000000000046}')
	#This class has Count() property - allow len(ob) to provide this
	def __len__(self):
		return self._ApplyTypes_(*(2, 2, (3, 0), (), "Count", None))
	#This class has a __len__ - this is needed so 'if object:' always returns TRUE.
	def __bool__(self):
		return True

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094C-0000-0000-C000-000000000046}", Rows )
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

Rows_vtables_dispatch_ = 1
Rows_vtables_ = [
	(( '_NewEnum' , 'prop' , ), -4, (-4, (), [ (16397, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 1024 , )),
	(( 'Count' , 'prop' , ), 2, (2, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'AllowBreakAcrossPages' , 'prop' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'AllowBreakAcrossPages' , 'prop' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'HeadingFormat' , 'prop' , ), 5, (5, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'HeadingFormat' , 'prop' , ), 5, (5, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBetweenColumns' , 'prop' , ), 6, (6, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBetweenColumns' , 'prop' , ), 6, (6, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 7, (7, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Height' , 'prop' , ), 7, (7, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'HeightRule' , 'prop' , ), 8, (8, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'HeightRule' , 'prop' , ), 8, (8, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 9, (9, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 9, (9, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'First' , 'prop' , ), 10, (10, (), [ (16393, 10, None, "IID('{00020950-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Last' , 'prop' , ), 11, (11, (), [ (16393, 10, None, "IID('{00020950-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 102, (102, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Item' , 'Index' , 'prop' , ), 0, (0, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{00020950-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Add' , 'BeforeRow' , 'prop' , ), 100, (100, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020950-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 256 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 199, (199, (), [ ], 1 , 1 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , ), 200, (200, (), [ ], 1 , 1 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'SetLeftIndent' , 'LeftIndent' , 'RulerStyle' , ), 202, (202, (), [ (4, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'SetHeight' , 'RowHeight' , 'HeightRule' , ), 203, (203, (), [ (4, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToTextOld' , 'Separator' , 'prop' , ), 16, (16, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 296 , (3, 0, None, None) , 64 , )),
	(( 'DistributeHeight' , ), 206, (206, (), [ ], 1 , 1 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToText' , 'Separator' , 'NestedTables' , 'prop' , ), 210, (210, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 312 , (3, 0, None, None) , 0 , )),
	(( 'WrapAroundText' , 'prop' , ), 12, (12, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'WrapAroundText' , 'prop' , ), 12, (12, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'DistanceTop' , 'prop' , ), 13, (13, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'DistanceTop' , 'prop' , ), 13, (13, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'DistanceBottom' , 'prop' , ), 14, (14, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'DistanceBottom' , 'prop' , ), 14, (14, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'DistanceLeft' , 'prop' , ), 20, (20, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'DistanceLeft' , 'prop' , ), 20, (20, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'DistanceRight' , 'prop' , ), 21, (21, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'DistanceRight' , 'prop' , ), 21, (21, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'HorizontalPosition' , 'prop' , ), 15, (15, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'HorizontalPosition' , 'prop' , ), 15, (15, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'VerticalPosition' , 'prop' , ), 17, (17, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'VerticalPosition' , 'prop' , ), 17, (17, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'RelativeHorizontalPosition' , 'prop' , ), 18, (18, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'RelativeHorizontalPosition' , 'prop' , ), 18, (18, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'RelativeVerticalPosition' , 'prop' , ), 19, (19, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'RelativeVerticalPosition' , 'prop' , ), 19, (19, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'AllowOverlap' , 'prop' , ), 22, (22, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'AllowOverlap' , 'prop' , ), 22, (22, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'NestingLevel' , 'prop' , ), 103, (103, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'TableDirection' , 'prop' , ), 104, (104, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'TableDirection' , 'prop' , ), 104, (104, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{0002094C-0000-0000-C000-000000000046}", Rows )
