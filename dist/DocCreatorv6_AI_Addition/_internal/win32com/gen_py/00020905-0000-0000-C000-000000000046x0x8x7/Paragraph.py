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
class Paragraph(DispatchBaseClass):
	CLSID = IID('{00020957-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def CloseUp(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0), (),)

	def Indent(self):
		return self._oleobj_.InvokeTypes(333, LCID, 1, (24, 0), (),)

	def IndentCharWidth(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(320, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def IndentFirstLineCharWidth(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(322, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def JoinList(self):
		return self._oleobj_.InvokeTypes(339, LCID, 1, (24, 0), (),)

	def ListAdvanceTo(self, Level1=0, Level2=0, Level3=0, Level4=0
			, Level5=0, Level6=0, Level7=0, Level8=0, Level9=0):
		return self._oleobj_.InvokeTypes(336, LCID, 1, (24, 0), ((2, 49), (2, 49), (2, 49), (2, 49), (2, 49), (2, 49), (2, 49), (2, 49), (2, 49)),Level1
			, Level2, Level3, Level4, Level5, Level6
			, Level7, Level8, Level9)

	# The method ListNumberOriginal is actually a property, but must be used as a method to correctly pass the arguments
	def ListNumberOriginal(self, Level=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(137, LCID, 2, (2, 0), ((2, 1),),Level
			)

	# Result is of type Paragraph
	def Next(self, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(324, LCID, 1, (9, 0), ((16396, 17),),Count
			)
		if ret is not None:
			ret = Dispatch(ret, 'Next', '{00020957-0000-0000-C000-000000000046}')
		return ret

	def OpenOrCloseUp(self):
		return self._oleobj_.InvokeTypes(303, LCID, 1, (24, 0), (),)

	def OpenUp(self):
		return self._oleobj_.InvokeTypes(302, LCID, 1, (24, 0), (),)

	def Outdent(self):
		return self._oleobj_.InvokeTypes(334, LCID, 1, (24, 0), (),)

	def OutlineDemote(self):
		return self._oleobj_.InvokeTypes(327, LCID, 1, (24, 0), (),)

	def OutlineDemoteToBody(self):
		return self._oleobj_.InvokeTypes(328, LCID, 1, (24, 0), (),)

	def OutlinePromote(self):
		return self._oleobj_.InvokeTypes(326, LCID, 1, (24, 0), (),)

	# Result is of type Paragraph
	def Previous(self, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(325, LCID, 1, (9, 0), ((16396, 17),),Count
			)
		if ret is not None:
			ret = Dispatch(ret, 'Previous', '{00020957-0000-0000-C000-000000000046}')
		return ret

	def Reset(self):
		return self._oleobj_.InvokeTypes(312, LCID, 1, (24, 0), (),)

	def ResetAdvanceTo(self):
		return self._oleobj_.InvokeTypes(337, LCID, 1, (24, 0), (),)

	def SelectNumber(self):
		return self._oleobj_.InvokeTypes(335, LCID, 1, (24, 0), (),)

	def SeparateList(self):
		return self._oleobj_.InvokeTypes(338, LCID, 1, (24, 0), (),)

	def Space1(self):
		return self._oleobj_.InvokeTypes(313, LCID, 1, (24, 0), (),)

	def Space15(self):
		return self._oleobj_.InvokeTypes(314, LCID, 1, (24, 0), (),)

	def Space2(self):
		return self._oleobj_.InvokeTypes(315, LCID, 1, (24, 0), (),)

	def TabHangingIndent(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(304, LCID, 1, (24, 0), ((2, 1),),Count
			)

	def TabIndent(self, Count=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(306, LCID, 1, (24, 0), ((2, 1),),Count
			)

	_prop_map_get_ = {
		"AddSpaceBetweenFarEastAndAlpha": (121, 2, (3, 0), (), "AddSpaceBetweenFarEastAndAlpha", None),
		"AddSpaceBetweenFarEastAndDigit": (122, 2, (3, 0), (), "AddSpaceBetweenFarEastAndDigit", None),
		"Alignment": (101, 2, (3, 0), (), "Alignment", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"AutoAdjustRightIndent": (124, 2, (3, 0), (), "AutoAdjustRightIndent", None),
		"BaseLineAlignment": (123, 2, (3, 0), (), "BaseLineAlignment", None),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"CharacterUnitFirstLineIndent": (128, 2, (4, 0), (), "CharacterUnitFirstLineIndent", None),
		"CharacterUnitLeftIndent": (127, 2, (4, 0), (), "CharacterUnitLeftIndent", None),
		"CharacterUnitRightIndent": (126, 2, (4, 0), (), "CharacterUnitRightIndent", None),
		"CollapseHeadingByDefault": (1204, 2, (11, 0), (), "CollapseHeadingByDefault", None),
		"CollapsedState": (1203, 2, (11, 0), (), "CollapsedState", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DisableLineHeightGrid": (125, 2, (3, 0), (), "DisableLineHeightGrid", None),
		# Method 'DropCap' returns object of type 'DropCap'
		"DropCap": (13, 2, (9, 0), (), "DropCap", '{00020956-0000-0000-C000-000000000046}'),
		"FarEastLineBreakControl": (117, 2, (3, 0), (), "FarEastLineBreakControl", None),
		"FirstLineIndent": (108, 2, (4, 0), (), "FirstLineIndent", None),
		# Method 'Format' returns object of type 'ParagraphFormat'
		"Format": (1102, 2, (13, 0), (), "Format", '{000209F4-0000-0000-C000-000000000046}'),
		"HalfWidthPunctuationOnTopOfLine": (120, 2, (3, 0), (), "HalfWidthPunctuationOnTopOfLine", None),
		"HangingPunctuation": (119, 2, (3, 0), (), "HangingPunctuation", None),
		"Hyphenation": (113, 2, (3, 0), (), "Hyphenation", None),
		"ID": (204, 2, (8, 0), (), "ID", None),
		"IsStyleSeparator": (134, 2, (11, 0), (), "IsStyleSeparator", None),
		"KeepTogether": (102, 2, (3, 0), (), "KeepTogether", None),
		"KeepWithNext": (103, 2, (3, 0), (), "KeepWithNext", None),
		"LeftIndent": (107, 2, (4, 0), (), "LeftIndent", None),
		"LineSpacing": (109, 2, (4, 0), (), "LineSpacing", None),
		"LineSpacingRule": (110, 2, (3, 0), (), "LineSpacingRule", None),
		"LineUnitAfter": (130, 2, (4, 0), (), "LineUnitAfter", None),
		"LineUnitBefore": (129, 2, (4, 0), (), "LineUnitBefore", None),
		"MirrorIndents": (135, 2, (3, 0), (), "MirrorIndents", None),
		"NoLineNumber": (105, 2, (3, 0), (), "NoLineNumber", None),
		"OutlineLevel": (202, 2, (3, 0), (), "OutlineLevel", None),
		"PageBreakBefore": (104, 2, (3, 0), (), "PageBreakBefore", None),
		"ParaID": (138, 2, (3, 0), (), "ParaID", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		# Method 'Range' returns object of type 'Range'
		"Range": (0, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'),
		"ReadingOrder": (203, 2, (3, 0), (), "ReadingOrder", None),
		"RightIndent": (106, 2, (4, 0), (), "RightIndent", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (116, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		"SpaceAfter": (112, 2, (4, 0), (), "SpaceAfter", None),
		"SpaceAfterAuto": (133, 2, (3, 0), (), "SpaceAfterAuto", None),
		"SpaceBefore": (111, 2, (4, 0), (), "SpaceBefore", None),
		"SpaceBeforeAuto": (132, 2, (3, 0), (), "SpaceBeforeAuto", None),
		"Style": (100, 2, (12, 0), (), "Style", None),
		# Method 'TabStops' returns object of type 'TabStops'
		"TabStops": (1103, 2, (9, 0), (), "TabStops", '{00020955-0000-0000-C000-000000000046}'),
		"TextID": (140, 2, (3, 0), (), "TextID", None),
		"TextboxTightWrap": (136, 2, (3, 0), (), "TextboxTightWrap", None),
		"WidowControl": (114, 2, (3, 0), (), "WidowControl", None),
		"WordWrap": (118, 2, (3, 0), (), "WordWrap", None),
	}
	_prop_map_put_ = {
		"AddSpaceBetweenFarEastAndAlpha": ((121, LCID, 4, 0),()),
		"AddSpaceBetweenFarEastAndDigit": ((122, LCID, 4, 0),()),
		"Alignment": ((101, LCID, 4, 0),()),
		"AutoAdjustRightIndent": ((124, LCID, 4, 0),()),
		"BaseLineAlignment": ((123, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"CharacterUnitFirstLineIndent": ((128, LCID, 4, 0),()),
		"CharacterUnitLeftIndent": ((127, LCID, 4, 0),()),
		"CharacterUnitRightIndent": ((126, LCID, 4, 0),()),
		"CollapseHeadingByDefault": ((1204, LCID, 4, 0),()),
		"CollapsedState": ((1203, LCID, 4, 0),()),
		"DisableLineHeightGrid": ((125, LCID, 4, 0),()),
		"FarEastLineBreakControl": ((117, LCID, 4, 0),()),
		"FirstLineIndent": ((108, LCID, 4, 0),()),
		"Format": ((1102, LCID, 4, 0),()),
		"HalfWidthPunctuationOnTopOfLine": ((120, LCID, 4, 0),()),
		"HangingPunctuation": ((119, LCID, 4, 0),()),
		"Hyphenation": ((113, LCID, 4, 0),()),
		"ID": ((204, LCID, 4, 0),()),
		"KeepTogether": ((102, LCID, 4, 0),()),
		"KeepWithNext": ((103, LCID, 4, 0),()),
		"LeftIndent": ((107, LCID, 4, 0),()),
		"LineSpacing": ((109, LCID, 4, 0),()),
		"LineSpacingRule": ((110, LCID, 4, 0),()),
		"LineUnitAfter": ((130, LCID, 4, 0),()),
		"LineUnitBefore": ((129, LCID, 4, 0),()),
		"MirrorIndents": ((135, LCID, 4, 0),()),
		"NoLineNumber": ((105, LCID, 4, 0),()),
		"OutlineLevel": ((202, LCID, 4, 0),()),
		"PageBreakBefore": ((104, LCID, 4, 0),()),
		"ReadingOrder": ((203, LCID, 4, 0),()),
		"RightIndent": ((106, LCID, 4, 0),()),
		"SpaceAfter": ((112, LCID, 4, 0),()),
		"SpaceAfterAuto": ((133, LCID, 4, 0),()),
		"SpaceBefore": ((111, LCID, 4, 0),()),
		"SpaceBeforeAuto": ((132, LCID, 4, 0),()),
		"Style": ((100, LCID, 4, 0),()),
		"TabStops": ((1103, LCID, 4, 0),()),
		"TextboxTightWrap": ((136, LCID, 4, 0),()),
		"WidowControl": ((114, LCID, 4, 0),()),
		"WordWrap": ((118, LCID, 4, 0),()),
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020957-0000-0000-C000-000000000046}", Paragraph )
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

Paragraph_vtables_dispatch_ = 1
Paragraph_vtables_ = [
	(( 'Range' , 'prop' , ), 0, (0, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Format' , 'prop' , ), 1102, (1102, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Format' , 'prop' , ), 1102, (1102, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'TabStops' , 'prop' , ), 1103, (1103, (), [ (16393, 10, None, "IID('{00020955-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'TabStops' , 'prop' , ), 1103, (1103, (), [ (9, 1, None, "IID('{00020955-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'DropCap' , 'prop' , ), 13, (13, (), [ (16393, 10, None, "IID('{00020956-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 100, (100, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 100, (100, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 101, (101, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Alignment' , 'prop' , ), 101, (101, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'KeepTogether' , 'prop' , ), 102, (102, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'KeepTogether' , 'prop' , ), 102, (102, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'KeepWithNext' , 'prop' , ), 103, (103, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'KeepWithNext' , 'prop' , ), 103, (103, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'PageBreakBefore' , 'prop' , ), 104, (104, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'PageBreakBefore' , 'prop' , ), 104, (104, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'NoLineNumber' , 'prop' , ), 105, (105, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'NoLineNumber' , 'prop' , ), 105, (105, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'RightIndent' , 'prop' , ), 106, (106, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'RightIndent' , 'prop' , ), 106, (106, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 107, (107, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'LeftIndent' , 'prop' , ), 107, (107, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'FirstLineIndent' , 'prop' , ), 108, (108, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'FirstLineIndent' , 'prop' , ), 108, (108, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacing' , 'prop' , ), 109, (109, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacing' , 'prop' , ), 109, (109, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacingRule' , 'prop' , ), 110, (110, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'LineSpacingRule' , 'prop' , ), 110, (110, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBefore' , 'prop' , ), 111, (111, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBefore' , 'prop' , ), 111, (111, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfter' , 'prop' , ), 112, (112, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfter' , 'prop' , ), 112, (112, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Hyphenation' , 'prop' , ), 113, (113, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'Hyphenation' , 'prop' , ), 113, (113, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'WidowControl' , 'prop' , ), 114, (114, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'WidowControl' , 'prop' , ), 114, (114, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 116, (116, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakControl' , 'prop' , ), 117, (117, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'FarEastLineBreakControl' , 'prop' , ), 117, (117, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 118, (118, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'WordWrap' , 'prop' , ), 118, (118, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'HangingPunctuation' , 'prop' , ), 119, (119, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'HangingPunctuation' , 'prop' , ), 119, (119, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'HalfWidthPunctuationOnTopOfLine' , 'prop' , ), 120, (120, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'HalfWidthPunctuationOnTopOfLine' , 'prop' , ), 120, (120, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndAlpha' , 'prop' , ), 121, (121, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndAlpha' , 'prop' , ), 121, (121, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndDigit' , 'prop' , ), 122, (122, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'AddSpaceBetweenFarEastAndDigit' , 'prop' , ), 122, (122, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'BaseLineAlignment' , 'prop' , ), 123, (123, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'BaseLineAlignment' , 'prop' , ), 123, (123, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'AutoAdjustRightIndent' , 'prop' , ), 124, (124, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'AutoAdjustRightIndent' , 'prop' , ), 124, (124, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'DisableLineHeightGrid' , 'prop' , ), 125, (125, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'DisableLineHeightGrid' , 'prop' , ), 125, (125, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'OutlineLevel' , 'prop' , ), 202, (202, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'OutlineLevel' , 'prop' , ), 202, (202, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'CloseUp' , ), 301, (301, (), [ ], 1 , 1 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'OpenUp' , ), 302, (302, (), [ ], 1 , 1 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'OpenOrCloseUp' , ), 303, (303, (), [ ], 1 , 1 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'TabHangingIndent' , 'Count' , ), 304, (304, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'TabIndent' , 'Count' , ), 306, (306, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'Reset' , ), 312, (312, (), [ ], 1 , 1 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'Space1' , ), 313, (313, (), [ ], 1 , 1 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'Space15' , ), 314, (314, (), [ ], 1 , 1 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'Space2' , ), 315, (315, (), [ ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'IndentCharWidth' , 'Count' , ), 320, (320, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'IndentFirstLineCharWidth' , 'Count' , ), 322, (322, (), [ (2, 1, None, None) , ], 1 , 1 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'Count' , 'prop' , ), 324, (324, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020957-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 640 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'Count' , 'prop' , ), 325, (325, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020957-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 648 , (3, 0, None, None) , 0 , )),
	(( 'OutlinePromote' , ), 326, (326, (), [ ], 1 , 1 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'OutlineDemote' , ), 327, (327, (), [ ], 1 , 1 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'OutlineDemoteToBody' , ), 328, (328, (), [ ], 1 , 1 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'Indent' , ), 333, (333, (), [ ], 1 , 1 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'Outdent' , ), 334, (334, (), [ ], 1 , 1 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitRightIndent' , 'prop' , ), 126, (126, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitRightIndent' , 'prop' , ), 126, (126, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitLeftIndent' , 'prop' , ), 127, (127, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitLeftIndent' , 'prop' , ), 127, (127, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitFirstLineIndent' , 'prop' , ), 128, (128, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'CharacterUnitFirstLineIndent' , 'prop' , ), 128, (128, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitBefore' , 'prop' , ), 129, (129, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitBefore' , 'prop' , ), 129, (129, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitAfter' , 'prop' , ), 130, (130, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'LineUnitAfter' , 'prop' , ), 130, (130, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'ReadingOrder' , 'prop' , ), 203, (203, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'ReadingOrder' , 'prop' , ), 203, (203, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 204, (204, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'ID' , 'prop' , ), 204, (204, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBeforeAuto' , 'prop' , ), 132, (132, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'SpaceBeforeAuto' , 'prop' , ), 132, (132, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfterAuto' , 'prop' , ), 133, (133, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'SpaceAfterAuto' , 'prop' , ), 133, (133, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'IsStyleSeparator' , 'prop' , ), 134, (134, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'SelectNumber' , ), 335, (335, (), [ ], 1 , 1 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'ListAdvanceTo' , 'Level1' , 'Level2' , 'Level3' , 'Level4' , 
			 'Level5' , 'Level6' , 'Level7' , 'Level8' , 'Level9' , 
			 ), 336, (336, (), [ (2, 49, '0', None) , (2, 49, '0', None) , (2, 49, '0', None) , (2, 49, '0', None) , 
			 (2, 49, '0', None) , (2, 49, '0', None) , (2, 49, '0', None) , (2, 49, '0', None) , (2, 49, '0', None) , ], 1 , 1 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'ResetAdvanceTo' , ), 337, (337, (), [ ], 1 , 1 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'SeparateList' , ), 338, (338, (), [ ], 1 , 1 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'JoinList' , ), 339, (339, (), [ ], 1 , 1 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
	(( 'MirrorIndents' , 'prop' , ), 135, (135, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'MirrorIndents' , 'prop' , ), 135, (135, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 896 , (3, 0, None, None) , 0 , )),
	(( 'TextboxTightWrap' , 'prop' , ), 136, (136, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'TextboxTightWrap' , 'prop' , ), 136, (136, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'ListNumberOriginal' , 'Level' , 'prop' , ), 137, (137, (), [ (2, 1, None, None) , 
			 (16386, 10, None, None) , ], 1 , 2 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'ParaID' , 'prop' , ), 138, (138, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 928 , (3, 0, None, None) , 64 , )),
	(( 'TextID' , 'prop' , ), 140, (140, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 936 , (3, 0, None, None) , 64 , )),
	(( 'CollapsedState' , 'prop' , ), 1203, (1203, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'CollapsedState' , 'prop' , ), 1203, (1203, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'CollapseHeadingByDefault' , 'prop' , ), 1204, (1204, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 960 , (3, 0, None, None) , 0 , )),
	(( 'CollapseHeadingByDefault' , 'prop' , ), 1204, (1204, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020957-0000-0000-C000-000000000046}", Paragraph )
