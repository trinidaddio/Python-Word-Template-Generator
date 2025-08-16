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
class _Font(DispatchBaseClass):
	CLSID = IID('{00020952-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{000209F5-0000-0000-C000-000000000046}')

	def Grow(self):
		return self._oleobj_.InvokeTypes(100, LCID, 1, (24, 0), (),)

	def Reset(self):
		return self._oleobj_.InvokeTypes(102, LCID, 1, (24, 0), (),)

	def SetAsTemplateDefault(self):
		return self._oleobj_.InvokeTypes(103, LCID, 1, (24, 0), (),)

	def Shrink(self):
		return self._oleobj_.InvokeTypes(101, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"AllCaps": (134, 2, (3, 0), (), "AllCaps", None),
		"Animation": (151, 2, (3, 0), (), "Animation", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"Bold": (130, 2, (3, 0), (), "Bold", None),
		"BoldBi": (160, 2, (3, 0), (), "BoldBi", None),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		"Color": (159, 2, (3, 0), (), "Color", None),
		"ColorIndex": (137, 2, (3, 0), (), "ColorIndex", None),
		"ColorIndexBi": (164, 2, (3, 0), (), "ColorIndexBi", None),
		"ContextualAlternates": (177, 2, (3, 0), (), "ContextualAlternates", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DiacriticColor": (165, 2, (3, 0), (), "DiacriticColor", None),
		"DisableCharacterSpaceGrid": (155, 2, (11, 0), (), "DisableCharacterSpaceGrid", None),
		"DoubleStrikeThrough": (136, 2, (3, 0), (), "DoubleStrikeThrough", None),
		# Method 'Duplicate' returns object of type 'Font'
		"Duplicate": (10, 2, (13, 0), (), "Duplicate", '{000209F5-0000-0000-C000-000000000046}'),
		"Emboss": (148, 2, (3, 0), (), "Emboss", None),
		"EmphasisMark": (154, 2, (3, 0), (), "EmphasisMark", None),
		"Engrave": (150, 2, (3, 0), (), "Engrave", None),
		# Method 'Fill' returns object of type 'FillFormat'
		"Fill": (170, 2, (9, 0), (), "Fill", '{000209C8-0000-0000-C000-000000000046}'),
		# Method 'Glow' returns object of type 'GlowFormat'
		"Glow": (167, 2, (9, 0), (), "Glow", '{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}'),
		"Hidden": (132, 2, (3, 0), (), "Hidden", None),
		"Italic": (131, 2, (3, 0), (), "Italic", None),
		"ItalicBi": (161, 2, (3, 0), (), "ItalicBi", None),
		"Kerning": (149, 2, (4, 0), (), "Kerning", None),
		"Ligatures": (174, 2, (3, 0), (), "Ligatures", None),
		# Method 'Line' returns object of type 'LineFormat'
		"Line": (171, 2, (9, 0), (), "Line", '{000209CA-0000-0000-C000-000000000046}'),
		"Name": (142, 2, (8, 0), (), "Name", None),
		"NameAscii": (157, 2, (8, 0), (), "NameAscii", None),
		"NameBi": (163, 2, (8, 0), (), "NameBi", None),
		"NameFarEast": (156, 2, (8, 0), (), "NameFarEast", None),
		"NameOther": (158, 2, (8, 0), (), "NameOther", None),
		"NumberForm": (175, 2, (3, 0), (), "NumberForm", None),
		"NumberSpacing": (176, 2, (3, 0), (), "NumberSpacing", None),
		"Outline": (147, 2, (3, 0), (), "Outline", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"Position": (143, 2, (3, 0), (), "Position", None),
		# Method 'Reflection' returns object of type 'ReflectionFormat'
		"Reflection": (168, 2, (9, 0), (), "Reflection", '{F01943FF-1985-445E-8602-8FB8F39CCA75}'),
		"Scaling": (145, 2, (3, 0), (), "Scaling", None),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (153, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		"Shadow": (146, 2, (3, 0), (), "Shadow", None),
		"Size": (141, 2, (4, 0), (), "Size", None),
		"SizeBi": (162, 2, (4, 0), (), "SizeBi", None),
		"SmallCaps": (133, 2, (3, 0), (), "SmallCaps", None),
		"Spacing": (144, 2, (4, 0), (), "Spacing", None),
		"StrikeThrough": (135, 2, (3, 0), (), "StrikeThrough", None),
		"StylisticSet": (178, 2, (3, 0), (), "StylisticSet", None),
		"Subscript": (138, 2, (3, 0), (), "Subscript", None),
		"Superscript": (139, 2, (3, 0), (), "Superscript", None),
		# Method 'TextColor' returns object of type 'ColorFormat'
		"TextColor": (173, 2, (9, 0), (), "TextColor", '{000209C6-0000-0000-C000-000000000046}'),
		# Method 'TextShadow' returns object of type 'ShadowFormat'
		"TextShadow": (169, 2, (9, 0), (), "TextShadow", '{000209CC-0000-0000-C000-000000000046}'),
		# Method 'ThreeD' returns object of type 'ThreeDFormat'
		"ThreeD": (172, 2, (9, 0), (), "ThreeD", '{000209D0-0000-0000-C000-000000000046}'),
		"Underline": (140, 2, (3, 0), (), "Underline", None),
		"UnderlineColor": (166, 2, (3, 0), (), "UnderlineColor", None),
	}
	_prop_map_put_ = {
		"AllCaps": ((134, LCID, 4, 0),()),
		"Animation": ((151, LCID, 4, 0),()),
		"Bold": ((130, LCID, 4, 0),()),
		"BoldBi": ((160, LCID, 4, 0),()),
		"Borders": ((1100, LCID, 4, 0),()),
		"Color": ((159, LCID, 4, 0),()),
		"ColorIndex": ((137, LCID, 4, 0),()),
		"ColorIndexBi": ((164, LCID, 4, 0),()),
		"ContextualAlternates": ((177, LCID, 4, 0),()),
		"DiacriticColor": ((165, LCID, 4, 0),()),
		"DisableCharacterSpaceGrid": ((155, LCID, 4, 0),()),
		"DoubleStrikeThrough": ((136, LCID, 4, 0),()),
		"Emboss": ((148, LCID, 4, 0),()),
		"EmphasisMark": ((154, LCID, 4, 0),()),
		"Engrave": ((150, LCID, 4, 0),()),
		"Fill": ((170, LCID, 4, 0),()),
		"Glow": ((167, LCID, 4, 0),()),
		"Hidden": ((132, LCID, 4, 0),()),
		"Italic": ((131, LCID, 4, 0),()),
		"ItalicBi": ((161, LCID, 4, 0),()),
		"Kerning": ((149, LCID, 4, 0),()),
		"Ligatures": ((174, LCID, 4, 0),()),
		"Line": ((171, LCID, 4, 0),()),
		"Name": ((142, LCID, 4, 0),()),
		"NameAscii": ((157, LCID, 4, 0),()),
		"NameBi": ((163, LCID, 4, 0),()),
		"NameFarEast": ((156, LCID, 4, 0),()),
		"NameOther": ((158, LCID, 4, 0),()),
		"NumberForm": ((175, LCID, 4, 0),()),
		"NumberSpacing": ((176, LCID, 4, 0),()),
		"Outline": ((147, LCID, 4, 0),()),
		"Position": ((143, LCID, 4, 0),()),
		"Reflection": ((168, LCID, 4, 0),()),
		"Scaling": ((145, LCID, 4, 0),()),
		"Shadow": ((146, LCID, 4, 0),()),
		"Size": ((141, LCID, 4, 0),()),
		"SizeBi": ((162, LCID, 4, 0),()),
		"SmallCaps": ((133, LCID, 4, 0),()),
		"Spacing": ((144, LCID, 4, 0),()),
		"StrikeThrough": ((135, LCID, 4, 0),()),
		"StylisticSet": ((178, LCID, 4, 0),()),
		"Subscript": ((138, LCID, 4, 0),()),
		"Superscript": ((139, LCID, 4, 0),()),
		"TextShadow": ((169, LCID, 4, 0),()),
		"ThreeD": ((172, LCID, 4, 0),()),
		"Underline": ((140, LCID, 4, 0),()),
		"UnderlineColor": ((166, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020952-0000-0000-C000-000000000046}", _Font )
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

_Font_vtables_dispatch_ = 1
_Font_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Duplicate' , 'prop' , ), 10, (10, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Bold' , 'prop' , ), 130, (130, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Bold' , 'prop' , ), 130, (130, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Italic' , 'prop' , ), 131, (131, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Italic' , 'prop' , ), 131, (131, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Hidden' , 'prop' , ), 132, (132, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Hidden' , 'prop' , ), 132, (132, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'SmallCaps' , 'prop' , ), 133, (133, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'SmallCaps' , 'prop' , ), 133, (133, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'AllCaps' , 'prop' , ), 134, (134, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'AllCaps' , 'prop' , ), 134, (134, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'StrikeThrough' , 'prop' , ), 135, (135, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'StrikeThrough' , 'prop' , ), 135, (135, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'DoubleStrikeThrough' , 'prop' , ), 136, (136, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'DoubleStrikeThrough' , 'prop' , ), 136, (136, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'ColorIndex' , 'prop' , ), 137, (137, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'ColorIndex' , 'prop' , ), 137, (137, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'Subscript' , 'prop' , ), 138, (138, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Subscript' , 'prop' , ), 138, (138, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'Superscript' , 'prop' , ), 139, (139, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'Superscript' , 'prop' , ), 139, (139, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Underline' , 'prop' , ), 140, (140, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Underline' , 'prop' , ), 140, (140, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'Size' , 'prop' , ), 141, (141, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'Size' , 'prop' , ), 141, (141, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'prop' , ), 142, (142, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Name' , 'prop' , ), 142, (142, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Position' , 'prop' , ), 143, (143, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'Position' , 'prop' , ), 143, (143, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Spacing' , 'prop' , ), 144, (144, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'Spacing' , 'prop' , ), 144, (144, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Scaling' , 'prop' , ), 145, (145, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'Scaling' , 'prop' , ), 145, (145, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'Shadow' , 'prop' , ), 146, (146, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Shadow' , 'prop' , ), 146, (146, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'Outline' , 'prop' , ), 147, (147, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'Outline' , 'prop' , ), 147, (147, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'Emboss' , 'prop' , ), 148, (148, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'Emboss' , 'prop' , ), 148, (148, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'Kerning' , 'prop' , ), 149, (149, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'Kerning' , 'prop' , ), 149, (149, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'Engrave' , 'prop' , ), 150, (150, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'Engrave' , 'prop' , ), 150, (150, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'Animation' , 'prop' , ), 151, (151, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 424 , (3, 0, None, None) , 64 , )),
	(( 'Animation' , 'prop' , ), 151, (151, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 432 , (3, 0, None, None) , 64 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 153, (153, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'EmphasisMark' , 'prop' , ), 154, (154, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'EmphasisMark' , 'prop' , ), 154, (154, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'DisableCharacterSpaceGrid' , 'prop' , ), 155, (155, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'DisableCharacterSpaceGrid' , 'prop' , ), 155, (155, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'NameFarEast' , 'prop' , ), 156, (156, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'NameFarEast' , 'prop' , ), 156, (156, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'NameAscii' , 'prop' , ), 157, (157, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'NameAscii' , 'prop' , ), 157, (157, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'NameOther' , 'prop' , ), 158, (158, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'NameOther' , 'prop' , ), 158, (158, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'Grow' , ), 100, (100, (), [ ], 1 , 1 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'Shrink' , ), 101, (101, (), [ ], 1 , 1 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'Reset' , ), 102, (102, (), [ ], 1 , 1 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'SetAsTemplateDefault' , ), 103, (103, (), [ ], 1 , 1 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'Color' , 'prop' , ), 159, (159, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 64 , )),
	(( 'Color' , 'prop' , ), 159, (159, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 584 , (3, 0, None, None) , 64 , )),
	(( 'BoldBi' , 'prop' , ), 160, (160, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'BoldBi' , 'prop' , ), 160, (160, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'ItalicBi' , 'prop' , ), 161, (161, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'ItalicBi' , 'prop' , ), 161, (161, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'SizeBi' , 'prop' , ), 162, (162, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'SizeBi' , 'prop' , ), 162, (162, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'NameBi' , 'prop' , ), 163, (163, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'NameBi' , 'prop' , ), 163, (163, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'ColorIndexBi' , 'prop' , ), 164, (164, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 656 , (3, 0, None, None) , 0 , )),
	(( 'ColorIndexBi' , 'prop' , ), 164, (164, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 664 , (3, 0, None, None) , 0 , )),
	(( 'DiacriticColor' , 'prop' , ), 165, (165, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 672 , (3, 0, None, None) , 0 , )),
	(( 'DiacriticColor' , 'prop' , ), 165, (165, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 680 , (3, 0, None, None) , 0 , )),
	(( 'UnderlineColor' , 'prop' , ), 166, (166, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 688 , (3, 0, None, None) , 0 , )),
	(( 'UnderlineColor' , 'prop' , ), 166, (166, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 696 , (3, 0, None, None) , 0 , )),
	(( 'Glow' , 'prop' , ), 167, (167, (), [ (16393, 10, None, "IID('{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}')") , ], 1 , 2 , 4 , 0 , 704 , (3, 0, None, None) , 0 , )),
	(( 'Glow' , 'prop' , ), 167, (167, (), [ (9, 1, None, "IID('{F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D}')") , ], 1 , 4 , 4 , 0 , 712 , (3, 0, None, None) , 0 , )),
	(( 'Reflection' , 'prop' , ), 168, (168, (), [ (16393, 10, None, "IID('{F01943FF-1985-445E-8602-8FB8F39CCA75}')") , ], 1 , 2 , 4 , 0 , 720 , (3, 0, None, None) , 0 , )),
	(( 'Reflection' , 'prop' , ), 168, (168, (), [ (9, 1, None, "IID('{F01943FF-1985-445E-8602-8FB8F39CCA75}')") , ], 1 , 4 , 4 , 0 , 728 , (3, 0, None, None) , 0 , )),
	(( 'TextShadow' , 'prop' , ), 169, (169, (), [ (16393, 10, None, "IID('{000209CC-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 736 , (3, 0, None, None) , 0 , )),
	(( 'TextShadow' , 'prop' , ), 169, (169, (), [ (9, 1, None, "IID('{000209CC-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 744 , (3, 0, None, None) , 0 , )),
	(( 'Fill' , 'prop' , ), 170, (170, (), [ (16393, 10, None, "IID('{000209C8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 752 , (3, 0, None, None) , 0 , )),
	(( 'Fill' , 'prop' , ), 170, (170, (), [ (9, 1, None, "IID('{000209C8-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'Line' , 'prop' , ), 171, (171, (), [ (16393, 10, None, "IID('{000209CA-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'Line' , 'prop' , ), 171, (171, (), [ (9, 1, None, "IID('{000209CA-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'ThreeD' , 'prop' , ), 172, (172, (), [ (16393, 10, None, "IID('{000209D0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 784 , (3, 0, None, None) , 0 , )),
	(( 'ThreeD' , 'prop' , ), 172, (172, (), [ (9, 1, None, "IID('{000209D0-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 792 , (3, 0, None, None) , 0 , )),
	(( 'TextColor' , 'prop' , ), 173, (173, (), [ (16393, 10, None, "IID('{000209C6-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'Ligatures' , 'prop' , ), 174, (174, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'Ligatures' , 'prop' , ), 174, (174, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 816 , (3, 0, None, None) , 0 , )),
	(( 'NumberForm' , 'prop' , ), 175, (175, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 824 , (3, 0, None, None) , 0 , )),
	(( 'NumberForm' , 'prop' , ), 175, (175, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'NumberSpacing' , 'prop' , ), 176, (176, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'NumberSpacing' , 'prop' , ), 176, (176, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 848 , (3, 0, None, None) , 0 , )),
	(( 'ContextualAlternates' , 'prop' , ), 177, (177, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 856 , (3, 0, None, None) , 0 , )),
	(( 'ContextualAlternates' , 'prop' , ), 177, (177, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 864 , (3, 0, None, None) , 0 , )),
	(( 'StylisticSet' , 'prop' , ), 178, (178, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 872 , (3, 0, None, None) , 0 , )),
	(( 'StylisticSet' , 'prop' , ), 178, (178, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 880 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020952-0000-0000-C000-000000000046}", _Font )
