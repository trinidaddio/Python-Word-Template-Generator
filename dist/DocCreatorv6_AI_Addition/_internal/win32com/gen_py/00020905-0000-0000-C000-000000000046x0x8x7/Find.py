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
class Find(DispatchBaseClass):
	CLSID = IID('{000209B0-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def ClearAllFuzzyOptions(self):
		return self._oleobj_.InvokeTypes(33, LCID, 1, (24, 0), (),)

	def ClearFormatting(self):
		return self._oleobj_.InvokeTypes(31, LCID, 1, (24, 0), (),)

	def ClearHitHighlight(self):
		return self._oleobj_.InvokeTypes(446, LCID, 1, (11, 0), (),)

	def Execute(self, FindText=defaultNamedOptArg, MatchCase=defaultNamedOptArg, MatchWholeWord=defaultNamedOptArg, MatchWildcards=defaultNamedOptArg
			, MatchSoundsLike=defaultNamedOptArg, MatchAllWordForms=defaultNamedOptArg, Forward=defaultNamedOptArg, Wrap=defaultNamedOptArg, Format=defaultNamedOptArg
			, ReplaceWith=defaultNamedOptArg, Replace=defaultNamedOptArg, MatchKashida=defaultNamedOptArg, MatchDiacritics=defaultNamedOptArg, MatchAlefHamza=defaultNamedOptArg
			, MatchControl=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(444, LCID, 1, (11, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FindText
			, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms
			, Forward, Wrap, Format, ReplaceWith, Replace
			, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl)

	def Execute2007(self, FindText=defaultNamedOptArg, MatchCase=defaultNamedOptArg, MatchWholeWord=defaultNamedOptArg, MatchWildcards=defaultNamedOptArg
			, MatchSoundsLike=defaultNamedOptArg, MatchAllWordForms=defaultNamedOptArg, Forward=defaultNamedOptArg, Wrap=defaultNamedOptArg, Format=defaultNamedOptArg
			, ReplaceWith=defaultNamedOptArg, Replace=defaultNamedOptArg, MatchKashida=defaultNamedOptArg, MatchDiacritics=defaultNamedOptArg, MatchAlefHamza=defaultNamedOptArg
			, MatchControl=defaultNamedOptArg, MatchPrefix=defaultNamedOptArg, MatchSuffix=defaultNamedOptArg, MatchPhrase=defaultNamedOptArg, IgnoreSpace=defaultNamedOptArg
			, IgnorePunct=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(447, LCID, 1, (11, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FindText
			, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms
			, Forward, Wrap, Format, ReplaceWith, Replace
			, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl, MatchPrefix
			, MatchSuffix, MatchPhrase, IgnoreSpace, IgnorePunct)

	def ExecuteOld(self, FindText=defaultNamedOptArg, MatchCase=defaultNamedOptArg, MatchWholeWord=defaultNamedOptArg, MatchWildcards=defaultNamedOptArg
			, MatchSoundsLike=defaultNamedOptArg, MatchAllWordForms=defaultNamedOptArg, Forward=defaultNamedOptArg, Wrap=defaultNamedOptArg, Format=defaultNamedOptArg
			, ReplaceWith=defaultNamedOptArg, Replace=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(30, LCID, 1, (11, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FindText
			, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms
			, Forward, Wrap, Format, ReplaceWith, Replace
			)

	def HitHighlight(self, FindText=defaultNamedNotOptArg, HighlightColor=defaultNamedOptArg, TextColor=defaultNamedOptArg, MatchCase=defaultNamedOptArg
			, MatchWholeWord=defaultNamedOptArg, MatchPrefix=defaultNamedOptArg, MatchSuffix=defaultNamedOptArg, MatchPhrase=defaultNamedOptArg, MatchWildcards=defaultNamedOptArg
			, MatchSoundsLike=defaultNamedOptArg, MatchAllWordForms=defaultNamedOptArg, MatchByte=defaultNamedOptArg, MatchFuzzy=defaultNamedOptArg, MatchKashida=defaultNamedOptArg
			, MatchDiacritics=defaultNamedOptArg, MatchAlefHamza=defaultNamedOptArg, MatchControl=defaultNamedOptArg, IgnoreSpace=defaultNamedOptArg, IgnorePunct=defaultNamedOptArg
			, HanjaPhoneticHangul=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(445, LCID, 1, (11, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FindText
			, HighlightColor, TextColor, MatchCase, MatchWholeWord, MatchPrefix
			, MatchSuffix, MatchPhrase, MatchWildcards, MatchSoundsLike, MatchAllWordForms
			, MatchByte, MatchFuzzy, MatchKashida, MatchDiacritics, MatchAlefHamza
			, MatchControl, IgnoreSpace, IgnorePunct, HanjaPhoneticHangul)

	def SetAllFuzzyOptions(self):
		return self._oleobj_.InvokeTypes(32, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"CorrectHangulEndings": (61, 2, (11, 0), (), "CorrectHangulEndings", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'Font' returns object of type 'Font'
		"Font": (11, 2, (13, 0), (), "Font", '{000209F5-0000-0000-C000-000000000046}'),
		"Format": (28, 2, (11, 0), (), "Format", None),
		"Forward": (10, 2, (11, 0), (), "Forward", None),
		"Found": (12, 2, (11, 0), (), "Found", None),
		# Method 'Frame' returns object of type 'Frame'
		"Frame": (26, 2, (9, 0), (), "Frame", '{0002092A-0000-0000-C000-000000000046}'),
		"HanjaPhoneticHangul": (109, 2, (11, 0), (), "HanjaPhoneticHangul", None),
		"Highlight": (24, 2, (3, 0), (), "Highlight", None),
		"IgnorePunct": (108, 2, (11, 0), (), "IgnorePunct", None),
		"IgnoreSpace": (107, 2, (11, 0), (), "IgnoreSpace", None),
		"LanguageID": (23, 2, (3, 0), (), "LanguageID", None),
		"LanguageIDFarEast": (29, 2, (3, 0), (), "LanguageIDFarEast", None),
		"LanguageIDOther": (60, 2, (3, 0), (), "LanguageIDOther", None),
		"MatchAlefHamza": (102, 2, (11, 0), (), "MatchAlefHamza", None),
		"MatchAllWordForms": (13, 2, (11, 0), (), "MatchAllWordForms", None),
		"MatchByte": (41, 2, (11, 0), (), "MatchByte", None),
		"MatchCase": (14, 2, (11, 0), (), "MatchCase", None),
		"MatchControl": (103, 2, (11, 0), (), "MatchControl", None),
		"MatchDiacritics": (101, 2, (11, 0), (), "MatchDiacritics", None),
		"MatchFuzzy": (40, 2, (11, 0), (), "MatchFuzzy", None),
		"MatchKashida": (100, 2, (11, 0), (), "MatchKashida", None),
		"MatchPhrase": (104, 2, (11, 0), (), "MatchPhrase", None),
		"MatchPrefix": (105, 2, (11, 0), (), "MatchPrefix", None),
		"MatchSoundsLike": (16, 2, (11, 0), (), "MatchSoundsLike", None),
		"MatchSuffix": (106, 2, (11, 0), (), "MatchSuffix", None),
		"MatchWholeWord": (17, 2, (11, 0), (), "MatchWholeWord", None),
		"MatchWildcards": (15, 2, (11, 0), (), "MatchWildcards", None),
		"NoProofing": (34, 2, (3, 0), (), "NoProofing", None),
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat'
		"ParagraphFormat": (18, 2, (13, 0), (), "ParagraphFormat", '{000209F4-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		# Method 'Replacement' returns object of type 'Replacement'
		"Replacement": (25, 2, (9, 0), (), "Replacement", '{000209B1-0000-0000-C000-000000000046}'),
		"Style": (19, 2, (12, 0), (), "Style", None),
		"Text": (22, 2, (8, 0), (), "Text", None),
		"Wrap": (27, 2, (3, 0), (), "Wrap", None),
	}
	_prop_map_put_ = {
		"CorrectHangulEndings": ((61, LCID, 4, 0),()),
		"Font": ((11, LCID, 4, 0),()),
		"Format": ((28, LCID, 4, 0),()),
		"Forward": ((10, LCID, 4, 0),()),
		"HanjaPhoneticHangul": ((109, LCID, 4, 0),()),
		"Highlight": ((24, LCID, 4, 0),()),
		"IgnorePunct": ((108, LCID, 4, 0),()),
		"IgnoreSpace": ((107, LCID, 4, 0),()),
		"LanguageID": ((23, LCID, 4, 0),()),
		"LanguageIDFarEast": ((29, LCID, 4, 0),()),
		"LanguageIDOther": ((60, LCID, 4, 0),()),
		"MatchAlefHamza": ((102, LCID, 4, 0),()),
		"MatchAllWordForms": ((13, LCID, 4, 0),()),
		"MatchByte": ((41, LCID, 4, 0),()),
		"MatchCase": ((14, LCID, 4, 0),()),
		"MatchControl": ((103, LCID, 4, 0),()),
		"MatchDiacritics": ((101, LCID, 4, 0),()),
		"MatchFuzzy": ((40, LCID, 4, 0),()),
		"MatchKashida": ((100, LCID, 4, 0),()),
		"MatchPhrase": ((104, LCID, 4, 0),()),
		"MatchPrefix": ((105, LCID, 4, 0),()),
		"MatchSoundsLike": ((16, LCID, 4, 0),()),
		"MatchSuffix": ((106, LCID, 4, 0),()),
		"MatchWholeWord": ((17, LCID, 4, 0),()),
		"MatchWildcards": ((15, LCID, 4, 0),()),
		"NoProofing": ((34, LCID, 4, 0),()),
		"ParagraphFormat": ((18, LCID, 4, 0),()),
		"Style": ((19, LCID, 4, 0),()),
		"Text": ((22, LCID, 4, 0),()),
		"Wrap": ((27, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{000209B0-0000-0000-C000-000000000046}", Find )
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

Find_vtables_dispatch_ = 1
Find_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'Forward' , 'prop' , ), 10, (10, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Forward' , 'prop' , ), 10, (10, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 11, (11, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 11, (11, (), [ (13, 1, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'Found' , 'prop' , ), 12, (12, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'MatchAllWordForms' , 'prop' , ), 13, (13, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'MatchAllWordForms' , 'prop' , ), 13, (13, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'MatchCase' , 'prop' , ), 14, (14, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'MatchCase' , 'prop' , ), 14, (14, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'MatchWildcards' , 'prop' , ), 15, (15, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'MatchWildcards' , 'prop' , ), 15, (15, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'MatchSoundsLike' , 'prop' , ), 16, (16, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'MatchSoundsLike' , 'prop' , ), 16, (16, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'MatchWholeWord' , 'prop' , ), 17, (17, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'MatchWholeWord' , 'prop' , ), 17, (17, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'MatchFuzzy' , 'prop' , ), 40, (40, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'MatchFuzzy' , 'prop' , ), 40, (40, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'MatchByte' , 'prop' , ), 41, (41, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'MatchByte' , 'prop' , ), 41, (41, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 18, (18, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 18, (18, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 19, (19, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 19, (19, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 22, (22, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 22, (22, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 23, (23, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 23, (23, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'Highlight' , 'prop' , ), 24, (24, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'Highlight' , 'prop' , ), 24, (24, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'Replacement' , 'prop' , ), 25, (25, (), [ (16393, 10, None, "IID('{000209B1-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'Frame' , 'prop' , ), 26, (26, (), [ (16393, 10, None, "IID('{0002092A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Wrap' , 'prop' , ), 27, (27, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'Wrap' , 'prop' , ), 27, (27, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'Format' , 'prop' , ), 28, (28, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'Format' , 'prop' , ), 28, (28, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 29, (29, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 29, (29, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 60, (60, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 60, (60, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'CorrectHangulEndings' , 'prop' , ), 61, (61, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'CorrectHangulEndings' , 'prop' , ), 61, (61, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'ExecuteOld' , 'FindText' , 'MatchCase' , 'MatchWholeWord' , 'MatchWildcards' , 
			 'MatchSoundsLike' , 'MatchAllWordForms' , 'Forward' , 'Wrap' , 'Format' , 
			 'ReplaceWith' , 'Replace' , 'prop' , ), 30, (30, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 11 , 408 , (3, 0, None, None) , 64 , )),
	(( 'ClearFormatting' , ), 31, (31, (), [ ], 1 , 1 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'SetAllFuzzyOptions' , ), 32, (32, (), [ ], 1 , 1 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'ClearAllFuzzyOptions' , ), 33, (33, (), [ ], 1 , 1 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'Execute' , 'FindText' , 'MatchCase' , 'MatchWholeWord' , 'MatchWildcards' , 
			 'MatchSoundsLike' , 'MatchAllWordForms' , 'Forward' , 'Wrap' , 'Format' , 
			 'ReplaceWith' , 'Replace' , 'MatchKashida' , 'MatchDiacritics' , 'MatchAlefHamza' , 
			 'MatchControl' , 'prop' , ), 444, (444, (), [ (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 15 , 440 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 34, (34, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 34, (34, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'MatchKashida' , 'prop' , ), 100, (100, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'MatchKashida' , 'prop' , ), 100, (100, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'MatchDiacritics' , 'prop' , ), 101, (101, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'MatchDiacritics' , 'prop' , ), 101, (101, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'MatchAlefHamza' , 'prop' , ), 102, (102, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'MatchAlefHamza' , 'prop' , ), 102, (102, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'MatchControl' , 'prop' , ), 103, (103, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'MatchControl' , 'prop' , ), 103, (103, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'MatchPhrase' , 'prop' , ), 104, (104, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'MatchPhrase' , 'prop' , ), 104, (104, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'MatchPrefix' , 'prop' , ), 105, (105, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'MatchPrefix' , 'prop' , ), 105, (105, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'MatchSuffix' , 'prop' , ), 106, (106, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'MatchSuffix' , 'prop' , ), 106, (106, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'IgnoreSpace' , 'prop' , ), 107, (107, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'IgnoreSpace' , 'prop' , ), 107, (107, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'IgnorePunct' , 'prop' , ), 108, (108, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'IgnorePunct' , 'prop' , ), 108, (108, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'HitHighlight' , 'FindText' , 'HighlightColor' , 'TextColor' , 'MatchCase' , 
			 'MatchWholeWord' , 'MatchPrefix' , 'MatchSuffix' , 'MatchPhrase' , 'MatchWildcards' , 
			 'MatchSoundsLike' , 'MatchAllWordForms' , 'MatchByte' , 'MatchFuzzy' , 'MatchKashida' , 
			 'MatchDiacritics' , 'MatchAlefHamza' , 'MatchControl' , 'IgnoreSpace' , 'IgnorePunct' , 
			 'HanjaPhoneticHangul' , 'prop' , ), 445, (445, (), [ (16396, 1, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 19 , 608 , (3, 0, None, None) , 0 , )),
	(( 'ClearHitHighlight' , 'prop' , ), 446, (446, (), [ (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'Execute2007' , 'FindText' , 'MatchCase' , 'MatchWholeWord' , 'MatchWildcards' , 
			 'MatchSoundsLike' , 'MatchAllWordForms' , 'Forward' , 'Wrap' , 'Format' , 
			 'ReplaceWith' , 'Replace' , 'MatchKashida' , 'MatchDiacritics' , 'MatchAlefHamza' , 
			 'MatchControl' , 'MatchPrefix' , 'MatchSuffix' , 'MatchPhrase' , 'IgnoreSpace' , 
			 'IgnorePunct' , 'prop' , ), 447, (447, (), [ (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16395, 10, None, None) , ], 1 , 1 , 4 , 20 , 624 , (3, 0, None, None) , 0 , )),
	(( 'HanjaPhoneticHangul' , 'prop' , ), 109, (109, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 632 , (3, 0, None, None) , 0 , )),
	(( 'HanjaPhoneticHangul' , 'prop' , ), 109, (109, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{000209B0-0000-0000-C000-000000000046}", Find )
