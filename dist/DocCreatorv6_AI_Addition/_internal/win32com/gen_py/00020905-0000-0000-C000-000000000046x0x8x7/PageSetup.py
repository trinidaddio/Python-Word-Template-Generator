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
class PageSetup(DispatchBaseClass):
	CLSID = IID('{00020971-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def SetAsTemplateDefault(self):
		return self._oleobj_.InvokeTypes(202, LCID, 1, (24, 0), (),)

	def TogglePortrait(self):
		return self._oleobj_.InvokeTypes(201, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"BookFoldPrinting": (1223, 2, (11, 0), (), "BookFoldPrinting", None),
		"BookFoldPrintingSheets": (1225, 2, (3, 0), (), "BookFoldPrintingSheets", None),
		"BookFoldRevPrinting": (1224, 2, (11, 0), (), "BookFoldRevPrinting", None),
		"BottomMargin": (101, 2, (4, 0), (), "BottomMargin", None),
		"CharsLine": (123, 2, (4, 0), (), "CharsLine", None),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		"DifferentFirstPageHeaderFooter": (116, 2, (3, 0), (), "DifferentFirstPageHeaderFooter", None),
		"FirstPageTray": (108, 2, (3, 0), (), "FirstPageTray", None),
		"FooterDistance": (113, 2, (4, 0), (), "FooterDistance", None),
		"Gutter": (104, 2, (4, 0), (), "Gutter", None),
		"GutterOnTop": (122, 2, (11, 0), (), "GutterOnTop", None),
		"GutterPos": (1222, 2, (3, 0), (), "GutterPos", None),
		"GutterStyle": (129, 2, (3, 0), (), "GutterStyle", None),
		"HeaderDistance": (112, 2, (4, 0), (), "HeaderDistance", None),
		"LayoutMode": (131, 2, (3, 0), (), "LayoutMode", None),
		"LeftMargin": (102, 2, (4, 0), (), "LeftMargin", None),
		# Method 'LineNumbering' returns object of type 'LineNumbering'
		"LineNumbering": (118, 2, (9, 0), (), "LineNumbering", '{00020972-0000-0000-C000-000000000046}'),
		"LinesPage": (124, 2, (4, 0), (), "LinesPage", None),
		"MirrorMargins": (111, 2, (3, 0), (), "MirrorMargins", None),
		"OddAndEvenPagesHeaderFooter": (115, 2, (3, 0), (), "OddAndEvenPagesHeaderFooter", None),
		"Orientation": (107, 2, (3, 0), (), "Orientation", None),
		"OtherPagesTray": (109, 2, (3, 0), (), "OtherPagesTray", None),
		"PageHeight": (106, 2, (4, 0), (), "PageHeight", None),
		"PageWidth": (105, 2, (4, 0), (), "PageWidth", None),
		"PaperSize": (120, 2, (3, 0), (), "PaperSize", None),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		"RightMargin": (103, 2, (4, 0), (), "RightMargin", None),
		"SectionDirection": (130, 2, (3, 0), (), "SectionDirection", None),
		"SectionStart": (114, 2, (3, 0), (), "SectionStart", None),
		"ShowGrid": (128, 2, (11, 0), (), "ShowGrid", None),
		"SuppressEndnotes": (117, 2, (3, 0), (), "SuppressEndnotes", None),
		# Method 'TextColumns' returns object of type 'TextColumns'
		"TextColumns": (119, 2, (9, 0), (), "TextColumns", '{00020973-0000-0000-C000-000000000046}'),
		"TopMargin": (100, 2, (4, 0), (), "TopMargin", None),
		"TwoPagesOnOne": (121, 2, (11, 0), (), "TwoPagesOnOne", None),
		"VerticalAlignment": (110, 2, (3, 0), (), "VerticalAlignment", None),
	}
	_prop_map_put_ = {
		"BookFoldPrinting": ((1223, LCID, 4, 0),()),
		"BookFoldPrintingSheets": ((1225, LCID, 4, 0),()),
		"BookFoldRevPrinting": ((1224, LCID, 4, 0),()),
		"BottomMargin": ((101, LCID, 4, 0),()),
		"CharsLine": ((123, LCID, 4, 0),()),
		"DifferentFirstPageHeaderFooter": ((116, LCID, 4, 0),()),
		"FirstPageTray": ((108, LCID, 4, 0),()),
		"FooterDistance": ((113, LCID, 4, 0),()),
		"Gutter": ((104, LCID, 4, 0),()),
		"GutterOnTop": ((122, LCID, 4, 0),()),
		"GutterPos": ((1222, LCID, 4, 0),()),
		"GutterStyle": ((129, LCID, 4, 0),()),
		"HeaderDistance": ((112, LCID, 4, 0),()),
		"LayoutMode": ((131, LCID, 4, 0),()),
		"LeftMargin": ((102, LCID, 4, 0),()),
		"LineNumbering": ((118, LCID, 4, 0),()),
		"LinesPage": ((124, LCID, 4, 0),()),
		"MirrorMargins": ((111, LCID, 4, 0),()),
		"OddAndEvenPagesHeaderFooter": ((115, LCID, 4, 0),()),
		"Orientation": ((107, LCID, 4, 0),()),
		"OtherPagesTray": ((109, LCID, 4, 0),()),
		"PageHeight": ((106, LCID, 4, 0),()),
		"PageWidth": ((105, LCID, 4, 0),()),
		"PaperSize": ((120, LCID, 4, 0),()),
		"RightMargin": ((103, LCID, 4, 0),()),
		"SectionDirection": ((130, LCID, 4, 0),()),
		"SectionStart": ((114, LCID, 4, 0),()),
		"ShowGrid": ((128, LCID, 4, 0),()),
		"SuppressEndnotes": ((117, LCID, 4, 0),()),
		"TextColumns": ((119, LCID, 4, 0),()),
		"TopMargin": ((100, LCID, 4, 0),()),
		"TwoPagesOnOne": ((121, LCID, 4, 0),()),
		"VerticalAlignment": ((110, LCID, 4, 0),()),
	}
	def __iter__(self):
		"Return a Python iterator for this object"
		try:
			ob = self._oleobj_.InvokeTypes(-4,LCID,3,(13, 10),())
		except pythoncom.error:
			raise TypeError("This object does not support enumeration")
		return win32com.client.util.Iterator(ob, None)

win32com.client.CLSIDToClass.RegisterCLSID( "{00020971-0000-0000-C000-000000000046}", PageSetup )
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

PageSetup_vtables_dispatch_ = 1
PageSetup_vtables_ = [
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'TopMargin' , 'prop' , ), 100, (100, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'TopMargin' , 'prop' , ), 100, (100, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'BottomMargin' , 'prop' , ), 101, (101, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'BottomMargin' , 'prop' , ), 101, (101, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'LeftMargin' , 'prop' , ), 102, (102, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'LeftMargin' , 'prop' , ), 102, (102, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'RightMargin' , 'prop' , ), 103, (103, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'RightMargin' , 'prop' , ), 103, (103, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'Gutter' , 'prop' , ), 104, (104, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Gutter' , 'prop' , ), 104, (104, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'PageWidth' , 'prop' , ), 105, (105, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'PageWidth' , 'prop' , ), 105, (105, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'PageHeight' , 'prop' , ), 106, (106, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'PageHeight' , 'prop' , ), 106, (106, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 107, (107, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 107, (107, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'FirstPageTray' , 'prop' , ), 108, (108, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'FirstPageTray' , 'prop' , ), 108, (108, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'OtherPagesTray' , 'prop' , ), 109, (109, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'OtherPagesTray' , 'prop' , ), 109, (109, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'VerticalAlignment' , 'prop' , ), 110, (110, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'VerticalAlignment' , 'prop' , ), 110, (110, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'MirrorMargins' , 'prop' , ), 111, (111, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'MirrorMargins' , 'prop' , ), 111, (111, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'HeaderDistance' , 'prop' , ), 112, (112, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'HeaderDistance' , 'prop' , ), 112, (112, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'FooterDistance' , 'prop' , ), 113, (113, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'FooterDistance' , 'prop' , ), 113, (113, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'SectionStart' , 'prop' , ), 114, (114, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'SectionStart' , 'prop' , ), 114, (114, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'OddAndEvenPagesHeaderFooter' , 'prop' , ), 115, (115, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'OddAndEvenPagesHeaderFooter' , 'prop' , ), 115, (115, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'DifferentFirstPageHeaderFooter' , 'prop' , ), 116, (116, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'DifferentFirstPageHeaderFooter' , 'prop' , ), 116, (116, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'SuppressEndnotes' , 'prop' , ), 117, (117, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'SuppressEndnotes' , 'prop' , ), 117, (117, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'LineNumbering' , 'prop' , ), 118, (118, (), [ (16393, 10, None, "IID('{00020972-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'LineNumbering' , 'prop' , ), 118, (118, (), [ (9, 1, None, "IID('{00020972-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'TextColumns' , 'prop' , ), 119, (119, (), [ (16393, 10, None, "IID('{00020973-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'TextColumns' , 'prop' , ), 119, (119, (), [ (9, 1, None, "IID('{00020973-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'PaperSize' , 'prop' , ), 120, (120, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'PaperSize' , 'prop' , ), 120, (120, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'TwoPagesOnOne' , 'prop' , ), 121, (121, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'TwoPagesOnOne' , 'prop' , ), 121, (121, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'GutterOnTop' , 'prop' , ), 122, (122, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 432 , (3, 0, None, None) , 64 , )),
	(( 'GutterOnTop' , 'prop' , ), 122, (122, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 440 , (3, 0, None, None) , 64 , )),
	(( 'CharsLine' , 'prop' , ), 123, (123, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'CharsLine' , 'prop' , ), 123, (123, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'LinesPage' , 'prop' , ), 124, (124, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'LinesPage' , 'prop' , ), 124, (124, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'ShowGrid' , 'prop' , ), 128, (128, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'ShowGrid' , 'prop' , ), 128, (128, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'TogglePortrait' , ), 201, (201, (), [ ], 1 , 1 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'SetAsTemplateDefault' , ), 202, (202, (), [ ], 1 , 1 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'GutterStyle' , 'prop' , ), 129, (129, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'GutterStyle' , 'prop' , ), 129, (129, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'SectionDirection' , 'prop' , ), 130, (130, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'SectionDirection' , 'prop' , ), 130, (130, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'LayoutMode' , 'prop' , ), 131, (131, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'LayoutMode' , 'prop' , ), 131, (131, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'GutterPos' , 'prop' , ), 1222, (1222, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'GutterPos' , 'prop' , ), 1222, (1222, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldPrinting' , 'prop' , ), 1223, (1223, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldPrinting' , 'prop' , ), 1223, (1223, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldRevPrinting' , 'prop' , ), 1224, (1224, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldRevPrinting' , 'prop' , ), 1224, (1224, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldPrintingSheets' , 'prop' , ), 1225, (1225, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'BookFoldPrintingSheets' , 'prop' , ), 1225, (1225, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020971-0000-0000-C000-000000000046}", PageSetup )
