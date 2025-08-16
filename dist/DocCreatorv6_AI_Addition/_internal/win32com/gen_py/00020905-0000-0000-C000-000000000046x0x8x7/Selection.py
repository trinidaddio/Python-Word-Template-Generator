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
class Selection(DispatchBaseClass):
	CLSID = IID('{00020975-0000-0000-C000-000000000046}')
	coclass_clsid = None

	def BoldRun(self):
		return self._oleobj_.InvokeTypes(602, LCID, 1, (24, 0), (),)

	def Calculate(self):
		return self._oleobj_.InvokeTypes(172, LCID, 1, (4, 0), (),)

	def ClearCharacterAllFormatting(self):
		return self._oleobj_.InvokeTypes(1031, LCID, 1, (24, 0), (),)

	def ClearCharacterDirectFormatting(self):
		return self._oleobj_.InvokeTypes(1033, LCID, 1, (24, 0), (),)

	def ClearCharacterStyle(self):
		return self._oleobj_.InvokeTypes(1032, LCID, 1, (24, 0), (),)

	def ClearFormatting(self):
		return self._oleobj_.InvokeTypes(1009, LCID, 1, (24, 0), (),)

	def ClearParagraphAllFormatting(self):
		return self._oleobj_.InvokeTypes(1039, LCID, 1, (24, 0), (),)

	def ClearParagraphDirectFormatting(self):
		return self._oleobj_.InvokeTypes(1040, LCID, 1, (24, 0), (),)

	def ClearParagraphStyle(self):
		return self._oleobj_.InvokeTypes(1030, LCID, 1, (24, 0), (),)

	def Collapse(self, Direction=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(101, LCID, 1, (24, 0), ((16396, 17),),Direction
			)

	# Result is of type Table
	def ConvertToTable(self, Separator=defaultNamedOptArg, NumRows=defaultNamedOptArg, NumColumns=defaultNamedOptArg, InitialColumnWidth=defaultNamedOptArg
			, Format=defaultNamedOptArg, ApplyBorders=defaultNamedOptArg, ApplyShading=defaultNamedOptArg, ApplyFont=defaultNamedOptArg, ApplyColor=defaultNamedOptArg
			, ApplyHeadingRows=defaultNamedOptArg, ApplyLastRow=defaultNamedOptArg, ApplyFirstColumn=defaultNamedOptArg, ApplyLastColumn=defaultNamedOptArg, AutoFit=defaultNamedOptArg
			, AutoFitBehavior=defaultNamedOptArg, DefaultTableBehavior=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(457, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Separator
			, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders
			, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow
			, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior
			)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTable', '{00020951-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Table
	def ConvertToTableOld(self, Separator=defaultNamedOptArg, NumRows=defaultNamedOptArg, NumColumns=defaultNamedOptArg, InitialColumnWidth=defaultNamedOptArg
			, Format=defaultNamedOptArg, ApplyBorders=defaultNamedOptArg, ApplyShading=defaultNamedOptArg, ApplyFont=defaultNamedOptArg, ApplyColor=defaultNamedOptArg
			, ApplyHeadingRows=defaultNamedOptArg, ApplyLastRow=defaultNamedOptArg, ApplyFirstColumn=defaultNamedOptArg, ApplyLastColumn=defaultNamedOptArg, AutoFit=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(162, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Separator
			, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders
			, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow
			, ApplyFirstColumn, ApplyLastColumn, AutoFit)
		if ret is not None:
			ret = Dispatch(ret, 'ConvertToTableOld', '{00020951-0000-0000-C000-000000000046}')
		return ret

	def Copy(self):
		return self._oleobj_.InvokeTypes(120, LCID, 1, (24, 0), (),)

	def CopyAsPicture(self):
		return self._oleobj_.InvokeTypes(167, LCID, 1, (24, 0), (),)

	def CopyFormat(self):
		return self._oleobj_.InvokeTypes(509, LCID, 1, (24, 0), (),)

	# Result is of type AutoTextEntry
	def CreateAutoTextEntry(self, Name=defaultNamedNotOptArg, StyleName=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(534, LCID, 1, (9, 0), ((8, 1), (8, 1)),Name
			, StyleName)
		if ret is not None:
			ret = Dispatch(ret, 'CreateAutoTextEntry', '{00020936-0000-0000-C000-000000000046}')
		return ret

	def CreateTextbox(self):
		return self._oleobj_.InvokeTypes(523, LCID, 1, (24, 0), (),)

	def Cut(self):
		return self._oleobj_.InvokeTypes(119, LCID, 1, (24, 0), (),)

	def Delete(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(127, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def DetectLanguage(self):
		return self._oleobj_.InvokeTypes(535, LCID, 1, (24, 0), (),)

	def EndKey(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(505, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def EndOf(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(108, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def EscapeKey(self):
		return self._oleobj_.InvokeTypes(506, LCID, 1, (24, 0), (),)

	def Expand(self, Unit=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(129, LCID, 1, (3, 0), ((16396, 17),),Unit
			)

	def ExportAsFixedFormat(self, OutputFileName=defaultNamedNotOptArg, ExportFormat=defaultNamedNotOptArg, OpenAfterExport=False, OptimizeFor=0
			, ExportCurrentPage=False, Item=0, IncludeDocProps=False, KeepIRM=True, CreateBookmarks=0
			, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1036, LCID, 1, (24, 0), ((8, 1), (3, 1), (11, 49), (3, 49), (11, 49), (3, 49), (11, 49), (11, 49), (3, 49), (11, 49), (11, 49), (11, 49), (16396, 17)),OutputFileName
			, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item
			, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts
			, UseISO19005_1, FixedFormatExtClassPtr)

	def ExportAsFixedFormat2(self, OutputFileName=defaultNamedNotOptArg, ExportFormat=defaultNamedNotOptArg, OpenAfterExport=False, OptimizeFor=0
			, ExportCurrentPage=False, Item=0, IncludeDocProps=False, KeepIRM=True, CreateBookmarks=0
			, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False, OptimizeForImageQuality=False, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1043, LCID, 1, (24, 0), ((8, 1), (3, 1), (11, 49), (3, 49), (11, 49), (3, 49), (11, 49), (11, 49), (3, 49), (11, 49), (11, 49), (11, 49), (11, 49), (16396, 17)),OutputFileName
			, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item
			, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts
			, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr)

	def ExportAsFixedFormat3(self, OutputFileName=defaultNamedNotOptArg, ExportFormat=defaultNamedNotOptArg, OpenAfterExport=False, OptimizeFor=0
			, ExportCurrentPage=False, Item=0, IncludeDocProps=False, KeepIRM=True, CreateBookmarks=0
			, DocStructureTags=True, BitmapMissingFonts=True, UseISO19005_1=False, OptimizeForImageQuality=False, ImproveExportTagging=False
			, FixedFormatExtClassPtr=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1044, LCID, 1, (24, 0), ((8, 1), (3, 1), (11, 49), (3, 49), (11, 49), (3, 49), (11, 49), (11, 49), (3, 49), (11, 49), (11, 49), (11, 49), (11, 49), (11, 49), (16396, 17)),OutputFileName
			, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item
			, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts
			, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr)

	def Extend(self, Character=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(300, LCID, 1, (24, 0), ((16396, 17),),Character
			)

	# The method GetXML is actually a property, but must be used as a method to correctly pass the arguments
	def GetXML(self, DataOnly=False):
		# Result is a Unicode object
		return self._oleobj_.InvokeTypes(314, LCID, 2, (8, 0), ((11, 49),),DataOnly
			)

	# Result is of type Range
	def GoTo(self, What=defaultNamedOptArg, Which=defaultNamedOptArg, Count=defaultNamedOptArg, Name=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(173, LCID, 1, (9, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17)),What
			, Which, Count, Name)
		if ret is not None:
			ret = Dispatch(ret, 'GoTo', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToEditableRange(self, EditorID=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(1027, LCID, 1, (9, 0), ((16396, 17),),EditorID
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToEditableRange', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToNext(self, What=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(174, LCID, 1, (9, 0), ((3, 1),),What
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToNext', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Range
	def GoToPrevious(self, What=defaultNamedNotOptArg):
		ret = self._oleobj_.InvokeTypes(175, LCID, 1, (9, 0), ((3, 1),),What
			)
		if ret is not None:
			ret = Dispatch(ret, 'GoToPrevious', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	def HomeKey(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(504, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def InRange(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(126, LCID, 1, (11, 0), ((9, 1),),Range
			)

	def InStory(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(125, LCID, 1, (11, 0), ((9, 1),),Range
			)

	# The method Information is actually a property, but must be used as a method to correctly pass the arguments
	def Information(self, Type=defaultNamedNotOptArg):
		return self._ApplyTypes_(401, 2, (12, 0), ((3, 1),), 'Information', None,Type
			)

	def InsertAfter(self, Text=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(104, LCID, 1, (24, 0), ((8, 1),),Text
			)

	def InsertBefore(self, Text=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(102, LCID, 1, (24, 0), ((8, 1),),Text
			)

	def InsertBreak(self, Type=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(122, LCID, 1, (24, 0), ((16396, 17),),Type
			)

	def InsertCaption(self, Label=defaultNamedNotOptArg, Title=defaultNamedOptArg, TitleAutoText=defaultNamedOptArg, Position=defaultNamedOptArg
			, ExcludeLabel=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(417, LCID, 1, (24, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),Label
			, Title, TitleAutoText, Position, ExcludeLabel)

	def InsertCaptionXP(self, Label=defaultNamedNotOptArg, Title=defaultNamedOptArg, TitleAutoText=defaultNamedOptArg, Position=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(166, LCID, 1, (24, 0), ((16396, 1), (16396, 17), (16396, 17), (16396, 17)),Label
			, Title, TitleAutoText, Position)

	def InsertCells(self, ShiftCells=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(214, LCID, 1, (24, 0), ((16396, 17),),ShiftCells
			)

	def InsertColumns(self):
		return self._oleobj_.InvokeTypes(529, LCID, 1, (24, 0), (),)

	def InsertColumnsRight(self):
		return self._oleobj_.InvokeTypes(538, LCID, 1, (24, 0), (),)

	def InsertCrossReference(self, ReferenceType=defaultNamedNotOptArg, ReferenceKind=defaultNamedNotOptArg, ReferenceItem=defaultNamedNotOptArg, InsertAsHyperlink=defaultNamedOptArg
			, IncludePosition=defaultNamedOptArg, SeparateNumbers=defaultNamedOptArg, SeparatorString=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(418, LCID, 1, (24, 0), ((16396, 1), (3, 1), (16396, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ReferenceType
			, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers
			, SeparatorString)

	def InsertCrossReference_2002(self, ReferenceType=defaultNamedNotOptArg, ReferenceKind=defaultNamedNotOptArg, ReferenceItem=defaultNamedNotOptArg, InsertAsHyperlink=defaultNamedOptArg
			, IncludePosition=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(165, LCID, 1, (24, 0), ((16396, 1), (3, 1), (16396, 1), (16396, 17), (16396, 17)),ReferenceType
			, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition)

	def InsertDateTime(self, DateTimeFormat=defaultNamedOptArg, InsertAsField=defaultNamedOptArg, InsertAsFullWidth=defaultNamedOptArg, DateLanguage=defaultNamedOptArg
			, CalendarType=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(444, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),DateTimeFormat
			, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType)

	def InsertDateTimeOld(self, DateTimeFormat=defaultNamedOptArg, InsertAsField=defaultNamedOptArg, InsertAsFullWidth=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(163, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17)),DateTimeFormat
			, InsertAsField, InsertAsFullWidth)

	def InsertFile(self, FileName=defaultNamedNotOptArg, Range=defaultNamedOptArg, ConfirmConversions=defaultNamedOptArg, Link=defaultNamedOptArg
			, Attachment=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(123, LCID, 1, (24, 0), ((8, 1), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),FileName
			, Range, ConfirmConversions, Link, Attachment)

	def InsertFormula(self, Formula=defaultNamedOptArg, NumberFormat=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(530, LCID, 1, (24, 0), ((16396, 17), (16396, 17)),Formula
			, NumberFormat)

	def InsertNewPage(self):
		return self._oleobj_.InvokeTypes(1041, LCID, 1, (24, 0), (),)

	def InsertParagraph(self):
		return self._oleobj_.InvokeTypes(160, LCID, 1, (24, 0), (),)

	def InsertParagraphAfter(self):
		return self._oleobj_.InvokeTypes(161, LCID, 1, (24, 0), (),)

	def InsertParagraphBefore(self):
		return self._oleobj_.InvokeTypes(212, LCID, 1, (24, 0), (),)

	def InsertRows(self, NumRows=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(528, LCID, 1, (24, 0), ((16396, 17),),NumRows
			)

	def InsertRowsAbove(self, NumRows=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(539, LCID, 1, (24, 0), ((16396, 17),),NumRows
			)

	def InsertRowsBelow(self, NumRows=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(537, LCID, 1, (24, 0), ((16396, 17),),NumRows
			)

	def InsertStyleSeparator(self):
		return self._oleobj_.InvokeTypes(1020, LCID, 1, (24, 0), (),)

	def InsertSymbol(self, CharacterNumber=defaultNamedNotOptArg, Font=defaultNamedOptArg, Unicode=defaultNamedOptArg, Bias=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(164, LCID, 1, (24, 0), ((3, 1), (16396, 17), (16396, 17), (16396, 17)),CharacterNumber
			, Font, Unicode, Bias)

	def InsertXML(self, XML=defaultNamedNotOptArg, Transform=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1028, LCID, 1, (24, 0), ((8, 1), (16396, 17)),XML
			, Transform)

	def IsEqual(self, Range=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(171, LCID, 1, (11, 0), ((9, 1),),Range
			)

	def ItalicRun(self):
		return self._oleobj_.InvokeTypes(603, LCID, 1, (24, 0), (),)

	def LtrPara(self):
		return self._oleobj_.InvokeTypes(606, LCID, 1, (24, 0), (),)

	def LtrRun(self):
		return self._oleobj_.InvokeTypes(601, LCID, 1, (24, 0), (),)

	def Move(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(109, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveDown(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(503, LCID, 1, (3, 0), ((16396, 17), (16396, 17), (16396, 17)),Unit
			, Count, Extend)

	def MoveEnd(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(111, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveEndUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(117, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveEndWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(114, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveLeft(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(500, LCID, 1, (3, 0), ((16396, 17), (16396, 17), (16396, 17)),Unit
			, Count, Extend)

	def MoveRight(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(501, LCID, 1, (3, 0), ((16396, 17), (16396, 17), (16396, 17)),Unit
			, Count, Extend)

	def MoveStart(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(110, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Count)

	def MoveStartUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(116, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveStartWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(113, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveUntil(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(115, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	def MoveUp(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(502, LCID, 1, (3, 0), ((16396, 17), (16396, 17), (16396, 17)),Unit
			, Count, Extend)

	def MoveWhile(self, Cset=defaultNamedNotOptArg, Count=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(112, LCID, 1, (3, 0), ((16396, 1), (16396, 17)),Cset
			, Count)

	# Result is of type Range
	def Next(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(105, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Unit
			, Count)
		if ret is not None:
			ret = Dispatch(ret, 'Next', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Field
	def NextField(self):
		ret = self._oleobj_.InvokeTypes(178, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'NextField', '{0002092F-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Revision
	def NextRevision(self, Wrap=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(531, LCID, 1, (9, 0), ((16396, 17),),Wrap
			)
		if ret is not None:
			ret = Dispatch(ret, 'NextRevision', '{00020981-0000-0000-C000-000000000046}')
		return ret

	def NextSubdocument(self):
		return self._oleobj_.InvokeTypes(514, LCID, 1, (24, 0), (),)

	def Paste(self):
		return self._oleobj_.InvokeTypes(121, LCID, 1, (24, 0), (),)

	def PasteAndFormat(self, Type=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1013, LCID, 1, (24, 0), ((3, 1),),Type
			)

	def PasteAppendTable(self):
		return self._oleobj_.InvokeTypes(1010, LCID, 1, (24, 0), (),)

	def PasteAsNestedTable(self):
		return self._oleobj_.InvokeTypes(533, LCID, 1, (24, 0), (),)

	def PasteExcelTable(self, LinkedToExcel=defaultNamedNotOptArg, WordFormatting=defaultNamedNotOptArg, RTF=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(1014, LCID, 1, (24, 0), ((11, 1), (11, 1), (11, 1)),LinkedToExcel
			, WordFormatting, RTF)

	def PasteFormat(self):
		return self._oleobj_.InvokeTypes(510, LCID, 1, (24, 0), (),)

	def PasteSpecial(self, IconIndex=defaultNamedOptArg, Link=defaultNamedOptArg, Placement=defaultNamedOptArg, DisplayAsIcon=defaultNamedOptArg
			, DataType=defaultNamedOptArg, IconFileName=defaultNamedOptArg, IconLabel=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(176, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),IconIndex
			, Link, Placement, DisplayAsIcon, DataType, IconFileName
			, IconLabel)

	# Result is of type Range
	def Previous(self, Unit=defaultNamedOptArg, Count=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(106, LCID, 1, (9, 0), ((16396, 17), (16396, 17)),Unit
			, Count)
		if ret is not None:
			ret = Dispatch(ret, 'Previous', '{0002095E-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Field
	def PreviousField(self):
		ret = self._oleobj_.InvokeTypes(177, LCID, 1, (9, 0), (),)
		if ret is not None:
			ret = Dispatch(ret, 'PreviousField', '{0002092F-0000-0000-C000-000000000046}')
		return ret

	# Result is of type Revision
	def PreviousRevision(self, Wrap=defaultNamedOptArg):
		ret = self._oleobj_.InvokeTypes(532, LCID, 1, (9, 0), ((16396, 17),),Wrap
			)
		if ret is not None:
			ret = Dispatch(ret, 'PreviousRevision', '{00020981-0000-0000-C000-000000000046}')
		return ret

	def PreviousSubdocument(self):
		return self._oleobj_.InvokeTypes(515, LCID, 1, (24, 0), (),)

	def ReadingModeGrowFont(self):
		return self._oleobj_.InvokeTypes(1037, LCID, 1, (24, 0), (),)

	def ReadingModeShrinkFont(self):
		return self._oleobj_.InvokeTypes(1038, LCID, 1, (24, 0), (),)

	def RtlPara(self):
		return self._oleobj_.InvokeTypes(605, LCID, 1, (24, 0), (),)

	def RtlRun(self):
		return self._oleobj_.InvokeTypes(600, LCID, 1, (24, 0), (),)

	def Select(self):
		return self._oleobj_.InvokeTypes(65535, LCID, 1, (24, 0), (),)

	def SelectCell(self):
		return self._oleobj_.InvokeTypes(536, LCID, 1, (24, 0), (),)

	def SelectColumn(self):
		return self._oleobj_.InvokeTypes(516, LCID, 1, (24, 0), (),)

	def SelectCurrentAlignment(self):
		return self._oleobj_.InvokeTypes(518, LCID, 1, (24, 0), (),)

	def SelectCurrentColor(self):
		return self._oleobj_.InvokeTypes(522, LCID, 1, (24, 0), (),)

	def SelectCurrentFont(self):
		return self._oleobj_.InvokeTypes(517, LCID, 1, (24, 0), (),)

	def SelectCurrentIndent(self):
		return self._oleobj_.InvokeTypes(520, LCID, 1, (24, 0), (),)

	def SelectCurrentSpacing(self):
		return self._oleobj_.InvokeTypes(519, LCID, 1, (24, 0), (),)

	def SelectCurrentTabs(self):
		return self._oleobj_.InvokeTypes(521, LCID, 1, (24, 0), (),)

	def SelectRow(self):
		return self._oleobj_.InvokeTypes(525, LCID, 1, (24, 0), (),)

	def SetRange(self, Start=defaultNamedNotOptArg, End=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(100, LCID, 1, (24, 0), ((3, 1), (3, 1)),Start
			, End)

	def Shrink(self):
		return self._oleobj_.InvokeTypes(301, LCID, 1, (24, 0), (),)

	def ShrinkDiscontiguousSelection(self):
		return self._oleobj_.InvokeTypes(1019, LCID, 1, (24, 0), (),)

	def Sort(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, SortColumn=defaultNamedOptArg, Separator=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, BidiSort=defaultNamedOptArg
			, IgnoreThe=defaultNamedOptArg, IgnoreKashida=defaultNamedOptArg, IgnoreDiacritics=defaultNamedOptArg, IgnoreHe=defaultNamedOptArg, LanguageID=defaultNamedOptArg
			, SubFieldNumber=defaultNamedOptArg, SubFieldNumber2=defaultNamedOptArg, SubFieldNumber3=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1023, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn
			, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida
			, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2
			, SubFieldNumber3)

	def Sort2000(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, SortColumn=defaultNamedOptArg, Separator=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, BidiSort=defaultNamedOptArg
			, IgnoreThe=defaultNamedOptArg, IgnoreKashida=defaultNamedOptArg, IgnoreDiacritics=defaultNamedOptArg, IgnoreHe=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(445, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn
			, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida
			, IgnoreDiacritics, IgnoreHe, LanguageID)

	def SortAscending(self):
		return self._oleobj_.InvokeTypes(169, LCID, 1, (24, 0), (),)

	def SortByHeadings(self, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, BidiSort=defaultNamedOptArg
			, IgnoreThe=defaultNamedOptArg, IgnoreKashida=defaultNamedOptArg, IgnoreDiacritics=defaultNamedOptArg, IgnoreHe=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(1042, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),SortFieldType
			, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida
			, IgnoreDiacritics, IgnoreHe, LanguageID)

	def SortDescending(self):
		return self._oleobj_.InvokeTypes(170, LCID, 1, (24, 0), (),)

	def SortOld(self, ExcludeHeader=defaultNamedOptArg, FieldNumber=defaultNamedOptArg, SortFieldType=defaultNamedOptArg, SortOrder=defaultNamedOptArg
			, FieldNumber2=defaultNamedOptArg, SortFieldType2=defaultNamedOptArg, SortOrder2=defaultNamedOptArg, FieldNumber3=defaultNamedOptArg, SortFieldType3=defaultNamedOptArg
			, SortOrder3=defaultNamedOptArg, SortColumn=defaultNamedOptArg, Separator=defaultNamedOptArg, CaseSensitive=defaultNamedOptArg, LanguageID=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(168, LCID, 1, (24, 0), ((16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17), (16396, 17)),ExcludeHeader
			, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2
			, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn
			, Separator, CaseSensitive, LanguageID)

	def SplitTable(self):
		return self._oleobj_.InvokeTypes(526, LCID, 1, (24, 0), (),)

	def StartOf(self, Unit=defaultNamedOptArg, Extend=defaultNamedOptArg):
		return self._oleobj_.InvokeTypes(107, LCID, 1, (3, 0), ((16396, 17), (16396, 17)),Unit
			, Extend)

	def ToggleCharacterCode(self):
		return self._oleobj_.InvokeTypes(1012, LCID, 1, (24, 0), (),)

	def TypeBackspace(self):
		return self._oleobj_.InvokeTypes(513, LCID, 1, (24, 0), (),)

	def TypeParagraph(self):
		return self._oleobj_.InvokeTypes(512, LCID, 1, (24, 0), (),)

	def TypeText(self, Text=defaultNamedNotOptArg):
		return self._oleobj_.InvokeTypes(507, LCID, 1, (24, 0), ((8, 1),),Text
			)

	def WholeStory(self):
		return self._oleobj_.InvokeTypes(524, LCID, 1, (24, 0), (),)

	_prop_map_get_ = {
		"Active": (403, 2, (11, 0), (), "Active", None),
		# Method 'Application' returns object of type 'Application'
		"Application": (1000, 2, (13, 0), (), "Application", '{000209FF-0000-0000-C000-000000000046}'),
		"BookmarkID": (308, 2, (3, 0), (), "BookmarkID", None),
		# Method 'Bookmarks' returns object of type 'Bookmarks'
		"Bookmarks": (75, 2, (9, 0), (), "Bookmarks", '{00020967-0000-0000-C000-000000000046}'),
		# Method 'Borders' returns object of type 'Borders'
		"Borders": (1100, 2, (9, 0), (), "Borders", '{0002093C-0000-0000-C000-000000000046}'),
		# Method 'Cells' returns object of type 'Cells'
		"Cells": (57, 2, (9, 0), (), "Cells", '{0002094A-0000-0000-C000-000000000046}'),
		# Method 'Characters' returns object of type 'Characters'
		"Characters": (53, 2, (9, 0), (), "Characters", '{0002095D-0000-0000-C000-000000000046}'),
		# Method 'ChildShapeRange' returns object of type 'ShapeRange'
		"ChildShapeRange": (1021, 2, (9, 0), (), "ChildShapeRange", '{000209B5-0000-0000-C000-000000000046}'),
		"ColumnSelectMode": (407, 2, (11, 0), (), "ColumnSelectMode", None),
		# Method 'Columns' returns object of type 'Columns'
		"Columns": (302, 2, (9, 0), (), "Columns", '{0002094B-0000-0000-C000-000000000046}'),
		# Method 'Comments' returns object of type 'Comments'
		"Comments": (56, 2, (9, 0), (), "Comments", '{00020940-0000-0000-C000-000000000046}'),
		# Method 'ContentControls' returns object of type 'ContentControls'
		"ContentControls": (1034, 2, (9, 0), (), "ContentControls", '{804CD967-F83B-432D-9446-C61A45CFEFF0}'),
		"Creator": (1001, 2, (3, 0), (), "Creator", None),
		# Method 'Document' returns object of type 'Document'
		"Document": (1003, 2, (13, 0), (), "Document", '{00020906-0000-0000-C000-000000000046}'),
		# Method 'Editors' returns object of type 'Editors'
		"Editors": (313, 2, (9, 0), (), "Editors", '{AED7E08C-14F0-4F33-921D-4C5353137BF6}'),
		"End": (4, 2, (3, 0), (), "End", None),
		# Method 'EndnoteOptions' returns object of type 'EndnoteOptions'
		"EndnoteOptions": (1025, 2, (9, 0), (), "EndnoteOptions", '{BF043168-F4DE-4E7C-B206-741A8B3EF71A}'),
		# Method 'Endnotes' returns object of type 'Endnotes'
		"Endnotes": (55, 2, (9, 0), (), "Endnotes", '{00020941-0000-0000-C000-000000000046}'),
		"EnhMetaFileBits": (315, 2, (12, 0), (), "EnhMetaFileBits", None),
		"ExtendMode": (406, 2, (11, 0), (), "ExtendMode", None),
		# Method 'Fields' returns object of type 'Fields'
		"Fields": (64, 2, (9, 0), (), "Fields", '{00020930-0000-0000-C000-000000000046}'),
		# Method 'Find' returns object of type 'Find'
		"Find": (262, 2, (9, 0), (), "Find", '{000209B0-0000-0000-C000-000000000046}'),
		"FitTextWidth": (1008, 2, (4, 0), (), "FitTextWidth", None),
		"Flags": (402, 2, (3, 0), (), "Flags", None),
		# Method 'Font' returns object of type 'Font'
		"Font": (5, 2, (13, 0), (), "Font", '{000209F5-0000-0000-C000-000000000046}'),
		# Method 'FootnoteOptions' returns object of type 'FootnoteOptions'
		"FootnoteOptions": (1024, 2, (9, 0), (), "FootnoteOptions", '{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}'),
		# Method 'Footnotes' returns object of type 'Footnotes'
		"Footnotes": (54, 2, (9, 0), (), "Footnotes", '{00020942-0000-0000-C000-000000000046}'),
		# Method 'FormFields' returns object of type 'FormFields'
		"FormFields": (65, 2, (9, 0), (), "FormFields", '{00020929-0000-0000-C000-000000000046}'),
		# Method 'FormattedText' returns object of type 'Range'
		"FormattedText": (2, 2, (9, 0), (), "FormattedText", '{0002095E-0000-0000-C000-000000000046}'),
		# Method 'Frames' returns object of type 'Frames'
		"Frames": (66, 2, (9, 0), (), "Frames", '{0002092B-0000-0000-C000-000000000046}'),
		# Method 'HTMLDivisions' returns object of type 'HTMLDivisions'
		"HTMLDivisions": (1011, 2, (9, 0), (), "HTMLDivisions", '{000209E8-0000-0000-C000-000000000046}'),
		"HasChildShapeRange": (1022, 2, (11, 0), (), "HasChildShapeRange", None),
		# Method 'HeaderFooter' returns object of type 'HeaderFooter'
		"HeaderFooter": (306, 2, (9, 0), (), "HeaderFooter", '{00020985-0000-0000-C000-000000000046}'),
		# Method 'Hyperlinks' returns object of type 'Hyperlinks'
		"Hyperlinks": (156, 2, (9, 0), (), "Hyperlinks", '{0002099C-0000-0000-C000-000000000046}'),
		"IPAtEndOfLine": (405, 2, (11, 0), (), "IPAtEndOfLine", None),
		# Method 'InlineShapes' returns object of type 'InlineShapes'
		"InlineShapes": (411, 2, (9, 0), (), "InlineShapes", '{000209A9-0000-0000-C000-000000000046}'),
		"IsEndOfRowMark": (307, 2, (11, 0), (), "IsEndOfRowMark", None),
		"LanguageDetected": (1007, 2, (11, 0), (), "LanguageDetected", None),
		"LanguageID": (153, 2, (3, 0), (), "LanguageID", None),
		"LanguageIDFarEast": (154, 2, (3, 0), (), "LanguageIDFarEast", None),
		"LanguageIDOther": (155, 2, (3, 0), (), "LanguageIDOther", None),
		"NoProofing": (1005, 2, (3, 0), (), "NoProofing", None),
		# Method 'OMaths' returns object of type 'OMaths'
		"OMaths": (316, 2, (9, 0), (), "OMaths", '{873E774B-926A-4CB1-878D-635A45187595}'),
		"Orientation": (410, 2, (3, 0), (), "Orientation", None),
		# Method 'PageSetup' returns object of type 'PageSetup'
		"PageSetup": (1101, 2, (9, 0), (), "PageSetup", '{00020971-0000-0000-C000-000000000046}'),
		# Method 'ParagraphFormat' returns object of type 'ParagraphFormat'
		"ParagraphFormat": (1102, 2, (13, 0), (), "ParagraphFormat", '{000209F4-0000-0000-C000-000000000046}'),
		# Method 'Paragraphs' returns object of type 'Paragraphs'
		"Paragraphs": (59, 2, (9, 0), (), "Paragraphs", '{00020958-0000-0000-C000-000000000046}'),
		"Parent": (1002, 2, (9, 0), (), "Parent", None),
		# Method 'ParentContentControl' returns object of type 'ContentControl'
		"ParentContentControl": (1035, 2, (9, 0), (), "ParentContentControl", '{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}'),
		"PreviousBookmarkID": (309, 2, (3, 0), (), "PreviousBookmarkID", None),
		# Method 'Range' returns object of type 'Range'
		"Range": (400, 2, (9, 0), (), "Range", '{0002095E-0000-0000-C000-000000000046}'),
		# Method 'Rows' returns object of type 'Rows'
		"Rows": (303, 2, (9, 0), (), "Rows", '{0002094C-0000-0000-C000-000000000046}'),
		# Method 'Sections' returns object of type 'Sections'
		"Sections": (58, 2, (9, 0), (), "Sections", '{0002095A-0000-0000-C000-000000000046}'),
		# Method 'Sentences' returns object of type 'Sentences'
		"Sentences": (52, 2, (9, 0), (), "Sentences", '{0002095B-0000-0000-C000-000000000046}'),
		# Method 'Shading' returns object of type 'Shading'
		"Shading": (61, 2, (9, 0), (), "Shading", '{0002093A-0000-0000-C000-000000000046}'),
		# Method 'ShapeRange' returns object of type 'ShapeRange'
		"ShapeRange": (1004, 2, (9, 0), (), "ShapeRange", '{000209B5-0000-0000-C000-000000000046}'),
		# Method 'SmartTags' returns object of type 'SmartTags'
		"SmartTags": (1015, 2, (9, 0), (), "SmartTags", '{000209EE-0000-0000-C000-000000000046}'),
		"Start": (3, 2, (3, 0), (), "Start", None),
		"StartIsActive": (404, 2, (11, 0), (), "StartIsActive", None),
		"StoryLength": (152, 2, (3, 0), (), "StoryLength", None),
		"StoryType": (7, 2, (3, 0), (), "StoryType", None),
		"Style": (8, 2, (12, 0), (), "Style", None),
		# Method 'Tables' returns object of type 'Tables'
		"Tables": (50, 2, (9, 0), (), "Tables", '{0002094D-0000-0000-C000-000000000046}'),
		"Text": (0, 2, (8, 0), (), "Text", None),
		# Method 'TopLevelTables' returns object of type 'Tables'
		"TopLevelTables": (1006, 2, (9, 0), (), "TopLevelTables", '{0002094D-0000-0000-C000-000000000046}'),
		"Type": (6, 2, (3, 0), (), "Type", None),
		"WordOpenXML": (317, 2, (8, 0), (), "WordOpenXML", None),
		# Method 'Words' returns object of type 'Words'
		"Words": (51, 2, (9, 0), (), "Words", '{0002095C-0000-0000-C000-000000000046}'),
		"XML": (314, 2, (8, 0), ((11, 49),), "XML", None),
		# Method 'XMLNodes' returns object of type 'XMLNodes'
		"XMLNodes": (310, 2, (9, 0), (), "XMLNodes", '{D36C1F42-7044-4B9E-9CA3-85919454DB04}'),
		# Method 'XMLParentNode' returns object of type 'XMLNode'
		"XMLParentNode": (311, 2, (9, 0), (), "XMLParentNode", '{09760240-0B89-49F7-A79D-479F24723F56}'),
	}
	_prop_map_put_ = {
		"Borders": ((1100, LCID, 4, 0),()),
		"ColumnSelectMode": ((407, LCID, 4, 0),()),
		"End": ((4, LCID, 4, 0),()),
		"ExtendMode": ((406, LCID, 4, 0),()),
		"FitTextWidth": ((1008, LCID, 4, 0),()),
		"Flags": ((402, LCID, 4, 0),()),
		"Font": ((5, LCID, 4, 0),()),
		"FormattedText": ((2, LCID, 4, 0),()),
		"LanguageDetected": ((1007, LCID, 4, 0),()),
		"LanguageID": ((153, LCID, 4, 0),()),
		"LanguageIDFarEast": ((154, LCID, 4, 0),()),
		"LanguageIDOther": ((155, LCID, 4, 0),()),
		"NoProofing": ((1005, LCID, 4, 0),()),
		"Orientation": ((410, LCID, 4, 0),()),
		"PageSetup": ((1101, LCID, 4, 0),()),
		"ParagraphFormat": ((1102, LCID, 4, 0),()),
		"Start": ((3, LCID, 4, 0),()),
		"StartIsActive": ((404, LCID, 4, 0),()),
		"Style": ((8, LCID, 4, 0),()),
		"Text": ((0, LCID, 4, 0),()),
	}
	# Default property for this class is 'Text'
	def __call__(self):
		return self._ApplyTypes_(*(0, 2, (8, 0), (), "Text", None))
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

win32com.client.CLSIDToClass.RegisterCLSID( "{00020975-0000-0000-C000-000000000046}", Selection )
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

Selection_vtables_dispatch_ = 1
Selection_vtables_ = [
	(( 'Text' , 'prop' , ), 0, (0, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 56 , (3, 0, None, None) , 0 , )),
	(( 'Text' , 'prop' , ), 0, (0, (), [ (8, 1, None, None) , ], 1 , 4 , 4 , 0 , 64 , (3, 0, None, None) , 0 , )),
	(( 'FormattedText' , 'prop' , ), 2, (2, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 72 , (3, 0, None, None) , 0 , )),
	(( 'FormattedText' , 'prop' , ), 2, (2, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 80 , (3, 0, None, None) , 0 , )),
	(( 'Start' , 'prop' , ), 3, (3, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 88 , (3, 0, None, None) , 0 , )),
	(( 'Start' , 'prop' , ), 3, (3, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 96 , (3, 0, None, None) , 0 , )),
	(( 'End' , 'prop' , ), 4, (4, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 104 , (3, 0, None, None) , 0 , )),
	(( 'End' , 'prop' , ), 4, (4, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 112 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 5, (5, (), [ (16397, 10, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 120 , (3, 0, None, None) , 0 , )),
	(( 'Font' , 'prop' , ), 5, (5, (), [ (13, 1, None, "IID('{000209F5-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 128 , (3, 0, None, None) , 0 , )),
	(( 'Type' , 'prop' , ), 6, (6, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 136 , (3, 0, None, None) , 0 , )),
	(( 'StoryType' , 'prop' , ), 7, (7, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 144 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 8, (8, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 152 , (3, 0, None, None) , 0 , )),
	(( 'Style' , 'prop' , ), 8, (8, (), [ (16396, 1, None, None) , ], 1 , 4 , 4 , 0 , 160 , (3, 0, None, None) , 0 , )),
	(( 'Tables' , 'prop' , ), 50, (50, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 168 , (3, 0, None, None) , 0 , )),
	(( 'Words' , 'prop' , ), 51, (51, (), [ (16393, 10, None, "IID('{0002095C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 176 , (3, 0, None, None) , 0 , )),
	(( 'Sentences' , 'prop' , ), 52, (52, (), [ (16393, 10, None, "IID('{0002095B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 184 , (3, 0, None, None) , 0 , )),
	(( 'Characters' , 'prop' , ), 53, (53, (), [ (16393, 10, None, "IID('{0002095D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 192 , (3, 0, None, None) , 0 , )),
	(( 'Footnotes' , 'prop' , ), 54, (54, (), [ (16393, 10, None, "IID('{00020942-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 200 , (3, 0, None, None) , 0 , )),
	(( 'Endnotes' , 'prop' , ), 55, (55, (), [ (16393, 10, None, "IID('{00020941-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 208 , (3, 0, None, None) , 0 , )),
	(( 'Comments' , 'prop' , ), 56, (56, (), [ (16393, 10, None, "IID('{00020940-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 216 , (3, 0, None, None) , 0 , )),
	(( 'Cells' , 'prop' , ), 57, (57, (), [ (16393, 10, None, "IID('{0002094A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 224 , (3, 0, None, None) , 0 , )),
	(( 'Sections' , 'prop' , ), 58, (58, (), [ (16393, 10, None, "IID('{0002095A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 232 , (3, 0, None, None) , 0 , )),
	(( 'Paragraphs' , 'prop' , ), 59, (59, (), [ (16393, 10, None, "IID('{00020958-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 240 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (16393, 10, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 248 , (3, 0, None, None) , 0 , )),
	(( 'Borders' , 'prop' , ), 1100, (1100, (), [ (9, 1, None, "IID('{0002093C-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 256 , (3, 0, None, None) , 0 , )),
	(( 'Shading' , 'prop' , ), 61, (61, (), [ (16393, 10, None, "IID('{0002093A-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 264 , (3, 0, None, None) , 0 , )),
	(( 'Fields' , 'prop' , ), 64, (64, (), [ (16393, 10, None, "IID('{00020930-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 272 , (3, 0, None, None) , 0 , )),
	(( 'FormFields' , 'prop' , ), 65, (65, (), [ (16393, 10, None, "IID('{00020929-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 280 , (3, 0, None, None) , 0 , )),
	(( 'Frames' , 'prop' , ), 66, (66, (), [ (16393, 10, None, "IID('{0002092B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 288 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 1102, (1102, (), [ (16397, 10, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 296 , (3, 0, None, None) , 0 , )),
	(( 'ParagraphFormat' , 'prop' , ), 1102, (1102, (), [ (13, 1, None, "IID('{000209F4-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 304 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (16393, 10, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 312 , (3, 0, None, None) , 0 , )),
	(( 'PageSetup' , 'prop' , ), 1101, (1101, (), [ (9, 1, None, "IID('{00020971-0000-0000-C000-000000000046}')") , ], 1 , 4 , 4 , 0 , 320 , (3, 0, None, None) , 0 , )),
	(( 'Bookmarks' , 'prop' , ), 75, (75, (), [ (16393, 10, None, "IID('{00020967-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 328 , (3, 0, None, None) , 0 , )),
	(( 'StoryLength' , 'prop' , ), 152, (152, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 336 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 153, (153, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 344 , (3, 0, None, None) , 0 , )),
	(( 'LanguageID' , 'prop' , ), 153, (153, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 352 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 154, (154, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 360 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDFarEast' , 'prop' , ), 154, (154, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 368 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 155, (155, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 376 , (3, 0, None, None) , 0 , )),
	(( 'LanguageIDOther' , 'prop' , ), 155, (155, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 384 , (3, 0, None, None) , 0 , )),
	(( 'Hyperlinks' , 'prop' , ), 156, (156, (), [ (16393, 10, None, "IID('{0002099C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 392 , (3, 0, None, None) , 0 , )),
	(( 'Columns' , 'prop' , ), 302, (302, (), [ (16393, 10, None, "IID('{0002094B-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 400 , (3, 0, None, None) , 0 , )),
	(( 'Rows' , 'prop' , ), 303, (303, (), [ (16393, 10, None, "IID('{0002094C-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 408 , (3, 0, None, None) , 0 , )),
	(( 'HeaderFooter' , 'prop' , ), 306, (306, (), [ (16393, 10, None, "IID('{00020985-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 416 , (3, 0, None, None) , 0 , )),
	(( 'IsEndOfRowMark' , 'prop' , ), 307, (307, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 424 , (3, 0, None, None) , 0 , )),
	(( 'BookmarkID' , 'prop' , ), 308, (308, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 432 , (3, 0, None, None) , 0 , )),
	(( 'PreviousBookmarkID' , 'prop' , ), 309, (309, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 440 , (3, 0, None, None) , 0 , )),
	(( 'Find' , 'prop' , ), 262, (262, (), [ (16393, 10, None, "IID('{000209B0-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 448 , (3, 0, None, None) , 0 , )),
	(( 'Range' , 'prop' , ), 400, (400, (), [ (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 456 , (3, 0, None, None) , 0 , )),
	(( 'Information' , 'Type' , 'prop' , ), 401, (401, (), [ (3, 1, None, None) , 
			 (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 464 , (3, 0, None, None) , 0 , )),
	(( 'Flags' , 'prop' , ), 402, (402, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 472 , (3, 0, None, None) , 0 , )),
	(( 'Flags' , 'prop' , ), 402, (402, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 480 , (3, 0, None, None) , 0 , )),
	(( 'Active' , 'prop' , ), 403, (403, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 488 , (3, 0, None, None) , 0 , )),
	(( 'StartIsActive' , 'prop' , ), 404, (404, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 496 , (3, 0, None, None) , 0 , )),
	(( 'StartIsActive' , 'prop' , ), 404, (404, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 504 , (3, 0, None, None) , 0 , )),
	(( 'IPAtEndOfLine' , 'prop' , ), 405, (405, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 512 , (3, 0, None, None) , 0 , )),
	(( 'ExtendMode' , 'prop' , ), 406, (406, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 520 , (3, 0, None, None) , 0 , )),
	(( 'ExtendMode' , 'prop' , ), 406, (406, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 528 , (3, 0, None, None) , 0 , )),
	(( 'ColumnSelectMode' , 'prop' , ), 407, (407, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 536 , (3, 0, None, None) , 0 , )),
	(( 'ColumnSelectMode' , 'prop' , ), 407, (407, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 544 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 410, (410, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 552 , (3, 0, None, None) , 0 , )),
	(( 'Orientation' , 'prop' , ), 410, (410, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 560 , (3, 0, None, None) , 0 , )),
	(( 'InlineShapes' , 'prop' , ), 411, (411, (), [ (16393, 10, None, "IID('{000209A9-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 568 , (3, 0, None, None) , 0 , )),
	(( 'Application' , 'prop' , ), 1000, (1000, (), [ (16397, 10, None, "IID('{000209FF-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 576 , (3, 0, None, None) , 0 , )),
	(( 'Creator' , 'prop' , ), 1001, (1001, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 584 , (3, 0, None, None) , 0 , )),
	(( 'Parent' , 'prop' , ), 1002, (1002, (), [ (16393, 10, None, None) , ], 1 , 2 , 4 , 0 , 592 , (3, 0, None, None) , 0 , )),
	(( 'Document' , 'prop' , ), 1003, (1003, (), [ (16397, 10, None, "IID('{00020906-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 600 , (3, 0, None, None) , 0 , )),
	(( 'ShapeRange' , 'prop' , ), 1004, (1004, (), [ (16393, 10, None, "IID('{000209B5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 608 , (3, 0, None, None) , 0 , )),
	(( 'Select' , ), 65535, (65535, (), [ ], 1 , 1 , 4 , 0 , 616 , (3, 0, None, None) , 0 , )),
	(( 'SetRange' , 'Start' , 'End' , ), 100, (100, (), [ (3, 1, None, None) , 
			 (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 624 , (3, 0, None, None) , 0 , )),
	(( 'Collapse' , 'Direction' , ), 101, (101, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 632 , (3, 0, None, None) , 0 , )),
	(( 'InsertBefore' , 'Text' , ), 102, (102, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 640 , (3, 0, None, None) , 0 , )),
	(( 'InsertAfter' , 'Text' , ), 104, (104, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 648 , (3, 0, None, None) , 0 , )),
	(( 'Next' , 'Unit' , 'Count' , 'prop' , ), 105, (105, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 656 , (3, 0, None, None) , 0 , )),
	(( 'Previous' , 'Unit' , 'Count' , 'prop' , ), 106, (106, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 2 , 664 , (3, 0, None, None) , 0 , )),
	(( 'StartOf' , 'Unit' , 'Extend' , 'prop' , ), 107, (107, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 672 , (3, 0, None, None) , 0 , )),
	(( 'EndOf' , 'Unit' , 'Extend' , 'prop' , ), 108, (108, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 680 , (3, 0, None, None) , 0 , )),
	(( 'Move' , 'Unit' , 'Count' , 'prop' , ), 109, (109, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 688 , (3, 0, None, None) , 0 , )),
	(( 'MoveStart' , 'Unit' , 'Count' , 'prop' , ), 110, (110, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 696 , (3, 0, None, None) , 0 , )),
	(( 'MoveEnd' , 'Unit' , 'Count' , 'prop' , ), 111, (111, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 704 , (3, 0, None, None) , 0 , )),
	(( 'MoveWhile' , 'Cset' , 'Count' , 'prop' , ), 112, (112, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 712 , (3, 0, None, None) , 0 , )),
	(( 'MoveStartWhile' , 'Cset' , 'Count' , 'prop' , ), 113, (113, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 720 , (3, 0, None, None) , 0 , )),
	(( 'MoveEndWhile' , 'Cset' , 'Count' , 'prop' , ), 114, (114, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 728 , (3, 0, None, None) , 0 , )),
	(( 'MoveUntil' , 'Cset' , 'Count' , 'prop' , ), 115, (115, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 736 , (3, 0, None, None) , 0 , )),
	(( 'MoveStartUntil' , 'Cset' , 'Count' , 'prop' , ), 116, (116, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 744 , (3, 0, None, None) , 0 , )),
	(( 'MoveEndUntil' , 'Cset' , 'Count' , 'prop' , ), 117, (117, (), [ 
			 (16396, 1, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 752 , (3, 0, None, None) , 0 , )),
	(( 'Cut' , ), 119, (119, (), [ ], 1 , 1 , 4 , 0 , 760 , (3, 0, None, None) , 0 , )),
	(( 'Copy' , ), 120, (120, (), [ ], 1 , 1 , 4 , 0 , 768 , (3, 0, None, None) , 0 , )),
	(( 'Paste' , ), 121, (121, (), [ ], 1 , 1 , 4 , 0 , 776 , (3, 0, None, None) , 0 , )),
	(( 'InsertBreak' , 'Type' , ), 122, (122, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 784 , (3, 0, None, None) , 0 , )),
	(( 'InsertFile' , 'FileName' , 'Range' , 'ConfirmConversions' , 'Link' , 
			 'Attachment' , ), 123, (123, (), [ (8, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 792 , (3, 0, None, None) , 0 , )),
	(( 'InStory' , 'Range' , 'prop' , ), 125, (125, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 800 , (3, 0, None, None) , 0 , )),
	(( 'InRange' , 'Range' , 'prop' , ), 126, (126, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 808 , (3, 0, None, None) , 0 , )),
	(( 'Delete' , 'Unit' , 'Count' , 'prop' , ), 127, (127, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 816 , (3, 0, None, None) , 0 , )),
	(( 'Expand' , 'Unit' , 'prop' , ), 129, (129, (), [ (16396, 17, None, None) , 
			 (16387, 10, None, None) , ], 1 , 1 , 4 , 1 , 824 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraph' , ), 160, (160, (), [ ], 1 , 1 , 4 , 0 , 832 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraphAfter' , ), 161, (161, (), [ ], 1 , 1 , 4 , 0 , 840 , (3, 0, None, None) , 0 , )),
	(( 'ConvertToTableOld' , 'Separator' , 'NumRows' , 'NumColumns' , 'InitialColumnWidth' , 
			 'Format' , 'ApplyBorders' , 'ApplyShading' , 'ApplyFont' , 'ApplyColor' , 
			 'ApplyHeadingRows' , 'ApplyLastRow' , 'ApplyFirstColumn' , 'ApplyLastColumn' , 'AutoFit' , 
			 'prop' , ), 162, (162, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 14 , 848 , (3, 0, None, None) , 64 , )),
	(( 'InsertDateTimeOld' , 'DateTimeFormat' , 'InsertAsField' , 'InsertAsFullWidth' , ), 163, (163, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 856 , (3, 0, None, None) , 64 , )),
	(( 'InsertSymbol' , 'CharacterNumber' , 'Font' , 'Unicode' , 'Bias' , 
			 ), 164, (164, (), [ (3, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 864 , (3, 0, None, None) , 0 , )),
	(( 'InsertCrossReference_2002' , 'ReferenceType' , 'ReferenceKind' , 'ReferenceItem' , 'InsertAsHyperlink' , 
			 'IncludePosition' , ), 165, (165, (), [ (16396, 1, None, None) , (3, 1, None, None) , (16396, 1, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 872 , (3, 0, None, None) , 64 , )),
	(( 'InsertCaptionXP' , 'Label' , 'Title' , 'TitleAutoText' , 'Position' , 
			 ), 166, (166, (), [ (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 3 , 880 , (3, 0, None, None) , 64 , )),
	(( 'CopyAsPicture' , ), 167, (167, (), [ ], 1 , 1 , 4 , 0 , 888 , (3, 0, None, None) , 0 , )),
	(( 'SortOld' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'SortColumn' , 'Separator' , 'CaseSensitive' , 'LanguageID' , 
			 ), 168, (168, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 14 , 896 , (3, 0, None, None) , 64 , )),
	(( 'SortAscending' , ), 169, (169, (), [ ], 1 , 1 , 4 , 0 , 904 , (3, 0, None, None) , 0 , )),
	(( 'SortDescending' , ), 170, (170, (), [ ], 1 , 1 , 4 , 0 , 912 , (3, 0, None, None) , 0 , )),
	(( 'IsEqual' , 'Range' , 'prop' , ), 171, (171, (), [ (9, 1, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , 
			 (16395, 10, None, None) , ], 1 , 1 , 4 , 0 , 920 , (3, 0, None, None) , 0 , )),
	(( 'Calculate' , 'prop' , ), 172, (172, (), [ (16388, 10, None, None) , ], 1 , 1 , 4 , 0 , 928 , (3, 0, None, None) , 0 , )),
	(( 'GoTo' , 'What' , 'Which' , 'Count' , 'Name' , 
			 'prop' , ), 173, (173, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 4 , 936 , (3, 0, None, None) , 0 , )),
	(( 'GoToNext' , 'What' , 'prop' , ), 174, (174, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 944 , (3, 0, None, None) , 0 , )),
	(( 'GoToPrevious' , 'What' , 'prop' , ), 175, (175, (), [ (3, 1, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 952 , (3, 0, None, None) , 0 , )),
	(( 'PasteSpecial' , 'IconIndex' , 'Link' , 'Placement' , 'DisplayAsIcon' , 
			 'DataType' , 'IconFileName' , 'IconLabel' , ), 176, (176, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 7 , 960 , (3, 0, None, None) , 0 , )),
	(( 'PreviousField' , 'prop' , ), 177, (177, (), [ (16393, 10, None, "IID('{0002092F-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 968 , (3, 0, None, None) , 0 , )),
	(( 'NextField' , 'prop' , ), 178, (178, (), [ (16393, 10, None, "IID('{0002092F-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 976 , (3, 0, None, None) , 0 , )),
	(( 'InsertParagraphBefore' , ), 212, (212, (), [ ], 1 , 1 , 4 , 0 , 984 , (3, 0, None, None) , 0 , )),
	(( 'InsertCells' , 'ShiftCells' , ), 214, (214, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 992 , (3, 0, None, None) , 0 , )),
	(( 'Extend' , 'Character' , ), 300, (300, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1000 , (3, 0, None, None) , 0 , )),
	(( 'Shrink' , ), 301, (301, (), [ ], 1 , 1 , 4 , 0 , 1008 , (3, 0, None, None) , 0 , )),
	(( 'MoveLeft' , 'Unit' , 'Count' , 'Extend' , 'prop' , 
			 ), 500, (500, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 3 , 1016 , (3, 0, None, None) , 0 , )),
	(( 'MoveRight' , 'Unit' , 'Count' , 'Extend' , 'prop' , 
			 ), 501, (501, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 3 , 1024 , (3, 0, None, None) , 0 , )),
	(( 'MoveUp' , 'Unit' , 'Count' , 'Extend' , 'prop' , 
			 ), 502, (502, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 3 , 1032 , (3, 0, None, None) , 0 , )),
	(( 'MoveDown' , 'Unit' , 'Count' , 'Extend' , 'prop' , 
			 ), 503, (503, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 3 , 1040 , (3, 0, None, None) , 0 , )),
	(( 'HomeKey' , 'Unit' , 'Extend' , 'prop' , ), 504, (504, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 1048 , (3, 0, None, None) , 0 , )),
	(( 'EndKey' , 'Unit' , 'Extend' , 'prop' , ), 505, (505, (), [ 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16387, 10, None, None) , ], 1 , 1 , 4 , 2 , 1056 , (3, 0, None, None) , 0 , )),
	(( 'EscapeKey' , ), 506, (506, (), [ ], 1 , 1 , 4 , 0 , 1064 , (3, 0, None, None) , 0 , )),
	(( 'TypeText' , 'Text' , ), 507, (507, (), [ (8, 1, None, None) , ], 1 , 1 , 4 , 0 , 1072 , (3, 0, None, None) , 0 , )),
	(( 'CopyFormat' , ), 509, (509, (), [ ], 1 , 1 , 4 , 0 , 1080 , (3, 0, None, None) , 0 , )),
	(( 'PasteFormat' , ), 510, (510, (), [ ], 1 , 1 , 4 , 0 , 1088 , (3, 0, None, None) , 0 , )),
	(( 'TypeParagraph' , ), 512, (512, (), [ ], 1 , 1 , 4 , 0 , 1096 , (3, 0, None, None) , 0 , )),
	(( 'TypeBackspace' , ), 513, (513, (), [ ], 1 , 1 , 4 , 0 , 1104 , (3, 0, None, None) , 0 , )),
	(( 'NextSubdocument' , ), 514, (514, (), [ ], 1 , 1 , 4 , 0 , 1112 , (3, 0, None, None) , 0 , )),
	(( 'PreviousSubdocument' , ), 515, (515, (), [ ], 1 , 1 , 4 , 0 , 1120 , (3, 0, None, None) , 0 , )),
	(( 'SelectColumn' , ), 516, (516, (), [ ], 1 , 1 , 4 , 0 , 1128 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentFont' , ), 517, (517, (), [ ], 1 , 1 , 4 , 0 , 1136 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentAlignment' , ), 518, (518, (), [ ], 1 , 1 , 4 , 0 , 1144 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentSpacing' , ), 519, (519, (), [ ], 1 , 1 , 4 , 0 , 1152 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentIndent' , ), 520, (520, (), [ ], 1 , 1 , 4 , 0 , 1160 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentTabs' , ), 521, (521, (), [ ], 1 , 1 , 4 , 0 , 1168 , (3, 0, None, None) , 0 , )),
	(( 'SelectCurrentColor' , ), 522, (522, (), [ ], 1 , 1 , 4 , 0 , 1176 , (3, 0, None, None) , 0 , )),
	(( 'CreateTextbox' , ), 523, (523, (), [ ], 1 , 1 , 4 , 0 , 1184 , (3, 0, None, None) , 0 , )),
	(( 'WholeStory' , ), 524, (524, (), [ ], 1 , 1 , 4 , 0 , 1192 , (3, 0, None, None) , 0 , )),
	(( 'SelectRow' , ), 525, (525, (), [ ], 1 , 1 , 4 , 0 , 1200 , (3, 0, None, None) , 0 , )),
	(( 'SplitTable' , ), 526, (526, (), [ ], 1 , 1 , 4 , 0 , 1208 , (3, 0, None, None) , 0 , )),
	(( 'InsertRows' , 'NumRows' , ), 528, (528, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1216 , (3, 0, None, None) , 0 , )),
	(( 'InsertColumns' , ), 529, (529, (), [ ], 1 , 1 , 4 , 0 , 1224 , (3, 0, None, None) , 0 , )),
	(( 'InsertFormula' , 'Formula' , 'NumberFormat' , ), 530, (530, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 2 , 1232 , (3, 0, None, None) , 0 , )),
	(( 'NextRevision' , 'Wrap' , 'prop' , ), 531, (531, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020981-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 1240 , (3, 0, None, None) , 0 , )),
	(( 'PreviousRevision' , 'Wrap' , 'prop' , ), 532, (532, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020981-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 1248 , (3, 0, None, None) , 0 , )),
	(( 'PasteAsNestedTable' , ), 533, (533, (), [ ], 1 , 1 , 4 , 0 , 1256 , (3, 0, None, None) , 0 , )),
	(( 'CreateAutoTextEntry' , 'Name' , 'StyleName' , 'prop' , ), 534, (534, (), [ 
			 (8, 1, None, None) , (8, 1, None, None) , (16393, 10, None, "IID('{00020936-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 0 , 1264 , (3, 0, None, None) , 0 , )),
	(( 'DetectLanguage' , ), 535, (535, (), [ ], 1 , 1 , 4 , 0 , 1272 , (3, 0, None, None) , 0 , )),
	(( 'SelectCell' , ), 536, (536, (), [ ], 1 , 1 , 4 , 0 , 1280 , (3, 0, None, None) , 0 , )),
	(( 'InsertRowsBelow' , 'NumRows' , ), 537, (537, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1288 , (3, 0, None, None) , 0 , )),
	(( 'InsertColumnsRight' , ), 538, (538, (), [ ], 1 , 1 , 4 , 0 , 1296 , (3, 0, None, None) , 0 , )),
	(( 'InsertRowsAbove' , 'NumRows' , ), 539, (539, (), [ (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1304 , (3, 0, None, None) , 0 , )),
	(( 'RtlRun' , ), 600, (600, (), [ ], 1 , 1 , 4 , 0 , 1312 , (3, 0, None, None) , 0 , )),
	(( 'LtrRun' , ), 601, (601, (), [ ], 1 , 1 , 4 , 0 , 1320 , (3, 0, None, None) , 0 , )),
	(( 'BoldRun' , ), 602, (602, (), [ ], 1 , 1 , 4 , 0 , 1328 , (3, 0, None, None) , 0 , )),
	(( 'ItalicRun' , ), 603, (603, (), [ ], 1 , 1 , 4 , 0 , 1336 , (3, 0, None, None) , 0 , )),
	(( 'RtlPara' , ), 605, (605, (), [ ], 1 , 1 , 4 , 0 , 1344 , (3, 0, None, None) , 0 , )),
	(( 'LtrPara' , ), 606, (606, (), [ ], 1 , 1 , 4 , 0 , 1352 , (3, 0, None, None) , 0 , )),
	(( 'InsertDateTime' , 'DateTimeFormat' , 'InsertAsField' , 'InsertAsFullWidth' , 'DateLanguage' , 
			 'CalendarType' , ), 444, (444, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 5 , 1360 , (3, 0, None, None) , 0 , )),
	(( 'Sort2000' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'SortColumn' , 'Separator' , 'CaseSensitive' , 'BidiSort' , 
			 'IgnoreThe' , 'IgnoreKashida' , 'IgnoreDiacritics' , 'IgnoreHe' , 'LanguageID' , 
			 ), 445, (445, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 19 , 1368 , (3, 0, None, None) , 64 , )),
	(( 'ConvertToTable' , 'Separator' , 'NumRows' , 'NumColumns' , 'InitialColumnWidth' , 
			 'Format' , 'ApplyBorders' , 'ApplyShading' , 'ApplyFont' , 'ApplyColor' , 
			 'ApplyHeadingRows' , 'ApplyLastRow' , 'ApplyFirstColumn' , 'ApplyLastColumn' , 'AutoFit' , 
			 'AutoFitBehavior' , 'DefaultTableBehavior' , 'prop' , ), 457, (457, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{00020951-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 16 , 1376 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 1005, (1005, (), [ (16387, 10, None, None) , ], 1 , 2 , 4 , 0 , 1384 , (3, 0, None, None) , 0 , )),
	(( 'NoProofing' , 'prop' , ), 1005, (1005, (), [ (3, 1, None, None) , ], 1 , 4 , 4 , 0 , 1392 , (3, 0, None, None) , 0 , )),
	(( 'TopLevelTables' , 'prop' , ), 1006, (1006, (), [ (16393, 10, None, "IID('{0002094D-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1400 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 1007, (1007, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1408 , (3, 0, None, None) , 0 , )),
	(( 'LanguageDetected' , 'prop' , ), 1007, (1007, (), [ (11, 1, None, None) , ], 1 , 4 , 4 , 0 , 1416 , (3, 0, None, None) , 0 , )),
	(( 'FitTextWidth' , 'prop' , ), 1008, (1008, (), [ (16388, 10, None, None) , ], 1 , 2 , 4 , 0 , 1424 , (3, 0, None, None) , 0 , )),
	(( 'FitTextWidth' , 'prop' , ), 1008, (1008, (), [ (4, 1, None, None) , ], 1 , 4 , 4 , 0 , 1432 , (3, 0, None, None) , 0 , )),
	(( 'ClearFormatting' , ), 1009, (1009, (), [ ], 1 , 1 , 4 , 0 , 1440 , (3, 0, None, None) , 0 , )),
	(( 'PasteAppendTable' , ), 1010, (1010, (), [ ], 1 , 1 , 4 , 0 , 1448 , (3, 0, None, None) , 0 , )),
	(( 'HTMLDivisions' , 'prop' , ), 1011, (1011, (), [ (16393, 10, None, "IID('{000209E8-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1456 , (3, 0, None, None) , 0 , )),
	(( 'SmartTags' , 'prop' , ), 1015, (1015, (), [ (16393, 10, None, "IID('{000209EE-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1464 , (3, 0, None, None) , 64 , )),
	(( 'ChildShapeRange' , 'prop' , ), 1021, (1021, (), [ (16393, 10, None, "IID('{000209B5-0000-0000-C000-000000000046}')") , ], 1 , 2 , 4 , 0 , 1472 , (3, 0, None, None) , 0 , )),
	(( 'HasChildShapeRange' , 'prop' , ), 1022, (1022, (), [ (16395, 10, None, None) , ], 1 , 2 , 4 , 0 , 1480 , (3, 0, None, None) , 0 , )),
	(( 'FootnoteOptions' , 'prop' , ), 1024, (1024, (), [ (16393, 10, None, "IID('{BEA85A24-D7DA-4F3D-B58C-ED90FB01D615}')") , ], 1 , 2 , 4 , 0 , 1488 , (3, 0, None, None) , 0 , )),
	(( 'EndnoteOptions' , 'prop' , ), 1025, (1025, (), [ (16393, 10, None, "IID('{BF043168-F4DE-4E7C-B206-741A8B3EF71A}')") , ], 1 , 2 , 4 , 0 , 1496 , (3, 0, None, None) , 0 , )),
	(( 'ToggleCharacterCode' , ), 1012, (1012, (), [ ], 1 , 1 , 4 , 0 , 1504 , (3, 0, None, None) , 0 , )),
	(( 'PasteAndFormat' , 'Type' , ), 1013, (1013, (), [ (3, 1, None, None) , ], 1 , 1 , 4 , 0 , 1512 , (3, 0, None, None) , 0 , )),
	(( 'PasteExcelTable' , 'LinkedToExcel' , 'WordFormatting' , 'RTF' , ), 1014, (1014, (), [ 
			 (11, 1, None, None) , (11, 1, None, None) , (11, 1, None, None) , ], 1 , 1 , 4 , 0 , 1520 , (3, 0, None, None) , 0 , )),
	(( 'ShrinkDiscontiguousSelection' , ), 1019, (1019, (), [ ], 1 , 1 , 4 , 0 , 1528 , (3, 0, None, None) , 0 , )),
	(( 'InsertStyleSeparator' , ), 1020, (1020, (), [ ], 1 , 1 , 4 , 0 , 1536 , (3, 0, None, None) , 0 , )),
	(( 'Sort' , 'ExcludeHeader' , 'FieldNumber' , 'SortFieldType' , 'SortOrder' , 
			 'FieldNumber2' , 'SortFieldType2' , 'SortOrder2' , 'FieldNumber3' , 'SortFieldType3' , 
			 'SortOrder3' , 'SortColumn' , 'Separator' , 'CaseSensitive' , 'BidiSort' , 
			 'IgnoreThe' , 'IgnoreKashida' , 'IgnoreDiacritics' , 'IgnoreHe' , 'LanguageID' , 
			 'SubFieldNumber' , 'SubFieldNumber2' , 'SubFieldNumber3' , ), 1023, (1023, (), [ (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 22 , 1544 , (3, 0, None, None) , 0 , )),
	(( 'XMLNodes' , 'prop' , ), 310, (310, (), [ (16393, 10, None, "IID('{D36C1F42-7044-4B9E-9CA3-85919454DB04}')") , ], 1 , 2 , 4 , 0 , 1552 , (3, 0, None, None) , 64 , )),
	(( 'XMLParentNode' , 'prop' , ), 311, (311, (), [ (16393, 10, None, "IID('{09760240-0B89-49F7-A79D-479F24723F56}')") , ], 1 , 2 , 4 , 0 , 1560 , (3, 0, None, None) , 64 , )),
	(( 'Editors' , 'prop' , ), 313, (313, (), [ (16393, 10, None, "IID('{AED7E08C-14F0-4F33-921D-4C5353137BF6}')") , ], 1 , 2 , 4 , 0 , 1568 , (3, 0, None, None) , 0 , )),
	(( 'XML' , 'DataOnly' , 'prop' , ), 314, (314, (), [ (11, 49, 'False', None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
	(( 'XML' , 'DataOnly' , 'prop' , ), 314, (314, (), [ (11, 49, 'False', None) , 
			 (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1576 , (3, 0, None, None) , 0 , )),
	(( 'EnhMetaFileBits' , 'prop' , ), 315, (315, (), [ (16396, 10, None, None) , ], 1 , 2 , 4 , 0 , 1584 , (3, 0, None, None) , 0 , )),
	(( 'GoToEditableRange' , 'EditorID' , 'prop' , ), 1027, (1027, (), [ (16396, 17, None, None) , 
			 (16393, 10, None, "IID('{0002095E-0000-0000-C000-000000000046}')") , ], 1 , 1 , 4 , 1 , 1592 , (3, 0, None, None) , 0 , )),
	(( 'InsertXML' , 'XML' , 'Transform' , ), 1028, (1028, (), [ (8, 1, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1600 , (3, 0, None, None) , 0 , )),
	(( 'InsertCaption' , 'Label' , 'Title' , 'TitleAutoText' , 'Position' , 
			 'ExcludeLabel' , ), 417, (417, (), [ (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 1608 , (3, 0, None, None) , 0 , )),
	(( 'InsertCrossReference' , 'ReferenceType' , 'ReferenceKind' , 'ReferenceItem' , 'InsertAsHyperlink' , 
			 'IncludePosition' , 'SeparateNumbers' , 'SeparatorString' , ), 418, (418, (), [ (16396, 1, None, None) , 
			 (3, 1, None, None) , (16396, 1, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , ], 1 , 1 , 4 , 4 , 1616 , (3, 0, None, None) , 0 , )),
	(( 'OMaths' , 'prop' , ), 316, (316, (), [ (16393, 10, None, "IID('{873E774B-926A-4CB1-878D-635A45187595}')") , ], 1 , 2 , 4 , 0 , 1624 , (3, 0, None, None) , 0 , )),
	(( 'WordOpenXML' , 'prop' , ), 317, (317, (), [ (16392, 10, None, None) , ], 1 , 2 , 4 , 0 , 1632 , (3, 0, None, None) , 0 , )),
	(( 'ClearParagraphStyle' , ), 1030, (1030, (), [ ], 1 , 1 , 4 , 0 , 1640 , (3, 0, None, None) , 0 , )),
	(( 'ClearCharacterAllFormatting' , ), 1031, (1031, (), [ ], 1 , 1 , 4 , 0 , 1648 , (3, 0, None, None) , 0 , )),
	(( 'ClearCharacterStyle' , ), 1032, (1032, (), [ ], 1 , 1 , 4 , 0 , 1656 , (3, 0, None, None) , 0 , )),
	(( 'ClearCharacterDirectFormatting' , ), 1033, (1033, (), [ ], 1 , 1 , 4 , 0 , 1664 , (3, 0, None, None) , 0 , )),
	(( 'ContentControls' , 'prop' , ), 1034, (1034, (), [ (16393, 10, None, "IID('{804CD967-F83B-432D-9446-C61A45CFEFF0}')") , ], 1 , 2 , 4 , 0 , 1672 , (3, 0, None, None) , 64 , )),
	(( 'ParentContentControl' , 'prop' , ), 1035, (1035, (), [ (16393, 10, None, "IID('{EE95AFE3-3026-4172-B078-0E79DAB5CC3D}')") , ], 1 , 2 , 4 , 0 , 1680 , (3, 0, None, None) , 64 , )),
	(( 'ExportAsFixedFormat' , 'OutputFileName' , 'ExportFormat' , 'OpenAfterExport' , 'OptimizeFor' , 
			 'ExportCurrentPage' , 'Item' , 'IncludeDocProps' , 'KeepIRM' , 'CreateBookmarks' , 
			 'DocStructureTags' , 'BitmapMissingFonts' , 'UseISO19005_1' , 'FixedFormatExtClassPtr' , ), 1036, (1036, (), [ 
			 (8, 1, None, None) , (3, 1, None, None) , (11, 49, 'False', None) , (3, 49, '0', None) , (11, 49, 'False', None) , 
			 (3, 49, '0', None) , (11, 49, 'False', None) , (11, 49, 'True', None) , (3, 49, '0', None) , (11, 49, 'True', None) , 
			 (11, 49, 'True', None) , (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1688 , (3, 0, None, None) , 0 , )),
	(( 'ReadingModeGrowFont' , ), 1037, (1037, (), [ ], 1 , 1 , 4 , 0 , 1696 , (3, 0, None, None) , 0 , )),
	(( 'ReadingModeShrinkFont' , ), 1038, (1038, (), [ ], 1 , 1 , 4 , 0 , 1704 , (3, 0, None, None) , 0 , )),
	(( 'ClearParagraphAllFormatting' , ), 1039, (1039, (), [ ], 1 , 1 , 4 , 0 , 1712 , (3, 0, None, None) , 0 , )),
	(( 'ClearParagraphDirectFormatting' , ), 1040, (1040, (), [ ], 1 , 1 , 4 , 0 , 1720 , (3, 0, None, None) , 0 , )),
	(( 'InsertNewPage' , ), 1041, (1041, (), [ ], 1 , 1 , 4 , 0 , 1728 , (3, 0, None, None) , 0 , )),
	(( 'SortByHeadings' , 'SortFieldType' , 'SortOrder' , 'CaseSensitive' , 'BidiSort' , 
			 'IgnoreThe' , 'IgnoreKashida' , 'IgnoreDiacritics' , 'IgnoreHe' , 'LanguageID' , 
			 ), 1042, (1042, (), [ (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , 
			 (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 9 , 1736 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat2' , 'OutputFileName' , 'ExportFormat' , 'OpenAfterExport' , 'OptimizeFor' , 
			 'ExportCurrentPage' , 'Item' , 'IncludeDocProps' , 'KeepIRM' , 'CreateBookmarks' , 
			 'DocStructureTags' , 'BitmapMissingFonts' , 'UseISO19005_1' , 'OptimizeForImageQuality' , 'FixedFormatExtClassPtr' , 
			 ), 1043, (1043, (), [ (8, 1, None, None) , (3, 1, None, None) , (11, 49, 'False', None) , (3, 49, '0', None) , 
			 (11, 49, 'False', None) , (3, 49, '0', None) , (11, 49, 'False', None) , (11, 49, 'True', None) , (3, 49, '0', None) , 
			 (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'False', None) , (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1744 , (3, 0, None, None) , 0 , )),
	(( 'ExportAsFixedFormat3' , 'OutputFileName' , 'ExportFormat' , 'OpenAfterExport' , 'OptimizeFor' , 
			 'ExportCurrentPage' , 'Item' , 'IncludeDocProps' , 'KeepIRM' , 'CreateBookmarks' , 
			 'DocStructureTags' , 'BitmapMissingFonts' , 'UseISO19005_1' , 'OptimizeForImageQuality' , 'ImproveExportTagging' , 
			 'FixedFormatExtClassPtr' , ), 1044, (1044, (), [ (8, 1, None, None) , (3, 1, None, None) , (11, 49, 'False', None) , 
			 (3, 49, '0', None) , (11, 49, 'False', None) , (3, 49, '0', None) , (11, 49, 'False', None) , (11, 49, 'True', None) , 
			 (3, 49, '0', None) , (11, 49, 'True', None) , (11, 49, 'True', None) , (11, 49, 'False', None) , (11, 49, 'False', None) , 
			 (11, 49, 'False', None) , (16396, 17, None, None) , ], 1 , 1 , 4 , 1 , 1752 , (3, 0, None, None) , 0 , )),
]

win32com.client.CLSIDToClass.RegisterCLSID( "{00020975-0000-0000-C000-000000000046}", Selection )
