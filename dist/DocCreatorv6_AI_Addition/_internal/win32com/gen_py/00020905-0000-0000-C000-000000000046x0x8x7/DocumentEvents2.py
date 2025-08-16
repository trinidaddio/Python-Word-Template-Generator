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

class DocumentEvents2:
	CLSID = CLSID_Sink = IID('{00020A02-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020906-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		        4 : "OnNew",
		        5 : "OnOpen",
		        6 : "OnClose",
		        7 : "OnSync",
		        8 : "OnXMLAfterInsert",
		        9 : "OnXMLBeforeDelete",
		       12 : "OnContentControlAfterAdd",
		       13 : "OnContentControlBeforeDelete",
		       14 : "OnContentControlOnExit",
		       15 : "OnContentControlOnEnter",
		       16 : "OnContentControlBeforeStoreUpdate",
		       17 : "OnContentControlBeforeContentUpdate",
		       18 : "OnBuildingBlockInsert",
		       19 : "OnContentControlNonContentChange",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnNew(self):
#	def OnOpen(self):
#	def OnClose(self):
#	def OnSync(self, SyncEventType=defaultNamedNotOptArg):
#	def OnXMLAfterInsert(self, NewXMLNode=defaultNamedNotOptArg, InUndoRedo=defaultNamedNotOptArg):
#	def OnXMLBeforeDelete(self, DeletedRange=defaultNamedNotOptArg, OldXMLNode=defaultNamedNotOptArg, InUndoRedo=defaultNamedNotOptArg):
#	def OnContentControlAfterAdd(self, NewContentControl=defaultNamedNotOptArg, InUndoRedo=defaultNamedNotOptArg):
#	def OnContentControlBeforeDelete(self, OldContentControl=defaultNamedNotOptArg, InUndoRedo=defaultNamedNotOptArg):
#	def OnContentControlOnExit(self, ContentControl=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnContentControlOnEnter(self, ContentControl=defaultNamedNotOptArg):
#	def OnContentControlBeforeStoreUpdate(self, ContentControl=defaultNamedNotOptArg, Content=defaultNamedNotOptArg):
#	def OnContentControlBeforeContentUpdate(self, ContentControl=defaultNamedNotOptArg, Content=defaultNamedNotOptArg):
#	def OnBuildingBlockInsert(self, Range=defaultNamedNotOptArg, Name=defaultNamedNotOptArg, Category=defaultNamedNotOptArg, BlockType=defaultNamedNotOptArg
#			, Template=defaultNamedNotOptArg):
#	def OnContentControlNonContentChange(self, ContentControl=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00020A02-0000-0000-C000-000000000046}", DocumentEvents2 )
