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

class ApplicationEvents3:
	CLSID = CLSID_Sink = IID('{00020A00-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{000209FF-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		1610612736 : "OnQueryInterface",
		1610612737 : "OnAddRef",
		1610612738 : "OnRelease",
		1610678272 : "OnGetTypeInfoCount",
		1610678273 : "OnGetTypeInfo",
		1610678274 : "OnGetIDsOfNames",
		1610678275 : "OnInvoke",
		        1 : "OnStartup",
		        2 : "OnQuit",
		        3 : "OnDocumentChange",
		        4 : "OnDocumentOpen",
		        6 : "OnDocumentBeforeClose",
		        7 : "OnDocumentBeforePrint",
		        8 : "OnDocumentBeforeSave",
		        9 : "OnNewDocument",
		       10 : "OnWindowActivate",
		       11 : "OnWindowDeactivate",
		       12 : "OnWindowSelectionChange",
		       13 : "OnWindowBeforeRightClick",
		       14 : "OnWindowBeforeDoubleClick",
		       15 : "OnEPostagePropertyDialog",
		       16 : "OnEPostageInsert",
		       17 : "OnMailMergeAfterMerge",
		       18 : "OnMailMergeAfterRecordMerge",
		       19 : "OnMailMergeBeforeMerge",
		       20 : "OnMailMergeBeforeRecordMerge",
		       21 : "OnMailMergeDataSourceLoad",
		       22 : "OnMailMergeDataSourceValidate",
		       23 : "OnMailMergeWizardSendToCustom",
		       24 : "OnMailMergeWizardStateChange",
		       25 : "OnWindowSize",
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
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnStartup(self):
#	def OnQuit(self):
#	def OnDocumentChange(self):
#	def OnDocumentOpen(self, Doc=defaultNamedNotOptArg):
#	def OnDocumentBeforeClose(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnDocumentBeforePrint(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnDocumentBeforeSave(self, Doc=defaultNamedNotOptArg, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnNewDocument(self, Doc=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):
#	def OnWindowSelectionChange(self, Sel=defaultNamedNotOptArg):
#	def OnWindowBeforeRightClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnWindowBeforeDoubleClick(self, Sel=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnEPostagePropertyDialog(self, Doc=defaultNamedNotOptArg):
#	def OnEPostageInsert(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeAfterMerge(self, Doc=defaultNamedNotOptArg, DocResult=defaultNamedNotOptArg):
#	def OnMailMergeAfterRecordMerge(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeBeforeMerge(self, Doc=defaultNamedNotOptArg, StartRecord=defaultNamedNotOptArg, EndRecord=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeBeforeRecordMerge(self, Doc=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceLoad(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeDataSourceValidate(self, Doc=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnMailMergeWizardSendToCustom(self, Doc=defaultNamedNotOptArg):
#	def OnMailMergeWizardStateChange(self, Doc=defaultNamedNotOptArg, FromState=defaultNamedNotOptArg, ToState=defaultNamedNotOptArg, Handled=defaultNamedNotOptArg):
#	def OnWindowSize(self, Doc=defaultNamedNotOptArg, Wn=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00020A00-0000-0000-C000-000000000046}", ApplicationEvents3 )
