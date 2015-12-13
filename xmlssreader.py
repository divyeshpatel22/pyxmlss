#!/user/bin/env python
import os
from lxml import etree
import sys
from xmlssbase import XmlssBase

class XmlssReader(XmlssBase):
	'''
	This is the Base xmlss reder class which contains the xmlss reading
	functionality related methods and properties.
	'''
	
	
	def readXmlss(self,fileName):
		'''to Read XMLSS file from particular file location.'''
		if self._debug:
			print "in readXmlss"
		if os.path.isfile(fileName):
			self._xmlssDoc = etree.parse(fileName,self._parser)
			print dir(self._xmlssDoc.getroot())
			print self._xmlssDoc.getroot().prefix
			print dir(self._xmlssDoc.getroot().tag)
			rootElem = etree.QName(self._xmlssDoc.getroot())
			if rootElem.localname != "Workbook":
				print("***E The File {0} is not and XMLSS file please provide valid file".format(fileName))
		else:
			print("***E The File {0} does not exist Please provide valid XMLSS file Path".format(fileName))
		if self._debug:
			print "out readXmlss"
			
	def getWorksheetList (self):
		'''this method will return the list of worksheet names available in the current workbok'''
		if self._debug:
			print "in getWorksheetList"
		lstWorkSheetElem = self._xmlssDoc.xpath("//ss:Worksheet",namespaces=self._xmlssXPathNameSpaceMap)
		lstWorkSheetName = []
		if len(lstWorkSheetElem) == 0 :
			print "***W No worksheet exist in the current workbook"
			if self._debug:
				print "out getWorksheetList"
			return lstWorkSheetName
		else:
			for workSheetElem in lstWorkSheetElem:
				worksheetName = workSheetElem.get(etree.QName(self._xmlssNameSpaceMap["ss"],"Name"))
				if self._debug:
					print worksheetName
				lstWorkSheetName.append(worksheetName)
			if self._debug:
				print "out getWorksheetList"
			return lstWorkSheetName
		
	def getCurrentWorksheet (self):
		'''get the name of the worksheet presently being used.'''
		
		if self._xmlssWorkSheetName != None:
			return self._xmlssWorkSheetName
		else:
			print "***W please use the api setCurrentWorksheet to set current workSheet"
			return None
			
	def getRowCount (self):
		'''It will give the total no. of rows available in the worksheet'''		
		totalRows = int (self._xmlssDoc.xpath("count(/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row".format(self._xmlssWorkSheetName),namespaces=self)._xmlssXPathNameSpaceMap)
		return totalRows

		
	def getActualRowCount (self):
		'''It will return the actual row count after substracting the merged rows'''
		
		totalRows = getRowCount()		
		lstMergedRows = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[@ss:MergeDown and position() = 1]".format(self._xmlssWorkSheetName),namespaces=self._xmlssXPathNameSpaceMap)
		
		for mergedRowElem in lstMergedRows:
			try:
				mergedRowCount = mergedRowElem.attrib[QName(self._xmlssNameSpaceMap["ss"],"MergeDown")]
				totalRows -= int(mergedRowCount)
			except:
				print "Unexpected error:", sys.exc_info()[0]
		return totalRows
		
	def getRow(self,rowIndex):		
		rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)
		
		if rowElem != None and len(rowElem) > 0:
			rowData=[]
			
		
		
		
		pass

	#Constructor to create instance of the XmlssBase

	# def __init__(debug=0):
		# _debug=debug

	
	
