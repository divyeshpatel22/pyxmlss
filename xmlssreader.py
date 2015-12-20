#!/user/bin/env python
import os
from lxml import etree
import sys
from xmlssbase import XmlssBase

class XmlssReader(XmlssBase):
	'''
	This is the Base xmlss reader class which contains the xmlss reading
	functionality related methods and properties.
	'''
	
	
	def readXmlss(self,fileName):
		'''to Read XMLSS file from particular file location.'''
		if self._debug:
			print "in readXmlss"
		if os.path.isfile(fileName):
			self._xmlssDoc = etree.parse(fileName,self._parser)
			#print dir(self._xmlssDoc.getroot())
			if self._debug:
				print self._xmlssDoc.getroot().prefix
			#print dir(self._xmlssDoc.getroot().tag)
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
		totalRows = int(self._xmlssDoc.xpath("count(/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row)".format(self._xmlssWorkSheetName),namespaces=self._xmlssXPathNameSpaceMap))		
		return totalRows

		
	def getActualRowCount (self):
		'''It will return the actual row count after substracting the merged rows'''
		
		totalRows = self.getRowCount()		
		lstMergedRows = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row/ss:Cell[@ss:MergeDown and position() = 1]".format(self._xmlssWorkSheetName),namespaces=self._xmlssXPathNameSpaceMap)
		
		for mergedRowElem in lstMergedRows:			
			try:				
				mergedRowCount = mergedRowElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown")]
				totalRows -= int(mergedRowCount)
			except:
				print "Unexpected error:", sys.exc_info()[0]
		return totalRows	
		
	def getRow(self,rowIndex):		
		rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)
		self._xmlssCurrentRowIndex = rowIndex
		finalRowData = []
		isMergedRow = 0
		if rowElem != None and len(rowElem) > 0:
			lstCellElem = rowElem[0].findall(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
			cellIndex = 1			
			for CellElem in lstCellElem:
				if etree.QName(self._xmlssNameSpaceMap["ss"],"Index") in CellElem.attrib:
					isMergedRow = 1
					curCellIndex = int(CellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Index")])
					for i in range(cellIndex,curCellIndex):
						finalRowData.append("")
				else :
					cellIndex+=1
					
				actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
				finalRowData.append(actData.text)
		else:
			print "***E Worksheet{0} is not having row at index {1}".format(self._xmlssWorkSheetName,rowIndex)
				
		if isMergedRow == 1:
			print "***W this row is part of merged row."
		self._xmlssNextRowIndex = rowIndex+1
		return finalRowData

	
	def getMergedRowIndexes(self):
		lstMergedRows = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[ss:Cell[@ss:MergeDown and position() = 1]]".format(self._xmlssWorkSheetName),namespaces=self._xmlssXPathNameSpaceMap)
		lstMergedRowIndexes = []
		for mergedRowElem in lstMergedRows:		
			mergedRowElemXpath = "count({0}/preceding-sibling::ss:Row)".format(self._xmlssDoc.getpath(mergedRowElem))
			if self._debug:
				print mergedRowElemXpath
			indexOfMegredRow =  int(self._xmlssDoc.xpath(mergedRowElemXpath,namespaces=self._xmlssXPathNameSpaceMap)) + 1
			lstMergedRowIndexes.append(indexOfMegredRow)
		
		return lstMergedRowIndexes
	
	def getMergedRow(self,rowIndex):
		'''This method is used to get the data of the merged row.'''
		if self._debug:					
			print "in getMergedRow"
		rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)
		self._xmlssCurrentRowIndex = rowIndex
		finalRowData = []
		isMergedRow = 0
		MergeRowCount = 0
		dictRowData = {}
		totalColumnCount = 0
		if rowElem != None and len(rowElem) > 0:			
			lstCellElem = rowElem[0].findall(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
			cellIndex = 1
			if cellIndex == 1 and etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown") in lstCellElem[0].attrib:
				MergeRowCount = int(lstCellElem[0].attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown")])
				for CellElem in lstCellElem:				
					#first proces the current row and the process remainig merged rows.
					actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
					dictRowData[cellIndex] =  actData.text
					cellIndex += 1
				totalColumnCount = cellIndex
			else:
				print "***W row at Index {0} is not a merged row getting single row value.".format(rowIndex)
				finalRowData = self.getRow(rowIndex)
				if self._debug:		
					print finalRowData
					print "out getMergedRow"
				return finalRowData
		else:
			print "***E Worksheet{0} is not having row at index {1}".format(self._xmlssWorkSheetName,rowIndex)
		if self._debug:
			print dictRowData
			
		#now add merged row data to dictionary
		for i in range(1,MergeRowCount + 1):
			curIndex = rowIndex + i
			childRowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,curIndex),namespaces=self._xmlssXPathNameSpaceMap)
			if childRowElem != None and len(childRowElem) > 0:
				lstCellElem = childRowElem[0].findall(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
				curCellIndex = 2 
				for CellElem in lstCellElem:
					if etree.QName(self._xmlssNameSpaceMap["ss"],"Index") in CellElem.attrib:
						#cell is having Index attribute so changes the current cell index
						curCellIndex = int(CellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Index")])
						if type(dictRowData[curCellIndex]) is list:
							actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
							dictRowData[curCellIndex].append(actData.text)
						else:
							tempDictData = dictRowData[curCellIndex]
							dictRowData[curCellIndex] = [tempDictData]
							actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
							dictRowData[curCellIndex].append(actData.text)
					else :
						#cell is not having Index attribute use the incremented cell index for curent value
						if type(dictRowData[curCellIndex]) is list:
							actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
							dictRowData[curCellIndex].append(actData.text)
						else:
							tempDictData = dictRowData[curCellIndex]
							dictRowData[curCellIndex] = [tempDictData]
							actData = CellElem.find(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
							dictRowData[curCellIndex].append(actData.text)
					curCellIndex+=1					
			else:
				print "***E Worksheet{0} is not having row at merged row index {1}".format(self._xmlssWorkSheetName,i)
		
		for colIndex in range(1,totalColumnCount):
			finalRowData.append(dictRowData[colIndex])	

		self._xmlssNextRowIndex = rowIndex + MergeRowCount + 1
		#now prepare final row data.
		if self._debug:		
			print finalRowData
			print "out getMergedRow"
		return finalRowData
		
	def getNextRow(self):
		if self._xmlssNextRowIndex == None:
			rowIndex = 1
			rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)			
			if rowElem != None and len(rowElem) > 0:
				CellElem = rowElem[0].find(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
				if etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown") in CellElem.attrib:
					return self.getMergedRow(rowIndex)
				else:
					return self.getRow(rowIndex)			
			else:
				print "***E Worksheet{0} is not having row at index {1}".format(self._xmlssWorkSheetName,rowIndex)
		else:
			rowIndex = self._xmlssNextRowIndex
			rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)			
			if rowElem != None and len(rowElem) > 0:
				CellElem = rowElem[0].find(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
				if etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown") in CellElem.attrib:
					return self.getMergedRow(rowIndex)
				else:
					return self.getRow(rowIndex)			
			else:
				print "***E Worksheet{0} is not having row at index {1}".format(self._xmlssWorkSheetName,rowIndex)
				
	def getRows(self,fromRowIndex,toRowIndex):
	
		if not (type(fromRowIndex) is int)  or  not ( type(toRowIndex) is int) or (fromRowIndex < 1) or (toRowIndex - fromRowIndex < 1) or (toRowIndex > self.getRowCount()):
			print "***E please provide valid value of fromRowIndex and toRowIndex "
			return None
		else:
			finalData = []
			rowIndex = fromRowIndex
			rowElem = self._xmlssDoc.xpath("/ss:Workbook/ss:Worksheet[@ss:Name='{0}']/ss:Table/ss:Row[position() = {1}]".format(self._xmlssWorkSheetName,rowIndex),namespaces=self._xmlssXPathNameSpaceMap)			
			if rowElem != None and len(rowElem) > 0:
				#get the initial row data 
				CellElem = rowElem[0].find(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
				if etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown") in CellElem.attrib:
					 finalData.append(self.getMergedRow(rowIndex))
				else:
					finalData.append(self.getRow(rowIndex))
				#now for further remainig rows
				for i in range(1,(toRowIndex-fromRowIndex)+1):
					finalData.append(self.getNextRow())
			else:
				print "***E Worksheet{0} is not having row at index {1}".format(self._xmlssWorkSheetName,rowIndex)
				
		return finalData
			
			#by default get the first row
			
		
	#remaining items are 
	#get merged row -- done
	#get next row -- done
	#getrows(i,j) 
	#get mergedRowIndexes. -- done

	#Constructor to create instance of the XmlssBase

	# def __init__(debug=0):
		# _debug=debug
	