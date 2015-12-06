#!/user/bin/env python
import os
from lxml import etree
from datetime import datetime
from dateutil.tz import tzlocal
from xmlssbase import XmlssBase

class XmlssWriter(XmlssBase):


	def createXmlss(self):
		''' This method is used to create basic structure of the XMLSS'''
		print self._xmlssNameSpaceMap
		elemRootWorkBook = etree.Element(self._xmlnsDefaultElem +"Workbook",nsmap=self._xmlssNameSpaceMap)

		self._xmlssDoc = etree.ElementTree(elemRootWorkBook)
		self.createDocumentProperties()
		#self.createStyle(styleName="default_style",hAlign="Left",vAlign="Top",wrapText=0,borderPosition="All",borderLineStyle="Continuous",borderWeight="1",borderColor="#000000",fontSize=11,fontName="Calibri",fontColor="#000000")
		self.createBasicStyles()
		print self._xmlssDoc

	
	def createDocumentProperties(self):
		'''This method is used to create document properties element int XMLSS'''
		workbookElem = self._xmlssDoc.getroot()
		docPropElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"Name"),xmlns = self._xmlnsO)
		temp=etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"Author"))
		temp.text = os.getlogin()
		docPropElem.append(temp)
		temp=etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"LastAuthor"))
		temp.text = os.getlogin()
		docPropElem.append(temp)
		temp=etree.Element( etree.QName(self._xmlssNameSpaceMap["o"],"Created"))
		curDateTime = datetime.now(tzlocal())
		temp.text = curDateTime.strftime("%a %B %d %H:%M:%S %Z %Y")
		docPropElem.append(temp)
		temp=etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"LastSaved"))
		temp.text = curDateTime.strftime("%a %B %d %H:%M:%S %Z %Y")
		docPropElem.append(temp)
		temp=etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"Company"))
		temp.text = os.getenv("COMPANYNAME")
		docPropElem.append(temp)
		temp=etree.Element(etree.QName(self._xmlssNameSpaceMap["o"],"Version"))
		temp.text ="12.00" 
		docPropElem.append(temp)
		workbookElem.append(docPropElem)
		
	
	def createBasicStyles(self):
		'''This method is used to create the default style for the XMLSS'''
		workbookElem = self._xmlssDoc.getroot()	
		stylesElem= self._xmlssDoc.xpath("//ss:Workbook/ss:Styles",namespaces=self._xmlssXPathNameSpaceMap) 

		#check that styles element exist and create one if it doesn't exist
		if len(stylesElem) == 0 :
			stylesElem=etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Styles"))
			workbookElem.append(stylesElem)

		#create style element		
		styleElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Style"))
		styleElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"ID")] ="Default"
		styleElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Name")] = "default_style"

		#to create Alignment.
		allignElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Alignment"))
		#hAlign="Left",vAlign="Top",wrapText=0
		allignElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Horizontal")] ="Left"
		allignElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Vertical")] ="Top"

		styleElem.append(allignElem)

		bordersElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Borders"))
		borderElem1 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
		borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Left"
		borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = "Continuous"
		borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = "1"
		borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "#000000"
		bordersElem.append(borderElem1)
		borderElem2 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
		borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Right"
		borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = "Continuous"
		borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = "1"
		borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "#000000"
		bordersElem.append(borderElem2)
		borderElem3 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
		borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Top"
		borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = "Continuous"
		borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = "1"
		borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "#000000"
		bordersElem.append(borderElem3)
		borderElem4 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
		borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Bottom"
		borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = "Continuous"
		borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = "1"
		borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "#000000"
		bordersElem.append(borderElem4)

		styleElem.append(bordersElem)

		fontElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Font"))
		fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Size")] = "11"
		fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"FontName")] = "Calibri"
		fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["x"],"Family")] = "Swiss"
		fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "#000000"

		styleElem.append(etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Interior")))
		styleElem.append(etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"NumberFormat")))

		protectionElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Protection"))
		protectionElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Protected")] = "0"
		styleElem.append(protectionElem)
		stylesElem.append(styleElem)
		
	    
	def createStyle(self,styleName,hAlign="Left",vAlign="Top",wrapText=0,borderPosition=None,borderLineStyle="Continuous",borderWeight="1",borderColor="#000000",fontBold=None,fontItalic=None,fontUnderLine=None,fontSize=None,fontName=None,fontColor=None,BGColor=None,styleParent=None,overWriteStyle=0):
		'''	createStyle accepts below mentioned args (its better to provide keyword arguments while calling this function.)
		styleName = Name of the style
		hAlign = Horizontal alignment of text in cell [<Left,Right,Center>Default Left]
		vAlign = Vertical alignment of text in cell [<Top,Bottom,Center>Default Top]
		wrapText = wrapping of text in the cell [<0,1>Default 0]
		borderPosition = position of the border in a cell [<Left,Top,Right,Bottom,All>]
		borderLineStyle = border style of the cell [<Continuous,Dash,DashDot,DashDotDot>]
		borderWight = thick ness of the border of the cell [<1,2,3,4>]
		borderColor = it will set the border color of the cell . It will accept the value in the form of #rrggbb format.
		fontBold = to make the text bold in the cell 
		fontItalic = to make text italic in the cell
		fontSize = size of the font in the cell [<integer value>]
		fontName = type of font name [<Valid font name>]
		fontColor = it will set the font color (foreground color). it will accept the value in the form of #rrggbb format.
		BGColor = it will set the background color of the cell.  it will accept the value in the form of #rrggbb format.
		styleParent = Parent style name (if any)
		overWrtieStyle = over write existing style with the same name
		'''

		workbookElem = self._xmlssDoc.getroot()
		stylesElem= self._xmlssDoc.xpath("//ss:Workbook/ss:Styles",namespaces=self._xmlssXPathNameSpaceMap) 

		#check that styles element exist and create one if it doesn't exist
		if len(stylesElem) == 0 :
			stylesElem=etree.Element("Styles")
			workbookElem.append(stylesElem)

		styleId="s"
		tempstyleId=1

		#to generate unique style id
		while 1:
		#print "//ss:Workbook/ss:Styles/ss:Style[@ss:ID={0}{1}]".format(styleId,tempstyleId)
			checkStyleId = self._xmlssDoc.xpath("//ss:Workbook/ss:Styles/ss:Style[@ss:ID={0}{1}]".format(styleId,tempstyleId),namespaces=self._xmlssXPathNameSpaceMap)
			if len(checkStyleId) != 0:
				tempstyleId += 1
			else:
				break
		#check that style with the same name exist or not

		checkStyleName = self._xmlssDoc.xpath("//ss:Workbook/ss:Styles/ss:Style[@ss:Name={0}]".format(styleName),namespaces=self._xmlssXPathNameSpaceMap)
		if len(checkStyleName) != 0 :
			if overWriteStyle == 0:
				print "***E Style Already Exist"
				return 1
			else:
				pass

		if borderPosition == None:
			if borderLineStyle != None or borderWeight != None or borderColor!= None:
				print "***E please provide border position"
				return 1

		#create style element

		styleElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Style"))
		styleElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"ID")] ="{0}{1}".format(styleId,tempstyleId)
		styleElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Name")] = "{0}".format(styleName)

		#to create Alignment.
		allignElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Alignment"))
		#hAlign="Left",vAlign="Top",wrapText=0
		allignElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Horizontal")] ="{0}".format(hAlign)
		allignElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Vertical")] ="{0}".format(vAlign)
		allignElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"WrapText")] ="{0}".format(wrapText)

		styleElem.append(allignElem)


		#to create borders
		#borderPosition=None,borderLineStyle="Continuous",borderWeight="1",borderColor="#000000"
		if borderPosition != None:
			bordersElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Borders"))
			if borderPosition == "Left":
				borderElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = borderPosition 
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem)
			elif borderPosition == "Right":
				borderElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = borderPosition 
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem)
			elif borderPosition == "Bottom":
				borderElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = borderPosition 
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem)
			elif borderPosition == "Top":
				borderElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = borderPosition 
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem)
			else :
				borderElem1 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Left"
				borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem1.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem1)
				borderElem2 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Right"
				borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem2.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem2)
				borderElem3 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Top"
				borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem3.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem3)
				borderElem4 = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Border"))
				borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Position")] = "Bottom"
				borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"LineStyle")] = borderLineStyle
				borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Weight")] = borderWeight
				borderElem4.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = borderColor
				bordersElem.append(borderElem4)
			styleElem.append(bordersElem)

		#add font to the Style
		#fontBold=0,fontItalic=0,fontSize=11,fontName="Calibri",fontColor=None
		if fontBold != None or fontItalic != None or fontUnderLine != None or fontSize != None or fontName != None or fontColor != None:
			fontElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Font"))
			if fontBold != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Bold")] = "{0}".format(fontBold)
			if fontItalic != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Italic")] = "{0}".format(fontItalic)
			if fontUnderLine != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"UnderLine")] = "{0}".format(fontUnderLine)
			if fontSize != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Size")] = "{0}".format(fontSize)
			if fontName != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"FontName")] = "{0}".format(fontName)
			if fontColor != None:
				fontElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = "{0}".format(fontColor)
			styleElem.append(fontElem)


		#for interior
		if BGColor != None:
			interiorElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Interior"))
			interiorElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Color")] = BGColor 
			interiorElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Pattern")] = "Solid" 
			styleElem.append(interiorElem)

		stylesElem.append(styleElem)
		
	def createWorkSheet(self,workSheetName,force=None):
		'''This method is used to create new worksheet.
			e.g. createWorkSheet(<WS Name>,0)
			e.g. createWorkSheet(<WS Name>,1)
		'''
		if force == None:
			#check that the worksheet already exist or not.Worksheet ss:Name="Revision History
			#print "//ss:Workbook/ss:Worksheet[@ss:Name='{0}']".format(workSheetName)
			
			checkWorksheet = self._xmlssDoc.xpath("//ss:Worksheet[@ss:Name='{0}']".format(workSheetName),namespaces=self._xmlssXPathNameSpaceMap)			
			#print etree.tostring(self._xmlssDoc)
			#print len(checkWorksheet) 
			if len(checkWorksheet) != 0:
				print "***E The worksheet with name {0} already exist".format(workSheetName)
			else:
				workbookElem = self._xmlssDoc.getroot()
				workSheetElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Worksheet"))
				workSheetElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Name")] = workSheetName
				
				#tableElem = etree.Element(self._xmlnsDefaultElem +"Table",nsmap=self._xmlssDefaultNSMap)
				tableElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Table"))
				workSheetElem.append(tableElem)
				workbookElem.append(workSheetElem)
				
		else:
			checkWorksheet = self._xmlssDoc.xpath("//ss:Worksheet[@ss:Name='{0}']".format(workSheetName),namespaces=self._xmlssXPathNameSpaceMap)
			self.deleteElem(checkWorksheet)
			workbookElem = self._xmlssDoc.getroot()
			workSheetElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Worksheet"))
			workSheetElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Name")] = workSheetName
			tableElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Table"))
			workSheetElem.append(tableElem)
			workbookElem.append(workSheetElem)
		
	def deleteElem(self,lstElementToDelete):
		'''This method is to delete a node in the xml file.'''
		for elems in lstElementToDelete:
			parentElem = elems.getparent()
			parentElem.remove(elems)
	
	def setCurrentWorksheet (self,workSheetName):
		'''This method will set the current worksheet on which further processing will be done.'''
		checkWorksheet = self._xmlssDoc.xpath("//ss:Worksheet[@ss:Name='{0}']".format(workSheetName),namespaces=self._xmlssXPathNameSpaceMap)
		if len(checkWorksheet) == 0 :
			print "***E Worksheet {0} doesn't exist in the workbook".format(workSheetName)
			self._xmlssWorkSheet = None
			self._xmlssCurrenWSTable = None
		else:
			self._xmlssWorkSheet = checkWorksheet[0]
			worksheetTableElem = self._xmlssWorkSheet.find("ss:Table",namespaces=self._xmlssXPathNameSpaceMap)
			print "hi this is table {0}".format(worksheetTableElem)
			self._xmlssCurrenWSTable= worksheetTableElem
			
	def listWorksheets (self):
		'''this method will return the list of worksheet names available in the current workbok'''
		lstWorkSheetElem = self._xmlssDoc.xpath("//ss:Worksheet",namespaces=self._xmlssXPathNameSpaceMap)
		lstWorkSheetName = []
		if len(lstWorkSheetElem) == 0 :
			print "***W No worksheet exist in the current workbook"
			return lstWorkSheetName
		else:
			for workSheetElem in lstWorkSheetElem:
				worksheetName = workSheetElem.get(etree.QName(self._xmlssNameSpaceMap["ss"],"Name"))
				print worksheetName
				lstWorkSheetName.append(worksheetName)
			return lstWorkSheetName

	def setColumnWidth(self,lstColumnWidths=[],lstColumnIndex=None):
		'''this method is used set the column width in the current worksheet'''

		if self._xmlssWorkSheet ==None :
			print "***E Please select current worksheet"
			return 1

		if len(lstColumnWidths) == 0:
			print "***E Please provide valid list of column widths"
			return 1
		#etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
#ataElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Type")] = "String"

		if lstColumnIndex == None:
			
			for colIndex in range(1,len(lstColumnWidths)+1):
				colElement = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Column"))
				colElement.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Width")] = "{0}".format(lstColumnWidths[colIndex -1 ])
				self._xmlssCurrenWSTable.insert(colIndex-1,colElement)

		else:
			lastsetColumnIndex = lstColumnIndex[-1]
			for colIndex in range(1,len(lstColumnWidths)+1):
				if colIndex <= len(lstColumnIndex):
					colElement = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Column"))
					colElement.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Width")] = "{0}".format(lstColumnWidths[colIndex -1 ])
					colElement.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Index")] = "{0}".format(lstColumnIndex[colIndex -1 ])
				else :
					lastsetColumnIndex += 1
					colElement = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Column"))
					colElement.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Width")] = "{0}".format(lstColumnWidths[colIndex -1 ])
					colElement.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Index")] = "{0}".format(lastsetColumnIndex)
				self._xmlssCurrenWSTable.insert(colIndex-1,colElement)
		
	
		
	def insertRow (self,lstData,styleName = None):
		'''this method is used to insert one dimensional list data in the current selected worksheet.'''
		if self._xmlssWorkSheet ==None :
			print "***E Please select current worksheet"
			return 1
			
		#Style Name check
		if styleName != None:
			lstStyleElems = self._xmlssDoc.xpath("ss:Workbook/ss:Styles/ss:Style[@ss:Name='{0}']".format(styleName),namespaces=self._xmlssXPathNameSpaceMap)
			if len(lstStyleElems) == 0: 
				print "***E Style {0} doesn't exist in the current workbook".format(styleName)
				return 1
			
		rowElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Row"))
		
		for dataValue in lstData:
			cellElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
			if styleName != None:
				cellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"StyleID")] = styleName
			dataElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
			dataElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Type")] = "String"
			dataElem.text = "{0}".format(dataValue)
			cellElem.append(dataElem)
			rowElem.append(cellElem)
		
		#finally insert new row to current worksheet table
		self._xmlssCurrenWSTable.append(rowElem)

	def insertMergedRow (self,lstData,styleName = None):
		'''This method is used to insert nested data till two dimensions and will merge the row based on the data.
			 the first value should not be nested. 
				e.g [[34,54],[45,65]] is not allowed
				but [34,[33,66,343,56],["dfsd",56,78]...] is allowed
		'''

		if self._xmlssWorkSheet ==None :
			print "***E Please select current worksheet"
			return 1
			
		#Style Name check
		if styleName != None:
			lstStyleElems = self._xmlssDoc.xpath("ss:Workbook/ss:Styles/ss:Style[@ss:Name='{0}']".format(styleName),namespaces=self._xmlssXPathNameSpaceMap)
			if len(lstStyleElems) == 0: 
				print "***E Style {0} doesn't exist in the current workbook".format(styleName)
				return 1

		if lstData != None and isinstance(lstData[0],list):
			print "***E please provice valid data to insert in the row."
			return 1
			

		#to check that the data is nested list or not.
		isDataNested = any(isinstance(i, list) for i in lstData)
		if isDataNested:
			dictMergedRows={}
			mergeRowIndex = 0
			firstRowElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Row"))
			firstRowFirstCell = None
			for colIndex in range(0,len(lstData)):
				print "***M colIndex {0}".format(colIndex)
				if isinstance(lstData[colIndex],list):
					lstNestedData = lstData[colIndex]
					#insert the first item of the entry of the list 
					cellElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
					if firstRowFirstCell == None:
						firstRowFirstCell = cellElem
					if styleName != None:
						cellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"StyleID")] = styleName

					dataElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
					dataElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Type")] = "String"
					dataElem.text = "{0}".format(lstNestedData[0])
					cellElem.append(dataElem)
					firstRowElem.append(cellElem)
					#create dictionary of the remanining items.
					lstNestedData = lstNestedData[1:]
					rowIndex = 1
					for actData in lstNestedData:
						if rowIndex in dictMergedRows:
							dictMergedRows[rowIndex].append(colIndex+1)
							dictMergedRows[rowIndex].append(actData)
						else:
							dictMergedRows[rowIndex] = []
							dictMergedRows[rowIndex].append(colIndex+1)
							dictMergedRows[rowIndex].append(actData)
						rowIndex+=1

					if mergeRowIndex < rowIndex:
						mergeRowIndex = rowIndex
				else:
					cellElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
					if firstRowFirstCell == None:
						firstRowFirstCell = cellElem
					if styleName != None:
						cellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"StyleID")] = styleName

					dataElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
					dataElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Type")] = "String"
					dataElem.text = "{0}".format(lstData[colIndex])
					cellElem.append(dataElem)
					firstRowElem.append(cellElem)
			#insert the first row 
			
			firstRowFirstCell.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"MergeDown")] = "{0}".format(mergeRowIndex)		
			self._xmlssCurrenWSTable.append(firstRowElem)			
			#now insert the remaining rows to the current worksheet table
			print dictMergedRows	
			for rows in range (1,mergeRowIndex):
				currentRow = dictMergedRows[rows]
				mergedRowElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Row"))
				for j in range (0,len(currentRow),2):
					cellIndex = currentRow[j]
					cellData = currentRow[j+1]
					cellElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Cell"))
					cellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Index")] = "{0}".format(cellIndex)				
					if styleName != None:
						cellElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"StyleID")] = styleName

					dataElem = etree.Element(etree.QName(self._xmlssNameSpaceMap["ss"],"Data"))
					dataElem.attrib[etree.QName(self._xmlssNameSpaceMap["ss"],"Type")] = "String"
					dataElem.text = "{0}".format(cellData)
					cellElem.append(dataElem)
					mergedRowElem.append(cellElem)
				self._xmlssCurrenWSTable.append(mergedRowElem)							
		else:
			self.insertRow(lstData)
		
		
		
		
	def writeXmlss(self,oFileName):
		''' This method is used to write XMLSS doument to particular file.'''		
		self._xmlssDoc.write(oFileName,pretty_print=True,xml_declaration="xml version='2.0' encoding='UTF-8'")
