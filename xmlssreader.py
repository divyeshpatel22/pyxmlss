#!/user/bin/env python
import os
from lxml import etree
from xmlssbase import XmlssBase

class XmlssReader(XmlssBase):
	'''
	This is the Base xmlss reder class which contains the xmlss reading
	functionality related methods and properties.
	'''
	
	
	def readXmlss(self,fileName):
		'''to Read XMLSS file from particular file location.'''
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

	#Constructor to create instance of the XmlssBase

	# def __init__(debug=0):
		# _debug=debug

	
	
