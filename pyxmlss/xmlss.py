#!/user/bin/env python
from xmlsswriter import XmlssWriter
from xmlssreader import XmlssReader


class Xmlss(XmlssReader,XmlssWriter):
	'''
	This is the final xmlss class having all the reading and writing
	functionality related methods and properties.
	'''
	#to Read XMLSS file from particular file location.
	def readXmlssDefault(self,fileName = None):
		if fileName == None:
			#create basic structure of xmlss
			self.createXmlss()
		else:
			self.readXmlss(fileName)			