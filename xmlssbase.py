#!/user/bin/env python
import os
from lxml import etree
from datetime import datetime
from dateutil.tz import tzlocal

class XmlssBase:
	'''
	This is the Base xmlss class which contains the xmlss document(root) element.
	it also contains other methods to read xmlss file if file name is provided
	while creating instance of the XMLSSRead/XMLSSWrite class which derives this
	base class.
	'''
	#XMLSS Document will be stored in this variable
	_xmlssDoc=None

	#to enable pretty print wrhile writing xmlss
	_parser = etree.XMLParser(remove_blank_text=True)

	#default xmlss namespace map
	_xmlnsDefault = "urn:schemas-microsoft-com:office:spreadsheet"
	_xmlnsDefaultElem = "{%s}"%_xmlnsDefault
	_xmlnsO="urn:schemas-microsoft-com:office:office"
	_xmnsX="urn:schemas-microsoft-com:office:excel"
	_xmlnsSS= _xmlnsDefault
	_xmlnsHTML="http://www.w3.org/TR/REC-html40"
	_xmlssNameSpaceMap =  { None :_xmlnsDefault , "o":_xmlnsO,"x":_xmnsX, "ss":_xmlnsSS,"html":_xmlnsHTML}
	_xmlssNameSpaceMapMine =  { None :_xmlnsDefault , "o":_xmlnsO,"x":_xmnsX,"html":_xmlnsHTML}
	_xmlssXPathNameSpaceMap =  {"o":_xmlnsO,"x":_xmnsX, "ss":_xmlnsSS,"html":_xmlnsHTML}
	_xmlssDocPropNSMap = {None:_xmlnsO}
	_xmlssDefaultNSMap = {None:"urn:schemas-microsoft-com:office:spreadsheet"}

	#XMLSS current worksheet element name will be stored in this variable
	_xmlssWorkSheet = None
	_xmlssWorkSheetName = None
	_xmlssCurrenWSTable = None
	_debug = 0
	
	#Constructor to create instance of the XmlssBase

	# def __init__(xmlFileName=None,debug=0):
		# _debug=debug	