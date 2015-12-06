#!/user/bin/env python

import xmlss
dd = xmlss.Xmlss()
#dd.readXmlss("temp.xml")
dd.createXmlss()
dd.createWorkSheet("MyFirstWorksheet")
dd.createWorkSheet("MyFirstWorksheet")
dd.createWorkSheet("MySecondWorksheet")
dd.createWorkSheet("MyFirstWorksheet",1)
lstWs = dd.listWorksheets()
print lstWs
dd.setCurrentWorksheet("MyFirstWorksheet")
lstData=[45,78,67]
dd.insertRow(lstData)
dd.setColumnWidth([150,250,45,30],[1,3])
lstData =  [ 34 , [33,66,343,56] , [ "sdf" ,56,78] ]  
dd.insertMergedRow(lstData)
dd.writeXmlss("temp.xml")
#default xmlss namespace map
#import os
#from lxml import etree
#from datetime import datetime
#from dateutil.tz import tzlocal
#from xmlssbase import XmlssBase

#parser = etree.XMLParser(remove_blank_text=True)
#fileName= "temp.xml"
#xmlssDoc = etree.parse(fileName,parser)
#root = xmlssDoc.getroot()
#_xmlnsDefault = "urn:schemas-microsoft-com:office:spreadsheet"
#_xmlnsO="urn:schemas-microsoft-com:office:office"
#_xmnsX="urn:schemas-microsoft-com:office:excel"
#_xmlnsSS= _xmlnsDefault
#_xmlnsHTML="http://www.w3.org/TR/REC-html40"
#_xmlssNameSpaceMap =  { None :_xmlnsDefault , "o":_xmlnsO,"x":_xmnsX, "ss":_xmlnsSS,"html":_xmlnsHTML}
#_xmlssXPathNameSpaceMap =  {"o":_xmlnsO,"x":_xmnsX, "ss":_xmlnsSS,"html":_xmlnsHTML}
#_xmlssDocPropNSMap = {None:_xmlnsO}

