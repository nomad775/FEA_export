Attribute VB_Name = "xmlData"
Sub xmlDataTest()
    
    Dim x As MSXML2.DOMDocument60
    Dim sectionElement As IXMLDOMElement
    Dim dataElement As IXMLDOMElement
    
    Dim node As IXMLDOMNode
    
    Set x = New DOMDocument60
    
    'x.documentElement.TagName = "FEA report"
    x.appendChild x.createElement("FEAReportData")
    
    Set sectionElement = x.createElement("studyOptions")
    Set dataElement = x.createElement("name")
    dataElement.Text = "Static-1"
    
    x.documentElement.appendChild sectionElement
    sectionElement.appendChild dataElement
    
    x.Save "D:\reportXML.xml"
    
End Sub



Sub materialToXML(theXMLDoc As DOMDocument60, materialData)
    
    Dim parentElement As IXMLDOMElement
    
    Set parentElement = theXMLDoc.getElementsByTagName("material").Item(0)
    Set childNode = parentElement.getElementsByTagName("material").Item(0).CloneNode(False)
    
End Sub
