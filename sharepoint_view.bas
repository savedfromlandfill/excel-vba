Attribute VB_Name = "sharepoint_view"
Option Explicit
' From: http://depressedpress.com/2014/04/05/accessing-sharepoint-lists-with-visual-basic-for-applications/


Sub SharePointConnection_View()

  ' Set credentials
  Dim CurUserName As String, CurPassword As String
  
  frmSharePointLogin.Show
  If frmSharePointLogin.Tag = "Cancel" Then
    Unload frmSharePointLogin
    Exit Sub
  End If
  CurUserName = frmSharePointLogin.txtLoginID.Text
  CurPassword = frmSharePointLogin.txtPassword.Text
  frmSharePointLogin.txtPassword.Text = ""
  Unload frmSharePointLogin


  ' Set SOAP/Webservice Parameters
  Dim SOAPURL_List As String, SOAPListName As String, SOAPViewName As String
  'SOAPURL_List = "http://yoursite/_vti_bin/Lists.asmx"
  SOAPURL_List = "http://ia.yoursite.local/_vti_bin/Lists.asmx"
  
  ' How to get List and View identifiers: http://depressedpress.com/2013/09/22/get-list-and-view-indentifiers-in-sharepoint/
  SOAPListName = "{xxxxx}"
  SOAPViewName = "{xxxxx}"
   
  ' SOAP Action URL
  Dim SOAPAction As String
  SOAPAction = "http://schemas.microsoft.com/sharepoint/soap/GetListItems"
  
  ' SOAP Envelope
  Dim SOAPEnvelope_Pre As String, SOAPEnvelope_Pst As String
  SOAPEnvelope_Pre = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
  "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
  "<soap:Body>"
  SOAPEnvelope_Pst = "</soap:Body>" & _
  "</soap:Envelope>"
   
  ' Complete the packet
  Dim SOAPMessage As String
  SOAPMessage = SOAPEnvelope_Pre & _
  " <GetListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" & _
  " <listName>" & SOAPListName & "</listName>" & _
  " <viewName>" & SOAPViewName & "</viewName>" & _
"<query><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>1690</Value></Eq></Where></Query></query>" & _
  " </GetListItems>" & _
  SOAPEnvelope_Pst
 '  " <rowLimit>50</rowLimit> " & _
' "<query><Query><Where><Lt><FieldRef Name='Company' /><Value Type='Text'>Test Company 1234</Value></Lt></Where></Query></query>" & _

  ' Create HTTP Object
  Dim Request As Object
  Set Request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  ' Call the service to get the List
  Request.Open "POST", SOAPURL_List, False, CurUserName, CurPassword
  Request.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
  Request.setRequestHeader "SOAPAction", SOAPAction
  Request.send (SOAPMessage)
  
  ' Init Vars
  Dim ReturnedRow
   
  ' Loop over returned rows to get keys for deletions

  For Each ReturnedRow In Request.responseXML.getElementsByTagName("z:row")
  ' Get the Current ID
    Debug.Print ReturnedRow.getAttribute("fieldname1") & ", " & _
                ReturnedRow.getAttribute("fieldname2") & "," & _
                ReturnedRow.getAttribute("fieldname3") & "," & _
                ReturnedRow.getAttribute("fieldname4") & "," & _
                ReturnedRow.getAttribute("fieldname5") & "," & _
                ReturnedRow.getAttribute("fieldname6")
  Next
   
  Debug.Print Request.responseText

End Sub


