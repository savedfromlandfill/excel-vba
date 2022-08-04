Attribute VB_Name = "sharepoint_update"
Option Explicit
' From: http://depressedpress.com/2014/04/05/accessing-sharepoint-lists-with-visual-basic-for-applications/


Sub SharePointConnection_Update()

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
  Dim SoapAction As String
  SoapAction = "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems"  ' CHANGE UPDATE/GET
  
  ' SOAP Envelope
  Dim SOAPEnvelope_Pre As String, SOAPEnvelope_Pst As String
  SOAPEnvelope_Pre = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
  "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
  "<soap:Body>"
  SOAPEnvelope_Pst = "</soap:Body>" & _
  "</soap:Envelope>"
   
' https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/bb249818(v%3Doffice.14)

' <Batch OnError="Continue" ListVersion="1" ViewName="270C0508-A54F-4387-8AD0-49686D685EB2">
'    <Method ID="1" Cmd="New">
'       <Field Name='ID'>New</Field>
'       <Field Name="Title">Value</Field>
'       <Field Name="Date_Column">2007-3-25</Field>
'       <Field Name="Date_Time_Column">
'           2006-1-11T09:15:30Z</Field>
'    </Method>
' </Batch>
 
  Dim FieldNameVar As String
  Dim ValueVar As String
  
  FieldNameVar = "Company"
  ValueVar = "Test Company 4321"

  Dim strBatchXml As String
  strBatchXml = "<Batch OnError='Continue'><Method ID='1' Cmd='Update'><Field Name='ID'>1690</Field><Field Name='" + _
  FieldNameVar + "'>" + _
  ValueVar + "</Field></Method></Batch>"

  ' Complete the packet
  Dim SoapMessage As String
  SoapMessage = SOAPEnvelope_Pre & _
  " <UpdateListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" & _
   "<listName>" & SOAPListName & "</listName><updates>" & _
  strBatchXml & _
  " </updates>" & _
  " </UpdateListItems>" & _
  SOAPEnvelope_Pst
  
  ' Create HTTP Object
  Dim Request As Object
  Set Request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
  ' Call the service to get the List
  Request.Open "POST", SOAPURL_List, False, CurUserName, CurPassword
  Request.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
  Request.setRequestHeader "SOAPAction", SoapAction
  Request.send (SoapMessage)
  
   
  Debug.Print Request.responseText

End Sub



