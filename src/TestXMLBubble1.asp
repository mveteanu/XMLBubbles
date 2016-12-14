<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
	Dim cn, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open "DSN=LIMSLactalis"
	Set rs = cn.Execute("SELECT ReportBlob FROM TBReportTemplate WHERE BlobID=1")
	
	Set XMLBubble = New clsXMLBubble
	XMLBubble.AddData "blob1", rs.Fields("ReportBlob").Value
	XMLBubble.SendToClient
	Set XMLBubble = Nothing
	
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
%>