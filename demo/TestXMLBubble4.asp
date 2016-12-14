<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
Call XMLRPCServer


Function Saluta(strName)
	Saluta = "Salut " & strName & " !"
End Function


Function Saluta2(nume, prenume)
	Saluta2 = "Salut " & prenume & " " & nume & " !"
End Function


Function Suma(a,b)
	Suma = a + b
End Function


Function Suma2(a)
	Dim s, i
	s = 0
	For Each i In a
		s = s + i
	Next
	Suma2 = s
End Function


Function SumaSimpa
	SumaSimpa = 2 + 3 
End Function


Function ConcatArrays(a,b)
	Dim re, i, i2
	
	Redim re(UBound(a)+UBound(b)+1)
	For i = 0 To UBound(a)
		re(i) = a(i)
	Next
	i2 = UBound(a)+1
	For i = 0 To UBound(b)
		re(i2+i) = b(i)
	Next
	ConcatArrays = re
End Function


Function SaveToFileOnServer(arBytearray, strFilename)
	Call BytearrayToFile(arBytearray, strFilename)
	SaveToFileOnServer = true
End Function


Function SaveToFileOnClient(BlobID)
	Dim cn, rs
	
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open "DSN=LIMSLactalis"
	Set rs = cn.Execute("SELECT ReportBlob FROM TBReportTemplate WHERE BlobID=" & BlobID)
	SaveToFileOnClient = rs.Fields("ReportBlob").Value
	rs.Close
	cn.Close
	Set rs = Nothing
	Set cn = Nothing
End Function


Function NullReturn
	NullReturn = Null
End Function


Sub Subrutina
	Dim a
	a = 2 + 3
End Sub
%>