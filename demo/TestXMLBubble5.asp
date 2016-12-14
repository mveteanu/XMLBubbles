<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
' Exemplu de solicitare de informatii de catre o pagina .asp de
' la alta pagina .asp localizata undeva pe Internet...
' See also: TestXMLBubble1.htm
Dim XMLBubble
	
Set XMLBubble = New clsXMLBubble
With XMLBubble
	.LoadFromURL "http://localhost/BRLocal_Local/TestXML/TestXMLBubble1.asp"
	If .Error.errorCode <> 0 Then
		Response.Write "Erori in pachetul XML :" & .Error.Reason
	Else
		BytearrayToFile .Field("blob1"), "d:\t\fisier.xls"
		Response.Write "Fisierul client a fost salvat cu succes!"
	End If
End With
Set XMLBubble = Nothing
%>

