<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
' Exemplu de trimitere de informatii 
' de catre o pagina .asp de din alta pagina .asp
' See also: TestXMLBubble2.htm
Dim XMLBubble
	
Set XMLBubble = New clsXMLBubble
XMLBubble.AddData "txt1", "Marian"
XMLBubble.AddData "txt2", "Veteanu"
XMLBubble.AddData "fis1", FileToBytearray("D:\t\SoapBubbles.bmp")
XMLBubble.SendTo "http://localhost/BRLocal_Local/TestXML/TestXMLBubble2.asp"

If XMLBubble.Status = 200 Then 
	Response.Write XMLBubble.ServerResponse.Text
Else
	Response.Write "Eroare transmisie: " & XMLBubble.Status.Text & "<br><hr><br>" & vbCrLf
	Response.Write XMLBubble.ServerResponse.Text
End If
Set XMLBubble = Nothing
%>