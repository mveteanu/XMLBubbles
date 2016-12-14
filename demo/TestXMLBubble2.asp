<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
	Set XMLBubble = New clsXMLBubble
	With XMLBubble
		.LoadFromClient
		If .Error.errorCode <> 0 Then
			Response.Write  "Erori receptie XML : " & .Error.Reason
		Else
			Response.Write "Datele persoanei " & .Field("txt1") & " " & .Field("txt2") & " au fost salvate"
			If .FieldCount("fis1") >= 1 Then
				BytearrayToFile .Field("fis1"), "d:\t\upload.bin"
				Response.Write "...impreuna cu imaginea."
			End If
		End If
	End With
	Set XMLBubble = Nothing
%>