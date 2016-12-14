<%@ Language=VBScript %>
<!-- #include file="bubbles.asp" -->
<%
' Exemplu apel XMLRPC intre 2 pagini .asp
' See also: TestXMLBubble4.htm
Const RemotePage = "http://localhost/BRLocal_Local/TestXML/TestXMLBubble4.asp"

Sub CallRemoteFunction(whatfn)
	On Error Resume Next
	Select Case whatfn
		Case 1 Response.Write XMLRPC(RemotePage, "Saluta", "Marian")
		Case 2 Response.Write XMLRPC(RemotePage, "Saluta2", Array("Marian", "Veteanu"))
		Case 3 Call XMLRPC(RemotePage, "SaveToFileOnServer", Array(FileToBytearray("D:\t\SoapBubbles.bmp"), "d:\t\uploaded.bin"))
		Case 4 BytearrayToFile XMLRPC(RemotePage, "SaveToFileOnClient", 1), "d:\t\blob.xls"
	End Select
	If Err.number <> 0 Then
		Response.Write "<font size=-3 color=red>Error (" & Err.number & "): " & Err.Description & "</font>"
		Err.Clear 
	End If
End Sub
%>
<HTML>
<head>
<style>
TABLE, BODY
{
	font-family:verdana;
	font-size:12px;
	background-color:white;
}
TD
{
	border-top:1px solid black;
}
</style>
</head>
<BODY>

<H1>XMLRPC Demo 2</H1>

<p>Prin <b>XMLRPC</b> se poate apela direct de pe client o functie server (dintr-o pagina .asp) si intoarce rezultatul
acesteia pe client (intr-o alta pagina .asp sau script client). Orice eventuala eroare generata de functia server, precum si posibilele erori
de comunicatie client-server sunt prinse si apoi ridicate pe client in vederea tratarii lor.</p>
<p>Parametrii de intrare si rezultatul intors de functia server poate fi un tip primitiv de date sau
array de tipuri primitive sau de alte array-uri (imbricare pe oricate niveluri). Tip primitiv
de date inseamna unul din urmatoarele tipuri: integer, long, single, double, date, string, boolean, byte si 
array de bytes (de ex. continutul unui fisier sau blob)</p>

<table border=0 cellspacing=0 cellpadding=5 width=100% style='border-bottom:1px solid black;'>
<tr bgcolor=#f0f0f0>
<td width=150 valign=center align=left>
Rezultat functie:<br>
<%CallRemoteFunction 1%>
</td>
<td valign=top align=left>
Subrutina server: <font color=blue><b>Function Saluta(strName)</b></font><br><br>
Descriere: Primeste la intrare un string si returneaza un alt string.<br><br>
Apel client: <font color=red><b>Response.Write XMLRPC("<%=RemotePage%>", "Saluta", "Marian")</b></font><br>
</td>
</tr>

<tr>
<td width=150 valign=center align=left>
Rezultat functie:<br>
<%CallRemoteFunction 2%>
</td>
<td valign=top align=left>
Subrutina server: <font color=blue><b>Function Saluta2(nume, prenume)</b></font><br><br>
Descriere: Primeste la intrare 2 string-uri si returneaza un alt string. De pe client la apel se vor da cele 2 string-uri intr-un array. <br><br>
Apel client: <font color=red><b>Response.Write XMLRPC("<%=RemotePage%>", "Saluta2", Array("Marian", "Veteanu"))</b></font><br>
</td>
</tr>

<tr bgcolor=#f0f0f0>
<td width=150 valign=center align=left>
Salvare pe server<br>
<%CallRemoteFunction 3%>
</td>
<td valign=top align=left>
Subrutina server: <font color=blue><b>Function SaveToFileOnServer(arBytearray, strFilename)</b></font><br><br>
Descriere: Functia primeste 2 parametrii: un array de bytes (continul unui fisier) si un string. Functia server va salva pe server acest bytearray intr-un fisier cu numele specificat. Practic functia realizeaza un UPLOAD!<br><br>
Apel client: <font color=red><b>Call XMLRPC("<%=RemotePage%>", "SaveToFileOnServer", Array(FileToBytearray("d:\t\s2.gif"), "d:\t\uploaded.bin"))</b></font><br>
</td>
</tr>

<tr>
<td width=150 valign=center align=left>
Salvare pe client<br>
<%CallRemoteFunction 4%>
</td>
<td valign=top align=left>
Subrutina server: <font color=blue><b>Function SaveToFileOnClient(BlobID)</b></font><br><br>
Descriere: Functia primeste un parametru reprezentand un ID dintr-o tabela de unde va extrage un blob pe care-l va returna. Acest blob este salvat pe client intr-un fisier. Practic functia realizeaza un DOWNLOAD!<br><br>
Apel client: <font color=red><b>BytearrayToFile XMLRPC("<%=RemotePage%>", "SaveToFileOnClient", 1), "d:\t\blob.xls"</b></font><br>
</td>
</tr>

</table>
</BODY>
</HTML>
