<%
' **********************************************************************************
' XMLBubbles - Biblioteca de clase si subrutine pentru asigurarea comunicarii intre
'              client si server (transferul text si binar) folosind pachete XML
'			   PS: Bubbles are not SOAP !!!
' Author     - Marian Veteanu
' Last date  - 19.aprilie.2002
' **********************************************************************************

' Clasa este folosita de clsXMLBubble pentru a intoarce
' status-ul trimiterii pachetului XML

Class clsStatus
	Private LocalCode
	Private LocalText

	Public Property Let Code(lCode)
		LocalCode = lCode
	End Property

	Public Default Property Get Code
		Code = LocalCode
	End Property
	
	Public Property Let Text(lText)
		LocalText = lText
	End Property

	Public Property Get Text
		Text = LocalText
	End Property
End Class


Class clsServerResponse
	Private LocalText
	Private LocalXML

	Public Property Let Text(lText)
		LocalText = lText
	End Property
	
	Public Default Property Get Text
		Text = LocalText
	End Property

	Public Property Let XML(lXML)
		LocalXML = lXML
	End Property
	
	Public Property Get XML
		XML = LocalXML
	End Property
End Class


Class clsXMLBubble
	Private XMLDom
	Private objStatus
	Private objServerResponse
	Private XMLHTTP
	Private FieldsPosDict
	Public bAsynchronous
	Public Error

	Private Sub Class_Initialize()
		bAsynchronous = false

		Set objStatus = New clsStatus
		Set objServerResponse = New clsServerResponse
		Set FieldsPosDict = CreateObject("Scripting.Dictionary")
		Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
		Set XMLDom = CreateObject("MSXML.DOMDocument")
		XMLDom.LoadXML "<?xml version='1.0' ?> <bubble/>"
		XMLDom.documentElement.setAttribute "xml:space", "preserve"
		XMLDom.documentElement.setAttribute "xmlns:dt", "urn:schemas-microsoft-com:datatypes"
	End Sub

	Private Sub Class_Terminate()
		Set objStatus = Nothing
		Set objServerResponse = Nothing
		If IsObject(Error) Then Set Error = Nothing
		Set XMLDom = Nothing
		Set XMLHTTP = Nothing
		Set FieldsPosDict = Nothing
	End Sub

	Private Function AddDataToNode(ParentDOMNode, strFieldID, strFieldContent)
		Dim DOMNode, NodeType, re
		
		re = false
		If IsArray(strFieldContent) and (VarType(strFieldContent)<>vbArray + vbByte) Then
			Set DOMNode = XMLDom.CreateElement(strFieldID)
			ParentDOMNode.AppendChild(DOMNode)
			For Each ArIt In strFieldContent
				Call AddDataToNode(DOMNode, "it", ArIt)
			Next
			Set DOMNode = Nothing
			re = true
		Else
			Select Case VarType(strFieldContent)
				Case vbInteger	NodeType = "i2"
				Case vbLong		NodeType = "i4"
				Case vbSingle	NodeType = "r4"
				Case vbDouble	NodeType = "r8"
				Case vbDate		NodeType = "dateTime"
				Case vbString	NodeType = "string"
				Case vbBoolean	NodeType = "boolean"
				Case vbByte		NodeType = "ui1"
				Case vbArray + vbByte  NodeType = "bin.base64"
				Case Else		NodeType = ""
			End Select
			If Len(NodeType)>0 Then
				Set DOMNode = XMLDom.CreateElement(strFieldID)
				DOMNode.DataType = NodeType
				DOMNode.NodeTypedValue = strFieldContent
				ParentDOMNode.AppendChild(DOMNode)
				Set DOMNode = Nothing
				re = true
			End If
		End If
		AddDataToNode = re
	End Function

	Public Function AddData(strFieldID, strFieldContent)
		AddData = AddDataToNode(XMLDom.DocumentElement, strFieldID, strFieldContent)
	End Function

	Public Sub SendTo(strURL)
		With XMLHTTP
			.Open "POST", strURL, bAsynchronous
			.setRequestHeader "ContentType", "text/xml"
			.Send XMLDom
		End With
	End Sub

	Public Property Let XML(strXML)
		XMLDom.LoadXML strXML
		Set Error = XMLDom.ParseError
	End Property

	Public Property Get XML
		XML = XMLDom.XML
	End Property

	Public Sub LoadFromClient
		XMLDom.Load(Request)
		Set Error = XMLDom.ParseError
	End Sub

	Public Sub LoadFromURL(strURL)
		XMLDom.Async = bAsynchronous
		XMLDom.Load(strURL)
		Set Error = XMLDom.ParseError
	End Sub

	Private Function FieldByNode(DOMNode)
		Dim DOMNodeChilds, ResultArray, i
		
		If DOMNode.ChildNodes.Length = 0 Then
			FieldByNode = DOMNode.nodeTypedValue
		ElseIf DOMNode.FirstChild.nodeType = 3 Then
			FieldByNode = DOMNode.nodeTypedValue
		Else
			DOMNodeChilds = DOMNode.ChildNodes.Length
			Redim ResultArray(DOMNodeChilds-1)
			For i = 0 To DOMNodeChilds - 1
				ResultArray(i) =  FieldByNode(DOMNode.ChildNodes(i))
			Next
			FieldByNode = ResultArray
		End If
	End Function

	Public Property Get Field(strFieldID)
		Dim DOMNode
		
		If FieldCount(strFieldID) = 0 Then Exit Property
		Set DOMNode = XMLDom.selectNodes("bubble/" & strFieldID)(FieldPos(strFieldID))
		Field = FieldByNode(DOMNode)
		Set DOMNode = Nothing
	End Property

	Public Function FieldCount(strFieldID)
		FieldCount = CLng(XMLDom.selectNodes("bubble/" & strFieldID).Length)
	End Function
	
	Public Function FieldPos(strFieldID)
		If FieldCount(strFieldID) = 0 Then Exit Function
		If not FieldsPosDict.Exists(strFieldID) Then Call FieldsPosDict.Add(strFieldID, 0)
		FieldPos = FieldsPosDict.Item(strFieldID)
	End Function
	
	Public Sub FieldNext(strFieldID)
		Dim CurentPos, LastPos
		
		LastPos = FieldCount(strFieldID) - 1
		If LastPos >=0 Then
			If not FieldsPosDict.Exists(strFieldID) Then Call FieldsPosDict.Add(strFieldID, 0)
			CurentPos = FieldsPosDict.Item(strFieldID)
			If CurentPos < LastPos Then FieldsPosDict.Item(strFieldID) = CurentPos + 1
		End If
	End Sub

	Public Sub FieldPrev(strFieldID)
		Dim CurentPos
		
		If FieldCount(strFieldID) = 0 Then Exit Sub
		If not FieldsPosDict.Exists(strFieldID) Then 
			Call FieldsPosDict.Add(strFieldID, 0)
		Else
			CurentPos = FieldsPosDict.Item(strFieldID)
			If CurentPos > 0 Then FieldsPosDict.Item(strFieldID) = CurentPos - 1
		End If
	End Sub

	Public Sub FieldFirst(strFieldID)
		If FieldCount(strFieldID) = 0 Then Exit Sub
		If not FieldsPosDict.Exists(strFieldID) Then 
			Call FieldsPosDict.Add(strFieldID, 0)
		Else
			FieldsPosDict.Item(strFieldID) = 0
		End If
	End Sub
	
	Public Sub FieldLast(strFieldID)
		Dim LastPos
		
		LastPos = FieldCount(strFieldID) - 1
		If LastPos >= 0 Then
			If not FieldsPosDict.Exists(strFieldID) Then 
				Call FieldsPosDict.Add(strFieldID, LastPos)
			Else
				FieldsPosDict.Item(strFieldID) = LastPos
			End If
		End If
	End Sub
	
	Public Property Get ReadyState
		ReadyState = XMLHTTP.readyState
	End Property

	Public Property Get Status
		With objStatus
			.Code = XMLHTTP.Status
			.Text = XMLHTTP.StatusText
		End With
		Set Status = objStatus
	End Property

	Public Property Get ServerResponse
		With objServerResponse
			.Text = XMLHTTP.ResponseText
			.XML  = XMLHTTP.ResponseXML.XML
		End With
		Set ServerResponse = objServerResponse
	End Property

	Public Sub SendToClient
		With Response
			.Buffer = true
			.ExpiresAbsolute = #1/1/1980#
			.AddHeader "cache-control", "no-store, must-revalidate, private" 
			.AddHeader "Pragma", "no-cache"
			.ContentType = "text/xml"
			.Write XMLDom.XML
		End With
	End Sub
End Class


Public Function FileToBytearray(strFilename)
	Dim ADOStream

	If Len(strFilename) > 0 Then
		Set ADOStream = Server.CreateObject("ADODB.Stream")
		With ADOStream
			.Type = adTypeBinary
			.Open
			.LoadFromFile(strFilename)
			FileToBytearray = .Read(adReadAll)
			.Close
		End With
		Set ADOStream = Nothing
	End If
End Function

Public Sub BytearrayToFile(arBytearray, strFilename)
	Dim ADOStream
		
	If (Len(strFilename)>0) and (VarType(arBytearray)=vbArray + vbByte) Then
		Set ADOStream = Server.CreateObject("ADODB.Stream")
		With ADOStream
			.Type = adTypeBinary
			.Open
			.Write arBytearray
			.SaveToFile strFilename, adSaveCreateOverWrite 
			.Close
		End With
		Set ADOStream = Nothing
	End If
End Sub

Sub XMLRPCServer
	Dim XMLBubble, XMLBubbleResp
	Dim MethodName, MethodParams, MethodResponse
	Dim Par, MethToEval
	
	Set XMLBubble = New clsXMLBubble
	Set XMLBubbleResp = New clsXMLBubble
	XMLBubble.LoadFromClient
	If XMLBubble.Error.errorCode <> 0 Then
		XMLBubbleResp.AddData "error", XMLBubble.Error.errorCode
		XMLBubbleResp.AddData "error", "XMLRPC - Bubble parsing error - " & XMLBubble.Error.Reason
	Else
		MethodName = XMLBubble.Field("method")
		If XMLBubble.FieldCount("params") > 0 Then 
			MethodParams = XMLBubble.Field("params")
			MethToEval = ""
			For Par = 0 To UBound(MethodParams)
				MethToEval = MethToEval & "MethodParams(" & Par & "),"
			Next
			If Len(MethToEval) > 0 Then MethToEval = Left(MethToEval, Len(MethToEval)-Len(","))
			MethToEval = MethodName & "(" & MethToEval & ")"
		Else
			MethToEval = MethodName
		End If
		On Error Resume Next
		MethodResponse = Eval(MethToEval)
		If Err.number <> 0 Then
			XMLBubbleResp.AddData "error", Err.number
			XMLBubbleResp.AddData "error", "XMLRPC - Remote function error - " & Err.Description 
			Err.Clear 
		Else
			XMLBubbleResp.AddData "result", MethodResponse
		End If
	End If
	XMLBubbleResp.SendToClient
	Set XMLBubble = Nothing
	Set XMLBubbleResp = Nothing
End Sub

Function XMLRPC(strURL, strFunctionName, arParameters)
	Dim XMLBubble, putParams
	Dim RemoteErrNr, RemoteErrDescr
	
	If IsEmpty(arParameters) Then
		putParams = false
	Else
		If not IsArray(arParameters) Then arParameters = Array(arParameters)
		putParams = true
	End If

	Set XMLBubble = New clsXMLBubble
	With XMLBubble
		.AddData "action", "RPC"
		.AddData "method", strFunctionName
		If putParams Then .AddData "params", arParameters
		.SendTo strURL
		If .Status = 200 Then
			.XML = .ServerResponse.XML
			If .FieldCount("error") > 0 Then
				RemoteErrNr = .Field("error")
				.FieldNext("error")
				RemoteErrDescr = .Field("error")
				Err.Raise RemoteErrNr,,RemoteErrDescr
			Else
				If .FieldCount("result") > 0 Then
					XMLRPC = .Field("result")
				Else
					Err.Raise 10001,, "XMLRPC - Remote function returned no result or an invalid result."
				End If
			End If
		Else
			Err.Raise 10000,, "XMLRPC - Error sending RPC request."
		End If
	End With
	Set XMLBubble = Nothing
End Function
%>