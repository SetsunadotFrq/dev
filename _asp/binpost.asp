<%
Dim objXMLHTTP, objADOStream, objFSO, objFile
Dim strURL, strFilePath, strResponse

strURL = "http://yourserver/api/upload"
strFilePath = "C:\path\to\your\file.txt"

Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.open "POST", strURL, False
objXMLHTTP.setRequestHeader "Content-Type", "multipart/form-data"

Set objADOStream = Server.CreateObject("ADODB.Stream")
objADOStream.Open
objADOStream.Type = 1 ' adTypeBinary
objADOStream.LoadFromFile strFilePath
objXMLHTTP.send objADOStream.Read

strResponse = objXMLHTTP.responseText

Response.Write "Server response: " & strResponse

Set objADOStream = Nothing
Set objXMLHTTP = Nothing
%>
