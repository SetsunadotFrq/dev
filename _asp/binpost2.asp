<%
Dim objXMLHTTP, objADOStream, objFSO, objFile, boundary
Dim strURL, strFilePath, strResponse, byteArray

boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
strURL = "http://yourserver/api/upload"
strFilePath = "C:\path\to\your\file.txt"

Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
objXMLHTTP.open "POST", strURL, False
objXMLHTTP.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary

' Create multipart/form-data body
Set objADOStream = Server.CreateObject("ADODB.Stream")
objADOStream.Type = 1 ' adTypeBinary
objADOStream.Open
objADOStream.WriteText "--" & boundary & vbCrLf
objADOStream.WriteText "Content-Disposition: form-data; name=""file""; filename=""" & _
    Mid(strFilePath, InStrRev(strFilePath, "\") + 1) & """" & vbCrLf
objADOStream.WriteText "Content-Type: application/octet-stream" & vbCrLf & vbCrLf

Set objFile = Server.CreateObject("ADODB.Stream")
objFile.Type = 1 ' adTypeBinary
objFile.Open
objFile.LoadFromFile strFilePath
objADOStream.Write objFile.Read
objFile.Close
Set objFile = Nothing

objADOStream.WriteText vbCrLf & "--" & boundary & "--" & vbCrLf
objADOStream.Position = 0

byteArray = objADOStream.Read
objADOStream.Close
Set objADOStream = Nothing

objXMLHTTP.send byteArray

strResponse = objXMLHTTP.responseText

Response.Write "Server response: " & strResponse

Set objXMLHTTP = Nothing
%>
