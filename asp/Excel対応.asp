<%
' サーバーAPIのURL
Dim apiUrl
apiUrl = "https://yourserver.com/api/upload"

' アップロードするファイルのパス
Dim filePath
filePath = Server.MapPath("yourfile.xlsx")

' ファイルの内容をバイナリモードで読み込む
Dim objFSO, objFile, byteData
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenAsBinaryStream(filePath, 1)
byteData = objFile.ReadAll()
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

' バウンダリとコンテンツを作成
Dim boundary
boundary = "---------------------------" & Right(CStr(Timer() * 1000), 10)
Dim CRLF
CRLF = vbCrLf
Dim postData
postData = "--" & boundary & CRLF & _
    "Content-Disposition: form-data; name=""file""; filename=""yourfile.xlsx""" & CRLF & _
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & CRLF & CRLF & _
    byteData & CRLF & "--" & boundary & "--" & CRLF

' XMLHTTPオブジェクトを作成
Dim xmlHttp
Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' HTTPリクエストを初期化
xmlHttp.Open "POST", apiUrl, False

' リクエストヘッダーを設定
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
xmlHttp.setRequestHeader "Content-Length", LenB(postData)

' ファイルの内容を送信
xmlHttp.Send postData

' レスポンスを取得
Dim status, responseText
status = xmlHttp.status
responseText = xmlHttp.responseText

' ステータスとレスポンスを表示
Response.Write "Status: " & status & "<br>"
Response.Write "Response from server: " & responseText & "<br>"

' オブジェクトを解放
Set xmlHttp = Nothing
%>