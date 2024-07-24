<%
On Error Resume Next

' サーバーAPIのURL
Dim apiUrl
apiUrl = "https://yourserver.com/api/upload"

' アップロードするファイルのパス
Dim filePath
filePath = Server.MapPath("yourfile.xlsx")

' ファイルの内容をバイナリモードで読み込む
Dim objFSO, objFile, byteData, byteDataArray
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenAsTextStream(filePath, 1, -1) ' バイナリモードでファイルを開く
If Err.Number <> 0 Then
    Response.Write "Error opening file: " & Err.Description & "<br>"
    Response.End
End If

byteData = objFile.ReadAll
byteDataArray = StrConv(byteData, vbFromUnicode) ' バイナリ配列に変換
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing

If IsEmpty(byteDataArray) Or LenB(byteDataArray) = 0 Then
    Response.Write "Error reading file or file is empty.<br>"
    Response.End
End If

' バウンダリとコンテンツを作成
Dim boundary, CRLF, postData
boundary = "---------------------------" & Right(CStr(Timer() * 1000), 10)
CRLF = vbCrLf

postData = "--" & boundary & CRLF & _
    "Content-Disposition: form-data; name=""file""; filename=""yourfile.xlsx""" & CRLF & _
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & CRLF & CRLF & _
    byteDataArray & CRLF & "--" & boundary & "--" & CRLF

' XMLHTTPオブジェクトを作成
Dim xmlHttp
Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' HTTPリクエストを初期化
xmlHttp.Open "POST", apiUrl, False

' リクエストヘッダーを設定
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary

' ファイルの内容を送信
xmlHttp.Send postData

' エラーチェック
If Err.Number <> 0 Then
    Response.Write "Error sending request: " & Err.Description & "<br>"
    Response.End
End If

' レスポンスを取得
Dim status, responseText
status = xmlHttp.status
responseText = xmlHttp.responseText

' ステータスとレスポンスを表示
Response.Write "Status: " & status & "<br>"
Response.Write "Response from server: " & responseText & "<br>"

' オブジェクトを解放
Set xmlHttp = Nothing

' エラーチェック
If Err.Number <> 0 Then
    Response.Write "Error after request: " & Err.Description & "<br>"
End If

On Error GoTo 0
%>