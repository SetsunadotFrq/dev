<%
On Error Resume Next

' サーバーAPIのURL
Dim apiUrl
apiUrl = "https://yourserver.com/api/upload"

' アップロードするファイルのパス
Dim filePath
filePath = Server.MapPath("yourfile.xlsx")

' ADODB.Streamオブジェクトを使ってファイルをバイナリモードで読み込む
Dim stream, byteData, boundary, CRLF
Set stream = Server.CreateObject("ADODB.Stream")
stream.Type = 1 ' バイナリデータとして扱う
stream.Open
stream.LoadFromFile(filePath)

If Err.Number <> 0 Then
    Response.Write "Error opening file: " & Err.Description & "<br>"
    Response.End
End If

byteData = stream.Read()
stream.Close
Set stream = Nothing

If IsEmpty(byteData) Or LenB(byteData) = 0 Then
    Response.Write "Error reading file or file is empty.<br>"
    Response.End
End If

' バウンダリとコンテンツを作成
boundary = "---------------------------" & Right(CStr(Timer() * 1000), 10)
CRLF = vbCrLf

' プレデータとポストデータのバイナリ表現を作成
Dim preData, postData
preData = "--" & boundary & CRLF & _
    "Content-Disposition: form-data; name=""file""; filename=""yourfile.xlsx""" & CRLF & _
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & CRLF & CRLF

postData = CRLF & "--" & boundary & "--" & CRLF

' プレデータとポストデータをバイナリストリームで書き込む
Dim preDataBytes, postDataBytes
preDataBytes = StrConv(preData, vbFromUnicode)
postDataBytes = StrConv(postData, vbFromUnicode)

Dim totalStream
Set totalStream = Server.CreateObject("ADODB.Stream")
totalStream.Type = 1 ' バイナリデータとして扱う
totalStream.Open

' 各部分をバイナリ形式で書き込む
totalStream.Write preDataBytes
totalStream.Write byteData
totalStream.Write postDataBytes

' ストリームの位置を先頭に戻す
totalStream.Position = 0
Dim totalData
totalData = totalStream.Read()
totalStream.Close
Set totalStream = Nothing

' デバッグ出力
Response.Write "Boundary: " & boundary & "<br>"
Response.Write "PreData length: " & LenB(preDataBytes) & "<br>"
Response.Write "File Data length: " & LenB(byteData) & "<br>"
Response.Write "PostData length: " & LenB(postDataBytes) & "<br>"
Response.Write "TotalData length: " & LenB(totalData) & "<br>"

' XMLHTTPオブジェクトを作成
Dim xmlHttp
Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' HTTPリクエストを初期化
xmlHttp.Open "POST", apiUrl, False

' リクエストヘッダーを設定
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
xmlHttp.setRequestHeader "Content-Length", LenB(totalData)

' ファイルの内容を送信
xmlHttp.Send totalData

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