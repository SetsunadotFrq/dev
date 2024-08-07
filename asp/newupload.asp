<%
On Error Resume Next

' サーバーAPIのURL
Dim apiUrl
apiUrl = "https://yourserver.com/api/upload"

' アップロードするファイルのパス
Dim filePath
filePath = Server.MapPath("yourfile.xlsx")

' ADODB.Streamオブジェクトを使ってファイルをバイナリモードで読み込む
Dim stream, byteData
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
Dim boundary, CRLF, preData, postData, totalData
boundary = "---------------------------" & Right(CStr(Timer() * 1000), 10)
CRLF = vbCrLf

preData = "--" & boundary & CRLF & _
    "Content-Disposition: form-data; name=""file""; filename=""yourfile.xlsx""" & CRLF & _
    "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" & CRLF & CRLF

postData = CRLF & "--" & boundary & "--" & CRLF

' プレデータとポストデータのバイト配列を作成
Dim preBytes, postBytes
preBytes = StrConv(preData, vbFromUnicode)
postBytes = StrConv(postData, vbFromUnicode)

' 総データのバイト長を計算
Dim totalLength
totalLength = LenB(preBytes) + LenB(byteData) + LenB(postBytes)

' ADODB.Streamオブジェクトを使ってデータを結合
Dim totalStream
Set totalStream = Server.CreateObject("ADODB.Stream")
totalStream.Type = 1 ' バイナリデータとして扱う
totalStream.Open
totalStream.Write preBytes
totalStream.Write byteData
totalStream.Write postBytes
totalStream.Position = 0
totalData = totalStream.Read()
totalStream.Close
Set totalStream = Nothing

' XMLHTTPオブジェクトを作成
Dim xmlHttp
Set xmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' HTTPリクエストを初期化
xmlHttp.Open "POST", apiUrl, False

' リクエストヘッダーを設定
xmlHttp.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
xmlHttp.setRequestHeader "Content-Length", totalLength

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