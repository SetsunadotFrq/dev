<%
' サーバーXMLHTTPオブジェクトを作成
Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' ターゲットURLを設定
url = "https://example.com/api/endpoint"

' リクエストを開く
xmlhttp.open "POST", url, false

' リクエストヘッダーを設定
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

' POSTデータを設定
postData = "param1=value1&param2=value2"

' リクエストを送信
xmlhttp.send postData

' レスポンスの取得
responseText = xmlhttp.responseText

' レスポンスを表示
Response.Write "Response: " & responseText

' オブジェクトの解放
Set xmlhttp = Nothing
%>