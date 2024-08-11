<%
Dim objXMLHTTP
Dim url
Dim response

' ASP.NET APIのURL
url = "http://yourdomain.com/api/file"

' MSXML2.ServerXMLHTTPオブジェクトの作成
Set objXMLHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")

' POSTリクエストの送信
objXMLHTTP.Open "POST", url, False
objXMLHTTP.setRequestHeader "Content-Type", "application/json"
objXMLHTTP.Send ""

' APIからのレスポンスを取得
response = objXMLHTTP.responseText

' ファイルパスを表示
Response.Write "File path returned by ASP.NET API: " & response

' オブジェクトのクリーンアップ
Set objXMLHTTP = Nothing
%>
