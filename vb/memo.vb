Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports System.Web.Http

Namespace YourNamespace
    Public Class UploadController
        Inherits ApiController

        <HttpPost>
        <Route("api/upload")>
        Public Async Function UploadFile() As Task(Of HttpResponseMessage)
            If Not Request.Content.IsMimeMultipartContent() Then
                Return Request.CreateResponse(HttpStatusCode.UnsupportedMediaType, "Unsupported media type")
            End If

            Try
                Dim streamProvider = New MultipartMemoryStreamProvider()
                Await Request.Content.ReadAsMultipartAsync(streamProvider)

                For Each file In streamProvider.Contents
                    Dim filename = file.Headers.ContentDisposition.FileName.Trim(""""c)
                    Dim buffer = Await file.ReadAsByteArrayAsync()
                    Dim filePath = Path.Combine("C:\Uploads", filename)

                    ' バッファを指定した文字コードで文字列にデコード
                    Dim content As String = Encoding.UTF8.GetString(buffer)

                    ' 文字列を指定した文字コードでバイト配列にエンコード
                    Dim encodedBytes As Byte() = Encoding.UTF8.GetBytes(content)

                    ' ファイルを書き込むためにファイルストリームを使用
                    Using fileStream As New FileStream(filePath, FileMode.Create, FileAccess.Write)
                        fileStream.Write(encodedBytes, 0, encodedBytes.Length)
                    End Using
                Next

                ' 成功メッセージを返す
                Return Request.CreateResponse(HttpStatusCode.OK, "File uploaded successfully")
            Catch ex As Exception
                ' エラーメッセージを返す
                Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message)
            End Try
        End Function
    End Class
End Namespace