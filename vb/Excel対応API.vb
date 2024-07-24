Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Web.Http

Public Class FileUploadController
    Inherits ApiController

    <HttpPost>
    <Route("api/upload")>
    Public Async Function UploadFile() As Task(Of HttpResponseMessage)
        If Not Request.Content.IsMimeMultipartContent() Then
            Return Request.CreateResponse(HttpStatusCode.UnsupportedMediaType, "Invalid media type.")
        End If

        Dim root As String = HttpContext.Current.Server.MapPath("~/App_Data")
        Dim provider As New MultipartFormDataStreamProvider(root)

        Try
            ' multipart/form-data の内容を読み込む
            Await Request.Content.ReadAsMultipartAsync(provider)

            ' ファイルが含まれているか確認
            If provider.FileData.Count = 0 Then
                Return Request.CreateResponse(HttpStatusCode.BadRequest, "No file uploaded.")
            End If

            ' アップロードされたファイルの情報を取得
            Dim fileData = provider.FileData(0)
            Dim fileName As String = Path.GetFileName(fileData.Headers.ContentDisposition.FileName.Trim(""""))
            Dim filePath As String = Path.Combine(root, fileName)

            ' ファイルを指定したパスに保存
            File.Move(fileData.LocalFileName, filePath)

            Return Request.CreateResponse(HttpStatusCode.OK, New With {.FilePath = filePath})
        Catch ex As Exception
            Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message)
        End Try
    End Function
End Class