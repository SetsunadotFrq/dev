Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Web.Http

Public Class FileUploadController
    Inherits ApiController

    <HttpPost>
    <Route("api/upload")>
    Public Async Function UploadFile() As Task(Of IHttpActionResult)
        If Not Request.Content.IsMimeMultipartContent() Then
            Return StatusCode(HttpStatusCode.UnsupportedMediaType)
        End If

        Try
            Dim root As String = HttpContext.Current.Server.MapPath("~/App_Data/uploads")
            Directory.CreateDirectory(root)

            Dim provider = New MultipartFormDataStreamProvider(root)
            Await Request.Content.ReadAsMultipartAsync(provider)

            For Each fileData In provider.FileData
                Dim fileName As String = fileData.Headers.ContentDisposition.FileName.Trim(""""c)
                Dim localFileName As String = fileData.LocalFileName
                Dim filePath As String = Path.Combine(root, fileName)

                File.Move(localFileName, filePath)
            Next

            Return Ok("File uploaded successfully")
        Catch ex As Exception
            Return InternalServerError(ex)
        End Try
    End Function
End Class
