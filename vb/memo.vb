Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Threading.Tasks
Imports System.Web.Http

Namespace YourNamespace
    Public Class UploadController
        Inherits ApiController

        <HttpPost>
        <Route("api/upload")>
        Public Async Function UploadFile() As Task(Of HttpResponseMessage)
            If Not Request.Content.IsMimeMultipartContent() Then
                Return Request.CreateResponse(HttpStatusCode.UnsupportedMediaType)
            End If

            Try
                Dim streamProvider = New MultipartMemoryStreamProvider()
                Await Request.Content.ReadAsMultipartAsync(streamProvider)

                For Each file In streamProvider.Contents
                    Dim filename = file.Headers.ContentDisposition.FileName.Trim(""""c)
                    Dim buffer = Await file.ReadAsByteArrayAsync()
                    Dim filePath = Path.Combine("C:\Uploads", filename)

                    File.WriteAllBytes(filePath, buffer)
                Next

                Return Request.CreateResponse(HttpStatusCode.OK)
            Catch ex As Exception
                Return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message)
            End Try
        End Function
    End Class
End Namespace