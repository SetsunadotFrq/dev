Imports Microsoft.AspNetCore.Http
Imports Microsoft.AspNetCore.Mvc
Imports System.IO
Imports System.Threading.Tasks

<Route("api/[controller]")>
<ApiController>
Public Class FileUploadController
    Inherits ControllerBase

    <HttpPost("upload")>
    Public Async Function UploadFile(<FromForm> file As IFormFile) As Task(Of IActionResult)
        If file Is Nothing OrElse file.Length = 0 Then
            Return BadRequest("No file uploaded.")
        End If

        ' ファイルの保存先パスを指定
        Dim filePath As String = Path.Combine(Path.GetTempPath(), file.FileName)

        ' ファイルをストリームにコピーして保存
        Using stream As New FileStream(filePath, FileMode.Create)
            Await file.CopyToAsync(stream)
        End Using

        Return Ok(New With { Key .FilePath = filePath })
    End Function
End Class