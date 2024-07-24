Imports System
Imports System.IO
Imports System.Web
Imports System.Web.Mvc

Namespace YourNamespace
    Public Class UploadController
        Inherits Controller

        <HttpPost>
        Public Function UploadFile(ByVal file As HttpPostedFileBase) As ActionResult
            If file Is Nothing OrElse file.ContentLength = 0 Then
                Return New HttpStatusCodeResult(HttpStatusCode.BadRequest, "No file uploaded.")
            End If

            Try
                Dim fileName As String = Path.GetFileName(file.FileName)
                Dim filePath As String = Path.Combine(Server.MapPath("~/Uploads"), fileName)

                ' ファイルを指定されたパスに保存
                file.SaveAs(filePath)

                ' 成功メッセージを返す
                Return Content("File uploaded successfully")
            Catch ex As Exception
                ' エラーメッセージを返す
                Return New HttpStatusCodeResult(HttpStatusCode.InternalServerError, ex.Message)
            End Try
        End Function
    End Class
End Namespace