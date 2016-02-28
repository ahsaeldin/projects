Imports System.Net

Public Class Form1
    Dim URLs As String()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim fileContents As String = My.Computer.FileSystem.ReadAllText("URLs.dat")
        URLs = fileContents.Split(vbLf)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim ahmed As New WebClient

        IO.Directory.CreateDirectory("Outs")


        For Each url As String In URLs
            Try
                Dim uriObi As New Uri(url)
                Dim whereCut As Integer = InStrRev(uriObi.AbsolutePath, "/")
                If whereCut <> -1 Then
                    Dim fileName As String = uriObi.AbsolutePath.Substring(whereCut)
                    ahmed.DownloadFile(url.Trim, Environment.CurrentDirectory + "\Outs\" + fileName)
                End If
            Catch ex As Exception

            End Try
        Next
    End Sub
End Class
