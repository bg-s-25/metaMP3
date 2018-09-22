'
' MetaMP3 Tag Editor -> Class RENAMEFILE
' Author: Bobby Georgiou
' Date: 2014
'
Public Class RenameFile

    Private Sub RenameFile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = System.IO.Path.GetFileNameWithoutExtension(Main.TextBox1.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            My.Computer.FileSystem.RenameFile(Main.TextBox1.Text, TextBox1.Text & ".mp3")
        Catch ex As System.ArgumentException
            MsgBox("You have entered illegal characters in this filename. Please check it again.", MsgBoxStyle.Exclamation, "Rename")
            Exit Sub
        Catch ex As System.IO.IOException
        End Try
        Application.Restart()
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then Button1.PerformClick()
    End Sub
End Class