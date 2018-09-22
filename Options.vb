'
' MetaMP3 Tag Editor -> Class OPTIONS
' Author: Bobby Georgiou
' Date: 2014
'
Imports System.IO

Public Class Options
    Private Sub Options_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If File.Exists(ARD & "\Options.ini") AndAlso ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultMP3") <> "" Then
            If My.Computer.FileSystem.DirectoryExists(ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultMP3")) Then
                TextBox1.Text = ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultMP3")
            Else
                TextBox1.Text = My.Computer.FileSystem.SpecialDirectories.MyMusic
            End If
        Else
            AddToIni(ARD & "\Options.ini", "Directories", "DefaultMP3", My.Computer.FileSystem.SpecialDirectories.MyMusic)
        End If

        If File.Exists(ARD & "\Options.ini") AndAlso ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultCovers") <> "" Then
            If My.Computer.FileSystem.DirectoryExists(ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultCovers")) Then
                TextBox1.Text = ReadFromIni(ARD & "\Options.ini", "Directories", "DefaultCovers")
            Else
                TextBox1.Text = My.Computer.FileSystem.SpecialDirectories.MyPictures
            End If
        Else
            AddToIni(ARD & "\Options.ini", "Directories", "DefaultCovers", My.Computer.FileSystem.SpecialDirectories.MyPictures)
        End If
    End Sub
End Class