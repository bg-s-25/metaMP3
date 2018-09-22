'
' MetaMP3 Tag Editor -> Module VARS
' Author: Bobby Georgiou
' Date: 2014
'
Imports System.IO
Imports System.Threading

Module Vars
    Public ARD As String = My.Application.Info.DirectoryPath
    Public LastImportDir As String
    Public LastArtworkDir As String
    Public CurrentlyEditing As String

    Sub AddToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
        Dim ini As New Setting.IniFile(Filename)
        ini.WriteValue(Section, Key, Value)
    End Sub

    Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
        Dim ini As New Setting.IniFile(Filename)
        ReadFromIni = ini.ReadValue(Section, Key)
    End Function

    Public Sub ImportDirectoryData()
        'import saved direcotry data from ini
        If File.Exists(ARD & "\LastDirs.ini") Then
            Try
                LastImportDir = ReadFromIni(ARD & "\LastDirs.ini", "DirectoryData", "LastImportDir")
            Catch ex As Exception
                LastImportDir = "" 'uses default
            End Try
            Try
                LastArtworkDir = ReadFromIni(ARD & "\LastDirs.ini", "DirectoryData", "LastArtworkDir")
            Catch ex As Exception
                LastArtworkDir = "" 'uses default
            End Try
        End If
    End Sub

    Public Sub ExportDirectoryData()
        AddToIni(ARD & "\LastDirs.ini", "DirectoryData", "LastImportDir", LastImportDir)
        AddToIni(ARD & "\LastDirs.ini", "DirectoryData", "LastArtworkDir", LastArtworkDir)
    End Sub

    Public Sub StartThread(ByVal SubName As String, Optional ByVal Parameter1 As Object = 0, Optional ByVal Parameter2 As Object = 0)
        Dim trd As Thread
        If SubName = "Tools.DisplayAction" Then
            trd = New Thread(Sub() ExportTags.DisplayAction())
        End If
        trd.IsBackground = True
        trd.Start()
    End Sub
End Module
