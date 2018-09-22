'
' MetaMP3 Tag Editor -> Class IMPORTTAGS
' Author: Bobby Georgiou
' Date: 2014
'
Imports System.IO

Public Class ImportTags
    Dim FileCount As Integer
    Dim Path1 As String
    Dim Path2 As String
    Dim IFilename, ITitle, IArtist, IAlbum, IAlbumArtist, IGenre, IYear, ITrack, ITrackCount, IDisc, IDiscCount, IComposer, IConductor, IComment, ILyrics As String
    Dim MP3FilePath As String
    Dim MP3File As TagLib.File
    Dim Max As Integer
    Dim ProgressNow As Integer
    Dim ErrorCount As Integer

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FolderBrowserDialog1.ShowDialog()
        Path1 = FolderBrowserDialog1.SelectedPath
        TextBox1.Text = Path1

        Dim FN As String = "na"

        Do Until FN = ""
            FN = ReadFromIni(FolderBrowserDialog1.SelectedPath & "\TagText.ini", FileCount + 1, "Filename")
            If FN <> "" Then FileCount += 1
        Loop
        Label2.Text = FileCount & " Data Entries Loaded"
        If FileCount > 0 Then Button3.Enabled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FolderBrowserDialog2.ShowDialog()
        Path2 = FolderBrowserDialog2.SelectedPath
        TextBox2.Text = Path2
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ProgressNow = 0
        Max = FileCount
        ProgressBar1.Maximum = Max
        If Not File.Exists(Path1 & "\TagText.ini") Then
            MsgBox("Could not find valid INI file to import tags.", MsgBoxStyle.Exclamation, "MetaMP3 Import Tags")
            Exit Sub
        End If
        ApplyTagsAction()
    End Sub

    Sub Log(ByVal Text As String)
        RichTextBox1.Text += vbCrLf & "[" & TimeOfDay.ToLongTimeString & "] " & Text
    End Sub

    Sub ApplyTagsAction()
        Log("Started importing tags.")
        For f = 1 To FileCount
            If File.Exists(Path2 & "\" & ReadFromIni(Path1 & "\TagText.ini", f, "Filename")) Then
                MP3FilePath = Path2 & "\" & ReadFromIni(Path1 & "\TagText.ini", f, "Filename")
                MP3File = TagLib.File.Create(MP3FilePath)
                MP3File.Tag.Title = ReadFromIni(Path1 & "\TagText.ini", f, "Title")
                MP3File.Tag.Artists = New String() {ReadFromIni(Path1 & "\TagText.ini", f, "Artist(s)")}
                MP3File.Tag.Album = ReadFromIni(Path1 & "\TagText.ini", f, "Album")
                MP3File.Tag.AlbumArtists = New String() {ReadFromIni(Path1 & "\TagText.ini", f, "Album Artist(s)")}
                MP3File.Tag.Genres = New String() {ReadFromIni(Path1 & "\TagText.ini", f, "Genre(s)")}
                MP3File.Tag.Year = ReadFromIni(Path1 & "\TagText.ini", f, "Year")
                MP3File.Tag.Track = ReadFromIni(Path1 & "\TagText.ini", f, "Track")
                MP3File.Tag.TrackCount = ReadFromIni(Path1 & "\TagText.ini", f, "TrackCount")
                MP3File.Tag.Disc = ReadFromIni(Path1 & "\TagText.ini", f, "Disc")
                MP3File.Tag.DiscCount = ReadFromIni(Path1 & "\TagText.ini", f, "DiscCount")
                MP3File.Tag.Composers = New String() {ReadFromIni(Path1 & "\TagText.ini", f, "Composer(s)")}
                MP3File.Tag.Conductor = ReadFromIni(Path1 & "\TagText.ini", f, "Conductor")
                MP3File.Tag.Comment = ReadFromIni(Path1 & "\TagText.ini", f, "Comment").Replace("\n", vbCrLf)
                MP3File.Tag.Lyrics = ReadFromIni(Path1 & "\TagText.ini", f, "Lyrics").Replace("\n", vbCrLf)
            Else
                Log("File '" & Path2 & "\" & ReadFromIni(Path1 & "\TagText.ini", f, "Filename") & "' does not exist. MetaMP3 could not restore tags to this file.")
                ErrorCount += 1
            End If
            ProgressNow += 1
            ProgressBar1.Value += 1
        Next
        MP3FilePath = ""
        MP3File = Nothing
        ProgressBar1.Value = 0
        Log("Process complete with " & ErrorCount & " logged errors.")
    End Sub
End Class