'
' MetaMP3 Tag Editor -> Class EXPORTTAGS
' Author: Bobby Georgiou
' Date: 2014
'
Imports System.IO

Public Class ExportTags
    Dim Path As String
    Dim filelist As ObjectModel.ReadOnlyCollection(Of String)
    Dim MP3File As TagLib.File
    Dim Max As Integer
    Dim ProgressNow As Integer
    Dim DispActionComplete As Boolean
    Dim ExportMode As Boolean '0 = all, 1 = text, 2 = covers

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Select directory
        FolderBrowserDialog1.ShowDialog()
        Path = FolderBrowserDialog1.SelectedPath
        TextBox1.Text = Path

        filelist = My.Computer.FileSystem.GetFiles(Path, FileIO.SearchOption.SearchTopLevelOnly, "*.mp3")

        Label2.Text = filelist.Count & " MP3s loaded"
        Max = filelist.Count
        If filelist.Count > 0 Then
            Button2.Enabled = True
        Else
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RichTextBox1.Text = ""
        DisplayAction()
    End Sub

    Sub DisplayAction()
        ' Display/Export tag info
        Dim CFilename As String
        Dim CTitle As String
        Dim CJoinedArtists As String
        Dim CJoinedAlbumArtists As String
        Dim CAlbum As String
        Dim CJoinedGenres As String
        Dim CYear As String
        Dim CTrack As String
        Dim CTrackCount As String
        Dim CDisc As String
        Dim CDiscCount As String
        Dim CJoinedComposers As String
        Dim CConductor As String
        Dim CComment As String
        Dim CLyrics As String

        Dim ExportString As String
        Dim CurFile As Integer = 1
        ExportString = "Directory: " & Path
        ProgressBar1.Maximum = Max
        ProgressNow = 0

        filelist = My.Computer.FileSystem.GetFiles(Path, FileIO.SearchOption.SearchTopLevelOnly, "*.mp3")

        RichTextBox1.Text = "MP3 Directory: " & Path & vbCrLf & "Date/Time of Export: " & Date.Today & " " & TimeOfDay.ToLongTimeString & vbCrLf & "File Count: " & Max
        For Each file As String In filelist
            MP3File = TagLib.File.Create(file)
            CFilename = file
            CTitle = MP3File.Tag.Title
            CJoinedArtists = MP3File.Tag.JoinedArtists
            CJoinedAlbumArtists = MP3File.Tag.JoinedAlbumArtists
            CAlbum = MP3File.Tag.Album
            CJoinedGenres = MP3File.Tag.JoinedGenres
            CYear = MP3File.Tag.Year
            CTrack = MP3File.Tag.Track
            CTrackCount = MP3File.Tag.TrackCount
            CDisc = MP3File.Tag.Disc
            CDiscCount = MP3File.Tag.DiscCount
            CJoinedComposers = MP3File.Tag.JoinedComposers
            CConductor = MP3File.Tag.Conductor
            If MP3File.Tag.Comment <> "" Then CComment = MP3File.Tag.Comment.Replace(vbCrLf, "\n")
            If MP3File.Tag.Lyrics <> "" Then CLyrics = MP3File.Tag.Lyrics.Replace(vbCrLf, "\n")

            RichTextBox1.Text = RichTextBox1.Text & vbCrLf
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "[" & CurFile & "]"
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "File=" & file
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Filename=" & System.IO.Path.GetFileName(file)
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Title=" & CTitle
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Artist(s)=" & CJoinedArtists
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Album=" & CAlbum
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Album Artist(s)=" & CJoinedAlbumArtists
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Genre(s)=" & CJoinedGenres
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Year=" & CYear
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Track=" & CTrack
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "TrackCount=" & CTrackCount
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Disc=" & CDisc
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "DiscCount=" & CDiscCount
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Composer(s)=" & CJoinedComposers
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Conductor=" & CConductor
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Comment=" & CComment
            RichTextBox1.Text = RichTextBox1.Text & vbCrLf & "Lyrics=" & CLyrics
            CurFile += 1
            ProgressNow += 1
            ProgressBar1.Visible = True
            ProgressBar1.Value = ProgressNow
        Next
        ProgressBar1.Value = 0
        ProgressBar1.Visible = False
        DispActionComplete = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button2.PerformClick()
        Button2.Enabled = False
        Button4.Enabled = False
        While DispActionComplete
            If RadioButton1.Checked = True Then
                ExportMode = 0
            ElseIf RadioButton2.Checked = True Then
                ExportMode = 1
            ElseIf RadioButton3.Checked = True Then
                ExportMode = 2
            End If
            FolderBrowserDialog2.ShowDialog()

            If ExportMode = 0 Or ExportMode = 1 Then RichTextBox1.SaveFile(FolderBrowserDialog2.SelectedPath & "\TagText.ini", RichTextBoxStreamType.UnicodePlainText)
            If ExportMode = 0 Or ExportMode = 2 Then
                Max = filelist.Count
                ProgressBar1.Maximum = Max
                ProgressNow = 0
                For Each file As String In filelist
                    MP3File = TagLib.File.Create(file)
                    If MP3File.Tag.Pictures.Length >= 1 Then
                        Dim bin As Byte() = DirectCast(MP3File.Tag.Pictures(0).Data.Data, Byte())
                        PictureBox1.Image = Image.FromStream(New MemoryStream(bin))
                    Else
                        PictureBox1.Image = Nothing
                    End If
                    If PictureBox1.Image IsNot Nothing Then PictureBox1.Image.Save(FolderBrowserDialog2.SelectedPath & "\" & System.IO.Path.GetFileNameWithoutExtension(file) & ".jpg")
                    ProgressBar1.Visible = True
                    ProgressNow += 1
                    ProgressBar1.Value = ProgressNow
                Next
                ProgressBar1.Value = 0
                ProgressBar1.Visible = False
            End If
            Dispose()
            Close()
            Exit While
        End While
    End Sub

    Private Sub ExportTags_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dispose()
    End Sub
End Class