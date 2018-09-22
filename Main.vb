'
' MetaMP3 Tag Editor -> Class MAIN
' Author: Bobby Georgiou
' Date: 2014
'
Imports System.IO
Imports System.Net

Public Class Main
    Dim MP3FilePath As String
    Dim MP3File As TagLib.File
    Dim MP3File2 As TagLib.File
    Dim ChangesMade As Boolean = False
    Dim InitialIndexChange As Boolean
    Dim CheckedCount As Integer
    Dim MultiChecked As Boolean
    Dim CheckedFilesPaths As New Collection()
    Dim ChangingInfo As Boolean 'AskForSave disregards changes made to controls if this is true
    Dim AllowCheckForChange As Boolean

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AllowCheckForChange = False
        Label17.Visible = False
        Label18.Visible = False
        DisableControls()
        ImportDirectoryData()
    End Sub

    Sub DisableControls()
        CheckedListBox1.ContextMenuStrip = Nothing
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        GroupBox3.Enabled = False
        ToolStripButton2.Enabled = False
        ToolStripButton3.Enabled = False
        ToolStripButton4.Enabled = False
        ToolStripButton5.Enabled = False
        ToolStripButton6.Enabled = False
        ToolStripButton7.Enabled = False
        ToolStripButton8.Enabled = False
        SaveToolStripMenuItem.Enabled = False
        RichTextBox1.Enabled = False
    End Sub

    Sub ClearControls(Optional ByVal AddSharedTags As Boolean = False)
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox7.Text = ""
        TextBox9.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox10.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox15.Text = ""
        TextBox6.Text = ""
        PictureBox1.Image = Nothing
        RichTextBox1.Text = ""
        TextBox11.Visible = False
        If AddSharedTags = True Then
            Dim FirstTitle As String = MF_TitleList.Item(0)
            For i = 1 To MF_TitleList.Count
                If MF_TitleList.Item(i - 1) <> FirstTitle Then Exit For
                If i = MF_TitleList.Count Then TextBox2.Text = FirstTitle
            Next
            Dim FirstArtist As String = MF_ArtistList.Item(0)
            For i = 1 To MF_ArtistList.Count
                If MF_ArtistList.Item(i - 1) <> FirstArtist Then Exit For
                If i = MF_ArtistList.Count Then TextBox3.Text = FirstArtist
            Next
            Dim FirstAlbum As String = MF_AlbumList.Item(0)
            For i = 1 To MF_AlbumList.Count
                If MF_AlbumList.Item(i - 1) <> FirstAlbum Then Exit For
                If i = MF_AlbumList.Count Then TextBox4.Text = FirstAlbum
            Next
            Dim FirstAlbumArtist As String = MF_AlbumArtistList.Item(0)
            For i = 1 To MF_AlbumArtistList.Count
                If MF_AlbumArtistList.Item(i - 1) <> FirstAlbumArtist Then Exit For
                If i = MF_AlbumArtistList.Count Then TextBox7.Text = FirstAlbumArtist
            Next
            Dim FirstYear As String = MF_YearList.Item(0)
            For i = 1 To MF_YearList.Count
                If MF_YearList.Item(i - 1) <> FirstYear Then Exit For
                If i = MF_YearList.Count Then TextBox9.Text = FirstYear
            Next
            Dim FirstTrack As String = MF_TrackList.Item(0)
            For i = 1 To MF_TrackList.Count
                If MF_TrackList.Item(i - 1) <> FirstTrack Then Exit For
                If i = MF_TrackList.Count Then TextBox12.Text = FirstTrack
            Next
            Dim FirstTrackOf As String = MF_TrackOfList.Item(0)
            For i = 1 To MF_TrackOfList.Count
                If MF_TrackOfList.Item(i - 1) <> FirstTrackOf Then Exit For
                If i = MF_TrackOfList.Count Then TextBox13.Text = FirstTrackOf
            Next
            Dim FirstDisc As String = MF_DiscList.Item(0)
            For i = 1 To MF_DiscList.Count
                If MF_DiscList.Item(i - 1) <> FirstDisc Then Exit For
                If i = MF_DiscList.Count Then TextBox14.Text = FirstDisc
            Next
            Dim FirstDiscOf As String = MF_DiscOfList.Item(0)
            For i = 1 To MF_DiscOfList.Count
                If MF_DiscOfList.Item(i - 1) <> FirstDiscOf Then Exit For
                If i = MF_DiscOfList.Count Then TextBox10.Text = FirstDiscOf
            Next
            Dim FirstGenre As String = MF_GenreList.Item(0)
            For i = 1 To MF_GenreList.Count
                If MF_GenreList.Item(i - 1) <> FirstGenre Then Exit For
                If i = MF_GenreList.Count Then TextBox5.Text = FirstGenre
            Next
            Dim FirstComposer As String = MF_ComposerList.Item(0)
            For i = 1 To MF_ComposerList.Count
                If MF_ComposerList.Item(i - 1) <> FirstComposer Then Exit For
                If i = MF_ComposerList.Count Then TextBox8.Text = FirstComposer
            Next
            Dim FirstConductor As String = MF_ConductorList.Item(0)
            For i = 1 To MF_ConductorList.Count
                If MF_ConductorList.Item(i - 1) <> FirstConductor Then Exit For
                If i = MF_ConductorList.Count Then TextBox15.Text = FirstConductor
            Next
            Dim FirstComment As String = MF_CommentList.Item(0)
            For i = 1 To MF_CommentList.Count
                If MF_CommentList.Item(i - 1) <> FirstComment Then Exit For
                If i = MF_CommentList.Count Then TextBox6.Text = FirstComment
            Next
            Dim FirstLyrics As String = MF_LyricsList.Item(0)
            For i = 1 To MF_LyricsList.Count
                If MF_LyricsList.Item(i - 1) <> FirstLyrics Then Exit For
                If i = MF_LyricsList.Count Then RichTextBox1.Text = FirstLyrics
            Next
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ToolStripButton1.PerformClick()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        InitialIndexChange = True
        Dim count As Integer = 0
        Dim AlreadyLoadedFile As Boolean = False
        For Each path As String In OpenFileDialog1.FileNames
            If ListBox2.Items.Contains(path) Then
                AlreadyLoadedFile = True
            Else
                CheckedListBox1.Items.Add(path)
                ListBox2.Items.Add(path)
            End If
        Next

        If AlreadyLoadedFile = True Then MsgBox("One or more of the files you attempted to add is already loaded for editing.", MsgBoxStyle.Information, "Import Notice")

        For i = 1 To CheckedListBox1.Items.Count
            CheckedListBox1.Items.Item(count) = Path.GetFileName(CheckedListBox1.Items.Item(count).ToString())
            count = count + 1
        Next
        CheckedListBox1.SelectedIndex = CheckedListBox1.Items.Count - 1

        LastImportDir = Path.GetDirectoryName(OpenFileDialog1.FileName)
    End Sub

    Sub AskForSave()
        If AllowCheckForChange = True And ChangesMade = True Then
            If MultiChecked = True Then
                Dim Result As MsgBoxResult = MsgBox("Save changes to the " & CheckedCount & " selected files?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Not Saved")
                If Result = MsgBoxResult.Yes Then SaveMP3File(False)
            Else
                Dim Result As MsgBoxResult = MsgBox("Save changes to '" & CurrentlyEditing & "' ?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Not Saved")
                If Result = MsgBoxResult.Yes Then SaveMP3File(False)
            End If

        End If
        ChangesMade = False
        AllowCheckForChange = False
    End Sub

    Dim MF_TitleList As New List(Of String)
    Dim MF_ArtistList As New List(Of String)
    Dim MF_AlbumArtistList As New List(Of String)
    Dim MF_GenreList As New List(Of String)
    Dim MF_AlbumList As New List(Of String)
    Dim MF_YearList As New List(Of String)
    Dim MF_TrackList As New List(Of String)
    Dim MF_TrackOfList As New List(Of String)
    Dim MF_DiscList As New List(Of String)
    Dim MF_DiscOfList As New List(Of String)
    Dim MF_ComposerList As New List(Of String)
    Dim MF_ConductorList As New List(Of String)
    Dim MF_CommentList As New List(Of String)
    Dim MF_LyricsList As New List(Of String)

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        If CheckedListBox1.Items.Count > 0 Then

            AskForSave()

            CheckedListBox1.ContextMenuStrip = ContextMenuStrip1
            GroupBox1.Enabled = True
            GroupBox2.Enabled = True
            GroupBox3.Enabled = True
            ToolStripButton2.Enabled = True
            ToolStripButton3.Enabled = True
            ToolStripButton4.Enabled = True
            ToolStripButton5.Enabled = True
            ToolStripButton6.Enabled = True
            ToolStripButton7.Enabled = True
            ToolStripButton8.Enabled = True
            SaveToolStripMenuItem.Enabled = True
            RichTextBox1.Enabled = True
            TextBox1.Text = ListBox2.Items.Item(CheckedListBox1.SelectedIndex).ToString()
            PictureBox1.ImageLocation = Nothing
            TextBox11.Text = ""
            ChangesMade = False
            CurrentlyEditing = CheckedListBox1.SelectedItem.ToString
        Else
            ClearDisable()
        End If

        'currentlyediting change to multiple

        If CheckedListBox1.CheckedIndices.Count > 1 Then
            AllowCheckForChange = False
            CheckedFilesPaths.Clear()
            For i = 1 To CheckedListBox1.CheckedIndices.Count
                Dim strnew As String = ListBox2.Items.Item(CheckedListBox1.CheckedIndices.IndexOf(CheckedListBox1.CheckedIndices.Item(i - 1))).ToString
                If CheckedFilesPaths.Contains(strnew) = False Then
                    CheckedFilesPaths.Add(strnew)
                End If
            Next
            Label1.Visible = False
            TextBox1.Visible = False
            CheckedCount = CheckedListBox1.CheckedIndices.Count
            Label17.Text = CheckedCount & " items selected"
            Label17.Visible = True
            Label18.Visible = True
            MultiChecked = True
            ToolStripButton8.Enabled = False
            ToolStripButton7.Enabled = False
            ToolStripButton6.Enabled = False
            If MF_TitleList.Count > 0 Then MF_TitleList.Clear()
            If MF_ArtistList.Count > 0 Then MF_ArtistList.Clear()
            If MF_AlbumList.Count > 0 Then MF_AlbumArtistList.Clear()
            If MF_AlbumArtistList.Count > 0 Then MF_AlbumList.Clear()
            If MF_TrackOfList.Count > 0 Then MF_TrackOfList.Clear()
            If MF_DiscList.Count > 0 Then MF_DiscList.Clear()
            If MF_DiscOfList.Count > 0 Then MF_DiscOfList.Clear()
            If MF_GenreList.Count > 0 Then MF_GenreList.Clear()
            If MF_TrackList.Count > 0 Then MF_TrackList.Clear()
            If MF_YearList.Count > 0 Then MF_YearList.Clear()
            If MF_ComposerList.Count > 0 Then MF_ComposerList.Clear()
            If MF_ConductorList.Count > 0 Then MF_ConductorList.Clear()
            If MF_CommentList.Count > 0 Then MF_CommentList.Clear()
            If MF_LyricsList.Count > 0 Then MF_LyricsList.Clear()
            For Each itm As String In CheckedListBox1.CheckedItems
                MP3FilePath = ListBox2.Items.Item(CheckedListBox1.Items.IndexOf(itm)).ToString
                MP3File = TagLib.File.Create(MP3FilePath)
                MF_TitleList.Add(MP3File.Tag.Title)
                MF_ArtistList.Add(MP3File.Tag.JoinedArtists)
                MF_AlbumList.Add(MP3File.Tag.Album)
                MF_AlbumArtistList.Add(MP3File.Tag.JoinedAlbumArtists)
                MF_YearList.Add(MP3File.Tag.Year)
                MF_TrackList.Add(MP3File.Tag.Track)
                MF_TrackOfList.Add(MP3File.Tag.TrackCount)
                MF_DiscList.Add(MP3File.Tag.Disc)
                MF_DiscOfList.Add(MP3File.Tag.DiscCount)
                MF_ComposerList.Add(MP3File.Tag.JoinedComposers)
                MF_ConductorList.Add(MP3File.Tag.Conductor)
                MF_GenreList.Add(MP3File.Tag.JoinedGenres)
                MF_CommentList.Add(MP3File.Tag.Comment)
                MF_LyricsList.Add(MP3File.Tag.Lyrics)
            Next
            ClearControls(True) 'add shared tags of all checked songs
        Else
            AllowCheckForChange = False
            MultiChecked = False
            CheckedCount = CheckedListBox1.CheckedIndices.Count
            Label17.Visible = False
            Label18.Visible = False
            Label1.Visible = True
            TextBox1.Visible = True
            ToolStripButton8.Enabled = True
            ToolStripButton7.Enabled = True
            ToolStripButton6.Enabled = True
        End If

        Label19.Visible = False
        Button3.Text = "Remove"
        ReadMetaData()

        ListBox1.Items.Clear()
        If CheckedFilesPaths.Count > 0 Then
            For i = 1 To CheckedFilesPaths.Count
                ListBox1.Items.Add(CheckedFilesPaths.Item(i))
            Next
        End If
    End Sub

    Public Sub ReadMetaData()
        ChangingInfo = True
        AllowCheckForChange = False
        If CheckedListBox1.Items.Count > 0 Then
            If CheckedListBox1.CheckedItems.Count <= 1 Then
                TextBox15.Text = ListBox2.Items.Item(CheckedListBox1.SelectedIndex).ToString()
                MP3FilePath = ListBox2.Items.Item(CheckedListBox1.SelectedIndex).ToString()
                MP3File = TagLib.File.Create(MP3FilePath)
                TextBox2.Text = MP3File.Tag.Title
                TextBox3.Text = MP3File.Tag.JoinedArtists
                TextBox4.Text = MP3File.Tag.Album
                TextBox7.Text = MP3File.Tag.JoinedAlbumArtists
                TextBox9.Text = MP3File.Tag.Year.ToString()
                TextBox12.Text = MP3File.Tag.Track.ToString()
                TextBox13.Text = MP3File.Tag.TrackCount.ToString()
                TextBox5.Text = MP3File.Tag.JoinedGenres
                TextBox6.Text = MP3File.Tag.Comment
                TextBox8.Text = MP3File.Tag.JoinedComposers
                TextBox14.Text = MP3File.Tag.Disc.ToString()
                TextBox10.Text = MP3File.Tag.DiscCount.ToString()
                TextBox15.Text = MP3File.Tag.Conductor
                RichTextBox1.Text = MP3File.Tag.Lyrics

                'display song art

                If MP3File.Tag.Pictures.Length >= 1 Then
                    Dim bin As Byte() = DirectCast(MP3File.Tag.Pictures(0).Data.Data, Byte())
                    PictureBox1.Image = Image.FromStream(New MemoryStream(bin)).GetThumbnailImage(200, 200, Nothing, System.IntPtr.Zero)
                    Button3.Enabled = True
                Else
                    PictureBox1.Image = Nothing
                    Button3.Enabled = False
                End If
            End If
        End If

        AllowCheckForChange = True
        ChangingInfo = False
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress, TextBox12.KeyPress, TextBox13.KeyPress, TextBox10.KeyPress, TextBox14.KeyPress
        If Char.IsNumber(e.KeyChar) = False And Microsoft.VisualBasic.AscW(e.KeyChar) <> 46 Then
            If Microsoft.VisualBasic.AscW(e.KeyChar) = 8 Then
                e.Handled = False
            ElseIf Microsoft.VisualBasic.AscW(e.KeyChar) = 13 Then
                Me.TextBox2.Select()
                e.Handled = False
            Else
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If Clipboard.ContainsImage = True Then
            Button2.Enabled = True
            PasteImageToolStripMenuItem.Enabled = True
        Else
            Button2.Enabled = False
            PasteImageToolStripMenuItem.Enabled = False
        End If

        If ChangesMade = False Then
            ToolStripButton5.Enabled = False
        Else
            ToolStripButton5.Enabled = True
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If LastArtworkDir <> "" Then OpenFileDialog2.InitialDirectory = LastArtworkDir
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If LastImportDir <> "" Then OpenFileDialog1.InitialDirectory = LastImportDir
        OpenFileDialog1.ShowDialog()
    End Sub

    Sub SaveMP3File(ByVal ShowSavedMsg As Boolean)
        If MultiChecked = False Then
            GeneralSaveProcedure()
        Else
            'multiple files to be saved
            For Each FilePath As String In CheckedFilesPaths
                MP3FilePath = FilePath
                MP3File = TagLib.File.Create(MP3FilePath)
                GeneralSaveProcedure()
            Next
        End If

        If ShowSavedMsg = True Then MsgBox("Saved.", MsgBoxStyle.Information, "Message")
    End Sub

    Sub GeneralSaveProcedure()
        MP3File.Tag.Title = TextBox2.Text
        MP3File.Tag.Artists = New String() {TextBox3.Text}
        MP3File.Tag.Album = TextBox4.Text
        MP3File.Tag.AlbumArtists = New String() {TextBox7.Text}
        If TextBox9.Text = "" Then
            MP3File.Tag.Year = Nothing
        Else
            MP3File.Tag.Year = TextBox9.Text
        End If
        If TextBox12.Text = "" Then
            MP3File.Tag.Track = 1
        Else
            MP3File.Tag.Track = TextBox12.Text
        End If
        If TextBox13.Text = "" Then
            MP3File.Tag.TrackCount = 0
        Else
            MP3File.Tag.TrackCount = TextBox13.Text
        End If
        If TextBox14.Text = "" Then
            MP3File.Tag.Disc = 0
        Else
            MP3File.Tag.Disc = TextBox14.Text
        End If
        If TextBox10.Text = "" Then
            MP3File.Tag.DiscCount = 0
        Else
            MP3File.Tag.DiscCount = TextBox10.Text
        End If
        MP3File.Tag.Genres = New String() {TextBox5.Text}
        MP3File.Tag.Composers = New String() {TextBox8.Text}
        MP3File.Tag.Conductor = TextBox15.Text
        MP3File.Tag.Comment = TextBox6.Text
        If Label19.Visible = True Then
            MP3File.Tag.Pictures(0).Data.Remove(0)
            MP3File.Tag.Pictures = Nothing
            PictureBox1.Image = Nothing
            Label19.Visible = False
        Else
            If PictureBox1.ImageLocation <> Nothing Then
                MP3File.Tag.Pictures = New TagLib.IPicture() {TagLib.Picture.CreateFromPath(PictureBox1.ImageLocation)}
                MP3File.Tag.Pictures(0).Type = TagLib.PictureType.FrontCover
            End If
        End If
        MP3File.Tag.Lyrics = RichTextBox1.Text
        MP3File.Save()
        Label19.Visible = False
        Button3.Text = "Remove"
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        SaveMP3File(True)
        ChangesMade = False
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        CheckedListBox1.Items.Remove(CheckedListBox1.SelectedItem)
        ListBox2.Items.Remove(CheckedListBox1.SelectedItem)
        ReadMetaData()
    End Sub

    Private Sub DeleteSelectedToolStripMenuItem_Click(sender As Object, e As EventArgs)
        CheckedListBox1.Items.Remove(CheckedListBox1.SelectedItem)
        ListBox2.Items.Remove(CheckedListBox1.SelectedItem)
        ReadMetaData()
        ChangesMade = False
    End Sub

    Sub ClearDisable()
        CheckedListBox1.Items.Clear()
        ListBox2.Items.Clear()
        CheckedFilesPaths.Clear()
        CheckedCount = 0
        Label17.Visible = False
        Label1.Visible = True
        TextBox1.Visible = True
        ClearControls()
        DisableControls()
        ChangesMade = False
        CurrentlyEditing = ""
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        ClearDisable()
    End Sub

    Private Sub UndoChangesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ReadMetaData()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        ReadMetaData()
    End Sub

    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        PictureBox1.ImageLocation = OpenFileDialog2.FileName
        TextBox11.Visible = True
        TextBox11.Text = Path.GetFileName(PictureBox1.ImageLocation)

        LastArtworkDir = Path.GetDirectoryName(OpenFileDialog2.FileName)
        ChangesMade = True
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        SaveMP3File(True)
        ChangesMade = False
    End Sub

    Private Sub RenameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenameToolStripMenuItem.Click
        ToolStripButton6.PerformClick()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        ToolStripButton7.PerformClick()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        RenameFile.ShowDialog()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        Dim RecycleFile As String = ListBox2.Items.Item(CheckedListBox1.SelectedIndex)
        ToolStripButton4.PerformClick()
        My.Computer.FileSystem.DeleteFile(RecycleFile, Microsoft.VisualBasic.FileIO.UIOption.AllDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        Shell("explorer /select," & ListBox2.Items.Item(CheckedListBox1.SelectedIndex), AppWinStyle.NormalFocus)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PictureBox1.Image = Clipboard.GetImage()
    End Sub

    Private Sub BrowseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BrowseToolStripMenuItem.Click
        Button4.PerformClick()
    End Sub

    Private Sub PasteImageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteImageToolStripMenuItem.Click
        Button2.PerformClick()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub CheckedListBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles CheckedListBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Shared ReadOnly SupportedExtensions As String() = {".mp3", ".m4a", ".mp4"} 'supported import formats [for drag & drop]

    Private Sub CheckedListBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles CheckedListBox1.DragDrop
        Dim Files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
        For Each FileName As String In Files
            Dim Extension As String = Path.GetExtension(FileName).ToLower
            If Array.IndexOf(SupportedExtensions, Extension) <> -1 Then
                CheckedListBox1.Items.Add(Path.GetFileName(FileName))
                ListBox2.Items.Add(FileName)
            End If
        Next
        ToolStripButton3.Enabled = True
    End Sub

    Private Sub SaveImageToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveImageToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        PictureBox1.Image.Save(SaveFileDialog1.FileName)
    End Sub

    Private Sub Main_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            ExportDirectoryData()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub Controls_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged, TextBox3.TextChanged, TextBox4.TextChanged,
        TextBox7.TextChanged, TextBox9.TextChanged, TextBox12.TextChanged, TextBox13.TextChanged, TextBox14.TextChanged, TextBox10.TextChanged,
        TextBox5.TextChanged, TextBox8.TextChanged, TextBox15.TextChanged, TextBox6.TextChanged, RichTextBox1.TextChanged
        'changes made set to true

        If InitialIndexChange = False And ChangingInfo = False And AllowCheckForChange = True Then
            ChangesMade = True
        Else
            InitialIndexChange = False
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Button3.Text = "Remove" Then
            Label19.Visible = True
            Button3.Text = "Undo"
        Else
            Label19.Visible = False
            Button3.Text = "Remove"
        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.ShowDialog()
    End Sub

    Private Sub DuplicateTagsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DuplicateTagsToolStripMenuItem.Click
        If CheckedListBox1.Items.Count >= 1 Then
            DuplicateTags.ShowDialog()
        Else
            MsgBox("No files have been loaded.")
        End If
    End Sub

    Private Sub ExportTagsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportTagsToolStripMenuItem.Click
        ExportTags.RadioButton1.Checked = True
        ExportTags.ShowDialog()
    End Sub

    Private Sub ExportTagsTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportTagsTextToolStripMenuItem.Click
        ExportTags.RadioButton2.Checked = True
        ExportTags.ShowDialog()
    End Sub

    Private Sub ExportTagsImagesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportTagsImagesToolStripMenuItem.Click
        ExportTags.RadioButton3.Checked = True
        ExportTags.ShowDialog()
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        Options.ShowDialog()
    End Sub

    Private Sub ImportTagsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportTagsToolStripMenuItem.Click
        ImportTags.ShowDialog()
    End Sub
End Class
