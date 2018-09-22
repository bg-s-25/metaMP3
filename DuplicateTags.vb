'
' MetaMP3 Tag Editor -> Class DUPLICATETAGS
' Author: Bobby Georgiou
' Date: 2014
'
Public Class DuplicateTags
    Dim MP3FilePath As String
    Dim MP3File As TagLib.File
    Dim MP3File2Path As String
    Dim MP3File2 As TagLib.File

    Private Sub DuplicateTags_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = Main.CheckedListBox1.SelectedItem.ToString
        CheckedListBox1.Items.AddRange(Main.CheckedListBox1.Items)
    End Sub

    Private Sub DuplicateTags_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MP3FilePath = Main.ListBox2.Items.Item(Main.CheckedListBox1.SelectedIndex).ToString
        MP3File = TagLib.File.Create(MP3FilePath)
        For Each itm As String In CheckedListBox1.CheckedItems
            MP3File2Path = Main.ListBox2.Items.Item(CheckedListBox1.Items.IndexOf(itm)).ToString
            MP3File2 = TagLib.File.Create(MP3File2Path)

            MP3File2.Tag.Title = MP3File.Tag.Title
            MP3File2.Tag.Artists = New String() {MP3File.Tag.JoinedArtists}
            MP3File2.Tag.Album = MP3File.Tag.Album
            MP3File2.Tag.AlbumArtists = New String() {MP3File.Tag.JoinedAlbumArtists}
            MP3File2.Tag.Year = MP3File.Tag.Year
            MP3File2.Tag.Track = MP3File.Tag.Track
            MP3File2.Tag.TrackCount = MP3File.Tag.TrackCount
            MP3File2.Tag.Genres = New String() {MP3File.Tag.JoinedGenres}
            MP3File2.Tag.Comment = MP3File.Tag.Comment
            MP3File2.Tag.Composers = New String() {MP3File.Tag.JoinedComposers}
            MP3File2.Tag.Disc = MP3File.Tag.Disc
            MP3File2.Tag.DiscCount = MP3File.Tag.DiscCount
            MP3File2.Tag.Conductor = MP3File.Tag.Conductor
            MP3File2.Tag.Lyrics = MP3File.Tag.Lyrics

            MP3File2.Tag.Pictures = New TagLib.IPicture() {MP3File.Tag.Pictures(0)}
            MP3File2.Tag.Pictures(0).Type = TagLib.PictureType.FrontCover

            MP3File2.Save()
        Next

        Me.Close()
    End Sub
End Class