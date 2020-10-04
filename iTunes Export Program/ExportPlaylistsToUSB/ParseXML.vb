Option Explicit On
Imports System.Collections
Imports System.Xml
Imports System.IO
Imports System.Uri

Module ParseXML

    Dim Writer As StreamWriter
    Public SongDictionary As New Dictionary(Of String, Song)
    Public PlaylistDictionary As New Dictionary(Of String, Playlist)
    Dim IllegalChars() As String = {"<", ">", ":", "/", "\", "|", "?", "*"}

    Sub Main()

        Dim ofd As New Windows.Forms.OpenFileDialog
        Dim User As String = Environment.UserName
        Dim GotPath As Boolean = False
        Dim iTunesPath As String

        Output("    ____        __      __                   __  ___            _          ")
        Output("   / __ \____ _/ /___  / /_  _____          /  |/  /___ _____ _(_)____     ")
        Output("  / /_/ / __ `/ / __ \/ __ \/ ___/         / /|_/ / __ `/ __ `/ / ___/     ")
        Output(" / _, _/ /_/ / / /_/ / / / (__  )         / /  / / /_/ / /_/ / / /__       ")
        Output("/_/ |_|\__,_/_/ .___/_/ /_/____/         /_/  /_/\__,_/\__, /_/\___/       ")
        Output("    ____  __ /_/        ___      __     __  ___      /____/_    _          ")
        Output("   / __ \/ /___ ___  __/ (_)____/ /_   /  |/  /___ ______/ /_  (_)___  ___ ")
        Output("  / /_/ / / __ `/ / / / / / ___/ __/  / /|_/ / __ `/ ___/ __ \/ / __ \/ _ \")
        Output(" / ____/ / /_/ / /_/ / / (__  ) /_   / /  / / /_/ / /__/ / / / / / / /  __/")
        Output("/_/   /_/\__,_/\__, /_/_/____/\__/  /_/  /_/\__,_/\___/_/ /_/_/_/ /_/\___/ ")
        Output("              /____/                           V0.0.1 - Back Everything Up!")
        Output()
        Output("FINDING iTUNES MUSIC LIBRARY...")

        'Find iTunes Install
        iTunesPath = "C:\Users\" & User & "\Music\iTunes\iTunes Music Library.xml"
        Do Until GotPath
            If File.Exists(iTunesPath) Then
                Output("Found Library: " & iTunesPath, 1)
                GotPath = True
            Else
                Output()
                Output("CAN'T FIND iTUNES MUSIC LIBRARY! SELECT 'iTunes Music Library.xml' wherever iTunes is installed")
                ofd.Title = "Open 'iTunes Music Library.xml'"
                ofd.InitialDirectory = "C:\"
                ofd.Filter = "xml Files|*.xml"
                If Not ofd.ShowDialog().ToString().Equals("OK") Then End
                iTunesPath = ofd.FileName
                Output("Found Library: " & iTunesPath, 1)
            End If
        Loop
        Output()

        'Create Log
        Dim ScratchPath = Path.GetDirectoryName(Path.GetDirectoryName(iTunesPath)) & "\"
        File.Delete(ScratchPath & "TransferLog.txt")
        Writer = New StreamWriter(ScratchPath & "TransferLog.txt") With {.AutoFlush = True}
        Writer.WriteLine(Now)
        Output("CREATED LOG AT: " & ScratchPath & "TransferLog.txt   (Just in case...)")
        Output()

        'Import Playlists
        Output("READING PLAYLISTS FROM iTUNES...")
        GetiTunesPlaylists(iTunesPath)

        Dim PlaylistForm As New SelectPlaylists
        PlaylistForm.ShowDialog()

        Writer.Dispose()

    End Sub

    Class Song

        Private IllegalChars() As String = {"<", ">", ":", "/", "\", "|", "?", "*", """"}

        Public Track_ID As String
        Public Name As String
        Public FileName As String
        Public Artist As String
        Public Album As String
        Public Year As Integer
        Public DateAdded As String
        Public BitRate As Integer
        Public FileSize As Long
        Public SongPath As String
        Public URLPath As Boolean

        Public Sub SetFileName()

            'Build Name from Track, Album & Artist
            FileName = Name & "__" & Artist & "__" & Album & Path.GetExtension(SongPath)

            'Remove Illegal Characters
            For Each s In IllegalChars : FileName = FileName.Replace(s, " ") : Next
            FileName = FileName.Replace("&#38;", "&")

        End Sub

        Public Playlists As New List(Of String)

    End Class
    Class Playlist

        Public Playlist_ID As String
        Public Name As String
        Public NumberTracks As Long
        Public Selected As Boolean
        Public PlaylistSize As Long

        Public Tracks As List(Of Song)

        'Calculate total size of playlist
        Public Sub EnumerateSize()
            PlaylistSize = 0
            NumberTracks = 0
            For Each s In Tracks
                NumberTracks += 1
                PlaylistSize += s.FileSize
            Next
        End Sub

        'Mark Each Song as belonging to the playlist
        Public Sub MarkSongs()
            For Each s In Tracks
                If Not s.Playlists.Contains(Playlist_ID) Then
                    s.Playlists.Add(Playlist_ID)
                End If
            Next
        End Sub

    End Class

    Function GetiTunesPlaylists(XMLPath As String)
        
        Dim tmpSong As New Song
        Dim tmpPlaylist As New Playlist
        Dim Reader As New StreamReader(XMLPath)
        Dim FilterList As New List(Of String)
        Dim RemoveList As New List(Of String)
        Dim Line As String = Reader.ReadLine
        Dim sName, sSongs, sSize As String
        Dim SizeMb, SizeGb, MaxNameLen, MaxSongs, NumTracks As Double
        Dim Music As Boolean = False

        'Get to beginning of tracks section
        If Not Line.Contains("<key>Tracks</key>") Then
            Line = Reader.ReadLine
            If Reader.EndOfStream Then
                Output("ERROR - XML File is not an iTunes Library File")
                Return False
            End If
        End If
        Line = Reader.ReadLine()

        'Read Songs into Dictionary
        Do Until Line.Contains("</dict>")

            tmpSong = New Song

            Do Until Line.Contains("</dict>")

                Line = Reader.ReadLine

                Select Case True
                        Case Line.Contains("<key>Track ID")
                            tmpSong.Track_ID = "ID_" & RemoveTags(Line)
                        Case Line.Contains("<key>Year")
                            tmpSong.Year = RemoveTags(Line)
                        Case Line.Contains("<key>Date Added")
                            tmpSong.DateAdded = RemoveTags(Line)
                        Case Line.Contains("<key>Bit Rate")
                            tmpSong.BitRate = RemoveTags(Line)
                        Case Line.Contains("<key>Size")
                            tmpSong.FileSize = RemoveTags(Line)
                        Case Line.Contains("<key>Name")
                            tmpSong.Name = RemoveTags(Line)
                        Case Line.Contains("<key>Album")
                            tmpSong.Album = RemoveTags(Line)
                        Case Line.Contains("<key>Artist")
                            tmpSong.Artist = RemoveTags(Line)
                        Case Line.Contains("<key>Location")
                            tmpSong.SongPath = RemoveTags(Line)
                            tmpSong.SongPath = UTF8toString(tmpSong.SongPath)
                        Case Line.Contains("Track Type") And Line.Contains("URL")
                        tmpSong.URLPath = True

                End Select

            Loop

            tmpSong.SetFileName()
            SongDictionary.Add(tmpSong.Track_ID, tmpSong)

            Line = Reader.ReadLine

        Loop

        'Get to beginning of Playlist section
        If Not Line.Contains("<key>Playlists</key>") Then
            Line = Reader.ReadLine
            If Reader.EndOfStream Then
                Output("ERROR - Failed to find Playlists in XML File")
                Return False
            End If
        End If
        Line = Reader.ReadLine()

        'Read Playlists into Dictionary
        Do Until Line.Contains("</array>")

            tmpPlaylist = New Playlist
            tmpPlaylist.Tracks = New List(Of Song)

            Do Until Line.Contains("</dict>")

                Line = Reader.ReadLine

                Select Case True
                    Case Line.Contains("<key>Playlist ID")
                        tmpPlaylist.Playlist_ID = "ID_" & RemoveTags(Line)
                    Case Line.Contains("<key>Name")
                        tmpPlaylist.Name = RemoveTags(Line).Replace("&#38;", "&")
                        For Each s In IllegalChars : tmpPlaylist.Name = tmpPlaylist.Name.Replace(s, " ") : Next
                    Case Line.Contains("<key>Playlist Items")
                        Line = Reader.ReadLine()
                        Line = Reader.ReadLine()
                        Do Until Line.Contains("</array>")
                            Line = Reader.ReadLine()
                            tmpPlaylist.Tracks.Add(SongDictionary("ID_" & RemoveTags(Line)))
                            Line = Reader.ReadLine()
                            Line = Reader.ReadLine()
                        Loop
                        Line = Reader.ReadLine()
                End Select

            Loop

            'Calculate number of Tracks
            tmpPlaylist.NumberTracks = tmpPlaylist.Tracks.Count

            'Add Playlist to Dictionary (By ID)
            PlaylistDictionary.Add(tmpPlaylist.Playlist_ID, tmpPlaylist)
            Line = Reader.ReadLine

        Loop

        'Filter iTunes Generated/Empty Playlists
        FilterList.Add("Library")
        FilterList.Add("Downloaded")
        For Each s In PlaylistDictionary.Keys
            If PlaylistDictionary(s).NumberTracks = 0 Then RemoveList.Add(s)
            If FilterList.Contains(PlaylistDictionary(s).Name) Then RemoveList.Add(s)
            If PlaylistDictionary(s).Name.Length > MaxNameLen Then MaxNameLen = PlaylistDictionary(s).Name.Length
            If PlaylistDictionary(s).NumberTracks > MaxSongs Then MaxSongs = PlaylistDictionary(s).NumberTracks
        Next
        For Each s In RemoveList
            PlaylistDictionary.Remove(s)
        Next

        Reader.Dispose()

        'Show Playlists in Console
        Output() : Output("FOUND " & PlaylistDictionary.Keys.Count & " PLAYLISTS!")
        For Each PlaylistID In PlaylistDictionary.Keys

            'Calculate Size of Playlist
            PlaylistDictionary(PlaylistID).EnumerateSize()

            'Mark Songs as Belonging to Playlist 
            PlaylistDictionary(PlaylistID).MarkSongs()

            'Name
            sName = PlaylistDictionary(PlaylistID).Name.PadRight(MaxNameLen + 5)

            'Number of Items in Playlist
            If sName.Contains("Recently Added") Then Music = True
            NumTracks = PlaylistDictionary(PlaylistID).NumberTracks
            sSongs = (NumTracks & IIf(Music, " Song", " Item") & IIf(NumTracks > 1, "s", "")).PadRight(CInt(Math.Log10(MaxSongs)) + 9)

            'Display Playlist Size
            SizeMb = Math.Round(PlaylistDictionary(PlaylistID).PlaylistSize / (1024 ^ 2), 0)
            SizeGb = Math.Round(PlaylistDictionary(PlaylistID).PlaylistSize / (1024 ^ 3), 2)
            sSize = IIf(SizeGb > 1, SizeGb & " Gb", SizeMb & " Mb")

            Output(sName & sSongs & sSize, 1)

        Next

        Return {SongDictionary, PlaylistDictionary}

    End Function

    Function RemoveTags(Line As String) As String

        Line = Line.Replace(vbTab, "")
        Dim tmpLine = Line

        tmpLine = tmpLine.Split("<//key><")(3)
        tmpLine = tmpLine.Split(">")(1)
        tmpLine = tmpLine.Replace("<*>", "")

        Return tmpLine

    End Function
    Function UTF8toString(Path As String) As String
        'Convert %## characters in string back from UTF-8

        Dim s As String
        'Dim sw0 As StringWriter
        'Dim sw1 As New StringWriter
        'Dim ec0 As New Text.UTF8Encoding
        'Dim n As Integer
        'Dim A() As Byte

        If Path.Contains("http:") Then Return Path
        If Path.Contains("https:") Then Return Path

        'Put Path into Windows Format
        Path = Path.Replace("/", "\")
        If Path.Contains(":") Then
            s = Right(Path.Split(":")(1), 1)
            Path = s & ":" & Path.Split(":")(2)
        End If

        Return Uri.UnescapeDataString(Path).Replace("&#38;", "&")

        ''Loop through characters in string
        'For i = 0 To Path.Length - 1

        '    '% indicates UTF8 Character needs to be converted
        '    If Path(i) = "%"c Then 'AndAlso Path.Length > i + 2 AndAlso (Path(i + 2) & Path(i + 2)) Like "##" Then

        '        'Extract 1st Hex value
        '        sw0 = New StringWriter
        '        sw0.Write(Path(i + 1))
        '        sw0.Write(Path(i + 2))
        '        i += 2

        '        'Extract 2nd Hex Value if needed
        '        If Path(i + 1) = "%"c Then 'AndAlso Path.Length > i + 2 AndAlso (Path(i + 2) & Path(i + 2)) Like "##" Then
        '            sw0.Write(Path(i + 2))
        '            sw0.Write(Path(i + 3))
        '            i += 3
        '        End If

        '        'Convert Hex to String
        '        s = sw0.ToString

        '        'Convert Hex String to Byte Array
        '        n = s.Length \ 2
        '        ReDim A(n - 1)
        '        For j = 0 To n - 1
        '            A(j) = Convert.ToByte(s.Substring(j * 2, 2), 16)
        '        Next

        '        'Encode byte Array as String
        '        sw1.Write(ec0.GetChars(A))

        '    Else

        '        'Otherwise Add character to string
        '        sw1.Write(Path(i))
        '    End If
        'Next

        ''Return Decoded string
        'Return sw1.ToString.Replace("&#38;", "&")

    End Function

    Sub Output(Optional Line As String = "", Optional Indent As Integer = 0, Optional NewLine As Boolean = True)

        Dim sIndent As String = ""
        For i = 0 To Indent : sIndent += "    " : Next : Line = sIndent + Line
        If Not Writer Is Nothing Then If NewLine Then Writer.WriteLine(Line) Else Writer.Write(Line)
        If NewLine Then Console.WriteLine(Line) Else Console.Write(Line)

    End Sub

End Module
