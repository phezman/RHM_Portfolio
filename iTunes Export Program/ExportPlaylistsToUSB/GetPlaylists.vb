Imports System.IO
Imports System.Threading
Imports System.Reflection

'' =======
'' Licence
'' =======
'' This program is free software: you can redistribute it and/or modify it under the terms
'' of the GNU General Public License as published by the Free Software Foundation, either
'' version 3 of the License, or (at your option) any later version.

'' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
'' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'' See the GNU General Public License for more details.

'' Please visit http://www.gnu.org/licenses/gpl-3.0-standalone.html to view the GNU GPLv3 licence.


Module GetPlaylists


    '    Dim IllegalChars() As String = {"<", ">", ":", "/", "\", "|", "?", "*"}


    '    Sub Main()

    '        Dim CopyDir As String ' = "C:\Users\R_Hew\Music\ExportPlaylistsToUSB\Scratch"
    '        Dim Title As String
    '        Dim Music As Boolean
    '        Dim FileList As Dictionary(Of String, String)
    '        Dim Playlists As New Dictionary(Of String, Dictionary(Of String, String))

    '        Writer = New StreamWriter(ScratchPath & "TransferLog.txt") With {.AutoFlush = True}


    '        Output("    ____        __      __                   __  ___            _          ")
    '        Output("   / __ \____ _/ /___  / /_  _____          /  |/  /___ _____ _(_)____     ")
    '        Output("  / /_/ / __ `/ / __ \/ __ \/ ___/         / /|_/ / __ `/ __ `/ / ___/     ")
    '        Output(" / _, _/ /_/ / / /_/ / / / (__  )         / /  / / /_/ / /_/ / / /__       ")
    '        Output("/_/ |_|\__,_/_/ .___/_/ /_/____/         /_/  /_/\__,_/\__, /_/\___/       ")
    '        Output("    ____  __ /_/        ___      __     __  ___      /____/_    _          ")
    '        Output("   / __ \/ /___ ___  __/ (_)____/ /_   /  |/  /___ ______/ /_  (_)___  ___ ")
    '        Output("  / /_/ / / __ `/ / / / / / ___/ __/  / /|_/ / __ `/ ___/ __ \/ / __ \/ _ \")
    '        Output(" / ____/ / /_/ / /_/ / / (__  ) /_   / /  / / /_/ / /__/ / / / / / / /  __/")
    '        Output("/_/   /_/\__,_/\__, /_/_/____/\__/  /_/  /_/\__,_/\___/_/ /_/_/_/ /_/\___/ ")
    '        Output("              /____/                           V0.0.1 - Back Everything Up!")
    '        Output()
    '        Output("SELECT FOLDER TO MIRROR ITUNES PLAYLISTS: ")

    '        'Select Directory to Copy into
    '        Dim ofd As New Windows.Forms.FolderBrowserDialog
    '        ofd.Description = "Choose Folder to Copy Into"
    '        ofd.ShowDialog() : CopyDir = ofd.SelectedPath & "\"
    '        If CopyDir = "" Then Exit Sub : Output(CopyDir) : Output()

    '        'Export Playlists from iTunes
    '        Dim CSVPath As String = ScratchPath & "iTunesXML.csv" 'ExtractCSV()
    '        If CSVPath Is Nothing Then Output("FAILED") : Exit Sub
    '        Output()

    '        'Build Collection of Playlists
    '        Output("PARSING CSV FOR PLAYLIST DATA")
    '        Dim Reader As New StreamReader(CSVPath)
    '        Reader.ReadLine() : Reader.ReadLine()
    '        Do Until Reader.EndOfStream
    '            Title = Reader.ReadLine
    '            If Title = """Exported""" Then Exit Do
    '            If Title = """Recently Added""" Then Music = True
    '            If Music And Not Title.StartsWith("Exported") Then : FileList = ReadBlock(Reader)
    '                Output("Playlist [" & FileList.Count & " Files] - " & Title, 1)
    '                For Each s In IllegalChars : Title = Title.Replace(s, " ") : Next
    '                Playlists.Add(Title.Replace("""", ""), FileList)
    '            Else : ReadBlock(Reader)
    '            End If
    '        Loop : Output()

    '        'Copy files to drive
    '        Output("MIRRORING PLAYLISTS IN " & CopyDir)
    '        CopyFiles(CopyDir, Playlists)

    '        Writer.Dispose()

    '    End Sub

    '    Function ExtractCSV() As String

    '        Dim iTunesXMLPath = "C:\Users\R_Hew\Music\iTunes\iTunes Music Library.xml"
    '        Dim iTunesExportScript = "TransformXML.vbs"
    '        Dim proc As New Process

    '        Output("EXTRACTING PLAYLISTS FROM ITUNES...")
    '        Output("Copying Playlist XML...  ", 1, False)
    '        File.Copy(iTunesXMLPath, ScratchPath & "iTunesXML.xml", True)
    '        Output("DONE")

    '        Output("Generating csv from Playlist File...  ", 1, False)
    '        proc.StartInfo.FileName = ScratchPath & "TransformXML.vbs"
    '        proc.StartInfo.Arguments = ScratchPath & "iTunesXML.xml"
    '        proc.Start()

    '        While Not proc.HasExited
    '            Thread.Sleep(100)
    '        End While

    '        File.Delete(ScratchPath & "iTunesXML.xml")

    '        If Not File.Exists(ScratchPath & "iTunesXML.csv") Then Output("FAILED!") : Return Nothing
    '        Output("COMPLETE") : Return ScratchPath & "iTunesXML.csv"

    '    End Function
    '    Function ReadBlock(Reader As StreamReader) As Dictionary(Of String, String)
    '        'Returns a collection of file paths for the given block

    '        Dim line As String
    '        Dim Song() As String
    '        ReadBlock = New Dictionary(Of String, String)

    '        Do Until Reader.EndOfStream

    '            'Read to end of playlist block
    '            line = Reader.ReadLine
    '            If line = "" Then Exit Do

    '            Try : Song = line.Split(vbTab)

    '                'Remove Illegal Characters in File Name
    '                For Each s In IllegalChars : Song(1) = Song(1).Replace(s, " ") : Next
    '                For Each s In IllegalChars : Song(2) = Song(2).Replace(s, " ") : Next

    '                'Add [Artist - Title, Path] to dictionary
    '                ReadBlock.Add(Song(2).Replace("""", "") & " - " & Song(1).Replace("""", ""), Song(4).Replace("""", ""))

    '            Catch : End Try

    '        Loop

    '    End Function

    Sub CopyFiles(CopyDir As String, Playlists As Dictionary(Of String, Dictionary(Of String, String)))

        Dim NewPath As String
        Dim SkipCount As Long
        Dim FailCount As Long
        Dim CopyCount As Long
        Dim TotalCopy As Long
        Dim TotalFail As Long
        Dim ErrorList As List(Of String)
        Dim PlaylistDir As String

        'For each playlist
        For i = 1 To Playlists.Count ' (only first 2 for now...)

            Output() : Output("SYNCHRONISING PLAYLIST """ & Playlists.Keys(i) & """", 1) : Output()

            'Check if the folder exists in copydrive/make it
            PlaylistDir = CopyDir & Playlists.Keys(i)
            If Not Directory.Exists(PlaylistDir) Then
                Directory.CreateDirectory(PlaylistDir)
                Output("CREATED FOLDER " & PlaylistDir, 2)
            Else : Output("FOLDER EXISTS " & PlaylistDir, 2)
            End If : Output() : PlaylistDir += "\"

            'Reset Counts
            SkipCount = 0
            CopyCount = 0
            FailCount = 0
            ErrorList = New List(Of String)

            'For each file in playlist
            For Each FileName In Playlists(Playlists.Keys(i))

                NewPath = PlaylistDir & FileName.Key & Path.GetExtension(FileName.Value)

                'If file already exists in drive then do nothing
                If File.Exists(NewPath) Then SkipCount += 1 : Continue For

                'Otherwise try to copy file
                Try
                    File.Copy(FileName.Value, NewPath)
                    Output("COPIED: " & FileName.Key, 2)
                    CopyCount += 1
                Catch ex As Exception
                    FailCount += 1
                    ErrorList.Add(FileName.Key & " [" & FileName.Value & "]")
                    Output(" --- ERROR: " & FileName.Key & "  (" & ex.Message & ")", 1)
                End Try

            Next

            'Summary per Playlist
            TotalCopy += CopyCount
            TotalFail += FailCount
            Output() : Output(CopyCount & " file(s) Copied!", 2)
            Output(SkipCount & " file(s) Required No Action", 2)

            If ErrorList.Count > 0 Then : Output(FailCount & " ERRORS COPYING FILES:", 1)
                For Each s In ErrorList : Output(s, 2) : Next
            End If

        Next

        'Final Reporting
        Output() : Output("FILE SYNCHRONISATION COMPLETE!")
        Output(TotalCopy & " file(s) copied, with " & TotalFail & " error(s)")
        Output() : Output("Press any key to exit") : Console.Read()

    End Sub




End Module
