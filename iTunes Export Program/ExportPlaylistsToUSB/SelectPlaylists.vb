Imports System
Imports System.IO
Imports System.Drawing
Imports System.Numerics
Imports System.Threading
Imports System.Windows.Forms

Public Class SelectPlaylists

    Dim IdenticalDir As Boolean
    Dim CopyDir As String
    Dim Cancelled As Boolean

    Public Sub New()

        Dim ListEntry As String

        ' This call is required by the designer.
        InitializeComponent()

        'Populate lb_Playlists
        For Each PlaylistKey In PlaylistDictionary.Keys

            ListEntry = PlaylistDictionary(PlaylistKey).Name.Replace("%", "#")
            ListEntry = ListEntry.PadRight(200) & "%" & PlaylistDictionary(PlaylistKey).Playlist_ID
            lb_Playlists.Items.Add(ListEntry)

        Next

        CopyDir = ""
        UpdateFreeSpace()
        Cancelled = True

    End Sub

    Private Sub cmd_ChangeDir_Click(sender As Object, e As EventArgs) Handles cmd_ChangeDir.Click

        'Select New Directory to Copy into
        Dim fbd As New Windows.Forms.FolderBrowserDialog
        fbd.Description = "Select Sync Folder"
        fbd.ShowDialog()

        CopyDir = fbd.SelectedPath & IIf(Strings.Right(fbd.SelectedPath, 1) = "\", "", "\")

        If CopyDir = "\" Then

            'Set Values if nothing chosen
            CopyDir = ""
            lbl_SpaceLeft.Text = "No Destination Selected!"

        End If

        cmd_SelectExisting_Click(Nothing, Nothing)

    End Sub

    Private Sub cmd_Cancel_Click(sender As Object, e As EventArgs) Handles cmd_Cancel.Click
        If Cancelled Then
            End
        Else
            Cancelled = True
        End If
    End Sub

    Private Sub cmd_SelectExisting_Click(sender As Object, e As EventArgs) Handles cmd_SelectExisting.Click

        For i = 0 To lb_Playlists.Items.Count - 1
            If Directory.Exists(CopyDir & PlaylistDictionary(lb_Playlists.Items(i).split("%")(1)).Name) Then
                lb_Playlists.SetSelected(i, True)
            End If
        Next

        UpdateFreeSpace()

    End Sub

    Private Sub lb_Playlists_Click(sender As Object, e As EventArgs) Handles lb_Playlists.Click

        UpdateFreeSpace()

    End Sub

    Private Function UpdateFreeSpace(Optional RenameFiles As Boolean = False) As Boolean
        'Returns true if there is space available for the transfer

        Dim sSelectedSize, sDirectorySpace, sRemoveSize As String
        Dim SelectedSizeGb, DirectorySpaceGb, RemoveSizeGb As Double
        Dim DirectorySpaceBytes As Long
        Dim SelectedSizeBytes As Long
        Dim OverlapSizeBytes As Long
        Dim RemovedSizeBytes As Long
        Dim ChangeCount, RemoveCount, RemoveDir As Integer
        Dim Found As Boolean

        Dim di As DriveInfo
        Dim diri As DirectoryInfo
        Dim fi As FileInfo

        Dim SelectedPlaylists As New List(Of Playlist)
        Dim tmpPlaylist As Playlist
        Dim PlaylistPath As String

        'Check destination
        If CopyDir = "" Then
            IdenticalDir = False
            lbl_SyncDestination.Text = "       No Destination Chosen!"
            cmd_CopyPlaylists.Text = "Sync Playlists!"
            cmd_CopyPlaylists.Enabled = False
            UpdateBar(1, 0)
            Return False
        Else
            cmd_CopyPlaylists.Enabled = True
            lbl_SyncDestination.Text = CopyDir
        End If

        'Find Total Size of Selected Playlists
        For Each sItem In lb_Playlists.SelectedItems
            tmpPlaylist = PlaylistDictionary(sItem.split("%")(1))
            SelectedSizeBytes += tmpPlaylist.PlaylistSize
            SelectedPlaylists.Add(tmpPlaylist)
        Next

        'Find Total Available Space in Selected Drive
        di = New DriveInfo(CopyDir)
        DirectorySpaceBytes = di.AvailableFreeSpace

        'Find Extra Directories
        For Each d In Directory.EnumerateDirectories(CopyDir)
            Found = False
            diri = New DirectoryInfo(d)
            If diri.Name = "System Volume Information" Then Continue For
            For Each pl In SelectedPlaylists
                If pl.Name = diri.Name Then
                    Found = True
                    Exit For
                End If
            Next
            If Not Found Then
                For Each f In diri.GetFiles : RemovedSizeBytes += f.Length : Next
                RemoveDir += 1
            End If
        Next

        'Check Playlists have been chosen to copy
        If SelectedPlaylists.Count = 0 Then
            IdenticalDir = False
            lbl_SpaceLeft.Text = "No Playlists Selected!"
            cmd_CopyPlaylists.Text = "Sync Playlists!"
            cmd_CopyPlaylists.Enabled = False
            UpdateBar(1, 0)
            Return False
        Else : cmd_CopyPlaylists.Enabled = True
        End If

        'Find Size of Existing Files in Folders
        For Each pl In SelectedPlaylists
            PlaylistPath = CopyDir & pl.Name & "\"
            If Directory.Exists(PlaylistPath) Then
                For Each s In pl.Tracks
                    If File.Exists(PlaylistPath & s.FileName) Then
                        fi = New FileInfo(PlaylistPath & s.FileName)
                        If s.FileSize = fi.Length Then
                            'If the file exists then add its size to the tally
                            OverlapSizeBytes += s.FileSize
                        Else
                            'If the filesize has changed, mark it for deletion
                            ChangeCount += 1
                            If RenameFiles Then Rename(PlaylistPath & s.FileName, PlaylistPath & "OLD_" & s.FileName)
                        End If
                    End If
                Next
            End If
        Next

        'Find Size of Files that will be removed from folder
        For Each pl In SelectedPlaylists
            PlaylistPath = CopyDir & pl.Name & "\"
            If Directory.Exists(PlaylistPath) Then
                For Each f In Directory.EnumerateFiles(PlaylistPath)
                    Found = False
                    For Each s In pl.Tracks
                        If s.FileName = Path.GetFileName(f) Then
                            Found = True
                            Exit For
                        End If
                    Next
                    If Not Found Then
                        RemoveCount += 1
                        fi = New FileInfo(f)
                        RemovedSizeBytes += fi.Length
                        If RenameFiles Then Rename(PlaylistPath & Path.GetFileName(f), PlaylistPath & "DELETE_" & Path.GetFileName(f))
                    End If
                Next
            End If
        Next

        'Update TextBox/Labels
        DirectorySpaceGb = Math.Round((DirectorySpaceBytes + RemovedSizeBytes) / 1024 ^ 3, 0)
        sDirectorySpace = IIf(DirectorySpaceGb > 1, DirectorySpaceGb & " Gb", Math.Round((DirectorySpaceBytes + RemovedSizeBytes) / (1024 ^ 2), 0) & " Mb")

        SelectedSizeGb = Math.Round((SelectedSizeBytes - OverlapSizeBytes) / 1024 ^ 3, 0)
        sSelectedSize = IIf(SelectedSizeGb > 1, SelectedSizeGb & " Gb", Math.Round((SelectedSizeBytes - OverlapSizeBytes) / (1024 ^ 2), 0) & " Mb")

        RemoveSizeGb = Math.Round((RemovedSizeBytes) / 1024 ^ 3, 0)
        sRemoveSize = IIf(RemoveSizeGb > 1, RemoveSizeGb & " Gb", Math.Round((RemovedSizeBytes) / (1024 ^ 2), 0) & " Mb")

        If DirectorySpaceBytes = 0 Then sDirectorySpace = "(No Destination Selected!)"

        lbl_SpaceLeft.Text = sSelectedSize & " to Copy (" & sDirectorySpace & " Available, " & sRemoveSize & " to Remove)"

        'Update Space Left Bar
        UpdateBar(Math.Round((DirectorySpaceBytes + RemovedSizeBytes + OverlapSizeBytes) / (1024 ^ 2), 0), Math.Round((SelectedSizeBytes) / (1024 ^ 2), 0))

        'Update whether the folders are identical
        If ChangeCount = 0 And
                RemoveCount = 0 And
                RemoveDir = 0 And
                SelectedSizeBytes > 0 And
                SelectedSizeBytes = OverlapSizeBytes Then
            cmd_CopyPlaylists.Text = "Playlists Are Identical!"
            UpdateFreeSpace = False
            IdenticalDir = True
        Else
            IdenticalDir = False
            cmd_CopyPlaylists.Text = "Sync Playlists!"
            UpdateFreeSpace = True
        End If

        'Return Value
        If UpdateFreeSpace And (DirectorySpaceBytes + RemovedSizeBytes) > (SelectedSizeBytes - OverlapSizeBytes) Then
            cmd_CopyPlaylists.Enabled = True
            Return True
        Else
            cmd_CopyPlaylists.Enabled = False
            Return False
        End If

    End Function
    Private Sub UpdateBar(DirectorySpaceMb As Long, SelectedSizeMb As Long)

        'Update Bar Value
        If DirectorySpaceMb = 0 Then
            pb_SpaceLeft.Maximum = 1
            pb_SpaceLeft.Value = 1
        Else
            pb_SpaceLeft.Maximum = DirectorySpaceMb
            If SelectedSizeMb > DirectorySpaceMb Then
                pb_SpaceLeft.Value = pb_SpaceLeft.Maximum
            Else
                pb_SpaceLeft.Value = SelectedSizeMb
            End If
        End If

        'Update Bar Colour
        If pb_SpaceLeft.Value = pb_SpaceLeft.Maximum Then
            pb_SpaceLeft.ForeColor = Color.Firebrick
        ElseIf pb_SpaceLeft.Value > 0.8 * pb_SpaceLeft.Maximum Then
            pb_SpaceLeft.ForeColor = Color.Chocolate
        Else
            pb_SpaceLeft.ForeColor = Color.ForestGreen
        End If

    End Sub

    Private Sub cmd_CopyPlaylists_Click(sender As Object, e As EventArgs) Handles cmd_CopyPlaylists.Click

        Dim Found As Boolean
        Dim DeleteDirs As Boolean
        Dim PlaylistDir As String
        Dim diri As DirectoryInfo
        Dim tmpPlaylist As Playlist
        Dim DeleteDirList As String = ""
        Dim UpdateList As New List(Of String)
        Dim CopyList As New List(Of String)
        Dim ErrorList As New List(Of String)
        Dim CopiedCount, RemoveCount, UpdateCount, ErrorCount As Integer
        Dim tmpCopiedCount, tmpRemoveCount, tmpUpdateCount, tmpErrorCount As Integer
        Dim RemoveDir As Integer

        Dim SelectedPlaylists As New List(Of Playlist)

        Cancelled = False
        TextOut(Reset:=True)
        TextOut()
        TextOut("CALCULATING SYNC...")
        TextOut()

        'Find Selected Playlists
        TextOut("SYNCING PLAYLISTS:")
        For Each sItem In lb_Playlists.SelectedItems
            tmpPlaylist = PlaylistDictionary(sItem.split("%")(1))
            SelectedPlaylists.Add(tmpPlaylist)
            TextOut(tmpPlaylist.Name, 1)
        Next
        TextOut()

        If UpdateFreeSpace() Then

            'Find Extra Directories
            For Each d In Directory.EnumerateDirectories(CopyDir)
                Found = False
                diri = New DirectoryInfo(d)
                If diri.Name = "System Volume Information" Then Continue For
                For Each pl In SelectedPlaylists
                    If pl.Name = diri.Name Then Found = True : Exit For
                Next
                If Not Found Then
                    DeleteDirList &= vbCr & diri.FullName
                    RemoveDir += 1
                End If
            Next

            If RemoveDir > 0 Then
                TextOut("REMOVE EXTRA FOLDER" & IIf(RemoveDir > 1, "S", "") & " IN " & CopyDir & "?")
                For Each s In DeleteDirList.Split(vbCr) : If Not s = "" Then TextOut(s, 1)
                Next
                TextOut()
            End If

            'Delete Extra Directories if Required
            If RemoveDir > 0 Then DeleteDirs = IIf(MsgBox("Remove Additional Folders From " & CopyDir & "?", vbYesNo, "Delete Folders?") = vbYes, True, False)
            If RemoveDir > 0 And DeleteDirs Then DeleteDirs = IIf(MsgBox("Are you sure? " & RemoveDir & " Folder" & IIf(RemoveDir > 1, "s", "") & " will be deleted!" & vbCr & DeleteDirList, vbYesNo, "Delete Folders?") = vbYes, True, False)
            If DeleteDirs And RemoveDir > 0 Then
                RemoveDir = 0
                For Each d In Directory.EnumerateDirectories(CopyDir)
                    Found = False
                    diri = New DirectoryInfo(d)
                    If diri.Name = "System Volume Information" Then Continue For
                    For Each pl In SelectedPlaylists
                        If pl.Name = diri.Name Then Found = True : Exit For
                    Next
                    If Not Found Then
                        TextOut("DELETED FOLDER:  " & d, 1)
                        Directory.Delete(d, True)
                        RemoveDir += 1
                    End If
                Next
                TextOut()
            End If

            UpdateFreeSpace(True)
            TextOut()
            TextOut("***********************************************************************")
            TextOut()
            TextOut("BEGINNING SYNC!")
            TextOut()

            'Copy the Shit Outta Them Playlists Boi
            For Each s In lb_Playlists.SelectedItems

                'Check if cancelled
                Application.DoEvents()
                If Cancelled Then TextOut("*** SYNC CANCELLED! ***",, True) : Exit Sub

                'Get Playlist Name and Make Collection of Tracks
                s = s.split("%")(1)
                tmpPlaylist = PlaylistDictionary(s)
                CopyList = New List(Of String)
                For Each Track In PlaylistDictionary(s).Tracks
                    CopyList.Add(Track.FileName)
                Next
                TextOut("SYNCING PLAYLIST " & tmpPlaylist.Name.ToUpper & " (" & CopyList.Count & ") TRACKS...")
                TextOut()

                'Reset Counters
                tmpCopiedCount = 0
                tmpUpdateCount = 0
                tmpRemoveCount = 0
                tmpErrorCount = 0

                'Make Folder if neccesary
                PlaylistDir = CopyDir & tmpPlaylist.Name
                If Not Directory.Exists(PlaylistDir) Then
                    Directory.CreateDirectory(PlaylistDir)
                    TextOut("CREATED FOLDER: " & PlaylistDir, 1)
                Else : TextOut("FOLDER EXISTS: " & PlaylistDir, 1)
                End If
                TextOut()
                diri = New DirectoryInfo(PlaylistDir)

                UpdateList = New List(Of String)

                'Remove files not in Playlist
                For Each file In diri.GetFiles

                    'Check if cancelled
                    Application.DoEvents()
                    If Cancelled Then TextOut("*** SYNC CANCELLED! ***",, True) : Exit Sub

                    If Not CopyList.Contains(file.Name) Then
                        Try
                            file.Delete()
                            If file.Name Like "DELETE_OLD_*" Then
                                UpdateList.Add(file.Name.Replace("DELETE_OLD_", ""))
                                TextOut("Updated:   " & file.Name.Replace("DELETE_OLD_", ""), 2)
                                tmpUpdateCount += 1
                            Else
                                TextOut("Deleted:   " & file.Name.Replace("DELETE_", ""), 2)
                                tmpRemoveCount += 1
                            End If

                        Catch ex As Exception
                            tmpErrorCount += 1
                            TextOut("ERROR DELETING: " & file.Name, 2)
                            ErrorList.Add(tmpPlaylist.Name & ": Error Deleting " & file.Name & " (" & ex.Message & ")")
                        End Try
                    End If
                Next

                'Copy Files if Needed
                For Each song In tmpPlaylist.Tracks

                    'Check if cancelled
                    Application.DoEvents()
                    If Cancelled Then TextOut("*** SYNC CANCELLED! ***",, True) : Exit Sub

                    If Not File.Exists(PlaylistDir & "\" & song.FileName) Then
                        Try
                            File.Copy(song.SongPath, PlaylistDir & "\" & song.FileName)
                            If Not UpdateList.Contains(song.FileName) Then
                                tmpCopiedCount += 1
                                TextOut("Copied Track:  " & song.FileName, 2)
                            End If
                        Catch ex As Exception
                            tmpErrorCount += 1
                            TextOut("ERROR COPYING: " & song.FileName, 2)
                            ErrorList.Add(tmpPlaylist.Name & ": Error Copying " & song.FileName & " (" & ex.Message & ")")
                        End Try

                    End If

                Next

                TextOut()
                TextOut(tmpPlaylist.Name & " UPDATED SUCCESSFULLY: ", 1)
                TextOut(tmpCopiedCount & " Tracks Copied", 2)
                TextOut(tmpUpdateCount & " Tracks Updated", 2)
                TextOut(tmpRemoveCount & " Tracks Deleted", 2)
                TextOut(tmpErrorCount & " Errors", 2)
                TextOut()

                CopiedCount += tmpCopiedCount
                UpdateCount += tmpUpdateCount
                RemoveCount += tmpRemoveCount
                ErrorCount += tmpErrorCount

            Next

            'Show Result
            TextOut()
            TextOut("***********************************************************************")
            TextOut()
            TextOut("SYNC COMPLETE!")
            TextOut(CopiedCount & " Total Tracks Copied", 1)
            TextOut(UpdateCount & " Total Tracks Updated", 1)
            TextOut(RemoveCount & " Total Tracks Deleted", 1)
            TextOut()
            TextOut(ErrorCount & " Error" & IIf(ErrorCount > 1, "s", "") & IIf(ErrorCount > 0, ":", ""), 1)
            For Each s In ErrorList : TextOut(s, 2) : Next
            TextOut()
            TextOut()

            MsgBox(lb_Playlists.SelectedItems.Count & " Playlists Synced Succesfully!" & vbCr & vbCr &
                   CopiedCount & " Files Copied" & vbCr &
                   RemoveCount & " Files Removed " & vbCr &
                   UpdateCount & " Files Updated " & vbCr &
                   ErrorCount & " Error(s)",, "Sync Complete!")

            UpdateFreeSpace()

        Else

            'Nuh-Uh
            MsgBox("Not Enough Space in " & CopyDir & " to copy files!")

        End If

        Cancelled = True

    End Sub

    Sub TextOut(Optional Line As String = "", Optional Indent As Integer = 0, Optional Reset As Boolean = False)

        Dim sIndent As String = " "
        Dim tmpLines(tbo_Output.Lines.Count - 1) As String

        For i = 0 To Indent - 1
            sIndent &= "  "
        Next

        tbo_Output.Text += sIndent & Line & vbCrLf
        If tbo_Output.Lines.Count >= 34 Then
            For i = 0 To tbo_Output.Lines.Count - 2
                tmpLines(i) = tbo_Output.Lines(i + 1)
            Next
            tbo_Output.Lines = tmpLines
        End If

        Console.WriteLine(sIndent & Line)

        If Reset Then tbo_Output.Text = vbCrLf & sIndent & Line

        Application.DoEvents()

    End Sub

End Class

Public Class VerticalProgressBar
    Inherits ProgressBar
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.Style = cp.Style Or &H4
            Return cp
        End Get
    End Property
End Class