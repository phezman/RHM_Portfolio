<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SelectPlaylists
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SelectPlaylists))
        Me.lb_Playlists = New System.Windows.Forms.ListBox()
        Me.cmd_CopyPlaylists = New System.Windows.Forms.Button()
        Me.cmd_Cancel = New System.Windows.Forms.Button()
        Me.cmd_ChangeDir = New System.Windows.Forms.Button()
        Me.lbl_SyncDestination = New System.Windows.Forms.Label()
        Me.tbo_Output = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lbl_SpaceLeft = New System.Windows.Forms.Label()
        Me.cmd_SelectExisting = New System.Windows.Forms.Button()
        Me.pb_SpaceLeft = New ExportPlaylistsToUSB.VerticalProgressBar()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lb_Playlists
        '
        Me.lb_Playlists.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lb_Playlists.Font = New System.Drawing.Font("Kozuka Gothic Pr6N M", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lb_Playlists.FormattingEnabled = True
        Me.lb_Playlists.ItemHeight = 23
        Me.lb_Playlists.Location = New System.Drawing.Point(12, 44)
        Me.lb_Playlists.Name = "lb_Playlists"
        Me.lb_Playlists.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lb_Playlists.Size = New System.Drawing.Size(373, 556)
        Me.lb_Playlists.TabIndex = 1
        '
        'cmd_CopyPlaylists
        '
        Me.cmd_CopyPlaylists.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmd_CopyPlaylists.Font = New System.Drawing.Font("Rockwell", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmd_CopyPlaylists.Location = New System.Drawing.Point(812, 606)
        Me.cmd_CopyPlaylists.Name = "cmd_CopyPlaylists"
        Me.cmd_CopyPlaylists.Size = New System.Drawing.Size(236, 36)
        Me.cmd_CopyPlaylists.TabIndex = 2
        Me.cmd_CopyPlaylists.Text = "Copy Playlists!"
        Me.cmd_CopyPlaylists.UseVisualStyleBackColor = False
        '
        'cmd_Cancel
        '
        Me.cmd_Cancel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmd_Cancel.Font = New System.Drawing.Font("Rockwell", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle))
        Me.cmd_Cancel.Location = New System.Drawing.Point(12, 606)
        Me.cmd_Cancel.Name = "cmd_Cancel"
        Me.cmd_Cancel.Size = New System.Drawing.Size(170, 36)
        Me.cmd_Cancel.TabIndex = 3
        Me.cmd_Cancel.Text = "Cancel"
        Me.cmd_Cancel.UseVisualStyleBackColor = False
        '
        'cmd_ChangeDir
        '
        Me.cmd_ChangeDir.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmd_ChangeDir.Font = New System.Drawing.Font("Rockwell", 11.0!, System.Drawing.FontStyle.Bold)
        Me.cmd_ChangeDir.Location = New System.Drawing.Point(925, 9)
        Me.cmd_ChangeDir.Name = "cmd_ChangeDir"
        Me.cmd_ChangeDir.Size = New System.Drawing.Size(123, 29)
        Me.cmd_ChangeDir.TabIndex = 4
        Me.cmd_ChangeDir.Text = "Destination"
        Me.cmd_ChangeDir.UseVisualStyleBackColor = False
        '
        'lbl_SyncDestination
        '
        Me.lbl_SyncDestination.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_SyncDestination.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_SyncDestination.Font = New System.Drawing.Font("Rockwell", 10.0!)
        Me.lbl_SyncDestination.Image = CType(resources.GetObject("lbl_SyncDestination.Image"), System.Drawing.Image)
        Me.lbl_SyncDestination.Location = New System.Drawing.Point(391, 9)
        Me.lbl_SyncDestination.Name = "lbl_SyncDestination"
        Me.lbl_SyncDestination.Size = New System.Drawing.Size(528, 29)
        Me.lbl_SyncDestination.TabIndex = 10
        Me.lbl_SyncDestination.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbo_Output
        '
        Me.tbo_Output.BackColor = System.Drawing.Color.Black
        Me.tbo_Output.Font = New System.Drawing.Font("Courier New", 11.0!)
        Me.tbo_Output.ForeColor = System.Drawing.Color.White
        Me.tbo_Output.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.tbo_Output.Location = New System.Drawing.Point(3, 3)
        Me.tbo_Output.Multiline = True
        Me.tbo_Output.Name = "tbo_Output"
        Me.tbo_Output.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.tbo_Output.Size = New System.Drawing.Size(606, 546)
        Me.tbo_Output.TabIndex = 11
        Me.tbo_Output.WordWrap = False
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.tbo_Output)
        Me.Panel1.Enabled = False
        Me.Panel1.Location = New System.Drawing.Point(391, 44)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(616, 556)
        Me.Panel1.TabIndex = 12
        '
        'lbl_SpaceLeft
        '
        Me.lbl_SpaceLeft.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lbl_SpaceLeft.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbl_SpaceLeft.Font = New System.Drawing.Font("Rockwell", 10.0!)
        Me.lbl_SpaceLeft.Image = CType(resources.GetObject("lbl_SpaceLeft.Image"), System.Drawing.Image)
        Me.lbl_SpaceLeft.Location = New System.Drawing.Point(12, 9)
        Me.lbl_SpaceLeft.Name = "lbl_SpaceLeft"
        Me.lbl_SpaceLeft.Size = New System.Drawing.Size(373, 29)
        Me.lbl_SpaceLeft.TabIndex = 13
        Me.lbl_SpaceLeft.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmd_SelectExisting
        '
        Me.cmd_SelectExisting.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.cmd_SelectExisting.Font = New System.Drawing.Font("Rockwell", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle))
        Me.cmd_SelectExisting.Location = New System.Drawing.Point(215, 606)
        Me.cmd_SelectExisting.Name = "cmd_SelectExisting"
        Me.cmd_SelectExisting.Size = New System.Drawing.Size(170, 36)
        Me.cmd_SelectExisting.TabIndex = 14
        Me.cmd_SelectExisting.Text = "Select Existing"
        Me.cmd_SelectExisting.UseVisualStyleBackColor = False
        '
        'pb_SpaceLeft
        '
        Me.pb_SpaceLeft.BackColor = System.Drawing.Color.Tan
        Me.pb_SpaceLeft.ForeColor = System.Drawing.Color.Green
        Me.pb_SpaceLeft.Location = New System.Drawing.Point(1013, 44)
        Me.pb_SpaceLeft.MarqueeAnimationSpeed = 200
        Me.pb_SpaceLeft.Name = "pb_SpaceLeft"
        Me.pb_SpaceLeft.Size = New System.Drawing.Size(35, 556)
        Me.pb_SpaceLeft.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pb_SpaceLeft.TabIndex = 9
        '
        'SelectPlaylists
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(1058, 651)
        Me.Controls.Add(Me.cmd_SelectExisting)
        Me.Controls.Add(Me.lbl_SpaceLeft)
        Me.Controls.Add(Me.lbl_SyncDestination)
        Me.Controls.Add(Me.pb_SpaceLeft)
        Me.Controls.Add(Me.cmd_ChangeDir)
        Me.Controls.Add(Me.cmd_Cancel)
        Me.Controls.Add(Me.cmd_CopyPlaylists)
        Me.Controls.Add(Me.lb_Playlists)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SelectPlaylists"
        Me.Text = "Playlist Exporter"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lb_Playlists As Windows.Forms.ListBox
    Friend WithEvents cmd_CopyPlaylists As Windows.Forms.Button
    Friend WithEvents cmd_Cancel As Windows.Forms.Button
    Friend WithEvents cmd_ChangeDir As Windows.Forms.Button
    Friend WithEvents pb_SpaceLeft As VerticalProgressBar
    Friend WithEvents lbl_SyncDestination As Windows.Forms.Label
    Friend WithEvents tbo_Output As Windows.Forms.TextBox
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents lbl_SpaceLeft As Windows.Forms.Label
    Friend WithEvents cmd_SelectExisting As Windows.Forms.Button
End Class
