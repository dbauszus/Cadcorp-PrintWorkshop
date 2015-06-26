<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dialog_Table
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Dialog_Table))
        Me.DGV_OverlaySchema = New System.Windows.Forms.DataGridView()
        Me.CB_Overlays = New System.Windows.Forms.ComboBox()
        Me.CB_FONTSIZE = New System.Windows.Forms.ComboBox()
        Me.CB_FONT = New System.Windows.Forms.ComboBox()
        Me.Btn_Opaque = New System.Windows.Forms.Button()
        Me.Btn_Frame = New System.Windows.Forms.Button()
        Me.btn_Place = New System.Windows.Forms.Button()
        CType(Me.DGV_OverlaySchema, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV_OverlaySchema
        '
        Me.DGV_OverlaySchema.AllowUserToAddRows = False
        Me.DGV_OverlaySchema.AllowUserToDeleteRows = False
        Me.DGV_OverlaySchema.AllowUserToOrderColumns = True
        Me.DGV_OverlaySchema.AllowUserToResizeColumns = False
        Me.DGV_OverlaySchema.AllowUserToResizeRows = False
        Me.DGV_OverlaySchema.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DGV_OverlaySchema.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.DGV_OverlaySchema.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DGV_OverlaySchema.BackgroundColor = System.Drawing.SystemColors.ControlLightLight
        Me.DGV_OverlaySchema.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.DGV_OverlaySchema.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.DGV_OverlaySchema.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV_OverlaySchema.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.DGV_OverlaySchema.Location = New System.Drawing.Point(12, 72)
        Me.DGV_OverlaySchema.MultiSelect = False
        Me.DGV_OverlaySchema.Name = "DGV_OverlaySchema"
        Me.DGV_OverlaySchema.RowHeadersVisible = False
        Me.DGV_OverlaySchema.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DGV_OverlaySchema.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DGV_OverlaySchema.Size = New System.Drawing.Size(539, 313)
        Me.DGV_OverlaySchema.TabIndex = 14
        '
        'CB_Overlays
        '
        Me.CB_Overlays.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CB_Overlays.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_Overlays.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_Overlays.FormattingEnabled = True
        Me.CB_Overlays.Location = New System.Drawing.Point(12, 12)
        Me.CB_Overlays.MaxDropDownItems = 5
        Me.CB_Overlays.Name = "CB_Overlays"
        Me.CB_Overlays.Size = New System.Drawing.Size(539, 24)
        Me.CB_Overlays.TabIndex = 15
        '
        'CB_FONTSIZE
        '
        Me.CB_FONTSIZE.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.CB_FONTSIZE.FormattingEnabled = True
        Me.CB_FONTSIZE.Location = New System.Drawing.Point(325, 42)
        Me.CB_FONTSIZE.MaxDropDownItems = 5
        Me.CB_FONTSIZE.Name = "CB_FONTSIZE"
        Me.CB_FONTSIZE.Size = New System.Drawing.Size(50, 24)
        Me.CB_FONTSIZE.TabIndex = 16
        '
        'CB_FONT
        '
        Me.CB_FONT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_FONT.FormattingEnabled = True
        Me.CB_FONT.Location = New System.Drawing.Point(74, 42)
        Me.CB_FONT.MaxDropDownItems = 5
        Me.CB_FONT.Name = "CB_FONT"
        Me.CB_FONT.Size = New System.Drawing.Size(245, 24)
        Me.CB_FONT.TabIndex = 17
        '
        'Btn_Opaque
        '
        Me.Btn_Opaque.FlatAppearance.BorderSize = 0
        Me.Btn_Opaque.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Opaque.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Opaque
        Me.Btn_Opaque.Location = New System.Drawing.Point(43, 42)
        Me.Btn_Opaque.Name = "Btn_Opaque"
        Me.Btn_Opaque.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Opaque.TabIndex = 21
        Me.Btn_Opaque.UseVisualStyleBackColor = True
        '
        'Btn_Frame
        '
        Me.Btn_Frame.FlatAppearance.BorderSize = 0
        Me.Btn_Frame.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Frame.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Frame
        Me.Btn_Frame.Location = New System.Drawing.Point(12, 42)
        Me.Btn_Frame.Name = "Btn_Frame"
        Me.Btn_Frame.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Frame.TabIndex = 22
        Me.Btn_Frame.UseVisualStyleBackColor = True
        '
        'btn_Place
        '
        Me.btn_Place.Location = New System.Drawing.Point(501, 42)
        Me.btn_Place.Name = "btn_Place"
        Me.btn_Place.Size = New System.Drawing.Size(50, 23)
        Me.btn_Place.TabIndex = 23
        Me.btn_Place.Text = "Place"
        Me.btn_Place.UseVisualStyleBackColor = True
        '
        'Dialog_Table
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(563, 397)
        Me.Controls.Add(Me.btn_Place)
        Me.Controls.Add(Me.Btn_Frame)
        Me.Controls.Add(Me.Btn_Opaque)
        Me.Controls.Add(Me.CB_FONT)
        Me.Controls.Add(Me.CB_FONTSIZE)
        Me.Controls.Add(Me.CB_Overlays)
        Me.Controls.Add(Me.DGV_OverlaySchema)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Dialog_Table"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Workshop - Schema Table"
        Me.TopMost = True
        CType(Me.DGV_OverlaySchema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DGV_OverlaySchema As System.Windows.Forms.DataGridView
    Friend WithEvents CB_Overlays As System.Windows.Forms.ComboBox
    Friend WithEvents CB_FONTSIZE As System.Windows.Forms.ComboBox
    Friend WithEvents CB_FONT As System.Windows.Forms.ComboBox
    Friend WithEvents Btn_Opaque As System.Windows.Forms.Button
    Friend WithEvents Btn_Frame As System.Windows.Forms.Button
    Friend WithEvents btn_Place As System.Windows.Forms.Button

End Class
