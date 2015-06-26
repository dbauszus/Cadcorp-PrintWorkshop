<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dialog_Label
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Dialog_Label))
        Me.LabelTextBox = New System.Windows.Forms.RichTextBox()
        Me.Btn_alignLeft = New System.Windows.Forms.Button()
        Me.Btn_alignMiddle = New System.Windows.Forms.Button()
        Me.Btn_alignRight = New System.Windows.Forms.Button()
        Me.Btn_txtItalic = New System.Windows.Forms.Button()
        Me.Btn_txtBold = New System.Windows.Forms.Button()
        Me.CB_LABELS = New System.Windows.Forms.ComboBox()
        Me.CB_FONT = New System.Windows.Forms.ComboBox()
        Me.CB_FONTSIZE = New System.Windows.Forms.ComboBox()
        Me.Btn_Edit = New System.Windows.Forms.Button()
        Me.btn_Place = New System.Windows.Forms.Button()
        Me.Btn_Reload = New System.Windows.Forms.Button()
        Me.Btn_Label = New System.Windows.Forms.Button()
        Me.Btn_Opaque = New System.Windows.Forms.Button()
        Me.Btn_Frame = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LabelTextBox
        '
        Me.LabelTextBox.DetectUrls = False
        Me.LabelTextBox.Location = New System.Drawing.Point(12, 103)
        Me.LabelTextBox.Name = "LabelTextBox"
        Me.LabelTextBox.ShortcutsEnabled = False
        Me.LabelTextBox.Size = New System.Drawing.Size(456, 340)
        Me.LabelTextBox.TabIndex = 0
        Me.LabelTextBox.Text = ""
        Me.LabelTextBox.WordWrap = False
        '
        'Btn_alignLeft
        '
        Me.Btn_alignLeft.FlatAppearance.BorderSize = 0
        Me.Btn_alignLeft.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_alignLeft.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_AlignLeft
        Me.Btn_alignLeft.Location = New System.Drawing.Point(12, 12)
        Me.Btn_alignLeft.Name = "Btn_alignLeft"
        Me.Btn_alignLeft.Size = New System.Drawing.Size(25, 25)
        Me.Btn_alignLeft.TabIndex = 1
        Me.Btn_alignLeft.UseVisualStyleBackColor = True
        '
        'Btn_alignMiddle
        '
        Me.Btn_alignMiddle.FlatAppearance.BorderSize = 0
        Me.Btn_alignMiddle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_alignMiddle.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_AlignCentre
        Me.Btn_alignMiddle.Location = New System.Drawing.Point(43, 12)
        Me.Btn_alignMiddle.Name = "Btn_alignMiddle"
        Me.Btn_alignMiddle.Size = New System.Drawing.Size(25, 25)
        Me.Btn_alignMiddle.TabIndex = 2
        Me.Btn_alignMiddle.UseVisualStyleBackColor = True
        '
        'Btn_alignRight
        '
        Me.Btn_alignRight.FlatAppearance.BorderSize = 0
        Me.Btn_alignRight.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_alignRight.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_AlignRight
        Me.Btn_alignRight.Location = New System.Drawing.Point(74, 12)
        Me.Btn_alignRight.Name = "Btn_alignRight"
        Me.Btn_alignRight.Size = New System.Drawing.Size(25, 25)
        Me.Btn_alignRight.TabIndex = 3
        Me.Btn_alignRight.UseVisualStyleBackColor = True
        '
        'Btn_txtItalic
        '
        Me.Btn_txtItalic.FlatAppearance.BorderSize = 0
        Me.Btn_txtItalic.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_txtItalic.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_TextItalic
        Me.Btn_txtItalic.Location = New System.Drawing.Point(136, 12)
        Me.Btn_txtItalic.Name = "Btn_txtItalic"
        Me.Btn_txtItalic.Size = New System.Drawing.Size(25, 25)
        Me.Btn_txtItalic.TabIndex = 5
        Me.Btn_txtItalic.UseVisualStyleBackColor = True
        '
        'Btn_txtBold
        '
        Me.Btn_txtBold.FlatAppearance.BorderSize = 0
        Me.Btn_txtBold.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_txtBold.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_TextBold
        Me.Btn_txtBold.Location = New System.Drawing.Point(105, 12)
        Me.Btn_txtBold.Name = "Btn_txtBold"
        Me.Btn_txtBold.Size = New System.Drawing.Size(25, 25)
        Me.Btn_txtBold.TabIndex = 4
        Me.Btn_txtBold.UseVisualStyleBackColor = True
        '
        'CB_LABELS
        '
        Me.CB_LABELS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_LABELS.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.CB_LABELS.FormattingEnabled = True
        Me.CB_LABELS.Location = New System.Drawing.Point(74, 44)
        Me.CB_LABELS.Name = "CB_LABELS"
        Me.CB_LABELS.Size = New System.Drawing.Size(394, 24)
        Me.CB_LABELS.TabIndex = 8
        '
        'CB_FONT
        '
        Me.CB_FONT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_FONT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CB_FONT.FormattingEnabled = True
        Me.CB_FONT.Location = New System.Drawing.Point(223, 14)
        Me.CB_FONT.MaxDropDownItems = 5
        Me.CB_FONT.Name = "CB_FONT"
        Me.CB_FONT.Size = New System.Drawing.Size(245, 24)
        Me.CB_FONT.TabIndex = 10
        '
        'CB_FONTSIZE
        '
        Me.CB_FONTSIZE.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.CB_FONTSIZE.FormattingEnabled = True
        Me.CB_FONTSIZE.Location = New System.Drawing.Point(167, 14)
        Me.CB_FONTSIZE.MaxDropDownItems = 5
        Me.CB_FONTSIZE.Name = "CB_FONTSIZE"
        Me.CB_FONTSIZE.Size = New System.Drawing.Size(50, 24)
        Me.CB_FONTSIZE.TabIndex = 11
        '
        'Btn_Edit
        '
        Me.Btn_Edit.FlatAppearance.BorderSize = 0
        Me.Btn_Edit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Edit.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Edit
        Me.Btn_Edit.Location = New System.Drawing.Point(12, 43)
        Me.Btn_Edit.Name = "Btn_Edit"
        Me.Btn_Edit.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Edit.TabIndex = 12
        Me.Btn_Edit.UseVisualStyleBackColor = True
        '
        'btn_Place
        '
        Me.btn_Place.Location = New System.Drawing.Point(418, 74)
        Me.btn_Place.Name = "btn_Place"
        Me.btn_Place.Size = New System.Drawing.Size(50, 23)
        Me.btn_Place.TabIndex = 13
        Me.btn_Place.Text = "Place"
        Me.btn_Place.UseVisualStyleBackColor = True
        '
        'Btn_Reload
        '
        Me.Btn_Reload.FlatAppearance.BorderSize = 0
        Me.Btn_Reload.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Reload.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Reload
        Me.Btn_Reload.Location = New System.Drawing.Point(43, 43)
        Me.Btn_Reload.Name = "Btn_Reload"
        Me.Btn_Reload.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Reload.TabIndex = 18
        Me.Btn_Reload.UseVisualStyleBackColor = True
        '
        'Btn_Label
        '
        Me.Btn_Label.FlatAppearance.BorderSize = 0
        Me.Btn_Label.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Label.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Label
        Me.Btn_Label.Location = New System.Drawing.Point(74, 72)
        Me.Btn_Label.Name = "Btn_Label"
        Me.Btn_Label.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Label.TabIndex = 19
        Me.Btn_Label.UseVisualStyleBackColor = True
        '
        'Btn_Opaque
        '
        Me.Btn_Opaque.FlatAppearance.BorderSize = 0
        Me.Btn_Opaque.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Opaque.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Opaque
        Me.Btn_Opaque.Location = New System.Drawing.Point(12, 72)
        Me.Btn_Opaque.Name = "Btn_Opaque"
        Me.Btn_Opaque.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Opaque.TabIndex = 20
        Me.Btn_Opaque.UseVisualStyleBackColor = True
        '
        'Btn_Frame
        '
        Me.Btn_Frame.FlatAppearance.BorderSize = 0
        Me.Btn_Frame.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Btn_Frame.Image = Global.PrintWorkshop.My.Resources.Resources.Btn_Frame
        Me.Btn_Frame.Location = New System.Drawing.Point(43, 72)
        Me.Btn_Frame.Name = "Btn_Frame"
        Me.Btn_Frame.Size = New System.Drawing.Size(25, 25)
        Me.Btn_Frame.TabIndex = 21
        Me.Btn_Frame.UseVisualStyleBackColor = True
        '
        'Dialog_Label
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(480, 455)
        Me.Controls.Add(Me.Btn_Frame)
        Me.Controls.Add(Me.Btn_Opaque)
        Me.Controls.Add(Me.Btn_Label)
        Me.Controls.Add(Me.Btn_Reload)
        Me.Controls.Add(Me.btn_Place)
        Me.Controls.Add(Me.Btn_Edit)
        Me.Controls.Add(Me.CB_FONTSIZE)
        Me.Controls.Add(Me.CB_FONT)
        Me.Controls.Add(Me.CB_LABELS)
        Me.Controls.Add(Me.Btn_txtItalic)
        Me.Controls.Add(Me.Btn_txtBold)
        Me.Controls.Add(Me.Btn_alignRight)
        Me.Controls.Add(Me.Btn_alignMiddle)
        Me.Controls.Add(Me.Btn_alignLeft)
        Me.Controls.Add(Me.LabelTextBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Dialog_Label"
        Me.ShowInTaskbar = False
        Me.Text = "Print Workshop - Labels"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LabelTextBox As System.Windows.Forms.RichTextBox
    Friend WithEvents Btn_alignLeft As System.Windows.Forms.Button
    Friend WithEvents Btn_alignMiddle As System.Windows.Forms.Button
    Friend WithEvents Btn_alignRight As System.Windows.Forms.Button
    Friend WithEvents Btn_txtItalic As System.Windows.Forms.Button
    Friend WithEvents Btn_txtBold As System.Windows.Forms.Button
    Friend WithEvents CB_LABELS As System.Windows.Forms.ComboBox
    Friend WithEvents CB_FONT As System.Windows.Forms.ComboBox
    Friend WithEvents CB_FONTSIZE As System.Windows.Forms.ComboBox
    Friend WithEvents Btn_Edit As System.Windows.Forms.Button
    Friend WithEvents btn_Place As System.Windows.Forms.Button
    Friend WithEvents Btn_Reload As System.Windows.Forms.Button
    Friend WithEvents Btn_Label As System.Windows.Forms.Button
    Friend WithEvents Btn_Opaque As System.Windows.Forms.Button
    Friend WithEvents Btn_Frame As System.Windows.Forms.Button
End Class
