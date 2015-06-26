<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Dialog_Book
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Dialog_Book))
        Me.CB_TEMPLATES = New System.Windows.Forms.ComboBox()
        Me.btnCreate = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TB_SCALE = New System.Windows.Forms.TextBox()
        Me.CB_GRIDS = New System.Windows.Forms.ComboBox()
        Me.LCB_TEMPLATES = New System.Windows.Forms.Label()
        Me.LTB_SCALE = New System.Windows.Forms.Label()
        Me.LCB_GRIDS = New System.Windows.Forms.Label()
        Me.CB_OVERLAP = New System.Windows.Forms.ComboBox()
        Me.LCB_OVERLAP = New System.Windows.Forms.Label()
        Me.CHK_AdaptiveOverlap = New System.Windows.Forms.CheckBox()
        Me.TB_GRID = New System.Windows.Forms.TextBox()
        Me.CHK_NEWGRID = New System.Windows.Forms.CheckBox()
        Me.ProgressBar = New System.Windows.Forms.ProgressBar()
        Me.CB_EXTENT = New System.Windows.Forms.ComboBox()
        Me.CB_ORDERBY = New System.Windows.Forms.ComboBox()
        Me.LCB_ORDERBY = New System.Windows.Forms.Label()
        Me.LCB_FORMULA = New System.Windows.Forms.Label()
        Me.CB_FORMULA = New System.Windows.Forms.ComboBox()
        Me.L_Pages = New System.Windows.Forms.Label()
        Me.Pages = New System.Windows.Forms.Label()
        Me.CHK_RENUMBER = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'CB_TEMPLATES
        '
        Me.CB_TEMPLATES.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_TEMPLATES.Enabled = False
        Me.CB_TEMPLATES.FormattingEnabled = True
        Me.CB_TEMPLATES.Location = New System.Drawing.Point(301, 27)
        Me.CB_TEMPLATES.Name = "CB_TEMPLATES"
        Me.CB_TEMPLATES.Size = New System.Drawing.Size(223, 21)
        Me.CB_TEMPLATES.TabIndex = 1
        '
        'btnCreate
        '
        Me.btnCreate.Location = New System.Drawing.Point(15, 194)
        Me.btnCreate.Name = "btnCreate"
        Me.btnCreate.Size = New System.Drawing.Size(259, 23)
        Me.btnCreate.TabIndex = 8
        Me.btnCreate.Text = "Create Grid / Book"
        Me.btnCreate.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(15, 39)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(259, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "GO"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TB_SCALE
        '
        Me.TB_SCALE.Enabled = False
        Me.TB_SCALE.Location = New System.Drawing.Point(301, 73)
        Me.TB_SCALE.Name = "TB_SCALE"
        Me.TB_SCALE.Size = New System.Drawing.Size(223, 20)
        Me.TB_SCALE.TabIndex = 3
        '
        'CB_GRIDS
        '
        Me.CB_GRIDS.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_GRIDS.Enabled = False
        Me.CB_GRIDS.FormattingEnabled = True
        Me.CB_GRIDS.Location = New System.Drawing.Point(15, 27)
        Me.CB_GRIDS.Name = "CB_GRIDS"
        Me.CB_GRIDS.Size = New System.Drawing.Size(259, 21)
        Me.CB_GRIDS.TabIndex = 6
        '
        'LCB_TEMPLATES
        '
        Me.LCB_TEMPLATES.AutoSize = True
        Me.LCB_TEMPLATES.Location = New System.Drawing.Point(298, 9)
        Me.LCB_TEMPLATES.Name = "LCB_TEMPLATES"
        Me.LCB_TEMPLATES.Size = New System.Drawing.Size(75, 13)
        Me.LCB_TEMPLATES.TabIndex = 0
        Me.LCB_TEMPLATES.Text = "Print Template"
        '
        'LTB_SCALE
        '
        Me.LTB_SCALE.AutoSize = True
        Me.LTB_SCALE.Location = New System.Drawing.Point(298, 55)
        Me.LTB_SCALE.Name = "LTB_SCALE"
        Me.LTB_SCALE.Size = New System.Drawing.Size(34, 13)
        Me.LTB_SCALE.TabIndex = 2
        Me.LTB_SCALE.Text = "Scale"
        '
        'LCB_GRIDS
        '
        Me.LCB_GRIDS.AutoSize = True
        Me.LCB_GRIDS.Location = New System.Drawing.Point(12, 9)
        Me.LCB_GRIDS.Name = "LCB_GRIDS"
        Me.LCB_GRIDS.Size = New System.Drawing.Size(49, 13)
        Me.LCB_GRIDS.TabIndex = 5
        Me.LCB_GRIDS.Text = "Grid Title"
        '
        'CB_OVERLAP
        '
        Me.CB_OVERLAP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_OVERLAP.FormattingEnabled = True
        Me.CB_OVERLAP.Items.AddRange(New Object() {"0", "5", "10", "15", "20", "25", "30", "40", "50"})
        Me.CB_OVERLAP.Location = New System.Drawing.Point(301, 121)
        Me.CB_OVERLAP.Name = "CB_OVERLAP"
        Me.CB_OVERLAP.Size = New System.Drawing.Size(223, 21)
        Me.CB_OVERLAP.TabIndex = 10
        Me.CB_OVERLAP.Visible = False
        '
        'LCB_OVERLAP
        '
        Me.LCB_OVERLAP.AutoSize = True
        Me.LCB_OVERLAP.Location = New System.Drawing.Point(298, 103)
        Me.LCB_OVERLAP.Name = "LCB_OVERLAP"
        Me.LCB_OVERLAP.Size = New System.Drawing.Size(72, 13)
        Me.LCB_OVERLAP.TabIndex = 11
        Me.LCB_OVERLAP.Text = "Page Overlap"
        Me.LCB_OVERLAP.Visible = False
        '
        'CHK_AdaptiveOverlap
        '
        Me.CHK_AdaptiveOverlap.AutoSize = True
        Me.CHK_AdaptiveOverlap.Location = New System.Drawing.Point(301, 148)
        Me.CHK_AdaptiveOverlap.Name = "CHK_AdaptiveOverlap"
        Me.CHK_AdaptiveOverlap.Size = New System.Drawing.Size(108, 17)
        Me.CHK_AdaptiveOverlap.TabIndex = 12
        Me.CHK_AdaptiveOverlap.Text = "Adaptive Overlap"
        Me.CHK_AdaptiveOverlap.UseVisualStyleBackColor = True
        Me.CHK_AdaptiveOverlap.Visible = False
        '
        'TB_GRID
        '
        Me.TB_GRID.Location = New System.Drawing.Point(15, 27)
        Me.TB_GRID.Name = "TB_GRID"
        Me.TB_GRID.Size = New System.Drawing.Size(259, 20)
        Me.TB_GRID.TabIndex = 14
        '
        'CHK_NEWGRID
        '
        Me.CHK_NEWGRID.AutoSize = True
        Me.CHK_NEWGRID.Enabled = False
        Me.CHK_NEWGRID.Location = New System.Drawing.Point(15, 54)
        Me.CHK_NEWGRID.Name = "CHK_NEWGRID"
        Me.CHK_NEWGRID.Size = New System.Drawing.Size(104, 17)
        Me.CHK_NEWGRID.TabIndex = 18
        Me.CHK_NEWGRID.Text = "Create New Grid"
        Me.CHK_NEWGRID.UseVisualStyleBackColor = True
        '
        'ProgressBar
        '
        Me.ProgressBar.Location = New System.Drawing.Point(15, 166)
        Me.ProgressBar.Name = "ProgressBar"
        Me.ProgressBar.Size = New System.Drawing.Size(259, 21)
        Me.ProgressBar.TabIndex = 20
        Me.ProgressBar.Visible = False
        '
        'CB_EXTENT
        '
        Me.CB_EXTENT.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_EXTENT.Enabled = False
        Me.CB_EXTENT.FormattingEnabled = True
        Me.CB_EXTENT.Items.AddRange(New Object() {"Extent of current map window"})
        Me.CB_EXTENT.Location = New System.Drawing.Point(15, 73)
        Me.CB_EXTENT.Name = "CB_EXTENT"
        Me.CB_EXTENT.Size = New System.Drawing.Size(259, 21)
        Me.CB_EXTENT.TabIndex = 23
        '
        'CB_ORDERBY
        '
        Me.CB_ORDERBY.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_ORDERBY.FormattingEnabled = True
        Me.CB_ORDERBY.Items.AddRange(New Object() {""})
        Me.CB_ORDERBY.Location = New System.Drawing.Point(301, 166)
        Me.CB_ORDERBY.Name = "CB_ORDERBY"
        Me.CB_ORDERBY.Size = New System.Drawing.Size(223, 21)
        Me.CB_ORDERBY.TabIndex = 29
        Me.CB_ORDERBY.Visible = False
        '
        'LCB_ORDERBY
        '
        Me.LCB_ORDERBY.AutoSize = True
        Me.LCB_ORDERBY.Location = New System.Drawing.Point(298, 149)
        Me.LCB_ORDERBY.Name = "LCB_ORDERBY"
        Me.LCB_ORDERBY.Size = New System.Drawing.Size(119, 13)
        Me.LCB_ORDERBY.TabIndex = 30
        Me.LCB_ORDERBY.Text = "Order pages by schema"
        Me.LCB_ORDERBY.Visible = False
        '
        'LCB_FORMULA
        '
        Me.LCB_FORMULA.AutoSize = True
        Me.LCB_FORMULA.Location = New System.Drawing.Point(298, 103)
        Me.LCB_FORMULA.Name = "LCB_FORMULA"
        Me.LCB_FORMULA.Size = New System.Drawing.Size(127, 13)
        Me.LCB_FORMULA.TabIndex = 32
        Me.LCB_FORMULA.Text = "Use schema for page title"
        Me.LCB_FORMULA.Visible = False
        '
        'CB_FORMULA
        '
        Me.CB_FORMULA.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CB_FORMULA.FormattingEnabled = True
        Me.CB_FORMULA.Items.AddRange(New Object() {""})
        Me.CB_FORMULA.Location = New System.Drawing.Point(301, 121)
        Me.CB_FORMULA.Name = "CB_FORMULA"
        Me.CB_FORMULA.Size = New System.Drawing.Size(223, 21)
        Me.CB_FORMULA.TabIndex = 31
        Me.CB_FORMULA.Visible = False
        '
        'L_Pages
        '
        Me.L_Pages.AutoSize = True
        Me.L_Pages.Location = New System.Drawing.Point(12, 124)
        Me.L_Pages.Name = "L_Pages"
        Me.L_Pages.Size = New System.Drawing.Size(43, 13)
        Me.L_Pages.TabIndex = 33
        Me.L_Pages.Text = "Pages: "
        '
        'Pages
        '
        Me.Pages.AutoSize = True
        Me.Pages.Location = New System.Drawing.Point(61, 124)
        Me.Pages.Name = "Pages"
        Me.Pages.Size = New System.Drawing.Size(0, 13)
        Me.Pages.TabIndex = 34
        '
        'CHK_RENUMBER
        '
        Me.CHK_RENUMBER.AutoSize = True
        Me.CHK_RENUMBER.Location = New System.Drawing.Point(15, 145)
        Me.CHK_RENUMBER.Name = "CHK_RENUMBER"
        Me.CHK_RENUMBER.Size = New System.Drawing.Size(110, 17)
        Me.CHK_RENUMBER.TabIndex = 35
        Me.CHK_RENUMBER.Text = "Re-number pages"
        Me.CHK_RENUMBER.UseVisualStyleBackColor = True
        '
        'Dialog_Book
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(538, 229)
        Me.Controls.Add(Me.CHK_RENUMBER)
        Me.Controls.Add(Me.Pages)
        Me.Controls.Add(Me.L_Pages)
        Me.Controls.Add(Me.LCB_FORMULA)
        Me.Controls.Add(Me.CB_FORMULA)
        Me.Controls.Add(Me.LCB_ORDERBY)
        Me.Controls.Add(Me.CB_ORDERBY)
        Me.Controls.Add(Me.CB_EXTENT)
        Me.Controls.Add(Me.ProgressBar)
        Me.Controls.Add(Me.CHK_NEWGRID)
        Me.Controls.Add(Me.TB_GRID)
        Me.Controls.Add(Me.CHK_AdaptiveOverlap)
        Me.Controls.Add(Me.LCB_OVERLAP)
        Me.Controls.Add(Me.CB_OVERLAP)
        Me.Controls.Add(Me.LCB_GRIDS)
        Me.Controls.Add(Me.LTB_SCALE)
        Me.Controls.Add(Me.LCB_TEMPLATES)
        Me.Controls.Add(Me.CB_GRIDS)
        Me.Controls.Add(Me.TB_SCALE)
        Me.Controls.Add(Me.btnCreate)
        Me.Controls.Add(Me.CB_TEMPLATES)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Dialog_Book"
        Me.ShowInTaskbar = False
        Me.Text = "Book Plot"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CB_TEMPLATES As System.Windows.Forms.ComboBox
    Friend WithEvents btnCreate As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TB_SCALE As System.Windows.Forms.TextBox
    Friend WithEvents CB_GRIDS As System.Windows.Forms.ComboBox
    Friend WithEvents LCB_TEMPLATES As System.Windows.Forms.Label
    Friend WithEvents LTB_SCALE As System.Windows.Forms.Label
    Friend WithEvents LCB_GRIDS As System.Windows.Forms.Label
    Friend WithEvents CB_OVERLAP As System.Windows.Forms.ComboBox
    Friend WithEvents LCB_OVERLAP As System.Windows.Forms.Label
    Friend WithEvents CHK_AdaptiveOverlap As System.Windows.Forms.CheckBox
    Friend WithEvents TB_GRID As System.Windows.Forms.TextBox
    Friend WithEvents CHK_NEWGRID As System.Windows.Forms.CheckBox
    Friend WithEvents CB_EXTENT As System.Windows.Forms.ComboBox
    Friend WithEvents ProgressBar As System.Windows.Forms.ProgressBar
    Friend WithEvents CB_ORDERBY As System.Windows.Forms.ComboBox
    Friend WithEvents LCB_ORDERBY As System.Windows.Forms.Label
    Friend WithEvents LCB_FORMULA As System.Windows.Forms.Label
    Friend WithEvents CB_FORMULA As System.Windows.Forms.ComboBox
    Friend WithEvents L_Pages As System.Windows.Forms.Label
    Friend WithEvents Pages As System.Windows.Forms.Label
    Friend WithEvents CHK_RENUMBER As System.Windows.Forms.CheckBox
End Class
