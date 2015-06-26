Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.Windows.Forms
Imports System.Data

Public Class Dialog_Table

    Private Table_font As Drawing.Font

    Private SchemaTable() As DataTable

    Public Sub New()

        InitializeComponent()

        'Font size dropdown
        CB_FONTSIZE.Items.Clear()
        For i = 8 To 72 Step 2

            CB_FONTSIZE.Items.Add(i)

        Next i
        CB_FONTSIZE.SelectedIndex = 1

        'Font dropdown
        CB_FONT.Items.Clear()
        Dim iIndex As Integer = 0
        For Each FontFamily In Drawing.FontFamily.Families

            If FontFamily.IsStyleAvailable(Drawing.FontStyle.Bold) = True And FontFamily.IsStyleAvailable(Drawing.FontStyle.Italic) = True Then

                CB_FONT.Items.Add(FontFamily.Name)
                If FontFamily.Name.ToString = "Arial" Then

                    CB_FONT.SelectedIndex = iIndex

                End If
                iIndex += 1

            End If

        Next

        'Overlay dropdown
        CB_Overlays.Items.Clear()
        PW.Handler = PW.SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$")
        PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", True)
        PW.OpenPhoto("lPhoto", 0)
        ReDim SchemaTable(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
        For i = 0 To PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

            CB_Overlays.Items.Add(PW.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$"))

        Next i
        PW.SIS.SwdClose(SIS_NOSAVE)
        PW.SIS.Dispose()
        PW.SIS = PW.APP.TakeoverManager
        PW.ActivateWindow(PW.Handler)

        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Me.Btn_Opaque, "Draw an opaque background.")
        tooltip.SetToolTip(Me.Btn_Frame, "Draw frame lines.")

    End Sub

    Private Sub CreateContextMenuStrip()
        Dim CMS As New ContextMenuStrip()
        For i = 0 To DGV_OverlaySchema.ColumnCount - 1

            Dim TSMI As New ToolStripMenuItem()
            TSMI.Text = DGV_OverlaySchema.Columns(i).HeaderText
            TSMI.Name = DGV_OverlaySchema.Columns(i).Name
            TSMI.Checked = True
            TSMI.CheckOnClick = True
            AddHandler TSMI.CheckedChanged, AddressOf TSMI_CheckChanged
            CMS.Items.Add(TSMI)

        Next
        For Each column As DataGridViewColumn In DGV_OverlaySchema.Columns()

            column.HeaderCell.ContextMenuStrip = CMS

        Next

    End Sub

    Private Sub TSMI_CheckChanged(ByVal sender As System.Object, ByVal e As EventArgs)

        Dim TSMI As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        If TSMI.Checked Then

            DGV_OverlaySchema.Columns(TSMI.Name).Visible = True

        Else

            DGV_OverlaySchema.Columns(TSMI.Name).Visible = False

        End If

    End Sub

    Private Sub CreateTableCache()

        Try

            PW.OpenPhoto("lPhoto", 0)
            Dim x1, y1, x2, y2, z As Double
            PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetViewExtent())
            PW.SIS.CreateRectLocus("Locus", x1, y1, x2, y2)
            PW.SIS.ChangeLocusTestMode("Locus", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
            Dim iOverlay As Integer = CB_Overlays.SelectedIndex
            If PW.SIS.ScanOverlay("ScanOverlay", iOverlay, "", "Locus") < 257 Then

                Dim listsize As Integer = PW.SIS.GetListSize("ScanOverlay")
                SchemaTable(iOverlay) = New DataTable

                'get fields
                Dim sFields As String = ""
                Dim sSchemaFormula As String = ""
                PW.SIS.GetOverlaySchema(iOverlay, "Schema")
                PW.SIS.LoadSchema("Schema")
                For i = 0 To PW.SIS.GetInt(SIS_OT_SCHEMA, 0, "_nColumns&") - 1

                    sSchemaFormula = PW.SIS.GetStr(SIS_OT_SCHEMACOLUMN, i, "_formula$")
                    sFields += sSchemaFormula & vbTab

                Next

                'get numrows
                PW.SIS.OpenListCursor("OverlayCursor", "ScanOverlay", sFields)
                Dim asStatsRows() As String = PW.SIS.GetCursorFieldStatistics("OverlayCursor", 0).Split(vbCrLf)
                Dim asRows() As String = asStatsRows(0).Split(vbTab)
                Dim iNumRows As Integer = asRows(1)

                'create table schema
                PW.SIS.LoadSchema("Schema")
                Dim sSchemaDesc As String
                Dim dColumn As DataColumn
                For i = 0 To PW.SIS.GetInt(SIS_OT_SCHEMA, 0, "_nColumns&") - 1

                    sSchemaDesc = PW.SIS.GetStr(SIS_OT_SCHEMACOLUMN, i, "_description$")
                    dColumn = New DataColumn(sSchemaDesc)
                    dColumn.DataType = System.Type.GetType("System.String")
                    SchemaTable(iOverlay).Columns.Add(dColumn)

                Next i

                'populate rows
                PW.SIS.MoveCursorToBegin("OverlayCursor")
                Dim dRow As DataRow
                Do

                    Dim asField() As String = PW.SIS.GetCursorFieldValues("OverlayCursor", vbTab, "").Split(vbTab)
                    dRow = SchemaTable(iOverlay).NewRow
                    For i As Integer = 0 To PW.SIS.GetInt(SIS_OT_SCHEMA, 0, "_nColumns&") - 1

                        sSchemaDesc = PW.SIS.GetStr(SIS_OT_SCHEMACOLUMN, i, "_description$")
                        dRow.Item(sSchemaDesc) = asField(i)

                    Next i
                    SchemaTable(iOverlay).Rows.Add(dRow)

                Loop Until PW.SIS.MoveCursor("OverlayCursor", 1) = 0

            Else

                MsgBox("Number of item records exceeds maximum (256).", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If
            PW.SIS.SwdClose(SIS_NOSAVE)
            PW.SIS.Dispose()
            PW.SIS = PW.APP.TakeoverManager
            PW.ActivateWindow(PW.Handler)
            PW.SIS.RemoveProperty(SIS_OT_WINDOW, 0, "Handler&")
        Catch
        End Try

    End Sub

    Private Sub CB_Overlays_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_Overlays.SelectedIndexChanged

        Try

            If SchemaTable(CB_Overlays.SelectedIndex) Is Nothing Then

                CreateTableCache()

            End If
            DGV_OverlaySchema.DataSource = Nothing
            DGV_OverlaySchema.DataSource = SchemaTable(CB_Overlays.SelectedIndex)
            CreateContextMenuStrip()

        Catch
        End Try

    End Sub

    Private Sub CB_FontSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONTSIZE.SelectedIndexChanged

        DGV_OverlaySchema.Font = New Drawing.Font(CB_FONT.Text, Single.Parse(CB_FONTSIZE.Text))
        DGV_OverlaySchema.Refresh()

    End Sub

    Private Sub CB_FONT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONT.SelectedIndexChanged

        DGV_OverlaySchema.Font = New Drawing.Font(CB_FONT.Text, Single.Parse(CB_FONTSIZE.Text))
        DGV_OverlaySchema.Refresh()

    End Sub

    Private Table_opaque As Boolean = False

    Private Sub Btn_Opaque_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Opaque.Click

        If Not Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard
            Table_opaque = True

        Else

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Flat
            Table_opaque = False

        End If

    End Sub

    Private Table_frame As Boolean = False

    Private Sub Btn_Frame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Frame.Click

        If Not Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Standard
            Table_frame = True

        Else

            Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Flat
            Table_frame = False

        End If

    End Sub

    Private Sub btn_Place_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Place.Click

        Try

            Dim iOverlay As Integer = CB_Overlays.SelectedIndex
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            PW.SIS.CreateGroup("")
            Dim ySpace As Double = 0.00044 * DGV_OverlaySchema.Font.Size
            Dim xSpace As Double = 0.002
            Dim iColumn As Integer = 0
            Dim iRow As Integer = 0
            For Each column As DataGridViewColumn In DGV_OverlaySchema.Columns()

                If column.Visible = True Then

                    iRow = 0
                    PW.SIS.CreateBoxText(xSpace, (-ySpace * iRow) - 0.0015, 0, DGV_OverlaySchema.Font.Size * 0.0003527, column.HeaderText.ToString)
                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                    PW.SIS.SetInt(SIS_OT_CURITEM, 0, "_text_bold&", True)
                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", CB_FONT.Text)
                    iRow += 1
                    For Each row As DataGridViewRow In DGV_OverlaySchema.Rows

                        If Not row.Cells(iColumn).Value.ToString = "" Then

                            PW.SIS.CreateBoxText(xSpace, (-ySpace * iRow) - 0.0015, 0, DGV_OverlaySchema.Font.Size * 0.0003527, row.Cells(iColumn).Value.ToString)
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", CB_FONT.Text)

                        End If
                        iRow += 1

                    Next
                    xSpace += (column.Width * 0.00025)

                End If
                iColumn += 1

            Next
            If Table_opaque = True Or Table_frame = True Then

                PW.SIS.CreateRectangle(0, 0, xSpace, (-ySpace * iRow) - 0.003)
                If Table_opaque = True Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 50</RGBA></Colour></Brush>")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Hollow</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")

                End If
                If Table_frame = True Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour><Width>" & Str(DGV_OverlaySchema.Font.Size * 2) & "</Width></Pen>")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")

                End If

            End If
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            PW.SIS.UpdateItem()
            Me.DialogResult = System.Windows.Forms.DialogResult.OK

        Catch
        End Try
        Me.Close()

    End Sub

End Class