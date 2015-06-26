Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System
Imports System.IO
Imports System.Data
Imports System.Windows.Forms
Imports System.Math

Public Class Dialog_Book

    Public Sub New()

        Try

            InitializeComponent()
            PW.Handler = PW.SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$")
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", True)

            'Create list from selected items
            PW.SIS.CreateListFromSelection("lSelectedItems")
            If PW.SIS.GetListSize("lSelectedItems") > 0 Then

                CB_EXTENT.Enabled = True
                CB_EXTENT.Items.Add("Extent of selected items")
                If PW.SIS.GetListSize("lSelectedItems") > 1 Then

                    CB_EXTENT.Items.Add("One page per item")
                    PW.SIS.OpenList("lSelectedItems", 0)
                    Dim iOverlay = PW.SIS.FindDatasetOverlay(PW.SIS.GetDataset(), -1, True)
                    If Not iOverlay = -1 Then

                        PW.SIS.GetOverlaySchema(iOverlay, "tmpSchema")
                        PW.SIS.LoadSchema("tmpSchema")
                        For i = 0 To PW.SIS.GetInt(SIS_OT_SCHEMA, 0, "_nColumns&") - 1

                            CB_FORMULA.Items.Add(PW.SIS.GetStr(SIS_OT_SCHEMACOLUMN, i, "_formula$"))
                            CB_ORDERBY.Items.Add(PW.SIS.GetStr(SIS_OT_SCHEMACOLUMN, i, "_formula$"))

                        Next

                    End If

                End If

            End If

            'Populate Templates
            Dim UniqueTemplateNames As New ArrayList
            Dim TemplateList As String = PW.SIS.NolCatalog("APrintTemplate", CInt(False))
            Dim TemplateNames() As String = Split(TemplateList, vbTab)
            For Each TemplateName As String In TemplateNames

                If TemplateName <> "" Then UniqueTemplateNames.Add(TemplateName)

            Next
            UniqueTemplateNames.Sort()
            For Each TemplateName As String In UniqueTemplateNames

                CB_TEMPLATES.Items.Add(TemplateName)

            Next

            'Populate Grids
            aGrids = New Integer(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1, 2) {}
            For i = 0 To PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                Try

                    If PW.SIS.GetStr(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_class$") = "AInternalDts" Or PW.SIS.GetStr(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_class$") = "ABdsDts" And PW.SIS.GetInt(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_nItems&") < 1000 Then

                        PW.SIS.CreateListFromOverlay(i, "lOverlay")
                        PW.SIS.OpenList("lOverlay", 0)
                        PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_mode&")
                        CB_GRIDS.Items.Add(PW.SIS.GetStr(SIS_OT_OVERLAY, i, "_name$"))
                        aGrids(CB_GRIDS.Items.Count - 1, 0) = i
                        aGrids(CB_GRIDS.Items.Count - 1, 1) = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
                        aGrids(CB_GRIDS.Items.Count - 1, 2) = PW.SIS.GetDataset()

                    End If

                Catch
                End Try

            Next i
            If CB_GRIDS.Items.Count > 0 Then CB_GRIDS.SelectedIndex = 0
            If CB_GRIDS.Items.Count > 0 Then CHK_NEWGRID.Enabled = True
            If CB_GRIDS.Items.Count > 1 Then CB_GRIDS.Enabled = True

            'Set New Grid
            If CB_GRIDS.Text = "" Then

                CHK_NEWGRID.Checked = True
                TB_GRID.Text = "New Grid"
                CB_TEMPLATES.SelectedIndex = 0
                TB_SCALE.Text = Round(PW.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#") / 5 / (10 ^ Int(Log10(PW.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#") / 5))), 0) * (10 ^ Int(Log10(PW.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#") / 5)))
                CB_OVERLAP.SelectedIndex = 0

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private x1, y1, x2, y2, z, x_photo, y_photo, x_photo_template, y_photo_template, x_Count, y_Count, x_adaptive, y_adaptive, quad As Double

    Private PaperFormat, PaperWidth, PaperHeight, TemplateIndexBrush, TemplateIndexPen, ViewExtent As String

    Private BookplotMode, GridOverlay, iScale, Overlap, TemplateFrame, TemplateTitle, TemplateScale, TemplatePage, TemplatePages, TemplatePageOfPP, TemplatePageNorth, TemplatePageEast, TemplatePageSouth, TemplatePageWest, TemplateIndex As Integer

    Private aGrids(,) As Integer

    Private Sub CB_GRIDS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_GRIDS.SelectedIndexChanged

        CB_GRIDS_Change()

    End Sub

    Private Sub CB_TEMPLATES_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_TEMPLATES.SelectedIndexChanged

        GetTemplate()
        If CHK_NEWGRID.Checked = True Then

            x_photo = x_photo_template * iScale
            y_photo = y_photo_template * iScale
            Pages.Text = CalcGrid()

        End If

    End Sub

    Private Sub TB_SCALE_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TB_SCALE.KeyPress

        Try

            If Asc(e.KeyChar) <> 8 Then
                If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                    e.Handled = True
                End If
            End If

        Catch ex As Exception
        End Try

    End Sub

    Private Sub TB_SCALE_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TB_SCALE.TextChanged

        Try

            iScale = Convert.ToInt32(TB_SCALE.Text)
            x_photo = x_photo_template * iScale
            y_photo = y_photo_template * iScale
            If CHK_NEWGRID.Checked = True Then Pages.Text = CalcGrid()

        Catch
        End Try

    End Sub

    Private Sub CB_GRIDS_Change()

        Try

            PW.SIS.OpenExistingDatasetItem(aGrids(CB_GRIDS.SelectedIndex, 2), aGrids(CB_GRIDS.SelectedIndex, 1))
            GridOverlay = aGrids(CB_GRIDS.SelectedIndex, 0)
            btnCreate.Text = "Create Bookplot"
            PW.SIS.CreateLayerFilter("fShade", "-frames +shade")
            PW.SIS.CreateLayerFilter("fFrames", "-shade +frames")
            PW.SIS.SetOverlayFilter(GridOverlay, "fFrames")
            PW.SIS.CreateListFromOverlay(GridOverlay, "lGrid")
            Pages.Text = PW.SIS.GetListSize("lGrid")
            TB_GRID.Text = PW.SIS.GetStr(SIS_OT_OVERLAY, GridOverlay, "_name$")
            TB_GRID.Visible = False
            BookplotMode = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_mode&")
            If BookplotMode = 0 Or BookplotMode = 1 Then

                CHK_RENUMBER.Enabled = True

            Else

                CHK_RENUMBER.Enabled = False
                CHK_RENUMBER.Checked = False

            End If
            CB_TEMPLATES.Text = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "BP_template$")
            If CheckTemplate() = False Then

                MsgBox("Template map frame doesn't match existing grid.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
                btnCreate.Enabled = False

            End If
            CB_TEMPLATES.Enabled = False
            PW.SIS.OpenList("lGrid", 0)
            TB_SCALE.Text = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_scale&").ToString
            TB_SCALE.Enabled = False
            CB_EXTENT.Visible = False
            LCB_FORMULA.Visible = False
            CB_FORMULA.Visible = False
            LCB_ORDERBY.Visible = False
            CB_ORDERBY.Visible = False
            CHK_NEWGRID.Checked = False

        Catch
        End Try

    End Sub

    Private Sub CHK_NEWGRID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHK_NEWGRID.CheckedChanged

        Try

            If CHK_NEWGRID.Checked = True Then

                TB_GRID.Visible = True
                CB_GRIDS.Visible = False
                CB_EXTENT.Visible = True
                CB_EXTENT.SelectedIndex = 0
                LCB_OVERLAP.Visible = True
                CB_OVERLAP.Visible = True
                CB_OVERLAP.SelectedIndex = 0
                CHK_AdaptiveOverlap.Visible = True
                CB_TEMPLATES.Enabled = True
                TB_SCALE.Enabled = True
                If CB_EXTENT.Items.Count > 1 Then

                    CB_EXTENT.Enabled = True
                    CB_EXTENT.SelectedIndex = 1

                End If
                CHK_RENUMBER.Visible = False
                CHK_RENUMBER.Enabled = False
                btnCreate.Text = "Create Grid"

            Else

                TB_GRID.Visible = False
                LCB_FORMULA.Visible = False
                CB_FORMULA.Visible = False
                LCB_ORDERBY.Visible = False
                CB_ORDERBY.Visible = False
                CHK_AdaptiveOverlap.Visible = False
                CB_OVERLAP.Visible = False
                LCB_OVERLAP.Visible = False
                CHK_RENUMBER.Visible = True
                CB_GRIDS.Visible = True
                CB_GRIDS_Change()

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub CB_EXTENT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_EXTENT.SelectedIndexChanged

        BookplotMode = CB_EXTENT.SelectedIndex

        Select Case BookplotMode

            Case 0, 1

                LCB_OVERLAP.Visible = True
                CB_OVERLAP.Visible = True
                CHK_AdaptiveOverlap.Visible = True
                LCB_FORMULA.Visible = False
                CB_FORMULA.Visible = False
                LCB_ORDERBY.Visible = False
                CB_ORDERBY.Visible = False
                Pages.Text = CalcGrid()

            Case 2

                LCB_OVERLAP.Visible = False
                CB_OVERLAP.Visible = False
                CHK_AdaptiveOverlap.Checked = False
                CHK_AdaptiveOverlap.Visible = False
                LCB_FORMULA.Visible = True
                CB_FORMULA.Visible = True
                LCB_ORDERBY.Visible = True
                CB_ORDERBY.Visible = True
                Pages.Text = CalcGrid()

        End Select

    End Sub

    Private Sub CHK_AdaptiveOverlap_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CHK_AdaptiveOverlap.CheckedChanged

        If CHK_AdaptiveOverlap.Checked = True Then Pages.Text = CalcGrid()

    End Sub

    Private Sub CB_OVERLAP_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_OVERLAP.SelectedIndexChanged

        If Convert.ToInt32(CB_OVERLAP.Text) > 0 Then

            Overlap = Convert.ToInt32(CB_OVERLAP.Text)
            Pages.Text = CalcGrid()

        ElseIf CHK_AdaptiveOverlap.Checked = False And Val(CB_OVERLAP.Text) = 0 Then

            Overlap = 0
            Pages.Text = CalcGrid()

        End If

    End Sub

    Private Sub GetTemplate()

        Try

            PW.SIS.DefineNolView("vTemplate")
            PW.SIS.Compose()
            PW.SIS.PlacePrintTemplate(CB_TEMPLATES.Text, 0, 1)
            PW.SIS.CreateListFromOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1, "lTemplate")
            PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetViewExtent())
            PaperWidth = Abs(x1 - x2) * 100
            PaperHeight = Abs(y1 - y2) * 100
            PaperFormat = Str(Abs(x1 - x2) * 100) & "x" & Str(Abs(y1 - y2) * 100) & "cm"
            PW.SIS.DeselectAll()
            TemplateTitle = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""Title""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplateTitle = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplateScale = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""Scale""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplateScale = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePage = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""SheetN""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePage = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePageOfPP = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""SheetN_of_NN""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePageOfPP = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePageNorth = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""p  North""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePageNorth = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePageEast = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""p  East""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePageEast = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePageSouth = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""p  South""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePageSouth = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplatePageWest = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""p  West""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplatePageWest = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")

            End If
            TemplateIndex = -1
            PW.SIS.CreatePropertyFilter("fFurniture", "Prompt$=""Index""")
            PW.SIS.ScanList("lFurniture", "lTemplate", "fFurniture", "")
            If PW.SIS.GetListSize("lFurniture") > 0 Then

                PW.SIS.OpenList("lFurniture", 0)
                TemplateIndex = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
                TemplateIndexBrush = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "_brush$")
                TemplateIndexPen = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "_pen$")

            End If
            PW.SIS.ScanList("lPhoto", "lTemplate", "fPhoto", "")
            PW.SIS.OpenList("lPhoto", 0)
            TemplateFrame = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
            PW.SIS.SelectItem()
            PW.SIS.DoCommand("AComBoundary")
            PW.SIS.DoCommand("AComFillGeometry")
            x_photo_template = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
            y_photo_template = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            PW.SIS.DefineNolShape("tmpShape", "lPhoto", PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#"), PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#"), 0, 1)
            PW.SIS.RemoveOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            PW.SIS.RecallNolView("vTemplate")

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Function CheckTemplate() As Boolean

        Try

            PW.SIS.OpenList("lGrid", 0)
            Dim grid_ratio = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#") / PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            Dim photo_ratio = x_photo_template / y_photo_template
            If Abs(grid_ratio - photo_ratio) < 0.001 Then

                Return True

            Else

                Return False

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)
            Return False

        End Try

    End Function

    Private Function CalcGrid()

        Try

            Select Case BookplotMode

                Case 0, 1

                    If BookplotMode = 1 And PW.SIS.GetListSize("lSelectedItems") > 0 Then

                        PW.SIS.SelectList("lSelectedItems")
                        PW.SIS.DoCommand("AComZoomSelect")

                    End If
                    PW.SIS.CreateInternalOverlay("tmp", 0)
                    PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
                    ViewExtent = PW.SIS.GetViewExtent()
                    PW.SIS.RecallNolView("vOrg")
                    PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, ViewExtent)
                    PW.SIS.CreateRectangle(x1, y1, x2, y2)
                    Dim sx = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                    Dim sy = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                    PW.SIS.CreatePoint(x1, y1, z, "", 0, 1)
                    x1 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                    y1 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                    PW.SIS.CreatePoint(x2, y2, z, "", 0, 1)
                    x2 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                    y2 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                    PW.SIS.RemoveOverlay(0)

                    'calculate overlap
                    Dim a = 4
                    Dim b = -((2 * x_photo) + (2 * y_photo))
                    Dim c = x_photo * y_photo * Overlap / 100
                    quad = (-b - Math.Sqrt(b ^ 2 - 4 * a * c)) / (2 * a)

                    'calculate count
                    x_Count = sx / x_photo
                    y_Count = sy / y_photo

                    'add overlap to count
                    x_Count = (sx + Floor(x_Count) * quad) / x_photo
                    y_Count = (sy + Floor(y_Count) * quad) / y_photo

                    'adaptive overlap
                    x_adaptive = 0
                    y_adaptive = 0
                    If CHK_AdaptiveOverlap.Checked = True And x_Count > 1 Then x_adaptive = (Ceiling(x_Count) * x_photo - Floor(x_Count) * quad - Abs(x1 - x2)) / Floor(x_Count)
                    If CHK_AdaptiveOverlap.Checked = True And y_Count > 1 Then y_adaptive = (Ceiling(y_Count) * y_photo - Floor(y_Count) * quad - Abs(y1 - y2)) / Floor(y_Count)

                    'grid size info
                    Return Ceiling(x_Count).ToString & " * " & Ceiling(y_Count).ToString

                Case 2

                    Return PW.SIS.GetListSize("lSelectedItems").ToString

                Case Else

                    Return Nothing

            End Select

        Catch ex As Exception

            MsgBox(ex.ToString)
            Return Nothing

        End Try

    End Function

    Private Sub getViewXandY()

        Try



        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click

        Try

            Convert.ToInt32(TB_SCALE.Text)
            Select Case btnCreate.Text

                Case "Create Grid"

                    CreateGrid()

                Case "Create Bookplot"

                    CreateBookplot()

            End Select

        Catch ex As Exception

            MsgBox("Scale must be a positive integer.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

        End Try

    End Sub

    Private Sub CreateGrid()

        Try

            'check for existing grid overlay
            If TB_GRID.Text = CB_GRIDS.Text Then

                If MsgBox("Overwrite " & CB_GRIDS.Text & " overlay?", MsgBoxStyle.OkCancel, Title:="Grid overwrite") = MsgBoxResult.Cancel Then Exit Sub
                PW.SIS.RemoveOverlay(GridOverlay)

            Else

                GridOverlay = PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&")

            End If
            PW.SIS.CreateInternalOverlay(TB_GRID.Text, GridOverlay)
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", GridOverlay)
            PW.SIS.SetFlt(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, GridOverlay, "_nDataset&"), "_scale#", 1)

            'check whether the new grid is page per item grid (mode = 2)
            Select Case CB_EXTENT.SelectedIndex

                Case 0, 1

                    Dim startX = x1 + x_photo / 2
                    Dim startY = y2 - y_photo / 2
                    PW.EmptyList("lFrames")
                    For y_loop = 0 To Ceiling(y_Count) - 1

                        For x_loop = 0 To Ceiling(x_Count) - 1

                            PW.SIS.CreatePoint(0, 0, 0, "tmpShape", 0, Convert.ToInt32(TB_SCALE.Text) + 1)
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", startX + (x_photo * x_loop) - (quad * x_loop) - (x_adaptive * x_loop))
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", startY - (y_photo * y_loop) + (quad * y_loop) + (y_adaptive * y_loop))
                            PW.SIS.UpdateItem()
                            PW.SIS.SelectItem()
                            PW.SIS.DoCommand("AComExplodeShape")
                            PW.SIS.OpenSel(0)
                            PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_scale&", Convert.ToInt32(TB_SCALE.Text))
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "BP_title$", TB_GRID.Text)
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "BP_template$", CB_TEMPLATES.Text)
                            PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_mode&", CB_EXTENT.SelectedIndex)
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Hollow</Style><Colour><RGBA>255 255 255 0</RGBA></Colour><BackgroundColour><RGBA>255 255 0 255</RGBA></BackgroundColour></Brush>")
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_layer$", "frames")
                            PW.SIS.UpdateItem()
                            PW.SIS.AddToList("lFrames")

                        Next

                    Next
                    Dim xi1, yi1, xi2, yi2 As Double
                    PW.SIS.SplitExtent(xi1, yi1, z, xi2, yi2, z, PW.SIS.GetListExtent("lFrames"))
                    PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, ViewExtent)
                    PW.SIS.MoveList("lFrames", -(Abs(xi2 - x2) / 2), +(Abs(yi1 - y1) / 2), 0, 0, 1)

                    'check for items
                    If CB_EXTENT.SelectedIndex = 1 Then

                        PW.SIS.CreateBoolean("lSelectedItems", SIS_BOOLEAN_OR)
                        PW.SIS.CreateLocusFromItem("Locus", SIS_GT_DISJOINT, SIS_GM_GEOMETRY)
                        PW.SIS.DeleteItem()
                        If PW.SIS.ScanList("lDisjoint", "lFrames", "", "Locus") > 0 Then PW.SIS.Delete("lDisjoint")

                    End If

                    'call function to number pages
                    NumberPages("lFrames", y_photo, x_photo)

                Case 2

                    Dim ppi_scale As Double
                    Dim page As Integer = 1
                    PW.EmptyList("lFrames")

                    'create list and add items in order to enable ordering of items
                    Dim lOrderbyList As New List(Of OrderbyObject)()
                    For i As Integer = 0 To PW.SIS.GetListSize("lSelectedItems") - 1

                        PW.SIS.OpenList("lSelectedItems", i)
                        Dim orderby As String
                        If Not CB_ORDERBY.Text = "" Then

                            Try

                                orderby = PW.SIS.Evaluate(SIS_OT_CURITEM, 0, CB_ORDERBY.Text)

                            Catch ex As Exception

                                orderby = ""

                            End Try

                            lOrderbyList.Add(New OrderbyObject(i, orderby))

                        End If

                        'get location and extent to determine fitting scale
                        Dim sx = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                        Dim sy = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")

                        'get the origin of multipolygons
                        Dim llX, llY, urX, urY As Double
                        PW.SIS.SplitExtent(llX, llY, z, urX, urY, z, PW.SIS.GetExtent())
                        PW.SIS.CreateRectangle(llX, llY, urX, urY)
                        Dim ox = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                        Dim oy = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                        PW.SIS.DeleteItem()

                        If sx = 0 And sy = 0 Then

                            ppi_scale = 1

                        ElseIf sx / x_photo > sy / y_photo Then

                            ppi_scale = sx / x_photo

                        Else

                            ppi_scale = sy / y_photo

                        End If
                        If ppi_scale < 1 Then ppi_scale = 1

                        'create and explode frame in order to assign bp properties to frame items
                        PW.SIS.CreatePoint(0, 0, 0, "tmpShape", 0, ppi_scale * Convert.ToInt32(TB_SCALE.Text) * PW.SIS.GetListItemFlt("lSelectedItems", i, "_scale#"))
                        PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", ox)
                        PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", oy)
                        PW.SIS.UpdateItem()
                        PW.SIS.SelectItem()
                        PW.SIS.DoCommand("AComExplodeShape")
                        PW.SIS.OpenSel(0)
                        PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_id&", i)
                        PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_page&", page)
                        page += 1
                        PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_scale&", Round((PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#") / x_photo) * Convert.ToInt32(TB_SCALE.Text)))
                        PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                        Try

                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "BP_title$", PW.SIS.GetListItemStr("lSelectedItems", i, CB_FORMULA.Text))

                        Catch

                            PW.SIS.SetStr(SIS_OT_CURITEM, 0, "BP_title$", TB_GRID.Text)

                        End Try
                        PW.SIS.SetStr(SIS_OT_CURITEM, 0, "BP_template$", CB_TEMPLATES.Text)
                        PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_mode&", CB_EXTENT.SelectedIndex)
                        PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Hollow</Style><Colour><RGBA>255 255 255 0</RGBA></Colour><BackgroundColour><RGBA>255 255 0 255</RGBA></BackgroundColour></Brush>")
                        PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                        PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_layer$", "frames")
                        PW.SIS.UpdateItem()
                        PW.SIS.AddToList("lFrames")
                        PW.SIS.GetListSize("lFrames")
                    Next

                    'order list items and assign page numbers
                    If Not CB_ORDERBY.Text = "" Then

                        Dim sorted = From s In lOrderbyList Order By s.OrderValue Select s
                        page = 1
                        For Each Frame As OrderbyObject In sorted

                            PW.SIS.CreatePropertyFilter("fID", "BP_id& =" & Str(Frame.ID))
                            Dim r = PW.SIS.ScanList("lFrame", "lFrames", "fID", "")
                            PW.SIS.OpenList("lFrame", 0)
                            PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_page&", page)
                            page += 1

                        Next

                    End If

            End Select
            Me.Close()

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Class OrderbyObject

        Public ID As String
        Public OrderValue As String
        Public Page As Integer
        Public Sub New(ByVal iID As String, ByVal sOrderValue As String)

            Me.ID = iID
            Me.OrderValue = sOrderValue

        End Sub

    End Class

    Private Sub NumberPages(ByVal list As String, ByVal y_step As Double, ByVal x_step As Double)

        Try

            'create list of grid items
            PW.SIS.DeselectAll()
            PW.SIS.SelectList(list)
            PW.SIS.CreateListFromSelection("list")

            'get the ox and oy extent of the list
            PW.SIS.CreateInternalOverlay("tmp", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            Dim xi1, yi1, xi2, yi2 As Double
            PW.SIS.SplitExtent(xi1, yi1, z, xi2, yi2, z, PW.SIS.GetListExtent("list"))
            PW.SIS.CreatePoint(xi1, yi1, z, "", 0, 1)
            xi1 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            yi1 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            PW.SIS.DeleteItem()
            PW.SIS.CreatePoint(xi2, yi2, z, "", 0, 1)
            xi2 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            yi2 = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            PW.SIS.DeleteItem()
            PW.SIS.CopyListItems("list")
            PW.SIS.CreateBoolean("list", SIS_BOOLEAN_OR)
            PW.SIS.Delete("list")
            PW.SIS.CreateListFromOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "list")
            PW.SIS.DecomposeGeometry("list")
            PW.SIS.CreateListFromOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "list")

            'scan first for from top (yi2) then from the left (xi1)
            Dim page As Integer = 1
            Dim y_scanline = yi2 - y_step / 2
            Do While y_scanline > yi1

                If PW.SIS.GetAxesType = SIS_AXES_SPHERICAL Then

                    PW.SIS.CreatePoint(0, 0, 0, "", 0, 1)
                    PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", xi1)
                    PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", y_scanline)
                    PW.SIS.UpdateItem()
                    Dim latStart = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                    Dim lonStart = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")
                    PW.SIS.DeleteItem()
                    PW.SIS.CreatePoint(0, 0, 0, "", 0, 1)
                    PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", xi2)
                    PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", y_scanline)
                    PW.SIS.UpdateItem()
                    Dim latEnd = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                    Dim lonEnd = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")
                    PW.SIS.DeleteItem()
                    PW.SIS.MoveTo(lonStart, latStart, 0)
                    PW.SIS.LineTo(lonEnd, latEnd, 0)

                Else

                    PW.SIS.MoveTo(xi1, y_scanline, 0)
                    PW.SIS.LineTo(xi2, y_scanline, 0)

                End If

                PW.SIS.StoreAsLine()
                PW.SIS.CreateLocusFromItem("linelocus", SIS_GM_GEOMETRY, SIS_GT_INTERSECT)
                PW.SIS.DeleteItem()
                If PW.SIS.ScanList("yscan", "list", "", "linelocus") > 0 Then

                    Dim x_scanline = xi1 + x_step / 2
                    Do While x_scanline < xi2

                        If PW.SIS.GetAxesType = SIS_AXES_SPHERICAL Then

                            PW.SIS.CreatePoint(0, 0, 0, "", 0, 1)
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", x_scanline)
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", yi1)
                            PW.SIS.UpdateItem()
                            Dim latStart = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                            Dim lonStart = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")
                            PW.SIS.DeleteItem()
                            PW.SIS.CreatePoint(0, 0, 0, "", 0, 1)
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", x_scanline)
                            PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", yi2)
                            PW.SIS.UpdateItem()
                            Dim latEnd = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                            Dim lonEnd = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")
                            PW.SIS.DeleteItem()
                            PW.SIS.MoveTo(lonStart, latStart, 0)
                            PW.SIS.LineTo(lonEnd, latEnd, 0)

                        Else

                            PW.SIS.MoveTo(x_scanline, yi1, 0)
                            PW.SIS.LineTo(x_scanline, yi2, 0)

                        End If

                        PW.SIS.StoreAsLine()
                        PW.SIS.CreateLocusFromItem("linelocus", SIS_GM_GEOMETRY, SIS_GT_INTERSECT)
                        PW.SIS.DeleteItem()
                        If PW.SIS.ScanList("xscan", "yscan", "", "linelocus") > 0 Then

                            For i = 0 To PW.SIS.GetListSize("xscan") - 1

                                PW.SIS.OpenList("xscan", i)
                                PW.SIS.CreateLocusFromItem("arealocus", SIS_GM_GEOMETRY, SIS_GT_INTERSECT)
                                If PW.SIS.ScanList("areascan", list, "", "arealocus") > 0 Then

                                    For ii = 0 To PW.SIS.GetListSize("areascan") - 1

                                        PW.SIS.OpenList("areascan", ii)
                                        PW.SIS.SetInt(SIS_OT_CURITEM, 0, "BP_page&", page)
                                        page += 1

                                    Next

                                End If

                            Next
                            PW.SIS.Delete("xscan")

                        End If
                        x_scanline += x_step

                    Loop

                End If
                y_scanline -= y_step

            Loop
            PW.SIS.RemoveOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"))

        Catch ex As Exception
        End Try

    End Sub

    Private Sub CreateBookplot()

        Try

            'renumber pages if necessary
            If CHK_RENUMBER.Checked = True Then NumberPages("lGrid", y_photo, x_photo)

            'run checks on the items in the grid list
            If Not CheckGrid() = True Then Exit Sub

            'save pdf file dialog
            Dim dlgFile As New SaveFileDialog
            dlgFile.AddExtension = True
            dlgFile.DefaultExt = ".pdf"
            dlgFile.Filter = "PDF Document(*.pdf)|*.pdf|JPEG Image(*.jpg)|*.jpg|GIF Image(*.gif)|*.gif"
            dlgFile.FilterIndex = 0
            dlgFile.OverwritePrompt = True
            dlgFile.Title = "Bookplot PDF"
            If Not dlgFile.ShowDialog() = Windows.Forms.DialogResult.OK Then Exit Sub
            Dim FileInfo As New FileInfo(dlgFile.FileName)
            Dim ExportFormat As String = FileInfo.Extension
            Dim FileName As String = FileInfo.Name.Remove(FileInfo.Name.Length - 4, 4)
            If FileIsLocked(FileInfo.FullName) = True Then Exit Sub

            'create table
            CreateTable()

            'creating bookplot
            PW.SIS.SetInt(SIS_OT_OVERLAY, GridOverlay, "_status&", 3)
            PW.SIS.SetOverlayFilter(GridOverlay, "fShade")
            PW.SIS.Compose()
            PW.SIS.SetOverlayFilter(GridOverlay, "")
            PW.SIS.DoCommand("AComFileNewSwd")
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            PW.SIS.PlacePrintTemplate(CB_TEMPLATES.Text, 0, Convert.ToInt32(TB_SCALE.Text))

            'remove furniture from PPI plots
            If BookplotMode = 2 Then RemoveFurniture()

            'call index creation function
            CreateIndex()

            'merge pages
            Dim MergeProcess As PdfMerge = New PdfMerge
            If FileInfo.Extension = ".pdf" Then My.Computer.FileSystem.CreateDirectory(FileInfo.DirectoryName & "\tmp_nolcatalogue")
            ProgressBar.Visible = True
            ProgressBar.Maximum = lBookplotFrames.Count
            ProgressBar.Value = 0
            Dim sorted = From s In lBookplotFrames Order By s.Page Select s
            Dim page As Integer = 0
            For Each Frame As BookplotFrame In sorted

                PW.SIS.OpenItem(TemplateFrame)
                PW.SetPhotoCentre(Frame.X, Frame.Y)

                'set ppi scale
                If BookplotMode = 2 Then PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoscale#", Frame.Scale)
                PW.SIS.UpdateItem()

                'call function to set page numbers
                page += 1
                SetFurnitureNew(Frame, page)
                Select Case FileInfo.Extension

                    Case ".jpg"

                        My.Computer.FileSystem.CreateDirectory(FileInfo.DirectoryName & "\" & FileName)
                        PW.SIS.ExportRaster("JPEG_GDALExporter", FileInfo.DirectoryName & "\" & FileName & "\" & FileName & Str(Frame.Page) & ".jpg", "WORLDFILE=FALSE,QUALITY=75,width=" & Str(PaperWidth * 100) & ",height=" & Str(PaperHeight * 100))
                        Dim myFile As String
                        For Each myFile In Directory.GetFiles(FileInfo.DirectoryName & "\" & FileName, "*.xml")
                            File.Delete(myFile)
                        Next

                    Case ".gif"

                        My.Computer.FileSystem.CreateDirectory(FileInfo.DirectoryName & "\" & FileName)
                        PW.SIS.ExportRaster("GIF_GDALExporter", FileInfo.DirectoryName & "\" & FileName & "\" & FileName & Str(Frame.Page) & ".gif", "WORLDFILE=FALSE,width=" & Str(PaperWidth * 100) & ",height=" & Str(PaperHeight * 100))
                        Dim myFile As String
                        For Each myFile In Directory.GetFiles(FileInfo.DirectoryName & "\" & FileName, "*.xml")
                            File.Delete(myFile)
                        Next

                    Case ".pdf"

                        PW.SIS.ExportPdf(FileInfo.DirectoryName & "\tmp_nolcatalogue\" & Str(Frame.Page) & ".pdf", PaperFormat, "")

                        'add page to the pdf merge stack
                        MergeProcess.AddDocument(FileInfo.DirectoryName & "\tmp_nolcatalogue\" & Str(Frame.Page) & ".pdf")

                    Case Else

                        MsgBox("Unsuported export format " & FileInfo.Extension, MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
                        PW.SIS.SwdClose(SIS_NOSAVE)
                        PW.ActivateWindow(PW.Handler)
                        PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)
                        Exit Sub

                End Select
                ProgressBar.Value += 1
                Me.Refresh()

            Next

            If FileInfo.Extension = ".pdf" Then

                MergeProcess.Merge(FileInfo.DirectoryName & "\tmp_nolcatalogue\merge.pdf", PaperWidth, PaperHeight)
                MergeProcess.AddPageLinks(FileInfo, lBookplotFrames, TemplateFrame, TemplateIndex)
                My.Computer.FileSystem.DeleteDirectory(FileInfo.DirectoryName & "\tmp_nolcatalogue", FileIO.DeleteDirectoryOption.DeleteAllContents)

            End If
            PW.SIS.SwdClose(SIS_NOSAVE)
            PW.ActivateWindow(PW.Handler)
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)
            Me.Close()

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Function CheckGrid() As Boolean

        Try

            For i = 0 To PW.SIS.GetListSize("lGrid") - 1

                PW.SIS.OpenList("lGrid", i)

                Dim Mode = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_mode&")
                Dim Page = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")
                Dim Scale = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_scale&")
                Dim Title = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "BP_title$")

            Next i

            Return True

        Catch ex As Exception

            MsgBox("The grid schema is incomplete.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
            Return False

        End Try

    End Function

    Public Function FileIsLocked(ByVal strFullFileName As String) As Boolean

        Dim fs As System.IO.FileStream
        Try

            fs = System.IO.File.Open(strFullFileName, IO.FileMode.OpenOrCreate, IO.FileAccess.Read, IO.FileShare.None)
            fs.Close()
            File.Delete(strFullFileName)
            Return False

        Catch ex As System.IO.IOException

            MsgBox("The chosen file is locked.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
            Return True

        End Try

    End Function

    Private Sub RemoveFurniture()

        Try
            PW.EmptyList("lFurniture")
            If Not TemplatePageNorth = -1 Then

                PW.SIS.OpenItem(TemplatePageNorth)
                PW.SIS.AddToList("lFurniture")
                TemplatePageNorth = -1

            End If
            If Not TemplatePageEast = -1 Then

                PW.SIS.OpenItem(TemplatePageEast)
                PW.SIS.AddToList("lFurniture")
                TemplatePageEast = -1

            End If
            If Not TemplatePageSouth = -1 Then

                PW.SIS.OpenItem(TemplatePageSouth)
                PW.SIS.AddToList("lFurniture")
                TemplatePageSouth = -1

            End If
            If Not TemplatePageWest = -1 Then

                PW.SIS.OpenItem(TemplatePageWest)
                PW.SIS.AddToList("lFurniture")
                TemplatePageWest = -1

            End If
            If Not TemplateIndex = -1 Then

                PW.SIS.OpenItem(TemplateIndex)
                PW.SIS.AddToList("lFurniture")
                TemplateIndex = -1

            End If
            PW.SIS.Delete("lFurniture")

        Catch
        End Try

    End Sub

    Public Class BookplotFrame

        Public ID As Integer
        Public X As Double
        Public Y As Double
        Public iX As Double
        Public iY As Double
        Public Page As Integer
        Public pUp As Integer
        Public pDown As Integer
        Public pLeft As Integer
        Public pRight As Integer
        Public Scale As Integer
        Public Title As String

        Public Sub New(ByVal ID As Integer, ByVal X As Double, ByVal Y As Double, ByVal iX As Double, ByVal iY As Double, ByVal Page As Integer, ByVal pUp As Integer, ByVal pDown As Integer, ByVal pLeft As Integer, ByVal pRight As Integer, ByVal Scale As Integer, ByVal Title As String)

            Me.ID = ID
            Me.X = X
            Me.Y = Y
            Me.iX = iX
            Me.iY = iY
            Me.Page = Page
            Me.pUp = pUp
            Me.pDown = pDown
            Me.pLeft = pLeft
            Me.pRight = pRight
            Me.Scale = Scale
            Me.Title = Title

        End Sub

    End Class

    Public lBookplotFrames As List(Of BookplotFrame)

    Private Sub CreateTable()

        Try

            lBookplotFrames = New List(Of BookplotFrame)
            Dim frame_ratio = x_photo_template / y_photo_template
            PW.SIS.CreateInternalOverlay("tmp", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            For i = 0 To PW.SIS.GetListSize("lGrid") - 1

                PW.SIS.OpenList("lGrid", i)
                Dim ID = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
                Dim X As Double
                Dim Y As Double
                If PW.SIS.GetAxesType = SIS_AXES_SPHERICAL Then

                    Y = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                    X = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")

                Else

                    X = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                    Y = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")

                End If
                Dim iX = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim iY = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Dim Page = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")
                Dim Scale = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_scale&")
                Dim Title = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "BP_title$")

                'check top neighbour
                PW.SIS.CreatePoint(X, Y, 0, "Triangle Up", 0, 1)
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") + (0.6 * y_photo))
                PW.SIS.UpdateItem()
                PW.SIS.CreateLocusFromItem("Locus", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                Dim pUp As Integer = -1
                If PW.SIS.ScanList("lNeighbour", "lGrid", "", "Locus") > 0 Then

                    PW.SIS.OpenList("lNeighbour", 0)
                    pUp = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")

                End If

                'check bottom neighbour
                PW.SIS.CreatePoint(X, Y, 0, "Triangle Down", 0, 1)
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oy#", PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#") - (0.6 * y_photo))
                PW.SIS.UpdateItem()
                PW.SIS.CreateLocusFromItem("Locus", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                Dim pDown As Integer = -1
                If PW.SIS.ScanList("lNeighbour", "lGrid", "", "Locus") > 0 Then

                    PW.SIS.OpenList("lNeighbour", 0)
                    pDown = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")

                End If

                'check left neighbour
                PW.SIS.CreatePoint(X, Y, 0, "Triangle Left", 0, 1)
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#") - (0.6 * x_photo))
                PW.SIS.UpdateItem()
                PW.SIS.CreateLocusFromItem("Locus", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                Dim pLeft As Integer = -1
                If PW.SIS.ScanList("lNeighbour", "lGrid", "", "Locus") > 0 Then

                    PW.SIS.OpenList("lNeighbour", 0)
                    pLeft = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")

                End If

                'check right neighbour
                PW.SIS.CreatePoint(X, Y, 0, "Triangle Right", 0, 1)
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_ox#", PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#") + (0.6 * x_photo))
                PW.SIS.UpdateItem()
                PW.SIS.CreateLocusFromItem("Locus", SIS_GT_INTERSECT, SIS_GM_GEOMETRY)
                Dim pRight As Integer = -1
                If PW.SIS.ScanList("lNeighbour", "lGrid", "", "Locus") > 0 Then

                    PW.SIS.OpenList("lNeighbour", 0)
                    pRight = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "BP_page&")

                End If

                'add frame to list
                lBookplotFrames.Add(New BookplotFrame(ID, X, Y, iX, iY, Page, pUp, pDown, pLeft, pRight, Scale, Title))

            Next i
            PW.SIS.RemoveOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"))

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub SetFurnitureNew(ByVal Frame As BookplotFrame, ByVal page As Integer)

        Try

            If Not TemplateTitle = -1 Then

                PW.SIS.OpenItem(TemplateTitle)
                Try

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Frame.Title)

                Catch

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", TB_GRID.Text)

                End Try
                PW.SIS.UpdateItem()

            End If
            If Not TemplateScale = -1 Then

                PW.SIS.OpenItem(TemplateScale)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.Scale))
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePage = -1 Then

                PW.SIS.OpenItem(TemplatePage)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.Page))
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePageOfPP = -1 Then

                PW.SIS.OpenItem(TemplatePageOfPP)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(page) & " of " & lBookplotFrames.Count)
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePageNorth = -1 Then

                PW.SIS.OpenItem(TemplatePageNorth)
                If Frame.pUp = -1 Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.pUp))

                End If
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePageSouth = -1 Then

                PW.SIS.OpenItem(TemplatePageSouth)
                If Frame.pDown = -1 Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.pDown))

                End If
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePageWest = -1 Then

                PW.SIS.OpenItem(TemplatePageWest)
                If Frame.pLeft = -1 Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.pLeft))

                End If
                PW.SIS.UpdateItem()

            End If
            If Not TemplatePageEast = -1 Then

                PW.SIS.OpenItem(TemplatePageEast)
                If Frame.pRight = -1 Then

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", "")

                Else

                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_text$", Str(Frame.pRight))

                End If
                PW.SIS.UpdateItem()

            End If
            If Not TemplateIndex = -1 Then

                PW.SIS.OpenList("lIndexFrames", page - 1)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", TemplateIndexPen)
                PW.SIS.UpdateItem()
                If Not page = 1 Then

                    PW.SIS.OpenList("lIndexFrames", page - 2)
                    PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                    PW.SIS.UpdateItem()

                End If

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub CreateIndex()

        Try

            If Not TemplateIndex = -1 Then

                PW.SIS.OpenItem(TemplateIndex)
                Dim x_index = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim y_index = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                Dim ox_index = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim oy_index = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                PW.SIS.DeleteItem()
                PW.SIS.OpenItem(TemplateFrame)
                Dim x_frame = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim y_frame = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                PW.EmptyList("lIndexFrames")

                For i As Integer = 0 To lBookplotFrames.Count - 1

                    PW.SIS.CreatePoint(lBookplotFrames(i).iX, lBookplotFrames(i).iY, 0, "tmpShape", 0, lBookplotFrames(i).Scale)
                    PW.SIS.UpdateItem()
                    PW.SIS.AddToList("lIndexFrames")

                Next

                PW.SIS.OpenList("lIndexFrames", 0)
                PW.SIS.AddToList("lIndexFrames")
                PW.SIS.ExplodeShape("lIndexFrames")
                PW.SIS.CreateListFromSelection("lIndexFrames")
                PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetListExtent("lIndexFrames"))
                Dim index_scale As Double
                If x_index / Abs(x1 - x2) > y_index / Abs(y1 - y2) Then

                    index_scale = y_index / Abs(y1 - y2)

                Else

                    index_scale = x_index / Abs(x1 - x2)

                End If
                PW.SIS.MoveList("lIndexFrames", 0, 0, 0, 0, index_scale)
                PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetListExtent("lIndexFrames"))
                PW.SIS.MoveList("lIndexFrames", ox_index - (x1 + x2) / 2, oy_index - (y1 + y2) / 2, 0, 0, 1)
                PW.SIS.SetListStr("lIndexFrames", "_brush$", TemplateIndexBrush)
                PW.SIS.SetListStr("lIndexFrames", "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

End Class