Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System
Imports System.IO
Imports System.Data
Imports System.Windows.Forms
Imports System.Math

<GisLinkProgram("PrintWorkshop")> Public Class PW

    Public Shared APP As SisApplication

    Private Shared _sis As Cadcorp.SIS.GisLink.Library.MapEditor

    Public Shared Property SIS As Cadcorp.SIS.GisLink.Library.MapEditor

        Get

            If _sis Is Nothing Then _sis = APP.TakeoverMapEditor
            Return _sis

        End Get
        Set(ByVal value As Cadcorp.SIS.GisLink.Library.MapEditor)

            _sis = value

        End Set

    End Property

    Private _currentRibbonButton As SisRibbonButton

    Public Sub New(ByVal SISApplication As SisApplication)
        APP = SISApplication
        SIS.CreateClassTreeFilter("fArea", "-Item +Photo +Area:")
        SIS.CreateClassTreeFilter("fPhoto", "-Item +Photo")
        SIS.CreateClassTreeFilter("fNoPhoto", "+Item -Photo")
        SIS.Dispose()
        SIS = Nothing

        Dim group As SisRibbonGroup = APP.RibbonGroup
        group.Text = "Print Workshop"

        'Place
        Dim SmallButton_PlaceTemplate As SisRibbonButton = New SisRibbonButton("Place Template", New SisClickHandler(AddressOf PlaceTemplate))
        SmallButton_PlaceTemplate.LargeImage = False
        SmallButton_PlaceTemplate.Icon = My.Resources.Btn_Place
        SmallButton_PlaceTemplate.Help = "Place Template"
        SmallButton_PlaceTemplate.Description = "Uses the current map extents and scale to create a print using a Print Template"
        SmallButton_PlaceTemplate.MinSelection = -1
        SmallButton_PlaceTemplate.Filter = "fNoPhoto"

        'Inset
        Dim SmallButton_Inset As SisRibbonButton = New SisRibbonButton("Inset Map", New SisClickHandler(AddressOf Inset))
        SmallButton_Inset.LargeImage = False
        SmallButton_Inset.Icon = My.Resources.Btn_Inset
        SmallButton_Inset.Help = "Inset"
        SmallButton_Inset.Description = "Creates new Map Frame or edits selected Map Frame"
        SmallButton_Inset.MinSelection = -1
        SmallButton_Inset.MaxSelection = 1
        SmallButton_Inset.Filter = "fArea"

        'Key Frames
        Dim SmallButton_KeyFrames As SisRibbonButton = New SisRibbonButton("Key Frame", New SisClickHandler(AddressOf KeyFrames))
        SmallButton_KeyFrames.LargeImage = False
        SmallButton_KeyFrames.Icon = My.Resources.Btn_KeyFrame
        SmallButton_KeyFrames.Help = "Key Frame"
        SmallButton_KeyFrames.Description = "Creates Key Frames in selected Map Frames"

        'Scalebar
        Dim SmallButton_ScaleBar As SisRibbonButton = New SisRibbonButton("Scale bar", New SisClickHandler(AddressOf Scalebar))
        SmallButton_ScaleBar.LargeImage = False
        SmallButton_ScaleBar.Icon = My.Resources.Btn_Scalebar
        SmallButton_ScaleBar.Help = "Automatic Scale Bar"
        SmallButton_ScaleBar.Description = "Creates an automatic Scale Bar"

        'Scaletext
        Dim SmallButton_ScaleText As SisRibbonButton = New SisRibbonButton("Scale text", New SisClickHandler(AddressOf Scaletext))
        SmallButton_ScaleText.LargeImage = False
        SmallButton_ScaleText.Icon = My.Resources.Btn_Scaletxt
        SmallButton_ScaleText.Help = "Scale Text"
        SmallButton_ScaleText.Description = "Creates Scale Text for selected Map Frame"
        SmallButton_ScaleText.MinSelection = 1
        SmallButton_ScaleText.MaxSelection = 1
        SmallButton_ScaleText.Filter = "fPhoto"

        'Label
        Dim SmallButton_Label As SisRibbonButton = New SisRibbonButton("Label", New SisClickHandler(AddressOf Label))
        SmallButton_Label.LargeImage = False
        SmallButton_Label.Icon = My.Resources.Btn_Labels
        SmallButton_Label.Help = "Create Label"
        SmallButton_Label.Description = "Creates a Map Label"

        'Table
        Dim SmallButton_Table As SisRibbonButton = New SisRibbonButton("Schema Table", New SisClickHandler(AddressOf Table))
        SmallButton_Table.LargeImage = False
        SmallButton_Table.Icon = My.Resources.Btn_Table
        SmallButton_Table.Help = "Table"
        SmallButton_Table.Description = "Adds a table of overlay features displayed in the selected Map Frame"
        SmallButton_Table.MinSelection = 1
        SmallButton_Table.MaxSelection = 1
        SmallButton_Table.Filter = "fPhoto"

        'Overlay Legend
        Dim SmallButton_OverlayLegend As SisRibbonButton = New SisRibbonButton("Legend", New SisClickHandler(AddressOf OverlayLegend))
        SmallButton_OverlayLegend.LargeImage = False
        SmallButton_OverlayLegend.Icon = My.Resources.Btn_Legend
        SmallButton_OverlayLegend.Help = "Overlay Legend"
        SmallButton_OverlayLegend.Description = "Creates a legend based on the Overlay Styles in the selected Map Frame"
        SmallButton_OverlayLegend.MinSelection = 1
        SmallButton_OverlayLegend.MaxSelection = 1
        SmallButton_OverlayLegend.Filter = "fPhoto"

        'Watermark
        Dim SmallButton_Watermark As SisRibbonButton = New SisRibbonButton("Watermark", New SisClickHandler(AddressOf Watermark))
        SmallButton_Watermark.LargeImage = False
        SmallButton_Watermark.Icon = My.Resources.Btn_Watermark
        SmallButton_Watermark.Help = "Watermark"
        SmallButton_Watermark.Description = "Add watermark to selected Map Frame"
        SmallButton_Watermark.MinSelection = 1
        SmallButton_Watermark.MaxSelection = 1
        SmallButton_Watermark.Filter = "fPhoto"

        Dim controlgroup_upper As SisRibbonControlGroup = New SisRibbonControlGroup
        controlgroup_upper.Controls.Add(SmallButton_PlaceTemplate)
        controlgroup_upper.Controls.Add(SmallButton_Inset)
        controlgroup_upper.Controls.Add(SmallButton_KeyFrames)
        controlgroup_upper.Controls.Add(SmallButton_ScaleBar)
        controlgroup_upper.Controls.Add(SmallButton_ScaleText)
        group.Controls.Add(controlgroup_upper)

        Dim controlgroup_lower As SisRibbonControlGroup = New SisRibbonControlGroup
        controlgroup_lower.Controls.Add(SmallButton_Label)
        controlgroup_lower.Controls.Add(SmallButton_Table)
        controlgroup_lower.Controls.Add(SmallButton_OverlayLegend)
        controlgroup_lower.Controls.Add(SmallButton_Watermark)
        group.Controls.Add(controlgroup_lower)

        'Book Plot
        Dim BigButton_Bookplot As SisRibbonButton = New SisRibbonButton("Book Plot", New SisClickHandler(AddressOf Bookplot))
        BigButton_Bookplot.LargeImage = True
        BigButton_Bookplot.Icon = My.Resources.BookPlot
        BigButton_Bookplot.MinSelection = -1
        BigButton_Bookplot.Help = "Book Plot"
        BigButton_Bookplot.Description = "Book Plot"
        group.Controls.Add(BigButton_Bookplot)

    End Sub

    Private Shared x1, x2, y1, y2, z As Double

    Private Sub Bookplot(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            If MapFrameCheck() = False Then

                SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                PW.SIS.DefineNolView("vOrg")
                Dim dialogBook As New Dialog_Book
                dialogBook.StartPosition = FormStartPosition.CenterParent
                dialogBook.ShowDialog()

            Else

                MsgBox("Cannot create book plot from a print template.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch
        End Try
        SIS.DeselectAll()
        PW.SIS.RecallNolView("vOrg")
        SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
        SIS.Redraw(SIS_CURRENTWINDOW)
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub PlaceTemplate(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            If MapFrameCheck() = False Then

                Dim dialogTemplate As New Dialog_Template
                dialogTemplate.StartPosition = FormStartPosition.CenterParent
                dialogTemplate.ShowDialog()

            Else

                MsgBox("Cannot place template on template.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch
        End Try
        SIS.DeselectAll()
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Public Shared Label_txt As String

    Public Shared Label_alignment As Integer

    Public Shared Label_font As Drawing.Font

    Public Shared Label_opaque As Boolean = False

    Public Shared Label_frame As Boolean = False

    Public Shared Label_lines As Boolean = False

    Private Sub Label(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
            SIS.CreateRectLocus("Locus", x1, y1, x2, y2)
            SIS.ChangeLocusTestMode("Locus", SIS_GT_INTERSECT, SIS_GM_EXTENTS)
            If SIS.Scan("Scan", "V", "fPhoto", "Locus") > 0 Then

                Dim dialogLabel As New Dialog_Label()
                If dialogLabel.ShowDialog() = DialogResult.OK Then

                    SIS.CreateGroup("")
                    SIS.CreateBoxText(0, 0, 0, Label_font.Size * 0.0003527, Label_txt)
                    SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", Label_font.Name)
                    SIS.SetInt(SIS_OT_CURITEM, 0, "_text_bold&", Label_font.Bold)
                    SIS.SetInt(SIS_OT_CURITEM, 0, "_text_italic&", Label_font.Italic)
                    SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                    SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", Label_alignment)
                    SIS.UpdateItem()
                    Try

                        APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceLabel_Snap)

                    Catch
                    End Try
                    APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceLabel_Snap))

                End If

            Else

                MsgBox("No Map Frame item in current window.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub PlaceLabel_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            SendKeys.SendWait("{ENTER}")
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceLabel_Snap)
            SIS.OpenSel(0)
            EmptyList("lLabel")
            SIS.AddToList("lLabel")
            Dim ox As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
            Dim oy As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
            Dim sx As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
            Dim sy As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            Dim offset As Double = (Label_font.Size * 0.0003527) * 0.17
            If Label_opaque = True Or Label_frame = True Then

                SIS.CreateRectangle(ox - sx / 2 - offset, oy - sy / 2 - offset, ox + sx / 2 + offset, oy + sy / 2 + offset)
                SIS.AddToList("lLabel")
                If Label_opaque = True Then

                    SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 50</RGBA></Colour></Brush>")

                Else

                    SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Hollow</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")

                End If
                If Label_frame = True Then

                    SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>128 128 128 0</RGBA></Colour><Width>" & Str(Label_font.Size * 3) & "</Width></Pen>")

                Else

                    SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")

                End If
                SIS.UpdateItem()

            End If

            If Label_lines = True Then

                Dim iReturn As Integer = 1
                Dim x, y As Double
                Do

                    Application.DoEvents()
                    iReturn = SIS.GetPosEx(x, y, z)
                    If x > ox - sx / 2 - offset And x < ox + sx / 2 + offset Then

                        If y > oy + sy / 2 Then

                            SIS.MoveTo(ox - sx / 2 - offset, oy + sy / 2 + offset, 0)
                            SIS.LineTo(ox + sx / 2 + offset, oy + sy / 2 + offset, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")
                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(x, oy + sy / 2 + offset, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        ElseIf y < oy - sy / 2 - offset Then

                            SIS.MoveTo(ox - sx / 2 - offset, oy - sy / 2 - offset, 0)
                            SIS.LineTo(ox + sx / 2 + offset, oy - sy / 2 - offset, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")
                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(x, oy - sy / 2 - offset, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        End If

                    ElseIf x < ox - sx / 2 - offset Then

                        SIS.MoveTo(ox - sx / 2 - offset, oy - sy / 2 - offset, 0)
                        SIS.LineTo(ox - sx / 2 - offset, oy + sy / 2 + offset, 0)
                        SIS.UpdateItem()
                        SIS.AddToList("lLines")
                        If y > oy - sy / 2 - offset And y < oy + sy / 2 Then

                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(ox - sx / 2 - offset, y, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        Else

                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(x, oy, 0)
                            SIS.LineTo(ox - sx / 2 - offset, oy, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        End If

                    ElseIf x > ox + sx / 2 + offset Then

                        SIS.MoveTo(ox + sx / 2 + offset, oy - sy / 2 - offset, 0)
                        SIS.LineTo(ox + sx / 2 + offset, oy + sy / 2 + offset, 0)
                        SIS.UpdateItem()
                        SIS.AddToList("lLines")
                        If y > oy - sy / 2 - offset And y < oy + sy / 2 Then

                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(ox + sx / 2 + offset, y, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        Else

                            SIS.MoveTo(x, y, 0)
                            SIS.LineTo(x, oy, 0)
                            SIS.LineTo(ox + sx / 2 + offset, oy, 0)
                            SIS.UpdateItem()
                            SIS.AddToList("lLines")

                        End If

                    End If
                    SIS.SetListStr("lLines", "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour><Width>" & Str(Label_font.Size * 3) & "</Width></Pen>")
                    SIS.CombineLists("lLabel", "lLabel", "lLines", SIS_BOOLEAN_XOR)
                    If iReturn = 0 Then

                        Exit Do

                    End If
                    Exit Do

                Loop

            End If
            SIS.CreateGroupFromItems("lLabel", True, "")
            SIS.OpenItem(SIS.GetInt(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_idNextItem&") - 1)
            SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 254)
            SIS.UpdateItem()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub Watermark(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            Dim sFile As String
            Dim lLevel As Long
            Dim OpenWatermarkDialog As New OpenFileDialog
            OpenWatermarkDialog.Filter = "Image Files(*.BMP;*.GIF;*.PNG)|*.BMP;*.GIF;*.PNG"
            OpenWatermarkDialog.Title = "Print Workshop - Add Watermark"
            If OpenWatermarkDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

                sFile = OpenWatermarkDialog.FileName
                SIS.OpenSel(0)
                lLevel = SIS.GetInt(SIS_OT_CURITEM, 0, "_level&")
                SIS.DoCommand("AComZoomSelect")
                SIS.ZoomView(0.7)
                SIS.PasteFrom(sFile, False)
                SIS.OpenSel(0)
                SIS.SetInt(SIS_OT_CURITEM, 0, "_transparent&", True)
                SIS.SetInt(SIS_OT_CURITEM, 0, "GBitmapAlphaTransparency&", 200)
                SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Hollow</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")
                SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 255)
                SIS.UpdateItem()
                SIS.DeselectAll()
                SIS.DoCommand("AComViewBack")
                SIS.DoCommand("AComViewBack")

            End If

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub Scalebar(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            If SIS.GetAxesType = SIS_AXES_SPHERICAL = True Then

                MsgBox("Scalebars can only be created within a cartographic projection.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
                SIS.Dispose()
                Exit Sub

            End If
            Dim bDoubleSnap As Boolean
            Dim reply As Integer
            Dim x, y, x1, y1, x2, y2, x3, y3 As Decimal
            Dim dBorder As Decimal = 0.0008
            Dim dThickness As Decimal = 0.002
            Do

                reply = SIS.GetPosEx(x, y, z)
                If reply = SIS_ARG_ESCAPE Then

                    bDoubleSnap = False
                    Exit Do

                ElseIf reply = SIS_ARG_BACKSPACE Then

                    bDoubleSnap = False

                ElseIf reply = SIS_ARG_POSITION Then

                    If bDoubleSnap Then

                        x2 = x
                        y2 = y
                        x = x2 - x1
                        y = y2 - y1
                        Dim Q As Integer
                        If Abs(x) > Abs(y) And x > 0 And y > 0 Then

                            Q = 1

                        ElseIf Abs(x) > Abs(y) And x > 0 And y < 0 Then

                            Q = 2

                        ElseIf Abs(x) > Abs(y) And x < 0 And y < 0 Then

                            Q = 3

                        ElseIf Abs(x) > Abs(y) And x < 0 And y > 0 Then

                            Q = 4

                        ElseIf Abs(y) > Abs(x) And y > 0 Then

                            Q = 5
                            Dim tmp As Decimal
                            tmp = x
                            x = y
                            y = tmp

                        ElseIf Abs(y) > Abs(x) And y < 0 Then

                            Q = 6
                            Dim tmp As Decimal
                            tmp = x
                            x = y
                            y = tmp

                        End If
                        x = Abs(x)
                        y = Abs(y)
                        x = x - y
                        Dim dScale As Decimal
                        Dim dScaleBySize As Decimal
                        Try

                            SIS.OpenClosestItem(x1, y1, z, 1, "H", "fPhoto")
                            dScale = SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoscale#")
                            dScaleBySize = dScale * Decimal.Round(x, 3)

                        Catch ex As Exception

                            dScale = 1
                            dScaleBySize = x
                            dBorder = 0.0008 * SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#")
                            dThickness = 0.002 * SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#")
                            SIS.CreateInternalOverlay("Scalebar", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
                            SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
                            SIS.SetFlt(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_scale#", SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#"))

                        End Try
                        Dim iMagnitude As Integer = 1
                        Do Until dScaleBySize < 10

                            iMagnitude = iMagnitude * 10
                            dScaleBySize = dScaleBySize / 10

                        Loop
                        Dim dScaleFactor As Double = 0
                        Select Case dScaleBySize

                            Case 1 To 1.5

                                dScaleFactor = 1

                            Case 1.5 To 2

                                dScaleFactor = 1.5

                            Case 2 To 2.5

                                dScaleFactor = 2

                            Case 2.5 To 5

                                dScaleFactor = 2.5

                            Case 5 To 7.5

                                dScaleFactor = 5

                            Case 7.5 To 10

                                dScaleFactor = 7.5

                            Case Else

                        End Select
                        Dim dScaleBarSize As Double = (dScaleFactor * iMagnitude) / dScale

                        'bar
                        SIS.CreateRectangle(x1, y1, x1 + dScaleBarSize, y1 + dThickness)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 2)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 0 0 50</RGBA></Colour></Brush>")
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                        SIS.UpdateItem()
                        SIS.AddToList("lScalebar")

                        'frame
                        SIS.CreateRectangle(x1 - dBorder, y1 - dBorder, x1 + dScaleBarSize + dBorder, y1 + dBorder + dThickness)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 1)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 75</RGBA></Colour></Brush>")
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Style>Null</Style><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                        SIS.UpdateItem()
                        SIS.AddToList("lScalebar")

                        'text
                        Dim sScaleText As String = Str(dScaleFactor * iMagnitude) + "m"
                        If iMagnitude > 1000 Then

                            sScaleText = Str(Abs(dScaleFactor * iMagnitude / 1000)) + "km"

                        End If
                        SIS.CreateText(x1 + (dScaleBarSize / 2), y1 + dThickness + dBorder, 0, sScaleText)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", 12)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 3)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignV&", SIS_BOTTOM)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_text_alignH&", SIS_CENTRE)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_text_bold&", True)
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_text_outline&", True)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_textFillBrush$", "<Brush><Style>Solid</Style><Colour><RGBA>0 0 0 50</RGBA></Colour></Brush>")
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_textOutlinePen$", "<Pen><Colour><RGBA>255 255 255 75</RGBA></Colour><Width>100</Width><RoundCaps>true</RoundCaps></Pen>")
                        SIS.UpdateItem()
                        SIS.SelectItem()
                        SIS.CallCommand("AComTextToBox")
                        SIS.AddToList("lScalebar")

                        'location
                        SIS.SplitExtent(x2, y2, z, x3, y3, z, SIS.GetListExtent("lScalebar"))

                        Select Case Q

                            Case 1

                                SIS.MoveList("lScalebar", (x1 + y) - x2, (y1 + y) - y2, z, 0, 1)

                            Case 2

                                SIS.MoveList("lScalebar", (x1 + y) - x2, (y1 - y) - y3, z, 0, 1)

                            Case 3

                                SIS.MoveList("lScalebar", (x1 - y) - x3, (y1 - y) - y3, z, 0, 1)

                            Case 4

                                SIS.MoveList("lScalebar", (x1 - y) - x3, (y1 + y) - y2, z, 0, 1)

                            Case 5

                                SIS.MoveList("lScalebar", (x1 - x3) / 2, (y1 + y) - y2, z, 0, 1)

                            Case 6

                                SIS.MoveList("lScalebar", (x1 - x3) / 2, (y1 - y) - y3, z, 0, 1)

                            Case Else

                        End Select
                        SIS.CreateGroupFromItems("lScalebar", True, "")
                        SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 253)
                        SIS.DeselectAll()
                        Exit Do

                    Else

                        bDoubleSnap = True
                        x1 = x
                        y1 = y

                    End If

                End If

            Loop

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private PhotoID As Integer = 0

    Private Sub Scaletext(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            SIS.OpenSel(0)
            For i As Integer = 0 To SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                If SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&") = SIS.GetDataset() Then

                    SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", i)

                End If

            Next
            PhotoID = SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
            Legend_font = New Drawing.Font("Arial", 10)
            Legend_opaque = False
            CreateScaletext()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub CreateScaletext()

        Try

            SIS.CreateGroup("")
            SIS.CreateText(0, 0, 0, "Scale 1:^(" & PhotoID.ToString & ".FormatFlt(_photoscale#,""%.0f""))")
            SIS.SetInt(SIS_OT_CURITEM, 0, "_point_height&", Legend_font.Size)
            SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", Legend_font.Name)
            SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
            SIS.SetStr(SIS_OT_CURITEM, 0, "_textOutlinePen$", "<Pen><Colour><RGBA>255 255 255 75</RGBA></Colour><Width>100</Width><RoundCaps>true</RoundCaps></Pen>")
            SIS.SetInt(SIS_OT_CURITEM, 0, "_text_outline&", True)
            SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 253)

            If Legend_opaque = True Then

                SIS.SetInt(SIS_OT_CURITEM, 0, "_text_outline&", False)
                SIS.SetInt(SIS_OT_CURITEM, 0, "_text_opaque&", True)
                SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 50</RGBA></Colour></Brush>")

            End If
            APP.AddTrigger("AComPlaceGroup::End", New SisTriggerHandler(AddressOf PlaceScaletext_End))
            APP.AddTrigger("AComPlaceGroup::KeyEnter", New SisTriggerHandler(AddressOf PlaceScaletext_KeyEnter))
            APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceScaletext_Snap))
            SIS.UpdateItem()

        Catch
        End Try

    End Sub

    Private Sub PlaceScaletext_End(ByVal sender As Object, ByVal e As SisTriggerArgs)

        APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceScaletext_End)
        APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceScaletext_KeyEnter)
        APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceScaletext_Snap)

    End Sub

    Private Sub PlaceScaletext_KeyEnter(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceScaletext_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceScaletext_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceScaletext_Snap)
            SendKeys.SendWait("{ESC}")
            Dim dialogOverlayLegendConfig As New Dialog_OverlayLegend()
            If dialogOverlayLegendConfig.ShowDialog() = DialogResult.OK Then

                CreateScaletext()

            End If

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub PlaceScaletext_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceScaletext_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceScaletext_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceScaletext_Snap)
            SendKeys.SendWait("{ENTER}")
            SIS.DeselectAll()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Shared Inset_SnapCount As Integer

    Private Sub Inset(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            EmptyList("lInsetMap")
            EmptyList("lPhoto")
            EmptyList("lInsetFrame")
            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
            SIS.CreateRectLocus("Locus", x1, y1, x2, y2)
            SIS.ChangeLocusTestMode("Locus", SIS_GT_INTERSECT, SIS_GM_EXTENTS)
            Dim frame As Boolean = False
            For i As Integer = 0 To SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                If SIS.GetStr(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_class$") = "AInternalDts" Then

                    If SIS.ScanOverlay("lPhoto", i, "fPhoto", "Locus") > 0 Then

                        frame = True
                        Handler = SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$")
                        SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", True)
                        Inset_SnapCount = 0
                        Inset_SnapCount = InsetMap()
                        SIS.DeselectAll()
                        SIS.DoCommand("AComRect")
                        System.Windows.Forms.Application.DoEvents()
                        APP.AddTrigger("AComRect::KeyBack", New SisTriggerHandler(AddressOf InsetMap_Keyback))
                        APP.AddTrigger("AComRect::Snap", New SisTriggerHandler(AddressOf InsetMap_Snap))
                        APP.AddTrigger("AMainWindow::SetCurWnd", New SisTriggerHandler(AddressOf InsetMap_SetCurWnd))
                        APP.AddTrigger("AComRect::End", New SisTriggerHandler(AddressOf InsetMap_End))
                        Exit For

                    End If

                End If

            Next
            If frame = False Then

                MsgBox("No Map Frame item in current window.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch

            RemoveInsetTrigger()

        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Function InsetMap() As Integer

        Try

            SIS.OpenSel(0)
            SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")
            SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
            SIS.UpdateItem()
            SIS.AddToList("lInsetMap")
            BringOnTop()
            Inset_SnapCount = 2

        Catch ex As Exception

            Inset_SnapCount = 0

        End Try
        Return Inset_SnapCount

    End Function

    Private Sub InsetMap_End(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            EmptyList("lInsetFrame")
            RemoveInsetTrigger()
            SIS.SwitchCommand("AComSelectSlide")

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub InsetMap_Keyback(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            If Inset_SnapCount = 1 Then Inset_SnapCount = 0
            If Inset_SnapCount = 3 Then Inset_SnapCount = 2

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub InsetMap_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            If Inset_SnapCount = 1 Then

                Inset_SnapCount = InsetMap()

            ElseIf Inset_SnapCount = 3 Then

                APP.RemoveTrigger("AComRect::End", (AddressOf InsetMap_End))
                SendKeys.SendWait("{ESC}")
                SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                SIS.OpenSel(0)
                SIS.AddToList("lInsetFrame")
                SIS.OpenList("lInsetMap", 0)
                Dim ox_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim oy_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Dim sx_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim sy_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                SIS.OpenList("lInsetFrame", 0)
                Dim ox_photo As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim oy_photo As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Dim sx_photo As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim sy_photo As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                Dim ScaleFactor As Double
                If sx_photo / sx_inset > sy_photo / sy_inset Then

                    ScaleFactor = sx_photo / sx_inset

                Else

                    ScaleFactor = sy_photo / sy_inset

                End If
                SIS.OpenList("lInsetMap", 0)
                If Not SIS.GetStr(SIS_OT_CURITEM, 0, "_class$") = "Photo" Then

                    SIS.OpenList("lInsetFrame", 0)
                    SIS.ScanGeometry("ScanList", SIS_GT_INTERSECT, SIS_GM_GEOMETRY, "fPhoto", "")
                    Dim iLevel As Integer = 0
                    Dim n As Integer = 0
                    For i = 0 To SIS.GetListSize("ScanList") - 1

                        SIS.OpenList("ScanList", i)
                        If SIS.GetInt(SIS_OT_CURITEM, 0, "_level&") >= iLevel Then

                            iLevel = SIS.GetInt(SIS_OT_CURITEM, 0, "_level&")
                            n = i

                        End If

                    Next i
                    SIS.OpenList("ScanList", n)
                    SIS.AddToList("lPhoto")
                    SIS.EmptyList("ScanList")

                Else

                    SIS.AddToList("lPhoto")

                End If
                SIS.OpenList("lPhoto", 0)
                Dim dScale = SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoscale#")
                Dim dRotation = SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoangleDeg#")
                Dim x_photo, y_photo As Double
                SIS.SplitPos(x_photo, y_photo, z, SIS.GetPhotoWorldPos(ox_photo, oy_photo))
                OpenPhoto("lPhoto", 0)
                SIS.SetViewExtent(x_photo - sx_photo, y_photo - sy_photo, 0, x_photo + sx_photo, y_photo + sy_photo, 0)
                SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayScale#", dScale)
                SIS.Compose()
                SIS.SwdClose(SIS_NOSAVE)
                ActivateWindow(Handler)
                SIS.CreatePhoto(ox_inset - sx_inset / 1.9, oy_inset - sy_inset / 1.9, ox_inset + sx_inset / 1.9, oy_inset + sy_inset / 1.9)
                SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")
                SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                SIS.SetFlt(SIS_OT_CURITEM, 0, "_photothreshold#", dScale * ScaleFactor)
                SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoscale#", dScale * ScaleFactor)
                SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoangleDeg#", dRotation)
                SIS.UpdateItem()
                SIS.AddToList("lInsetMap")
                BringOnTop()
                SIS.CreateBoolean("lInsetMap", SIS_BOOLEAN_AND)

                SIS.Delete("lInsetMap")
                SIS.Delete("lInsetFrame")
                RemoveInsetTrigger()
                SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
                SIS.SwitchCommand("AComSelectSlide")
                SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)

            Else

                Inset_SnapCount += 1

            End If

        Catch

            RemoveInsetTrigger()

        End Try
        SIS.DeselectAll()
        SIS.Redraw(SIS_CURRENTWINDOW)
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub InsetMap_SetCurWnd(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
            SIS.CreateRectLocus("Locus", x1, y1, x2, y2)
            SIS.ChangeLocusTestMode("Locus", SIS_GT_INTERSECT, SIS_GM_EXTENTS)
            Dim frame As Boolean = False
            For i As Integer = 0 To SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                If SIS.GetStr(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, i, "_nDataset&"), "_class$") = "AInternalDts" Then

                    If SIS.ScanOverlay("lPhoto", i, "fPhoto", "Locus") > 0 Then

                        frame = True

                    End If

                End If

            Next
            If frame = False Then

                SIS.DeselectAll()
                SIS.Compose()
                Dim sx_map, sy_map, map_scale, photo_scale As Double
                SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                map_scale = SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#")
                SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
                SIS.CreateInternalOverlay("tmpOverlay", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
                SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
                SIS.CreateRectangle(x1, y1, x2, y2)
                sx_map = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                sy_map = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                SIS.RemoveOverlay(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
                ActivateWindow(Handler)
                APP.RemoveTrigger("AComRect::End", AddressOf InsetMap_End)
                SendKeys.SendWait("{ESC}")
                SIS.OpenList("lInsetMap", 0)
                Dim ox_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                Dim oy_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                Dim sx_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                Dim sy_inset As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                If sx_map / sx_inset > sy_map / sy_inset Then

                    photo_scale = sx_map / sx_inset

                Else

                    photo_scale = sy_map / sy_inset

                End If
                SIS.CreatePhoto(ox_inset - sx_inset / 1.9, oy_inset - sy_inset / 1.9, ox_inset + sx_inset / 1.9, oy_inset + sy_inset / 1.9)
                SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 0</RGBA></Colour></Brush>")
                SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                SIS.SetFlt(SIS_OT_CURITEM, 0, "_photothreshold#", photo_scale)
                SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoscale#", photo_scale)
                SIS.UpdateItem()
                SIS.AddToList("lInsetMap")
                BringOnTop()
                SIS.CreateBoolean("lInsetMap", SIS_BOOLEAN_AND)

                SIS.Delete("lInsetMap")
                SIS.Delete("lInsetFrame")
                RemoveInsetTrigger()
                SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
                SIS.SwitchCommand("AComSelectSlide")
                SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)

            Else

                ActivateWindow(Handler)
                SendKeys.SendWait("{ESC}")
                RemoveInsetTrigger()
                MsgBox("Cannot populate inset from a map window which contains a Map Frame.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch
        End Try
        SIS.DeselectAll()
        SIS.Redraw(SIS_CURRENTWINDOW)
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub RemoveInsetTrigger()

        Try

            APP.RemoveTrigger("AComRect::Snap", AddressOf InsetMap_Snap)

        Catch
        End Try
        Try

            APP.RemoveTrigger("AComRect::End", AddressOf InsetMap_End)

        Catch
        End Try
        Try

            APP.RemoveTrigger("AComRect::KeyBack", AddressOf InsetMap_Keyback)

        Catch
        End Try
        Try

            APP.RemoveTrigger("AMainWindow::SetCurWnd", AddressOf InsetMap_SetCurWnd)

        Catch
        End Try

    End Sub

    Private Sub KeyFrames(ByVal sender As Object, ByVal e As SisClickArgs)
        SIS = e.MapEditor
        Try

            If MapFrameCheck() > 1 Then

                Dim dialogBook As New Dialog_Keyframes
                dialogBook.StartPosition = FormStartPosition.CenterParent
                dialogBook.ShowDialog()

            Else

                MsgBox("At least two Map Frames are required for this function.", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")

            End If

        Catch ex As Exception
        End Try
        SIS.DeselectAll()
        SIS.Redraw(SIS_CURRENTWINDOW)
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Public Shared Legend_font As Drawing.Font

    Public Shared Legend_opaque As Boolean = False

    Private Shared OverlayName() As String

    Private Shared OverlayStatus() As Integer

    Private Shared OverlayPenOverride() As Integer

    Private Shared OverlayPen() As String

    Private Shared OverlayBrushOverride() As Integer

    Private Shared OverlayBrush() As String

    Private Shared OverlaySymbolOverride() As Integer

    Private Shared OverlaySymbol() As String

    Private Sub OverlayLegend(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            'get overlay styles
            Handler = SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$")
            SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", True)
            SIS.CreateListFromSelection("lPhoto")
            OpenPhoto("lPhoto", 0)
            OverlayName = New String(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlayStatus = New Integer(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlayPenOverride = New Integer(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlayPen = New String(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlayBrushOverride = New Integer(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlayBrush = New String(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlaySymbolOverride = New Integer(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            OverlaySymbol = New String(SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1) {}
            For i As Integer = 0 To SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                OverlayName(i) = SIS.GetStr(SIS_OT_OVERLAY, i, "_name$")
                OverlayStatus(i) = SIS.GetInt(SIS_OT_OVERLAY, i, "_status&")
                OverlayPenOverride(i) = SIS.GetInt(SIS_OT_OVERLAY, i, "_bPenOverride&")
                OverlayPen(i) = SIS.GetStr(SIS_OT_OVERLAY, i, "_pen$")
                OverlayBrushOverride(i) = SIS.GetInt(SIS_OT_OVERLAY, i, "_bBrushOverride&")
                OverlayBrush(i) = SIS.GetStr(SIS_OT_OVERLAY, i, "_brush$")
                OverlaySymbolOverride(i) = SIS.GetInt(SIS_OT_OVERLAY, i, "_bShapeOverride&")
                OverlaySymbol(i) = SIS.GetStr(SIS_OT_OVERLAY, i, "_shape$")

            Next i
            SIS.SwdClose(SIS_NOSAVE)
            ActivateWindow(Handler)
            SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)
            Legend_font = New Drawing.Font("Arial", 10)
            Legend_opaque = False
            CreateOverlayLegend()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub CreateOverlayLegend()

        Try

            Dim xSpace As Double = 0.00044 * Legend_font.Size
            Dim ySpace As Double = 0.00044 * Legend_font.Size
            Dim square As Double = 0.00035 * Legend_font.Size
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)

            'symbols
            Dim lx, ux, ly, uy As Double
            For i As Integer = 0 To OverlayName.Length - 1

                If OverlaySymbolOverride(i) = True And OverlayStatus(i) > 0 Then

                    SIS.CreatePoint(0, 0, 0, OverlaySymbol(i), 0, 1)
                    SIS.SelectItem()
                    SIS.DoCommand("AComExplodeShape")
                    SIS.CreateListFromSelection("lSymbol")
                    SIS.SplitExtent(lx, ly, z, ux, uy, z, SIS.GetListExtent("lSymbol"))
                    Dim scale As Double = 0
                    If (uy - ly) > (ux - lx) Then scale = square / (uy - ly) Else scale = square / (ux - lx)
                    SIS.MoveList("lSymbol", -lx, -uy, 0, 0, 1)
                    SIS.MoveList("lSymbol", 0, 0, 0, 0, scale)
                    Dim listsize = SIS.GetListSize("lSymbol") - 1
                    For ii As Integer = 0 To SIS.GetListSize("lSymbol") - 1

                        SIS.OpenList("lSymbol", ii)
                        If SIS.GetStr(SIS_OT_CURITEM, 0, "_pen$") = "" Then SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", OverlayPen(i))
                        If SIS.GetStr(SIS_OT_CURITEM, 0, "_brush$") = "" Then SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", OverlayBrush(i))

                    Next ii
                    SIS.UpdateItem()
                    SIS.DefineNolShape("tmpShape_" & i, "lSymbol", 0, 0, 0, 1)
                    SIS.Delete("lSymbol")

                End If

            Next i

            'create legend group
            SIS.CreateGroup("")
            Dim iActual As Integer = 0
            For i As Integer = 0 To OverlayName.Length - 1

                If OverlayStatus(i) > 0 Then

                    If OverlayPenOverride(i) = True Or OverlayBrushOverride(i) = True Or OverlaySymbolOverride(i) = True Then

                        If OverlayPenOverride(i) = True Or OverlayBrushOverride(i) = True Then

                            SIS.CreateRectangle(0, -ySpace * iActual, square, (-ySpace * iActual) - square)
                            If OverlayPenOverride(i) = True Then SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", OverlayPen(i)) Else SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>255 255 255 255</RGBA></Colour></Pen>")
                            If OverlayBrushOverride(i) = True Then SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", OverlayBrush(i)) Else SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 255</RGBA></Colour></Brush>")

                        End If
                        If OverlaySymbolOverride(i) = True Then SIS.CreatePoint(xSpace, -ySpace * iActual, 0, "tmpShape_" & i, 0, 1)
                        SIS.CreateBoxText(xSpace * 2, -ySpace * iActual, 0, Legend_font.Size * 0.0003527, OverlayName(i))
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_font$", Legend_font.Name)
                        SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>0 0 0 0</RGBA></Colour></Pen>")
                        iActual += 1

                    End If

                End If

            Next i
            SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            APP.AddTrigger("AComPlaceGroup::End", New SisTriggerHandler(AddressOf PlaceOverlayLegend_End))
            APP.AddTrigger("AComPlaceGroup::KeyEnter", New SisTriggerHandler(AddressOf PlaceOverlayLegend_KeyEnter))
            APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceOverlayLegend_Snap))
            SIS.CloseItem()

        Catch
        End Try

    End Sub

    Private Sub PlaceOverlayLegend_End(ByVal sender As Object, ByVal e As SisTriggerArgs)

        APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceOverlayLegend_End)
        APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceOverlayLegend_KeyEnter)
        APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceOverlayLegend_Snap)

    End Sub

    Private Sub PlaceOverlayLegend_KeyEnter(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceOverlayLegend_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceOverlayLegend_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceOverlayLegend_Snap)
            SendKeys.SendWait("{ESC}")
            Dim dialogOverlayLegendConfig As New Dialog_OverlayLegend()
            If dialogOverlayLegendConfig.ShowDialog() = DialogResult.OK Then CreateOverlayLegend()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub PlaceOverlayLegend_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceOverlayLegend_End)
            APP.RemoveTrigger("AComPlaceGroup::KeyEnter", AddressOf PlaceOverlayLegend_KeyEnter)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceOverlayLegend_Snap)
            SendKeys.SendWait("{ENTER}")
            SIS.CreateListFromSelection("lLegend")
            SIS.DoCommand("AComExplodeShape")
            SIS.CreateListFromSelection("lSymbols")
            SIS.CombineLists("lLegend", "lLegend", "lSymbols", SIS_BOOLEAN_OR)
            If Legend_opaque = True Then

                SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetListExtent("lLegend"))
                SIS.CreateRectangle(x1 - 0.0001 * Legend_font.Size, y1 - 0.0001 * Legend_font.Size, x2, y2 + 0.0001 * Legend_font.Size)
                SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 50</RGBA></Colour></Brush>")
                SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", "<Pen><Colour><RGBA>255 255 255 255</RGBA></Colour></Pen>")
                SIS.UpdateItem()
                SIS.AddToList("lLegend")

            End If
            SIS.CreateGroupFromItems("lLegend", True, "")
            SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 253)
            SIS.CloseItem()
            SIS.DeselectAll()
            SIS.Redraw(SIS_CURRENTWINDOW)

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub Table(ByVal sender As Object, ByVal e As SisClickArgs)

        SIS = e.MapEditor
        Try

            PW.SIS.CreateListFromSelection("lPhoto")
            Dim dialogTableConfig As New Dialog_Table()
            If dialogTableConfig.ShowDialog() = DialogResult.OK Then

                APP.AddTrigger("AComPlaceGroup::End", New SisTriggerHandler(AddressOf PlaceTable_End))
                APP.AddTrigger("AComPlaceGroup::Snap", New SisTriggerHandler(AddressOf PlaceTable_Snap))

            End If

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Private Sub PlaceTable_End(ByVal sender As Object, ByVal e As SisTriggerArgs)

        APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceTable_End)
        APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceTable_Snap)

    End Sub

    Private Sub PlaceTable_Snap(ByVal sender As Object, ByVal e As SisTriggerArgs)

        SIS = e.MapEditor
        Try

            APP.RemoveTrigger("AComPlaceGroup::End", AddressOf PlaceTable_End)
            APP.RemoveTrigger("AComPlaceGroup::Snap", AddressOf PlaceTable_Snap)
            SendKeys.SendWait("{ENTER}")
            SIS.CreateListFromSelection("lTable")
            SIS.CreateGroupFromItems("lTable", True, "")
            SIS.OpenItem(SIS.GetInt(SIS_OT_DATASET, SIS.GetInt(SIS_OT_OVERLAY, SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_idNextItem&") - 1)
            SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", 254)
            SIS.UpdateItem()
            SIS.DeselectAll()

        Catch
        End Try
        SIS.Dispose()
        SIS = Nothing

    End Sub

    Public Shared Sub BringOnTop()

        Try

            SIS.AddToList("ScanItem")
            SIS.ScanGeometry("ScanList", SIS_GT_CROSSBY, SIS_GM_GEOMETRY, "", "")
            Dim iLevel As Integer = 0
            For i = 0 To SIS.GetListSize("ScanList") - 1

                SIS.OpenList("ScanList", i)
                If SIS.GetInt(SIS_OT_CURITEM, 0, "_level&") >= iLevel Then

                    iLevel = SIS.GetInt(SIS_OT_CURITEM, 0, "_level&") + 1

                End If

            Next i
            SIS.OpenList("ScanItem", 0)
            SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", iLevel)
            SIS.UpdateItem()
            SIS.EmptyList("ScanItem")

        Catch
        End Try

    End Sub

    Public Shared Handler As String = ""

    Public Shared Sub ActivateWindow(ByVal WndHandler As String)

        Try

            SIS.Dispose()
            SIS = Nothing
            SIS = APP.TakeoverMapEditor
            For i As Integer = 0 To SIS.GetNumWnd() - 1

                SIS.ActivateWnd(i)
                System.Windows.Forms.Application.DoEvents()
                If SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$") = WndHandler Then

                    Try

                        If SIS.GetInt(SIS_OT_WINDOW, 0, "Handler&") = True Then

                            Exit For

                        End If

                    Catch
                    End Try

                End If

            Next

        Catch
        End Try

    End Sub

    Public Shared Sub OpenPhoto(ByVal PhotoList As String, ByVal Index As Integer)

        Try

            SIS.OpenList(PhotoList, Index)
            Dim dScale As Double = SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoscale#")
            SIS.CreateComposition("composition")
            SIS.SwdNew()
            SIS.FillSwdFromComposition("composition")
            SIS.SetFlt(SIS_OT_WINDOW, 0, "_displayScale#", dScale)

        Catch
        End Try

    End Sub

    Public Shared Sub SetPhotoCentre(ByVal X As Double, ByVal Y As Double)

        Try

            Dim x_PhotoCentre, y_PhotoCentre, x1, y1, x2, y2, z As Double
            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetExtent())
            SIS.SplitPos(x_PhotoCentre, y_PhotoCentre, z, SIS.GetPhotoWorldPos((x1 + x2) / 2, (y1 + y2) / 2))
            SIS.SetPhotoWorldCentre(x_PhotoCentre + (x_PhotoCentre - X), y_PhotoCentre + (y_PhotoCentre - Y), z)

        Catch
        End Try

    End Sub

    Public Shared Function MapFrameCheck()

        Try

            Dim check = False
            SIS.SplitExtent(x1, y1, z, x2, y2, z, SIS.GetViewExtent())
            SIS.CreateRectLocus("Locus", x1, y1, x2, y2)
            SIS.ChangeLocusTestMode("Locus", SIS_GT_INTERSECT, SIS_GM_EXTENTS)
            Return SIS.Scan("lPhoto", "I", "fPhoto", "Locus")

        Catch

            Dim check = False
            Return check

        End Try

    End Function

    Public Shared Sub EmptyList(ByVal List As String)

        Try

            SIS.EmptyList(List)

        Catch
        End Try

    End Sub

End Class