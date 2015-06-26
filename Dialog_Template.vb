Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants

Public Class Dialog_Template

    Public Sub New()

        Try

            InitializeComponent()
            Dim UniqueTemplateNames As New ArrayList
            Dim TemplateList As String = PW.SIS.NolCatalog("APrintTemplate", CInt(False))
            Dim TemplateNames() As String = Split(TemplateList, vbTab)
            For Each TemplateName As String In TemplateNames

                If TemplateName <> "" Then UniqueTemplateNames.Add(TemplateName)

            Next
            UniqueTemplateNames.Sort()
            For Each TemplateName As String In UniqueTemplateNames

                CB_Templates.Items.Add(TemplateName)

            Next
            CB_Templates.SelectedIndex = 0

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Try

            Dim x1, y1, x2, y2, z, photo_scale As Double
            Dim map_scale = PW.SIS.GetFlt(SIS_OT_WINDOW, 0, "_displayScale#")
            PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetViewExtent())
            PW.SIS.CreateInternalOverlay("tmpOverlay", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            PW.SIS.CreateRectangle(x1, y1, x2, y2)
            PW.SIS.UpdateItem()
            Dim x_map = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
            Dim y_map = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            PW.SIS.RemoveOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            PW.SIS.Compose()
            PW.SIS.DefineNolView("View")
            PW.SIS.PlacePrintTemplate(CB_Templates.Text, 0, 1)
            PW.SIS.CreateListFromOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1, "lTemplate")
            PW.SIS.ScanList("lPhoto", "lTemplate", "fPhoto", "")
            PW.SIS.OpenList("lPhoto", 0)
            Dim x_photo = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
            Dim y_photo = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
            PW.SIS.RemoveOverlay(PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
            PW.SIS.RecallNolView("View")
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            If x_map / x_photo > y_map / y_photo Then

                photo_scale = x_map / x_photo

            Else

                photo_scale = y_map / y_photo

            End If
            Dim scale As Double
            Dim scales As Double() = {1, 1.25, 1.5, 2, 2.5, 5, 7.5}
            For Each scale In scales

                Do Until scale > map_scale

                    scale = scale * 10
                    Select Case scale

                        Case (map_scale - 1) To (map_scale + 1)

                            photo_scale = scale

                        Case Else

                    End Select

                Loop

            Next
            PW.SIS.DoCommand("AComFileNewSwd")
            PW.SIS.PlacePrintTemplate(CB_Templates.Text, 0, photo_scale)

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
        Me.Close()

    End Sub

End Class