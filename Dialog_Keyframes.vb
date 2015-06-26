Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.Windows.Forms

Public Class Dialog_Keyframes

    Public Sub New()

        Try

            InitializeComponent()

        Catch
        End Try

    End Sub

    Private Sub Keyframes() Handles Me.Shown

        Try

            Dim x1, y1, x2, y2, z As Double
            PW.Handler = PW.SIS.GetStr(SIS_OT_SYSTEM, 0, "_MainWndTitle$")
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", True)
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
            'create arrays to store map frame properties
            Dim ID() As Integer = New Integer(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim overlay() As Integer = New Integer(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim scale() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim pen() As String = New String(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim brush() As String = New String(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim level() As Integer = New Integer(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim angle() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim oX() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim oY() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim sX() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim sY() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim oLON() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            Dim oLAT() As Double = New Double(PW.SIS.GetListSize("lPhoto") - 1) {}
            'loop through map frame list, store properties, create frame shape
            ProgressBar.Maximum = PW.SIS.GetListSize("lPhoto") * 2
            ProgressBar.Value = 0
            For i = 0 To PW.SIS.GetListSize("lPhoto") - 1

                ProgressBar.Value += 1
                PW.SIS.OpenList("lPhoto", i)
                PW.EmptyList("lBool")
                PW.SIS.AddToList("lBool")
                'store map frame properties in arrays
                ID(i) = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_id&")
                overlay(i) = PW.SIS.FindDatasetOverlay(PW.SIS.GetDataset(), 0, True)
                scale(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoscale#")
                angle(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_photoangleDeg#")
                pen(i) = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "_pen$")
                brush(i) = PW.SIS.GetStr(SIS_OT_CURITEM, 0, "_brush$")
                level(i) = PW.SIS.GetInt(SIS_OT_CURITEM, 0, "_level&")
                oX(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_ox#")
                oY(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oy#")
                sX(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sx#")
                sY(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_sy#")
                'open map frame and store map centre as lat/lon
                PW.OpenPhoto("lPhoto", i)
                PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                PW.SIS.SplitExtent(x1, y1, z, x2, y2, z, PW.SIS.GetViewExtent())
                PW.SIS.CreateInternalOverlay("tmp", 0)
                PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
                PW.SIS.SetDatasetPrj(PW.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "Latitude/Longitude.OGC.WGS_1984")
                PW.SIS.CreatePoint((x1 + x2) / 2, (y1 + y2) / 2, 0, "", 0, 1)
                oLON(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLon#")
                oLAT(i) = PW.SIS.GetFlt(SIS_OT_CURITEM, 0, "_oLat#")
                PW.SIS.SwdClose(SIS_NOSAVE)
                PW.ActivateWindow(PW.Handler)
                'create shape from map frame geometry
                PW.SIS.CreateInternalOverlay("tmp", 0)
                PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
                PW.SIS.CreateRectangle(oX(i) - sX(i) / 1.9, oY(i) - sY(i) / 1.9, oX(i) + sX(i) / 1.9, oY(i) + sY(i) / 1.9)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", pen(i))
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", "<Brush><Style>Solid</Style><Colour><RGBA>255 255 255 255</RGBA></Colour></Brush>")
                PW.SIS.UpdateItem()
                PW.SIS.AddToList("lBool")
                PW.SIS.CreateBoolean("lBool", SIS_BOOLEAN_AND)
                PW.EmptyList("lShape")
                PW.SIS.AddToList("lShape")
                PW.SIS.DefineNolShape("tmpShape" & Str(ID(i)), "lShape", oX(i), oY(i), 0, 1)
                PW.SIS.RemoveOverlay(0)

            Next i
            'loop through map frame list, add key frames as overlays within the map frame 
            For i = 0 To ID.Count - 1

                ProgressBar.Value += 1
                PW.OpenPhoto("lPhoto", i)
                PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", False)
                For o As Integer = 0 To PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1

                    Try

                        If PW.SIS.GetStr(SIS_OT_OVERLAY, o, "_name$") = "PrintWorkshop_KeyFrame" Then

                            PW.SIS.RemoveOverlay(o)
                            o -= 1

                        End If

                    Catch
                    End Try

                Next o
                For ii = 0 To ID.Count - 1

                    'create key frame layer with the exploded geometry from the key frame shape
                    If Not i = ii Then

                        'create a temporary wgs84 overlay for the key frame shape
                        PW.SIS.CreateInternalOverlay("tmpWGS84overlay", 0)
                        PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", 0)
                        PW.SIS.SetDatasetPrj(PW.SIS.GetInt(SIS_OT_OVERLAY, 0, "_nDataset&"), "Latitude/Longitude.OGC.WGS_1984")
                        PW.SIS.SetFlt(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_scale#", 1)
                        PW.SIS.CreatePoint(0, 0, 0, "tmpShape" & Str(ID(ii)), 0.0, scale(ii))
                        PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oLon#", oLON(ii))
                        PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_oLat#", oLAT(ii))
                        PW.SIS.UpdateItem()
                        PW.EmptyList("lShape")
                        PW.SIS.AddToList("lShape")
                        'create an overlay in the native map frame projection and replicate map frame shape
                        PW.SIS.CreateInternalOverlay("PrintWorkshop_KeyFrame", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&"))
                        PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&") - 1)
                        PW.SIS.SetFlt(SIS_OT_DATASET, PW.SIS.GetInt(SIS_OT_OVERLAY, PW.SIS.GetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&"), "_nDataset&"), "_scale#", 1)
                        PW.SIS.CopyListItems("lShape")
                        'remove the temporary wgs84 overlay and set key frame properties
                        PW.SIS.RemoveOverlay(0)
                        PW.SIS.OpenList("lShape", 0)
                        PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_angleDeg#", angle(ii))
                        PW.SIS.SelectItem()
                        PW.SIS.UpdateItem()
                        PW.SIS.DoCommand("AComExplodeShape")

                    End If

                Next ii
                'compose the map window with key frame overlays
                PW.SIS.Compose()
                PW.SIS.SwdClose(SIS_NOSAVE)
                PW.ActivateWindow(PW.Handler)
                'create a new map frame from the composition with key frames and apply properties
                PW.SIS.OpenList("lPhoto", i)
                PW.EmptyList("lBool")
                PW.SIS.AddToList("lBool")
                PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_nDefaultOverlay&", overlay(i))
                PW.SIS.CreatePhoto(oX(i) - sX(i) / 1.9, oY(i) - sY(i) / 1.9, oX(i) + sX(i) / 1.9, oY(i) + sY(i) / 1.9)
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_pen$", pen(i))
                PW.SIS.SetStr(SIS_OT_CURITEM, 0, "_brush$", brush(i))
                PW.SIS.SetInt(SIS_OT_CURITEM, 0, "_level&", level(i))
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_photothreshold#", scale(i))
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoscale#", scale(i))
                PW.SIS.SetFlt(SIS_OT_CURITEM, 0, "_photoangleDeg#", angle(i))
                PW.SIS.UpdateItem()
                'create boolean from new and old map frame
                PW.SIS.AddToList("lBool")
                PW.SIS.CreateBoolean("lBool", SIS_BOOLEAN_AND)
                PW.SIS.OpenList("lBool", 1)
                PW.SIS.DeleteItem()

            Next i
            PW.SIS.Delete("lPhoto")
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "_bRedraw&", True)
            PW.SIS.SetInt(SIS_OT_WINDOW, 0, "Handler&", False)

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try
        Me.Close()

    End Sub

End Class