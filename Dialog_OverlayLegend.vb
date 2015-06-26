Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System
Imports System.Windows.Forms
Imports VB = Microsoft.VisualBasic

Public Class Dialog_OverlayLegend

    Public Sub New()

        InitializeComponent()

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

        'Font size dropdown
        CB_FONTSIZE.Items.Clear()
        For i = 8 To 72 Step 2

            CB_FONTSIZE.Items.Add(i)

        Next i
        CB_FONTSIZE.SelectedIndex = 1
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Me.Btn_Opaque, "Draw an opaque background.")

    End Sub

    Private Sub Btn_Opaque_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Opaque.Click

        If Not Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard
            PW.Legend_opaque = True

        Else

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Flat
            PW.Legend_opaque = False

        End If

    End Sub

    Private Sub CB_FONT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONT.SelectedIndexChanged

        Try

            PW.Legend_font = New Drawing.Font(CB_FONT.Text, Single.Parse(CB_FONTSIZE.Text))

        Catch
        End Try

    End Sub

    Private Sub CB_FontSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONTSIZE.SelectedIndexChanged

        Try

            PW.Legend_font = New Drawing.Font(CB_FONT.Text, Single.Parse(CB_FONTSIZE.Text))

        Catch
        End Try

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click

        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()

    End Sub

End Class