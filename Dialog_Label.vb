Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System
Imports System.Windows.Forms
Imports VB = Microsoft.VisualBasic

Public Class Dialog_Label

    Private sFile As String = AppDomain.CurrentDomain.BaseDirectory & "\Addins\Print Workshop\Labels.txt"

    Private asLabels() As String

    Private FontStyle As Drawing.FontStyle

    Public Sub New()

        InitializeComponent()
        If Not My.Computer.FileSystem.FileExists(sFile) Then

            MsgBox("Cannot open \Addins\Print Workshop\Labels.txt", MsgBoxStyle.Exclamation, "Cadcorp SIS Print Workshop")
            Me.Close()

        End If

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
        PW.Label_opaque = False
        PW.Label_frame = False
        PW.Label_lines = False

        'Label dropdown
        LoadLabels()

        'Load first label into textbox
        LabelTextBox.Text = asLabels(CB_LABELS.SelectedIndex)

        'Apply styles
        ApplyStyles()

        'initialize context menu
        Dim RTBcontextMenu As New ContextMenu()
        Dim contextX As New MenuItem("Cut")
        AddHandler contextX.Click, AddressOf Cut_Click
        Dim contextC As New MenuItem("Copy")
        AddHandler contextC.Click, AddressOf Copy_Click
        Dim contextP As New MenuItem("Paste")
        AddHandler contextP.Click, AddressOf Paste_Click
        RTBcontextMenu.MenuItems.Add(contextX)
        RTBcontextMenu.MenuItems.Add(contextC)
        RTBcontextMenu.MenuItems.Add(contextP)
        LabelTextBox.ContextMenu = RTBcontextMenu
        Dim tooltip As New ToolTip
        tooltip.SetToolTip(Me.Btn_Edit, "Open Notepad to edit Labels text file.")
        tooltip.SetToolTip(Me.Btn_Reload, "Reload changes from Labels text file.")
        tooltip.SetToolTip(Me.Btn_Label, "Draw lines from label to a user defined location.")
        tooltip.SetToolTip(Me.Btn_Opaque, "Draw an opaque background.")
        tooltip.SetToolTip(Me.Btn_Frame, "Draw frame lines.")

    End Sub

    Private Sub Cut_Click()

        If LabelTextBox.SelectedText <> "" Then

            LabelTextBox.Cut()

        End If

    End Sub

    Private Sub Copy_Click()

        If LabelTextBox.SelectedText <> "" Then

            LabelTextBox.Copy()

        End If

    End Sub

    Private Sub Paste_Click()

        If Clipboard.GetText() <> Nothing Then

            Clipboard.SetText(Clipboard.GetText())
            LabelTextBox.Paste()

        End If
        ApplyStyles()

    End Sub

    Private Sub LoadLabels()

        Try

            Dim sLabel As String = ""
            Dim iCount As Integer = 0
            Dim asFileContent() As String = Split(My.Computer.FileSystem.ReadAllText(sFile), vbCrLf)
            CB_LABELS.Items.Clear()
            For Each Line In asFileContent

                If Not Line = "\" Then

                    If sLabel = "" Then

                        sLabel = Line

                    Else

                        sLabel = sLabel & " " & vbCrLf & Line

                    End If

                Else

                    ReDim Preserve asLabels(iCount)
                    asLabels(iCount) = sLabel
                    CB_LABELS.Items.Add(VB.Left(sLabel, 70) & "...")
                    sLabel = ""
                    iCount += 1

                End If

            Next
            CB_LABELS.SelectedIndex = 0

        Catch
        End Try

    End Sub

    Private Sub ApplyStyles()

        Try

            LabelTextBox.Font = New Drawing.Font(CB_FONT.Text, Single.Parse(CB_FONTSIZE.Text), FontStyle)
            LabelTextBox.SelectAll()
            Select Case PW.Label_alignment

                Case 0

                    LabelTextBox.SelectionAlignment = Windows.Forms.HorizontalAlignment.Left

                Case 2

                    LabelTextBox.SelectionAlignment = Windows.Forms.HorizontalAlignment.Right

                Case 6

                    LabelTextBox.SelectionAlignment = Windows.Forms.HorizontalAlignment.Center

                Case Else

            End Select
            LabelTextBox.Refresh()

        Catch
        End Try

    End Sub

    Private Sub Btn_Edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Edit.Click

        Try

            Diagnostics.Process.Start("notepad.exe", sFile)

        Catch
        End Try

    End Sub

    Private Sub Btn_Reload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Reload.Click

        LoadLabels()

    End Sub

    Private Sub CB_FONT_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONT.SelectedIndexChanged

        ApplyStyles()

    End Sub

    Private Sub CB_FontSize_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_FONTSIZE.SelectedIndexChanged

        ApplyStyles()

    End Sub

    Private Sub CB_LABELS_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CB_LABELS.SelectedIndexChanged

        LabelTextBox.Text = asLabels(CB_LABELS.SelectedIndex)
        ApplyStyles()

    End Sub

    Private Sub Btn_txtBold_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_txtBold.Click

        If LabelTextBox.Font.Bold = True Then

            FontStyle = LabelTextBox.Font.Style - Drawing.FontStyle.Bold
            Btn_txtBold.FlatStyle = Windows.Forms.FlatStyle.Flat

        Else

            FontStyle = LabelTextBox.Font.Style + Drawing.FontStyle.Bold
            Btn_txtBold.FlatStyle = Windows.Forms.FlatStyle.Standard

        End If
        ApplyStyles()

    End Sub

    Private Sub Btn_txtItalic_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_txtItalic.Click

        If LabelTextBox.Font.Italic = True Then

            FontStyle = LabelTextBox.Font.Style - Drawing.FontStyle.Italic
            Btn_txtItalic.FlatStyle = Windows.Forms.FlatStyle.Flat

        Else

            FontStyle = LabelTextBox.Font.Style + Drawing.FontStyle.Italic
            Btn_txtItalic.FlatStyle = Windows.Forms.FlatStyle.Standard

        End If
        ApplyStyles()

    End Sub

    Private Sub Btn_alignLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_alignLeft.Click

        Btn_alignLeft.FlatStyle = Windows.Forms.FlatStyle.Standard
        Btn_alignMiddle.FlatStyle = Windows.Forms.FlatStyle.Flat
        Btn_alignRight.FlatStyle = Windows.Forms.FlatStyle.Flat
        PW.Label_alignment = 0
        ApplyStyles()

    End Sub

    Private Sub Btn_alignMiddle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_alignMiddle.Click

        Btn_alignLeft.FlatStyle = Windows.Forms.FlatStyle.Flat
        Btn_alignMiddle.FlatStyle = Windows.Forms.FlatStyle.Standard
        Btn_alignRight.FlatStyle = Windows.Forms.FlatStyle.Flat
        PW.Label_alignment = 6
        ApplyStyles()

    End Sub

    Private Sub Btn_alignRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_alignRight.Click

        Btn_alignLeft.FlatStyle = Windows.Forms.FlatStyle.Flat
        Btn_alignMiddle.FlatStyle = Windows.Forms.FlatStyle.Flat
        Btn_alignRight.FlatStyle = Windows.Forms.FlatStyle.Standard
        PW.Label_alignment = 2
        ApplyStyles()

    End Sub

    Private Sub Btn_Label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Label.Click

        If Not Btn_Label.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Label.FlatStyle = Windows.Forms.FlatStyle.Standard
            PW.Label_lines = True

        Else

            Btn_Label.FlatStyle = Windows.Forms.FlatStyle.Flat
            PW.Label_lines = False

        End If

    End Sub

    Private Sub Btn_Opaque_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Opaque.Click

        If Not Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Standard
            PW.Label_opaque = True

        Else

            Btn_Opaque.FlatStyle = Windows.Forms.FlatStyle.Flat
            PW.Label_opaque = False

        End If

    End Sub

    Private Sub Btn_Frame_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Frame.Click

        If Not Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Standard Then

            Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Standard
            PW.Label_frame = True

        Else

            Btn_Frame.FlatStyle = Windows.Forms.FlatStyle.Flat
            PW.Label_frame = False

        End If

    End Sub

    Private Sub btn_Place_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Place.Click

        PW.Label_txt = LabelTextBox.Text.Replace(vbLf, vbCrLf)
        PW.Label_font = LabelTextBox.Font
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

End Class