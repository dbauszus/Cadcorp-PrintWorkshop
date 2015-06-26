Imports Cadcorp.SIS.GisLink.Library
Imports Cadcorp.SIS.GisLink.Library.Constants
Imports System.Text
Imports System.Windows.Forms

Module sisLib

    Public Function sisGetInt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As Integer

        Try
            Return sis.GetInt(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetInt - " & ex.Message)
            Return 0
        End Try

    End Function

    Public Function sisGetFlt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As Double

        Try
            Return sis.GetFlt(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetFlt - " & ex.Message)
            Return 0.0
        End Try

    End Function

    Public Function sisGetStr(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As String

        Try
            Return sis.GetStr(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetStr - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisGetProperty(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As String

        Try
            Return CStr(sis.GetProperty(objectType, nObject, propertyName))
        Catch ex As Exception
            Debug.Print("sisGetProperty - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Sub sisSetStr(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String, ByVal propertyValue As String)

        Try
            sis.SetStr(objectType, nObject, propertyName, propertyValue)
        Catch ex As Exception
            Debug.Print("sisSetStr - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetInt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String, ByVal propertyValue As Integer)

        Try
            sis.SetInt(objectType, nObject, propertyName, propertyValue)
        Catch ex As Exception
            Debug.Print("sisSetInt - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetFlt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String, ByVal propertyValue As Double)

        Try
            sis.SetFlt(objectType, nObject, propertyName, propertyValue)
        Catch ex As Exception
            Debug.Print("sisSetFlt - " & ex.Message)
        End Try

    End Sub

    Public Function sisEvaluateInt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As Integer

        Try
            Return sis.EvaluateInt(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisEvaluateInt - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisEvaluateFlt(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As Double

        Try
            Return sis.EvaluateFlt(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisEvaluateFlt - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisEvaluate(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As String

        Try
            Return sis.Evaluate(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisEvaluate - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisEvaluateStr(ByVal sis As MapManager, ByVal objectType As Short, ByVal nObject As Integer, ByVal propertyName As String) As String

        Try
            Return sis.EvaluateStr(objectType, nObject, propertyName)
        Catch ex As Exception
            Debug.Print("sisEvaluateStr - " & ex.Message)
            Return ""
        End Try

    End Function


    Public Function sisGetListSize(ByVal sis As MapManager, ByVal listName As String) As Integer

        Try
            Return sis.GetListSize(listName)
        Catch ex As Exception
            Debug.Print("sisGetListSize - " & ex.Message)
            Return 0
        End Try

    End Function

    Public Sub sisEmptyList(ByVal sis As MapManager, ByVal listName As String)

        Try
            sis.EmptyList(listName)
        Catch ex As Exception
            Debug.Print("sisEmptyList - " & ex.Message)
        End Try

    End Sub


    Public Sub sisDelete(ByVal sis As MapManager, ByVal listName As String)

        Try
            sis.Delete(listName)
        Catch ex As Exception
            Debug.Print("sisDelete - " & ex.Message)
        End Try

    End Sub

    Public Function sisGetAxesType(ByVal sis As MapManager) As Integer

        Try
            Return sis.GetAxesType()
        Catch ex As Exception
            Debug.Print("sisGetAxesType - " & ex.Message)
            Return SIS_AXES_CARTESIAN
        End Try

    End Function

    Public Sub sisCompose(ByVal sis As MapManager)

        Try
            sis.Compose()
        Catch ex As Exception
            Debug.Print("sisCompose - " & ex.Message)
        End Try

    End Sub

    Public Function sisGetNumSel(ByVal sis As MapManager) As Integer

        Try
            Return sis.GetNumSel()
        Catch ex As Exception
            Debug.Print("sisGetNumSel - " & ex.Message)
            Return 0
        End Try

    End Function

    Public Function sisGetProfileStr(ByVal sis As MapManager, ByVal propertyName As String, ByVal defaultValue As String) As String

        Try
            Return sis.GetProfileStr(propertyName, defaultValue)
        Catch ex As Exception
            Debug.Print("sisGetProfileStr - " & ex.Message)
            Return defaultValue
        End Try

    End Function

    Public Sub sisSetProfileStr(ByVal sis As MapManager, ByVal propertyName As String, ByVal value As String)

        Try
            sis.SetProfileStr(propertyName, value)
        Catch ex As Exception
            Debug.Print("sisSetProfileStr - " & ex.Message)
        End Try

    End Sub

    Public Function sisGetDatasetExtent(ByVal sis As MapManager, ByVal datasetNo As Integer) As String

        Try
            Return sis.GetDatasetExtent(datasetNo)
        Catch ex As Exception
            Debug.Print("sisGetDatasetExtent - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisSplitExtent(ByVal sis As MapManager, ByRef x1 As Double, ByRef y1 As Double, ByRef z1 As Double, ByRef x2 As Double, ByRef y2 As Double, ByRef z2 As Double, ByVal strExtent As String) As Boolean

        Try
            Return sis.SplitExtent(x1, y1, z1, x2, y2, z2, strExtent) = CInt(True)
        Catch ex As Exception
            Debug.Print("sisSplitExtent - " & ex.Message)
            x1 = -1
            y1 = -1
            z1 = -1
            x2 = -1
            y2 = -1
            z2 = -1
            Return False
        End Try

    End Function

    Public Function sisSplitPos(ByVal sis As MapManager, ByRef x As Double, ByRef y As Double, ByRef z As Double, ByVal strPos As String) As Boolean

        Try
            Return sis.SplitPos(x, y, z, strPos) = CInt(True)
        Catch ex As Exception
            Debug.Print("sisSplitPos - " & ex.Message)
            x = -1
            y = -1
            z = -1
            Return False
        End Try
        Return False

    End Function

    Public Function sisGetExtent(ByVal sis As MapManager) As String

        Try
            Return sis.GetExtent()
        Catch ex As Exception
            Debug.Print("sisGetExtent - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisGetPhotoWorldPos(ByVal sis As MapManager, ByVal paperX As Double, ByVal paperY As Double) As String

        Try
            Return sis.GetPhotoWorldPos(paperX, paperY)
        Catch ex As Exception
            Debug.Print("sisGetPhotoWorldPos - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisGetDisplayExtent(ByVal sis As MapManager) As String

        Try
            Return sis.GetDisplayExtent()
        Catch ex As Exception
            Debug.Print("sisGetDisplayExtent - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Function sisGetListExtent(ByVal sis As MapManager, ByVal listName As String) As String

        Try
            Return sis.GetListExtent(listName)
        Catch ex As Exception
            Debug.Print("sisGetListExtent - " & ex.Message)
            Return ""
        End Try

    End Function

    Public Sub sisCreateListFromSelection(ByVal sis As MapManager, ByVal listName As String)

        Try
            sis.CreateListFromSelection(listName)
        Catch ex As Exception
            Debug.Print("sisCreateListFromSelection - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetPhotoWorldCentre(ByVal sis As MapManager, ByVal x As Double, ByVal y As Double, ByVal z As Double)

        Try
            sis.SetPhotoWorldCentre(x, y, z)
        Catch ex As Exception
            Debug.Print("sisSetPhotoWorldCentre - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreatePhoto(ByVal sis As MapManager, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

        Try
            sis.CreatePhoto(x1, y1, x2, y2)
        Catch ex As Exception
            Debug.Print("sisCreatePhoto - " & ex.Message)
        End Try

    End Sub

    Public Sub sisUpdateItem(ByVal sis As MapManager)

        Try
            sis.UpdateItem()
        Catch ex As Exception
            Debug.Print("sisUpdateItem - " & ex.Message)
        End Try

    End Sub

    Public Function sisCreatePropertyFilter(ByVal sis As MapManager, ByVal filterName As String, ByVal formula As String) As Boolean

        Try
            sis.CreatePropertyFilter(filterName, formula)
            Return True
        Catch ex As Exception
            Debug.Print("sisCreatePropertyFilter - " & ex.Message)
            Return False
        End Try

    End Function

    Public Function sisScan(ByVal sis As MapManager, ByVal returnedList As String, ByVal status As String, ByVal filterName As String, ByVal locusName As String) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.Scan(returnedList, status, filterName, locusName)
        Catch ex As Exception
            Debug.Print("sisScan - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Function sisScanOverlay(ByVal sis As MapManager, ByVal returnedList As String, ByVal overlayPos As Integer, ByVal filterName As String, ByVal locusName As String) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.ScanOverlay(returnedList, overlayPos, filterName, locusName)
        Catch ex As Exception
            Debug.Print("sisScanOverlay - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Function sisScanDataset(ByVal sis As MapManager, ByVal returnedList As String, ByVal datasetNo As Integer, ByVal filterName As String, ByVal locusName As String) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.ScanDataset(returnedList, datasetNo, filterName, locusName)
        Catch ex As Exception
            Debug.Print("sisScanDataset - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Function sisScanList(ByVal sis As MapManager, ByVal returnedList As String, ByVal listIn As String, ByVal filterName As String, ByVal locusName As String) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.ScanList(returnedList, listIn, filterName, locusName)
        Catch ex As Exception
            Debug.Print("sisScanList - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function


    Public Sub sisOpenList(ByVal sis As MapManager, ByVal listName As String, ByVal itemPosition As Integer)

        Try
            sis.OpenList(listName, itemPosition)
        Catch ex As Exception
            Debug.Print("sisOpenList - " & ex.Message)
        End Try

    End Sub

    Public Sub sisOpenSel(ByVal sis As MapManager, ByVal itemPosition As Integer)

        Try
            sis.OpenSel(itemPosition)
        Catch ex As Exception
            Debug.Print("sisOpenSel - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreateListFromOverlay(ByVal sis As MapManager, ByVal OverlayPos As Integer, ByVal listName As String)

        Try
            sis.CreateListFromOverlay(OverlayPos, listName)
        Catch ex As Exception
            Debug.Print("sisCreateListFromOverlay - " & ex.Message)
        End Try

    End Sub

    Public Sub sisGetOverlaySchema(ByVal sis As MapManager, ByVal OverlayPos As Integer, ByVal SchemaName As String)

        Try
            sis.GetOverlaySchema(OverlayPos, SchemaName)
        Catch ex As Exception
            Debug.Print("sisGetOverlaySchema - " & ex.Message)
        End Try

    End Sub

    Public Sub sisLoadSchema(ByVal sis As MapManager, ByVal SchemaName As String)

        Try
            sis.LoadSchema(SchemaName)
        Catch ex As Exception
            Debug.Print("sisLoadSchema - " & ex.Message)
        End Try

    End Sub

    Public Sub sisGetViewPrj(ByVal sis As MapManager, ByVal ProjectionName As String)

        Try
            sis.GetViewPrj(ProjectionName)
        Catch ex As Exception
            Debug.Print("sisGetViewPrj - " & ex.Message)
        End Try

    End Sub

    Public Function sisGetImplicitNolObject(ByVal sis As MapManager, ByVal objectClass As String, ByVal objectName As String) As String

        Dim returnVal As String = ""

        Try
            returnVal = sis.GetImplicitNolObject(objectClass, objectName)
        Catch ex As Exception
            Debug.Print("sisGetImplicitNolObject - " & ex.Message)
            returnVal = ""
        End Try

        Return returnVal

    End Function

    Public Sub sisCreateInternalOverlay(ByVal sis As MapManager, ByVal overlayName As String, ByVal overlayPosition As Integer)

        Try
            sis.CreateInternalOverlay(overlayName, overlayPosition)
        Catch ex As Exception
            Debug.Print("sisCreateInternalOverlay - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreateRectangle(ByVal sis As MapManager, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

        Try
            sis.CreateRectangle(x1, y1, x2, y2)
        Catch ex As Exception
            Debug.Print("sisCreateRectangle - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreateGroup(ByVal sis As MapManager, ByVal groupType As String)

        Try
            sis.CreateGroup(groupType)
        Catch ex As Exception
            Debug.Print("sisCreateGroup - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreateGroupFromItems(ByVal sis As MapManager, ByVal listName As String, ByVal deleteItems As Boolean, ByVal groupType As String)

        Try
            sis.CreateGroupFromItems(listName, CInt(deleteItems), groupType)
        Catch ex As Exception
            Debug.Print("sisCreateGroupFromItems - " & ex.Message)
        End Try

    End Sub

    Public Sub sisDefineNolItem(ByVal sis As MapEditor, ByVal item As String)

        Try
            sis.DefineNolItem(item)
        Catch ex As Exception
            Debug.Print("sisDefineNolItem - " & ex.Message)
            MsgBox("DefineNolItem requires a Map Editor licence or higher", MsgBoxStyle.Exclamation, "Cadcorp SIS")
        End Try

    End Sub


    Public Sub sisCreateLabelTheme(ByVal sis As MapManager, ByVal formula As String)

        Try
            sis.CreateLabelTheme(formula)
        Catch ex As Exception
            Debug.Print("sisCreateLabelTheme - " & ex.Message)
        End Try

    End Sub

    Public Sub sisStoreTheme(ByVal sis As MapManager, ByVal themeName As String)

        Try
            sis.StoreTheme(themeName)
        Catch ex As Exception
            Debug.Print("sisStoreTheme - " & ex.Message)
        End Try

    End Sub


    Public Sub sisInsertOverlayTheme(ByVal sis As MapManager, ByVal overlayPos As Integer, ByVal themeName As String, ByVal themePos As Integer)

        Try
            sis.InsertOverlayTheme(overlayPos, themeName, themePos)
        Catch ex As Exception
            Debug.Print("sisInsertOverlayTheme - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetListStr(ByVal sis As MapManager, ByVal listName As String, ByVal propertyName As String, ByVal value As String)

        Try
            sis.SetListStr(listName, propertyName, value)
        Catch ex As Exception
            Debug.Print("sisSetListStr - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetListInt(ByVal sis As MapManager, ByVal listName As String, ByVal propertyName As String, ByVal value As Integer)

        Try
            sis.SetListInt(listName, propertyName, value)
        Catch ex As Exception
            Debug.Print("sisSetListInt - " & ex.Message)
        End Try

    End Sub


    Public Sub sisSetListFlt(ByVal sis As MapManager, ByVal listName As String, ByVal propertyName As String, ByVal value As Double)

        Try
            sis.SetListFlt(listName, propertyName, value)
        Catch ex As Exception
            Debug.Print("sisSetListFlt - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetListFormula(ByVal sis As MapManager, ByVal listName As String, ByVal propertyName As String, ByVal formula As String)

        Try
            sis.SetListFormula(listName, propertyName, formula)
        Catch ex As Exception
            Debug.Print("sisSetListFormula - " & ex.Message)
        End Try

    End Sub


    Public Sub sisCreateKeyMap(ByVal sis As MapManager, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal listName As String, ByVal backdropName As String, ByVal viewName As String)

        Try
            sis.CreateKeyMap(x1, y1, x2, y2, listName, backdropName, viewName)
        Catch ex As Exception
            Debug.Print("sisCreateKeyMap - " & ex.Message)
        End Try

    End Sub

    Public Sub sisAddToList(ByVal sis As MapManager, ByVal listName As String)

        Try
            sis.AddToList(listName)
        Catch ex As Exception
            Debug.Print("sisAddToList - " & ex.Message)
        End Try

    End Sub

    Public Sub sisCreateBoolean(ByVal sis As MapEditor, ByVal listName As String, ByVal booleanOperator As Integer)

        Try
            sis.CreateBoolean(listName, booleanOperator)
        Catch ex As Exception
            Debug.Print("sisCreateBoolean - " & ex.Message)
        End Try

    End Sub


    Public Sub sisCreateRectLocus(ByVal sis As MapEditor, ByVal locusName As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double)

        Try
            sis.CreateRectLocus(locusName, x1, y1, x2, y2)
        Catch ex As Exception
            Debug.Print("sisCreateRectLocus - " & ex.Message)
        End Try

    End Sub


    Public Sub sisChangeLocusTestMode(ByVal sis As MapEditor, ByVal locusName As String, ByVal geomTest As Integer, ByVal geomMode As Integer)

        Try
            sis.ChangeLocusTestMode(locusName, geomTest, geomMode)
        Catch ex As Exception
            Debug.Print("sisChangeLocusTestMode - " & ex.Message)
        End Try

    End Sub

    Public Function sisExportJpeg(ByVal sis As MapManager, ByVal filename As String, ByVal width As Integer, ByVal height As Integer) As Boolean

        Try
            sis.ExportJpeg(filename, width, height)
            Return True
        Catch ex As Exception
            Debug.Print("ExportJpeg - " & ex.Message)
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SIS - ExportJpeg failed")
            Return False
        End Try

    End Function

    Public Sub sisRegisterGroupType(ByVal sis As MapManager, ByVal groupType As String)

        Try
            sis.RegisterGroupType(groupType)
        Catch ex As Exception
            Debug.Print("sisRegisterGroupType - " & ex.Message)
        End Try

    End Sub

    Public Sub sisDeselectAll(ByVal sis As MapManager)

        Try
            sis.DeselectAll()
        Catch ex As Exception
            Debug.Print("sisDeselectAll - " & ex.Message)
        End Try

    End Sub

    Public Function sisGetDataset(ByVal sis As MapManager) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.GetDataset()
        Catch ex As Exception
            Debug.Print("sisGetDataset - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Function sisGetListItemStr(ByVal sis As MapManager, ByVal listName As String, ByVal itemPosition As Integer, ByVal propertyName As String) As String

        Dim returnVal As String = ""

        Try
            returnVal = sis.GetListItemStr(listName, itemPosition, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetListItemStr - " & ex.Message)
            returnVal = ""
        End Try

        Return returnVal

    End Function

    Public Function sisGetListItemInt(ByVal sis As MapManager, ByVal listName As String, ByVal itemPosition As Integer, ByVal propertyName As String) As Integer

        Dim returnVal As Integer = 0

        Try
            returnVal = sis.GetListItemInt(listName, itemPosition, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetListItemInt - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Function sisGetListItemFlt(ByVal sis As MapManager, ByVal listName As String, ByVal itemPosition As Integer, ByVal propertyName As String) As Double

        Dim returnVal As Double = 0

        Try
            returnVal = sis.GetListItemFlt(listName, itemPosition, propertyName)
        Catch ex As Exception
            Debug.Print("sisGetListItemFlt - " & ex.Message)
            returnVal = 0
        End Try

        Return returnVal

    End Function

    Public Sub sisCreateText(ByVal sis As MapManager, ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal text As String)

        Try
            sis.CreateText(x, y, z, text)
        Catch ex As Exception
            Debug.Print("sisCreateText - " & ex.Message)
        End Try

    End Sub

    Public Sub sisSetViewExtent(ByVal sis As MapManager, ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double)

        Try
            sis.SetViewExtent(x1, y1, z1, x2, y2, z2)
        Catch ex As Exception
            Debug.Print("sisSetViewExtent - " & ex.Message)
        End Try

    End Sub

    Public Sub sisDefineNolView(ByVal sis As MapManager, ByVal viewName As String)

        Try
            sis.DefineNolView(viewName)
        Catch ex As Exception
            Debug.Print("sisDefineNolView - " & ex.Message)
        End Try

    End Sub

    Public Sub sisMoveTo(ByVal sis As MapManager, ByVal x As Double, ByVal y As Double, ByVal z As Double)

        Try
            sis.MoveTo(x, y, z)
        Catch ex As Exception
            Debug.Print("sisMoveTo - " & ex.Message)
        End Try

    End Sub

    Public Sub sisLineTo(ByVal sis As MapManager, ByVal x As Double, ByVal y As Double, ByVal z As Double)

        Try
            sis.LineTo(x, y, z)
        Catch ex As Exception
            Debug.Print("sisLineTo - " & ex.Message)
        End Try

    End Sub

    Public Sub sisStoreAsLine(ByVal sis As MapManager)

        Try
            sis.StoreAsLine()
        Catch ex As Exception
            Debug.Print("sisStoreAsLine - " & ex.Message)
        End Try

    End Sub

    Public Sub sisStoreAsArea(ByVal sis As MapManager)

        Try
            sis.StoreAsArea()
        Catch ex As Exception
            Debug.Print("sisStoreAsArea - " & ex.Message)
        End Try

    End Sub

    ' ======================================================================
    ' ======================================================================

    Public Sub populateOverlayListView(ByVal sis As MapManager, ByVal lv As ListView)

        lv.Items.Clear()

        Dim numOverlays As Integer = sisGetInt(sis, SIS_OT_WINDOW, 0, "_nOverlay&")
        For i As Integer = 0 To numOverlays - 1
            Dim overlayName As String = sisGetStr(sis, SIS_OT_OVERLAY, i, "_name$")

            Dim lvi As New ListViewItem
            lvi.Text = overlayName

            'Dim numDataset As Integer = sisGetInt(sis, SIS_OT_OVERLAY, i, "_nDataset&")
            'Dim datasetClass As String = sisGetStr(sis, SIS_OT_DATASET, numDataset, "_class$")

            'If datasetClass.Contains("Cursor") Then lvi.ForeColor = Color.Red

            lv.Items.Add(lvi)
        Next

    End Sub


    Public Sub PopulatePrintTemplateCombo(ByVal sis As MapManager, ByVal cb As ComboBox)

        Dim UniqueTemplateNames As New ArrayList

        Dim TemplateList As String = sis.NolCatalog("APrintTemplate", CInt(False))

        Dim TemplateNames() As String = Split(TemplateList, vbTab)
        For Each TemplateName As String In TemplateNames
            If TemplateName <> "" Then
                'If Not IsInArrayList(UniqueTemplateNames, TemplateName) Then
                UniqueTemplateNames.Add(TemplateName)
                'End If
            End If
        Next

        UniqueTemplateNames.Sort()

        cb.Items.Clear()
        For Each TemplateName As String In UniqueTemplateNames
            cb.Items.Add(TemplateName)
        Next

    End Sub

    Public Function GetOverlayExtent(ByVal sis As MapManager, ByVal OverlayNo As Integer) As String

        Dim DatasetNo As Integer = sisGetInt(sis, SIS_OT_OVERLAY, OverlayNo, "_nDataset&")

        If DatasetNo > 0 Then
            Return sisGetDatasetExtent(sis, DatasetNo)
        Else
            Return ""
        End If

    End Function

    Public Function GetPhotoWorldExtent(ByVal sis As MapManager) As String

        If sisGetStr(sis, SIS_OT_CURITEM, 0, "_class$") <> "Photo" Then Return ""

        Dim photoExtent As String = sisGetExtent(sis)
        If photoExtent = "" Then Return ""

        Dim x1, y1, z1, x2, y2, z2 As Double
        sisSplitExtent(sis, x1, y1, z1, x2, y2, z2, photoExtent)

        'calculate extent necessary to fill photo
        Dim photoOrigin As String = sisGetPhotoWorldPos(sis, x1, y1)
        Dim photoOtherCorner As String = sisGetPhotoWorldPos(sis, x2, y2)
        Dim dX1Photo, dY1Photo, dX2Photo, dY2Photo As Double
        sisSplitPos(sis, dX1Photo, dY1Photo, z1, photoOrigin)
        sisSplitPos(sis, dX2Photo, dY2Photo, z2, photoOtherCorner)

        Dim sb As New StringBuilder
        sb.Append(CStr(dX1Photo).Replace(",", "."))     ' allow for , as decimal separator
        sb.Append(",")
        sb.Append(CStr(dY1Photo).Replace(",", "."))
        sb.Append(",0,")
        sb.Append(CStr(dX2Photo).Replace(",", "."))
        sb.Append(",")
        sb.Append(CStr(dY2Photo).Replace(",", "."))
        sb.Append(",0")

        Return sb.ToString

    End Function

    Public Function GetPhotoWorldCentre(ByVal sis As MapManager) As String
        'Returns a comma-separated string of X1,Y1,Z1
        ' Can be split using GisSplitPos
        '   Operates on the current Photo item

        If sisGetStr(sis, SIS_OT_CURITEM, 0, "_class$") <> "Photo" Then Return ""

        Dim x1, y1, z1, x2, y2, z2 As Double

        Dim photoExtent As String = sisGetExtent(sis)
        sisSplitExtent(sis, x1, y1, z1, x2, y2, z2, photoExtent)

        Dim photoCentre As String = sisGetPhotoWorldPos(sis, (x1 + x2) / 2, (y1 + y2) / 2)
        Dim dX1Photo, dY1Photo As Double
        sisSplitPos(sis, dX1Photo, dY1Photo, z1, photoCentre)

        Dim sb As New StringBuilder
        sb.Append(CStr(dX1Photo).Replace(",", "."))     ' allow for , as decimal separator
        sb.Append(",")
        sb.Append(CStr(dY1Photo).Replace(",", "."))
        sb.Append(",0")

        Return sb.ToString

    End Function

    Public Sub SetRealPhotoCentre(ByVal sis As MapManager, ByVal X As Double, ByVal Y As Double)
        'set photo world centre correctly
        ' not like GisSetPhotoWorldCentre !!!

        Dim CurX As Double
        Dim CurY As Double
        Dim CurZ As Double
        Dim dXdiff As Double
        Dim dYdiff As Double

        sis.SplitPos(CurX, CurY, CurZ, GetPhotoWorldCentre(sis))

        dXdiff = (CurX - X)
        dYdiff = (CurY - Y)
        sisSetPhotoWorldCentre(sis, CurX + dXdiff, CurY + dYdiff, CurZ)

    End Sub

    Public Function GetOverlayNoFromName(ByVal Sis As MapManager, ByVal OverlayName As String) As Integer

        GetOverlayNoFromName = -1

        Dim nOverlays As Integer = Sis.GetInt(SIS_OT_WINDOW, 0, "_nOverlay&")

        For i As Integer = 0 To nOverlays - 1

            Dim Name As String = Sis.GetStr(SIS_OT_OVERLAY, i, "_name$")
            If Name = OverlayName Then
                GetOverlayNoFromName = i
                Exit Function
            End If

        Next

    End Function

    Public Function sisGetMetresConversion(ByVal sis As MapManager) As Double

        ' get conversion factor to metres

        'needed if Transverse Mercator with non-metre units, e.g. NAD27

        sisGetViewPrj(sis, "TempPrj")
        Dim WKT As String = sisGetImplicitNolObject(sis, "APrj", "TempPrj")
        Dim iPos As Integer = CShort(InStr(WKT, "UNIT"))
        If iPos > 0 Then
            Dim iPos2 As Short = CShort(InStr(iPos + 1, WKT, "UNIT"))
            If iPos2 > 0 Then
                Dim LastPart As String = Mid(WKT, iPos2)
                Dim sUnits() As String = Split(LastPart, ",")
                Return Val(sUnits(1))
            Else
                Return 1.0
            End If
        Else
            Return 1.0
        End If

    End Function

    Public Function createInternalOverlay(ByVal sis As MapManager, ByVal overlayName As String) As Integer

        ' create an internal overlay, make it current and return its overlay position

        Dim overlayPos As Integer = sisGetInt(sis, SIS_OT_WINDOW, 0, "_nOverlay&")
        sisCreateInternalOverlay(sis, overlayName, overlayPos)
        sisSetInt(sis, SIS_OT_WINDOW, 0, "_nDefaultOverlay&", overlayPos)
        Return overlayPos

    End Function


    Public Sub PopulateStyleCombo(ByVal sis As MapManager, ByVal objectClass As String, ByVal cb As ComboBox, ByVal selectedObject As String, ByVal folderIgnoreList As String)

        If Not (objectClass = "ABrush" Or objectClass = "APen") Then
            MsgBox("Object class '" & objectClass & "' not recognised in PopulateStyleCombo", MsgBoxStyle.Critical, "Error!")
            Exit Sub
        End If

        Dim foldersToIgnore As New List(Of String)
        If folderIgnoreList <> "" Then
            Dim folders() As String = folderIgnoreList.Split("|")
            For Each foldername As String In folders
                foldersToIgnore.Add(foldername)
            Next
        End If

        Dim UniqueObjectNames As New ArrayList
        Dim selectedObjectExists As Boolean = False

        Dim objectList As String = sis.NolCatalog(objectClass, CInt(False))

        Dim ObjectNames() As String = Split(objectList, vbTab)
        For Each ObjectName As String In ObjectNames
            If ObjectName <> "" Then
                Dim inFolderToIgnore As Boolean = False
                If foldersToIgnore.Count > 0 Then
                    If ObjectName.Contains(".") Then
                        Dim dotPosition As Integer = ObjectName.IndexOf(".")
                        Dim folderName As String = ObjectName.Substring(0, dotPosition)
                        For Each folderToIgnore In foldersToIgnore
                            If folderName = folderToIgnore Then inFolderToIgnore = True
                        Next
                    End If
                End If

                If Not inFolderToIgnore Then
                    UniqueObjectNames.Add(ObjectName)
                    If ObjectName = selectedObject Then selectedObjectExists = True
                End If
            End If
        Next

        UniqueObjectNames.Sort()

        cb.Items.Clear()
        For Each ObjectName As String In UniqueObjectNames
            cb.Items.Add(ObjectName)
        Next

        If selectedObjectExists Then cb.Text = selectedObject

    End Sub
End Module
