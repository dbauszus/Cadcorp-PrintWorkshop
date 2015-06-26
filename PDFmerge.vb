Imports System.IO
Imports System.Data
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class PdfMerge

    Private ReadOnly m_documents As List(Of PdfReader)

    Public Sub New()

        m_documents = New List(Of PdfReader)()

    End Sub

    Public Sub AddPageLinks(ByVal FileInfo As FileInfo, ByVal lBookplotFrames As List(Of Dialog_Book.BookplotFrame), ByVal TemplateFrame As Integer, ByVal TemplateIndex As Integer)

        Try

            PW.SIS.OpenItem(TemplateFrame)
            Dim llX, illX, llY, illY, urX, iurX, urY, iurY, vllX, vllY, vurX, vurY, Z As Double
            PW.SIS.SplitExtent(llX, llY, Z, urX, urY, Z, PW.SIS.GetExtent())
            PW.SIS.SplitExtent(vllX, vllY, Z, vurX, vurY, Z, PW.SIS.GetViewExtent())
            llX -= vllX
            urX -= vllX
            llY -= vllY
            urY -= vllY
            Dim Reader As New PdfReader(FileInfo.DirectoryName & "\tmp_nolcatalogue\merge.pdf")
            Dim Stamper As New PdfStamper(Reader, New FileStream(FileInfo.DirectoryName & "\" & FileInfo.Name, FileMode.Create, FileAccess.Write, FileShare.None))
            Dim Writer As PdfWriter = Stamper.Writer
            Dim Action As PdfAction
            Dim iPage As Integer = 1
            Dim sorted = From s In lBookplotFrames Order By s.Page Select s
            For Each Frame As Dialog_Book.BookplotFrame In sorted

                If Not TemplateIndex = -1 Then

                    For ii As Integer = 0 To PW.SIS.GetListSize("lIndexFrames") - 1

                        PW.SIS.OpenList("lIndexFrames", ii)
                        PW.SIS.SplitExtent(illX, illY, Z, iurX, iurY, Z, PW.SIS.GetExtent())
                        illX -= vllX
                        iurX -= vllX
                        illY -= vllY
                        iurY -= vllY
                        Action = PdfAction.GotoLocalPage(ii + 1, New PdfDestination(PdfDestination.FIT), Writer)
                        Stamper.AddAnnotation(New PdfAnnotation(Writer, illX * 2834.6456664, illY * 2834.6456664, iurX * 2834.6456664, iurY * 2834.6456664, Action), iPage)

                    Next

                End If

                If Not Frame.pUp = -1 Then

                    Action = PdfAction.GotoLocalPage(TargetPage(Frame.pUp, lBookplotFrames), New PdfDestination(PdfDestination.FIT), Writer)
                    Stamper.AddAnnotation(New PdfAnnotation(Writer, llX * 2834.6456664, (urY - 0.01) * 2834.6456664, urX * 2834.6456664, urY * 2834.6456664, Action), iPage)

                End If

                If Not Frame.pDown = -1 Then

                    Action = PdfAction.GotoLocalPage(TargetPage(Frame.pDown, lBookplotFrames), New PdfDestination(PdfDestination.FIT), Writer)
                    Stamper.AddAnnotation(New PdfAnnotation(Writer, llX * 2834.6456664, llY * 2834.6456664, urX * 2834.6456664, (llY + 0.01) * 2834.6456664, Action), iPage)

                End If

                If Not Frame.pLeft = -1 Then

                    Action = PdfAction.GotoLocalPage(TargetPage(Frame.pLeft, lBookplotFrames), New PdfDestination(PdfDestination.FIT), Writer)
                    Stamper.AddAnnotation(New PdfAnnotation(Writer, llX * 2834.6456664, llY * 2834.6456664, (llX + 0.01) * 2834.6456664, urY * 2834.6456664, Action), iPage)

                End If

                If Not Frame.pRight = -1 Then

                    Action = PdfAction.GotoLocalPage(TargetPage(Frame.pRight, lBookplotFrames), New PdfDestination(PdfDestination.FIT), Writer)
                    Stamper.AddAnnotation(New PdfAnnotation(Writer, (urX - 0.01) * 2834.6456664, llY * 2834.6456664, urX * 2834.6456664, urY * 2834.6456664, Action), iPage)

                End If

                iPage += 1

            Next

            Stamper.Close()

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

    Private Function TargetPage(ByVal Page As Integer, ByVal lBookplotFrames As List(Of Dialog_Book.BookplotFrame))

        Try

            Dim iPage As Integer = 1
            Dim sorted = From s In lBookplotFrames Order By s.Page Select s
            For Each Frame As Dialog_Book.BookplotFrame In sorted

                If Frame.Page = Page Then Return iPage
                iPage += 1

            Next

        Catch ex As Exception

            MsgBox(ex.ToString)
            Return 1

        End Try

    End Function

    Public Sub AddDocument(ByVal filename As String)

        m_documents.Add(New PdfReader(filename))

    End Sub

    Public Sub Merge(ByVal outputFilename As String, ByVal PageWidth As Double, ByVal PageHeight As Double)

        Try

            Dim outputStream As New FileStream(outputFilename, FileMode.Create)
            Dim pgsize As New iTextSharp.text.Rectangle(PageWidth * 28.346456664, PageHeight * 28.346456664)
            Dim doc As New iTextSharp.text.Document(pgsize)
            Dim newDocument As New Document(pgsize)
            Dim pdfWriter As PdfWriter = pdfWriter.GetInstance(newDocument, outputStream)
            newDocument.Open()
            Dim pdfContentByte As PdfContentByte = pdfWriter.DirectContent
            For Each pdfReader As PdfReader In m_documents

                For page As Integer = 1 To pdfReader.NumberOfPages

                    newDocument.NewPage()
                    Dim importedPage As PdfImportedPage = pdfWriter.GetImportedPage(pdfReader, page)
                    pdfContentByte.AddTemplate(importedPage, 0, 0)

                Next

            Next
            outputStream.Flush()
            If newDocument IsNot Nothing Then

                newDocument.Close()

            End If
            outputStream.Close()

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub

End Class