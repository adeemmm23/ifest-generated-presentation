Sub Main()
    ProcessPrizeCategory 391, "top3"
    ProcessPrizeCategory 271, "top10"
    ProcessPrizeCategory 291, "gold"
    ProcessPrizeCategory 293, "silver"
    ProcessPrizeCategory 388, "bronze"
    ProcessPrizeCategory 398, "honorable"
End Sub
Sub ProcessPrizeCategory(id As Long, prize As String)

    Dim filePath As String
    filePath = ActivePresentation.Path & "\files\" & prize & ".csv"
    
    Dim xlApp As Object
    Dim wb As Object
    Dim ws As Object
    
    On Error GoTo PathError

    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    Set wb = xlApp.Workbooks.Open(filePath)
    Set ws = wb.Sheets(1)

    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
    
    For i = 1 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            cleanedText = ""
            lastCol = ws.Cells(i, ws.Columns.Count).End(-4159).Column
        
            For j = 1 To lastCol
                If Trim(ws.Cells(i, j).Value) <> "" Then
                    If cleanedText = "" Then
                        cleanedText = ws.Cells(i, j).Value
                    Else
                        cleanedText = cleanedText & vbNewLine & ws.Cells(i, j).Value
                    End If
                End If
            Next j

            GenerateSlide id, cleanedText
        End If
    Next i
    
    wb.Close False
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing

    DeleteSlide id
    Exit Sub

PathError:
    MsgBox "File not found or Excel failed to open: " & filePath

End Sub
Sub GenerateSlide(id As Long, text As Variant)
    Dim newSlide As SlideRange
    Dim currentText As Shape
    Set newSlide = ActivePresentation.Slides.FindBySlideID(id).Duplicate
    Set currentText = newSlide.Shapes("names")
    currentText.TextFrame.TextRange.text = text
End Sub
Sub DeleteSlide(id As Long)
    ActivePresentation.Slides.FindBySlideID(id).Delete
End Sub
