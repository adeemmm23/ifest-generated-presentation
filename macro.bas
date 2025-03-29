Sub Main()
    ProcessPrizeCategory 391, "top3"
    ProcessPrizeCategory 271, "top10"
    ProcessPrizeCategory 291, "gold"
    ProcessPrizeCategory 293, "silver"
    ProcessPrizeCategory 388, "bronze"
    ProcessPrizeCategory 398, "honorable"
End Sub
Sub ProcessPrizeCategory(id As Long, prize As String)
    Const ROW_DELIMITER As String = vbCrLf
    Const COL_DELIMITER As String = ","
    Const SEPARATOR As String = " & "
    
    Dim filePath As String: filePath = ActivePresentation.Path & "\files\" & prize & ".csv"
    On Error GoTo PathError
    Dim sArr: sArr = TextFileToArray(filePath, ROW_DELIMITER)
    If IsEmpty(sArr) Then Exit Sub
    
    Dim Data(): Data = GetSplitArray(sArr, COL_DELIMITER)
    
    Dim i As Long
    For i = LBound(Data, 1) To UBound(Data, 1) - LBound(Data, 1) + 1
        If Not IsEmpty(Data(i, 1)) Then
            Dim cleanedText As String
            cleanedText = Replace(Data(i, 1), SEPARATOR, vbNewLine)
            GenerateSlide id, cleanedText
        End If
    Next i
    DeleteSlide id
    Exit Sub
PathError:
    MsgBox "Please check the " + prize + " path, if you don't want " + prize + " just make an empty file '" + prize + ".csv' to remove this error"
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
Function TextFileToArray( _
    ByVal filePath As String, _
    Optional ByVal LineSeparator As String = vbLf) _
As Variant

    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    ' Open file as UTF-8
    With objStream
        .Type = 2 ' adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile filePath
    End With

    Dim sArr() As String
    sArr = Split(objStream.ReadText, LineSeparator)
    
    objStream.Close
    Set objStream = Nothing

    ' Remove trailing empty lines
    Dim n As Long
    For n = UBound(sArr) To LBound(sArr) Step -1
        If Len(sArr(n)) > 0 Then Exit For
    Next n
    
    If n < LBound(sArr) Then Exit Function
    If n < UBound(sArr) Then ReDim Preserve sArr(0 To n)
    
    TextFileToArray = sArr

End Function
Function GetSplitArray( _
    ByVal SourceArray As Variant, _
    Optional ByVal ColumnDelimiter As String = ",") _
As Variant

    Dim rDiff As Long: rDiff = 1 - LBound(SourceArray)
    Dim rCount As Long: rCount = UBound(SourceArray) + rDiff
    Dim cCount As Long: cCount = 1

    Dim Data(): ReDim Data(1 To rCount, 1 To cCount)

    Dim rArr() As String, r As Long, c As Long, cc As Long, rString As String
    
    For r = 1 To rCount
        rString = SourceArray(r - rDiff)
        If Len(rString) > 0 Then
            rArr = Split(rString, ColumnDelimiter)
            cc = UBound(rArr) + 1
            If cc > cCount Then
                cCount = cc
                ReDim Preserve Data(1 To rCount, 1 To cCount)
            End If
            For c = 1 To cc
                Data(r, c) = rArr(c - 1)
            Next c
        End If
    Next r

    GetSplitArray = Data

End Function