Attribute VB_Name = "Module2"
Sub OutputKeywords()
    Dim srcWs As Worksheet
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim nextRow As Long
    Dim threshold As Double
    
    Set srcWs = ThisWorkbook.Worksheets("10�ʈȓ��Ƀ����N�C�����Ă���KW")
    Set targetWs = ThisWorkbook.Worksheets("�l�����ׂ�KW�ꗗ")
    
    lastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    nextRow = 3
    
    threshold = targetWs.Cells(2, "B").value ' B2�Z���̒l���������l�Ƃ��Ďg�p
    
    For i = 3 To lastRow
        If srcWs.Cells(i, "B").value >= threshold Then
            targetWs.Cells(nextRow, "A").value = srcWs.Cells(i, "A").value
            targetWs.Cells(nextRow, "B").value = srcWs.Cells(i, "B").value
            targetWs.Cells(nextRow, "B").NumberFormat = "0.0%" ' �Z���̏������p�[�Z���e�[�W�\���ɕύX
            nextRow = nextRow + 1
        End If
    Next i
    
    With targetWs.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B3:B" & nextRow - 1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange Range("A2:B" & nextRow - 1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


