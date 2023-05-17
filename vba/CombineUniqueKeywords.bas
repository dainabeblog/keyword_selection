Sub CombineUniqueKeywords()
    Dim ws As Worksheet
    Dim targetWs As Worksheet
    Dim lastRow As Long
    Dim lastTargetRow As Long
    Dim r As Range
    Dim targetRng As Range
    Dim uniqueValues As Collection
    Dim value As Variant
    Dim headerValue As Variant

    ' Set the target worksheet with the specified name.
    On Error Resume Next
    Set targetWs = ThisWorkbook.Worksheets("10位以内にランクインしているKW")
    On Error GoTo 0

    If targetWs Is Nothing Then
        MsgBox "The specified worksheet does not exist. Please check the name and try again."
        Exit Sub
    End If

    ' Store the value of the A1 cell
    headerValue = targetWs.Cells(1, "A").Value

    ' Create a new Collection object to store unique values.
    Set uniqueValues = New Collection

    ' Loop through all worksheets in the workbook.
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the sheet name is one of the specified names.
        If ws.Name Like "[1-9]" Or ws.Name = "10" Then
            ' Find the last row in the current worksheet's A column.
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

            ' Check if the last row is 1 and if the cell is empty
            If lastRow = 1 And ws.Cells(1, "A").Value = "" Then
                lastRow = 0
            End If

            ' Loop through all cells in the current worksheet's A column.
            If lastRow > 1 Then
                For Each r In ws.Range("A2:A" & lastRow)
                    ' If the cell is not empty
                    If Not IsEmpty(r.Value) Then
                        ' Replace all half-width spaces with full-width spaces.
                        r.Value = Replace(CStr(r.Value), " ", "　")

                        ' If the value is not in the collection, add it.
                        On Error Resume Next
                        uniqueValues.Add r.Value, CStr(r.Value)
                        On Error GoTo 0
                    End If
                Next r
            End If
        End If
    Next ws

    ' Clear target worksheet's A column starting from the third row.
    targetWs.Range("A3:A" & targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).Row).ClearContents

    ' Set the starting row for the output.
    lastTargetRow = 3

    ' Write the unique values from the collection to the target worksheet's A column starting from the third row.
    Set targetRng = targetWs.Range("A" & lastTargetRow)
    For Each value In uniqueValues
        targetWs.Cells(lastTargetRow, 1).Value = value
        lastTargetRow = lastTargetRow + 1
    Next value

    ' Set the A1 cell value back to its original value
    targetWs.Cells(1, "A").Value = headerValue
End Sub
