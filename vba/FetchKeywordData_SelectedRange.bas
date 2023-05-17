Sub FetchKeywordData_SelectedRange()
    Dim ws As Worksheet
    Dim targetWs As Worksheet
    Dim keywordCell As Range
    Dim sheetNumber As Variant
    Dim keyword As String
    Dim lastRow As Long
    Dim foundCell As Range
    Dim i As Long
    Dim lastKeywordRow As Long
    Dim selectedRange As Range
    
    ' 指定された名前のターゲットワークシートを設定します。
    Set targetWs = ThisWorkbook.Worksheets("10位以内にランクインしているKW")
    
    ' ターゲットシートのA列で最後の行を見つけます。
    lastKeywordRow = targetWs.Cells(targetWs.Rows.Count, "A").End(xlUp).Row
    
    ' 選択された範囲を取得します。
    Set selectedRange = targetWs.Application.Selection
    
    ' マクロの実行速度を上げます。
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    ' A列のすべてのキーワードをループします。
    For Each keywordCell In targetWs.Range("A3:A" & lastKeywordRow)
        ' 選択された範囲の列をループします。
        For Each cell In selectedRange
            ' セルからシート番号を取得します。
            sheetNumber = targetWs.Cells(2, cell.Column).value
            
            ' シート番号が有効かどうかを確認します（1〜10の間）。
            If IsNumeric(sheetNumber) And sheetNumber >= 1 And sheetNumber <= 10 Then
                ' 対応するワークシートを取得します。
                Set ws = ThisWorkbook.Worksheets(CStr(sheetNumber))
                
                ' ターゲットシートからキーワードを取得します。
                keyword = keywordCell.value
                
                ' 現在のワークシートのA列で最後の行を見つけます。
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                
                ' 現在のシートのA列でキーワードを見つけます。
                Set foundCell = ws.Range("A1:A" & lastRow).Find(keyword, LookIn:=xlValues, LookAt:=xlWhole)
                
                ' キーワードが見つかった場合、同じ行のH列から値を取得し、ターゲットシートに書き込みます。
                If Not foundCell Is Nothing Then
                    targetWs.Cells(keywordCell.Row, cell.Column).value = ws.Cells(foundCell.Row, "H").value
                End If
            End If
        Next cell
    Next keywordCell
    
    ' 設定をリセットします。
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub


