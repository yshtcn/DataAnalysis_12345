Attribute VB_Name = "AutoUpdate"
Sub Calling()
Attribute Calling.VB_ProcData.VB_Invoke_Func = "r\n14"
    CopyValues
    DeleteRowsAndColumns
    ImportData
    Add7Days
    SortData
End Sub

Sub CopyValues()
    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("数据分析")
    
    For i = 2 To 40
        ws.Range("R" & i).Value = ws.Range("B" & i).Value
        ws.Range("S" & i).Value = ws.Range("C" & i).Value
        ws.Range("T" & i).Value = ws.Range("D" & i).Value
        ws.Range("U" & i).Value = ws.Range("E" & i).Value
        ws.Range("V" & i).Value = ws.Range("F" & i).Value
        ws.Range("W" & i).Value = ws.Range("G" & i).Value
        ws.Range("X" & i).Value = ws.Range("H" & i).Value
    Next i

    For i = 1 To 40
        ws.Range("Y" & i).Value = ws.Range("K" & i).Value
        ws.Range("Z" & i).Value = ws.Range("L" & i).Value
        ws.Range("AA" & i).Value = ws.Range("M" & i).Value
        ws.Range("AB" & i).Value = ws.Range("N" & i).Value
        ws.Range("AC" & i).Value = ws.Range("O" & i).Value
        ws.Range("AD" & i).Value = ws.Range("P" & i).Value
        ws.Range("AE" & i).Value = ws.Range("Q" & i).Value
    Next i

    Set ws = ThisWorkbook.Sheets("数据分析社区篇")
    
     For i = 3 To 28
        ws.Range("R" & i).Value = ws.Range("B" & i).Value
        ws.Range("S" & i).Value = ws.Range("C" & i).Value
        ws.Range("T" & i).Value = ws.Range("D" & i).Value
        ws.Range("U" & i).Value = ws.Range("E" & i).Value
        ws.Range("V" & i).Value = ws.Range("F" & i).Value
        ws.Range("W" & i).Value = ws.Range("G" & i).Value
        ws.Range("X" & i).Value = ws.Range("H" & i).Value
    Next i

    For i = 2 To 28
        ws.Range("Y" & i).Value = ws.Range("K" & i).Value
        ws.Range("Z" & i).Value = ws.Range("L" & i).Value
        ws.Range("AA" & i).Value = ws.Range("M" & i).Value
        ws.Range("AB" & i).Value = ws.Range("N" & i).Value
        ws.Range("AC" & i).Value = ws.Range("O" & i).Value
        ws.Range("AD" & i).Value = ws.Range("P" & i).Value
        ws.Range("AE" & i).Value = ws.Range("Q" & i).Value
    Next i

    Set ws = ThisWorkbook.Sheets("数据分析单位篇")
    
     For i = 29 To 37
        ws.Range("R" & i).Value = ws.Range("B" & i).Value
        ws.Range("S" & i).Value = ws.Range("C" & i).Value
        ws.Range("T" & i).Value = ws.Range("D" & i).Value
        ws.Range("U" & i).Value = ws.Range("E" & i).Value
        ws.Range("V" & i).Value = ws.Range("F" & i).Value
        ws.Range("W" & i).Value = ws.Range("G" & i).Value
        ws.Range("X" & i).Value = ws.Range("H" & i).Value
    Next i

     For i = 2 To 2
        ws.Range("Y" & i).Value = ws.Range("K" & i).Value
        ws.Range("Z" & i).Value = ws.Range("L" & i).Value
        ws.Range("AA" & i).Value = ws.Range("M" & i).Value
        ws.Range("AB" & i).Value = ws.Range("N" & i).Value
        ws.Range("AC" & i).Value = ws.Range("O" & i).Value
        ws.Range("AD" & i).Value = ws.Range("P" & i).Value
        ws.Range("AE" & i).Value = ws.Range("Q" & i).Value
    Next i

    For i = 29 To 37
        ws.Range("Y" & i).Value = ws.Range("K" & i).Value
        ws.Range("Z" & i).Value = ws.Range("L" & i).Value
        ws.Range("AA" & i).Value = ws.Range("M" & i).Value
        ws.Range("AB" & i).Value = ws.Range("N" & i).Value
        ws.Range("AC" & i).Value = ws.Range("O" & i).Value
        ws.Range("AD" & i).Value = ws.Range("P" & i).Value
        ws.Range("AE" & i).Value = ws.Range("Q" & i).Value
    Next i

End Sub

Sub DeleteRowsAndColumns()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("数据") '将“数据”更改为你要删除行的工作表名称
    
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    
    '清空第二行的A到AG列
    With ws
        .Range("A2:AG2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    
    Application.ScreenUpdating = True '恢复屏幕更新
    
End Sub

Sub ImportData()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    
    '弹出文件选择窗口，让用户选择要导入的Excel文件
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '如果用户未选择文件，退出该子过程
    If selectedFile = False Then Exit Sub
    
    '打开选定的文件
    Workbooks.Open selectedFile
    
    '获取打开的工作簿的第一个工作表
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '获取最后一行的行号
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '将第二行到最后一行的A到AG列复制到当前工作表名为“数据”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("数据")
    sourceSheet.Range("A2:AG" & LastRow).Copy targetSheet.Range("A2")
    
    '关闭打开的工作簿
    ActiveWorkbook.Close False


    
    ' 获取最后一行的行号
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' 将 AH2 公式填充到整个 AH 列
    Range("AH2").AutoFill Destination:=Range("AH2:AH" & LastRow), Type:=xlFillDefault
    Range("AI2").AutoFill Destination:=Range("AI2:AI" & LastRow), Type:=xlFillDefault
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & LastRow), Type:=xlFillDefault
    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub

Sub Add7Days()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("数据分析社区篇").Range("K2:Q2")
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = cell.Value + 7
        End If
    Next cell
    Set rng = ThisWorkbook.Sheets("数据分析单位篇").Range("K2:Q2")
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = cell.Value + 7
        End If
    Next cell
End Sub

Sub SortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("数据分析社区篇")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("D3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A3:E27")
        
        '应用排序
        .Header = xlNo '包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
    
    Set ws = ThisWorkbook.Worksheets("数据分析单位篇")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("D3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A3:D9")
        
        '应用排序
        .Header = xlNo '包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
End Sub







