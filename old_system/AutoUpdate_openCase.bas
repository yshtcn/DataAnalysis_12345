Attribute VB_Name = "模块1"
Sub Calling()
Attribute Calling.VB_ProcData.VB_Invoke_Func = "r\n14"
    CopyValues
    DeleteRowsAndColumns
    ImportData
    SortData
End Sub

Sub CopyValues()
    Dim i As Integer
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("数据分析")
    
    For i = 2 To 36
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next i


    Set ws = ThisWorkbook.Sheets("数据分析社区篇")
    
     For i = 3 To 27
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next

     ws.Range("H" & 36).Value = ws.Range("B" & 36).Value
    
    Set ws = ThisWorkbook.Sheets("数据分析单位篇")
    
     For i = 28 To 38
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next i
     
End Sub

Sub DeleteRowsAndColumns()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("数据") '将“数据”更改为你要删除行的工作表名称
    
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    
    '清空第二行的A到AG列
    With ws
        .Range("A2:AO2").ClearContents
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
    
  
    '获取打开的工作簿的第一个工作表
    Set sourceSheet = ThisWorkbook.Sheets(5)
    
    '获取最后一行的行号
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '将第二行到最后一行的A到AG列复制到当前工作表名为“数据”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("数据")
    sourceSheet.Range("A2:AO" & LastRow).Copy targetSheet.Range("A2")
    
    
    ' 获取最后一行的行号
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' 将 AH2 公式填充到整个 AH 列
    Range("AP2").AutoFill Destination:=Range("AP2:AP" & LastRow), Type:=xlFillDefault
    Range("AQ2").AutoFill Destination:=Range("AQ2:AQ" & LastRow), Type:=xlFillDefault

    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub

Sub SortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("数据分析社区篇")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("D3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A3:K27")
        
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
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A28:K34")
        
        '应用排序
        .Header = xlNo '不包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
End Sub



