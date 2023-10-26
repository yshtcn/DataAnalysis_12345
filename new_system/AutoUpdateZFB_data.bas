Attribute VB_Name = "AutoUpdateZFB"
Sub CallingZFB()
    DeleteRowsAndColumnsZFB1
    DeleteRowsAndColumnsZFB2
    ImportDataZFB
    MatchDataZFB
    AutoFillZFB
    NumberRowsZFB
    UnHideRowsZFB
    SortDataZFB
    HideRowsZFB
End Sub


Sub DeleteRowsAndColumnsZFB1()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("执法办") '将“数据”更改为你要删除行的工作表名称
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    '清空第二行的A到AG列
    With ws
        .Range("A2:H2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("2:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub

Sub DeleteRowsAndColumnsZFB2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("执法办匹配") '将“数据”更改为你要删除行的工作表名称
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    '清空第二行的A到AG列
    With ws
        .Range("A2:H2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("2:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub ImportDataZFB()
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
    If IsEmpty(sourceSheet.Range("B2")) Then
        LastRow = 1
    Else
        LastRow = sourceSheet.Cells(Rows.Count, 2).End(xlUp).Row
    End If
    
    '将第二行到最后一行的B到H列复制到当前工作表名为“执法办”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("执法办")
    sourceSheet.Range("B2:H" & LastRow).Copy targetSheet.Range("B1")
    
    '关闭打开的工作簿
    ActiveWorkbook.Close False
    
    ' 获取最后一行的行号
    LastRow = targetSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub


Sub NumberRowsZFB()

    ' 引用“执法办”表格
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("执法办")
    
    ' 定义变量
    Dim i As Integer
    Dim LastRow As Long
    
    ' 获取最后一行
    LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' 循环遍历每一行
    For i = 2 To LastRow
        ' 如果该行的 B 列有数据，则在 A 列填入序号
        If ws.Cells(i, "B").Value <> "" Then
            ws.Cells(i, "A").Value = i - 1
        End If
    Next i
    
End Sub



Sub MatchDataZFB()
    Application.ScreenUpdating = False '暂停刷新屏幕，加快速度
    
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim matchSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("综合查询")
    Set targetSheet = ThisWorkbook.Sheets("执法办匹配")
    Set matchSheet = ThisWorkbook.Sheets("执法办")
    
    '第一步：查找并复制数据
    Dim sourceLastRow As Long
    sourceLastRow = sourceSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Dim targetLastRow As Long
    targetLastRow = targetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Dim i As Long, j As Long
    For i = 3 To sourceLastRow
        If InStr(1, sourceSheet.Range("AH" & i).Value, "综合行政执法") > 0 Or InStr(1, sourceSheet.Range("AH" & i).Value, "城市管理科") > 0 Then
            targetLastRow = targetLastRow + 1
            targetSheet.Range("A" & targetLastRow).Value = sourceSheet.Range("C" & i).Value
            targetSheet.Range("B" & targetLastRow).Value = sourceSheet.Range("AH" & i).Value
        End If
    Next i
    
    '第二步：查找并复制数据
    Dim matchLastRow As Long
    matchLastRow = matchSheet.Cells(Rows.Count, "B").End(xlUp).Row
    targetLastRow = targetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To targetLastRow
        For j = 2 To matchLastRow
            If targetSheet.Range("A" & i).Value = matchSheet.Range("B" & j).Value Then
                targetSheet.Range("C" & i).Value = matchSheet.Range("E" & j).Value
                Exit For '找到了就退出循环，加快速度
            End If
        Next j
    Next i
    
    
    Application.ScreenUpdating = True '恢复刷新屏幕
    MsgBox "执法办匹配完成！"
End Sub


Sub AutoFillZFB()
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To LastRow
        If Range("B" & i) <> "" And Range("C" & i) = "" Then
            If InStr(Range("B" & i), "执法办") > 0 Then
                Range("C" & i) = "未分拨"
            ElseIf InStr(Range("B" & i), "城市管理科") > 0 Then
                Range("C" & i) = "城市管理科"
            End If
        End If
    Next i
    
    MsgBox "自动填充成功！"
End Sub


Sub SortDataZFB()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("执法队待办件情况")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("C3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A3:D12")
        
        '应用排序
        .Header = xlNo '包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
    
        MsgBox "自动排序成功！"
End Sub

Sub UnHideRowsZFB()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("执法队待办件情况")
    
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果该行被隐藏，则取消隐藏
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
End Sub


Sub HideRowsZFB()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("执法队待办件情况")
    
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    
    ' 循环每一行，如果C列为0，则隐藏该行
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
End Sub

