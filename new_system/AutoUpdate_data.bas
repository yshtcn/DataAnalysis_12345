Attribute VB_Name = "AutoUpdate"
Sub Calling()
    Worksheets("新平台工单情况").Range("B3").Value = "更新中"
    DeleteRowsAndColumns
    ImportData
    UnHideRows
    SortData
    HideRows
    setNowTime
    RemoveBackgroundColor
    ChangeBackgroundColor
    Worksheets("新平台工单情况").Range("B3").Formula = "=SUMPRODUCT(--(TEXT('综合查询'!Y:Y,""yyyy-mm-dd"")=TEXT(A3,""yyyy-mm-dd"")))"
End Sub



Sub CopyValues()
    '选择表格
    Worksheets("新平台工单情况").Activate
    
    '复制A3到F3的值到A4到F4
    Range("A3:F3").Copy
    Range("A4:F4").PasteSpecial xlPasteValues
End Sub





Sub 快速删除综合查询数据()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("综合查询") '将“数据”更改为你要删除行的工作表名称
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    '清空第二行的A到AG列
    With ws
        .Range("A3:AG3").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 3 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub 快速删除导出累积数据表()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("导出累积数据") '将“数据”更改为你要删除行的工作表名称
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    '清空第二行的A到AG列
    With ws
        .Range("A2:K2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub 快速删除导出区间数据表()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("导出区间数据") '将“数据”更改为你要删除行的工作表名称
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    '清空第二行的A到AG列
    With ws
        .Range("A2:K2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '检查是否有超过2行的数据
            '删除第三行及其以下所有行
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '恢复屏幕更新
End Sub


Sub 导入累积数据()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '弹出文件选择窗口，让用户选择要导入的Excel文件
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '如果用户未选择文件，退出该子过程
    If selectedFile = False Then Exit Sub

    
    '打开工作表
    Workbooks.Open selectedFile
    
    '获取打开的工作簿的第一个工作表
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '获取最后一行的行号
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '将第二行到最后一行的A到BG列复制到当前工作表名为“数据”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("导出累积数据")
    sourceSheet.Range("A2:K" & LastRow).Copy targetSheet.Range("A2")
    
    '关闭打开的工作簿
    ActiveWorkbook.Close False
    
    '获取最后一行的行号
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub


Sub 导入区间数据()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '弹出文件选择窗口，让用户选择要导入的Excel文件
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '如果用户未选择文件，退出该子过程
    If selectedFile = False Then Exit Sub

    
    '打开工作表
    Workbooks.Open selectedFile
    
    '获取打开的工作簿的第一个工作表
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '获取最后一行的行号
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '将第二行到最后一行的A到BG列复制到当前工作表名为“数据”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("导出区间数据")
    sourceSheet.Range("A2:K" & LastRow).Copy targetSheet.Range("A2")
    
    '关闭打开的工作簿
    ActiveWorkbook.Close False
    
    '获取最后一行的行号
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub

Sub 导入综合查询()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '弹出文件选择窗口，让用户选择要导入的Excel文件
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '如果用户未选择文件，退出该子过程
    If selectedFile = False Then Exit Sub
    
    '弹出密码输入框，让用户输入密码
    pw = InputBox("请输入密码：")
    
    '如果用户未输入密码，退出该子过程
    If pw = "" Then Exit Sub
    
    '尝试使用输入的密码打开选定的文件，如果密码不正确，给出提示并退出该子过程
    On Error Resume Next
    Workbooks.Open selectedFile, , , , pw
    If Err.Number <> 0 Then
        MsgBox "密码不正确。"
        Exit Sub
    End If
    On Error GoTo 0
    
    '获取打开的工作簿的第一个工作表
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '获取最后一行的行号
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '将第二行到最后一行的A到BG列复制到当前工作表名为“数据”的相同位置
    Set targetSheet = ThisWorkbook.Sheets("综合查询")
    sourceSheet.Range("A3:BG" & LastRow).Copy targetSheet.Range("A3")
    
    '关闭打开的工作簿
    ActiveWorkbook.Close False
    
    '获取最后一行的行号
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '提醒用户导入已完成
    MsgBox "数据已成功导入。"
End Sub




Sub SortDataWithP()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("社区待办件情况")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("E3"), _
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
    
    Set ws = ThisWorkbook.Worksheets("单位待办件情况")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("F3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A3:F11")
        
        '应用排序
        .Header = xlNo '不包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
End Sub



Sub 部门社区表排序()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("社区待办件情况")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("C4"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A4:C28")
        
        '应用排序
        .Header = xlNo '包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
    
    Set ws = ThisWorkbook.Worksheets("单位待办件情况")
    
    With ws.Sort
        .SortFields.Clear '清除排序字段
        
        '添加排序字段
        .SortFields.Add Key:=ws.Range("D4"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '设置排序范围
        .SetRange ws.Range("A4:D14")
        
        '应用排序
        .Header = xlNo '不包含表头
        .MatchCase = False '不区分大小写
        .Orientation = xlTopToBottom '排序方向：从上到下
        .SortMethod = xlPinYin '按拼音排序
        .Apply
    End With
End Sub


Sub setNowTime()
    Worksheets("新平台工单情况").Range("A3").Value = Date - 1
End Sub

Sub HideRowsWithP()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("单位待办件情况")
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果该行被隐藏，则取消隐藏
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果C列为0，则隐藏该行
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 And ws.Cells(i, "E").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("社区待办件情况")
    
    
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' 循环每一行，如果B列为0，则隐藏该行
    For i = 3 To LastRow
        If ws.Cells(i, "B").Value = 0 And ws.Cells(i, "D").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
End Sub
Sub UnHideRows()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("单位待办件情况")
    
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果该行被隐藏，则取消隐藏
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    Set ws = ThisWorkbook.Sheets("社区待办件情况")
    
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
End Sub





Sub HideRows()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("单位待办件情况")
    
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果该行被隐藏，则取消隐藏
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' 循环每一行，如果C列为0，则隐藏该行
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
    ' 定义工作表
    Set ws = ThisWorkbook.Sheets("社区待办件情况")
    
    
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' 获取最后一行
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' 循环每一行，如果B列为0，则隐藏该行
    For i = 3 To LastRow
        If ws.Cells(i, "B").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
End Sub

Sub ChangeCellContentForWeek()
    Sheets("新平台工单情况").Range("B2").Value = "本周受理工单（件）"
    Sheets("新平台工单情况").Range("C2").Value = "剩余待办结（件）"
    Sheets("单位待办件情况").Range("B1").Value = "图二：部门剩余待办结工单"
    Sheets("单位待办件情况").Range("C2").Value = "剩余待办结工单数(件）"
    Sheets("社区待办件情况").Range("A1").Value = "图三：社区剩余待办结工单"
    Sheets("社区待办件情况").Range("B2").Value = "剩余待办结工单数(件）"
End Sub



Sub ChangeCellContentForDay()
    Sheets("新平台工单情况").Range("B2").Value = "受理工单（件）"
    Sheets("新平台工单情况").Range("C2").Value = "待办结（件）"
    Sheets("单位待办件情况").Range("B1").Value = "图二：部门待办结工单"
    Sheets("单位待办件情况").Range("C2").Value = "待办结工单数(件）"
    Sheets("社区待办件情况").Range("A1").Value = "图三：社区待办结工单"
    Sheets("社区待办件情况").Range("B2").Value = "待办结工单数(件）"
End Sub

Sub UnmergeAndCenter()
    With Worksheets("单位待办件情况").Range("B1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("单位待办件情况").Range("B1:D1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Worksheets("社区待办件情况").Range("A1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("社区待办件情况").Range("A1:C1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub UnmergeAndCenterWithP()
    With Worksheets("单位待办件情况").Range("B1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("单位待办件情况").Range("B1:F1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Worksheets("社区待办件情况").Range("E1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("社区待办件情况").Range("A1:E1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub RemoveBackgroundColor()
    Worksheets("单位待办件情况").Range("C3:F17").Interior.ColorIndex = xlNone
    Worksheets("社区待办件情况").Range("B3:E27").Interior.ColorIndex = xlNone
End Sub

Sub ChangeBackgroundColor()
    Worksheets("单位待办件情况").Range("C11:F12").Interior.Color = RGB(255, 153, 204)
    Worksheets("社区待办件情况").Range("B27:E27").Interior.Color = RGB(255, 153, 204)
End Sub






Sub DeleteRowsAndCopyData()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim LastRow As Long, i As Long, j As Long
    
    '获取“执法办综合匹配”和“综合查询”两个工作表
    Set ws1 = ThisWorkbook.Sheets("执法办综合匹配")
    Set ws2 = ThisWorkbook.Sheets("综合查询")
    
    '删除“执法办综合匹配”第二行到最后一行
    LastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    If LastRow > 1 Then
        ws1.Range("A2:A" & LastRow).EntireRow.Delete
    End If
    
    '查找“综合查询”中包含“综合行政执法办”的行，并将C列和AD列复制到“执法办综合匹配”中
    LastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    If LastRow > 2 Then
        LastRow = ws2.Cells(LastRow, "A").End(xlUp).Row
    ElseIf LastRow <= 2 Then
        LastRow = ws2.Cells(3, "A").End(xlDown).Row
    End If
    For i = 3 To LastRow
        If InStr(ws2.Range("AD" & i).Value, "综合行政执法办") > 0 Then
            j = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row + 1
            ws1.Range("A" & j).Value = ws2.Range("C" & i).Value
            ws1.Range("B" & j).Value = ws2.Range("AD" & i).Value
        End If
    Next i
End Sub

Sub DeleteRows()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim lastRow1 As Long, lastRow2 As Long
    Dim dateThreshold As Date
    Dim i As Long
    
    Worksheets("新平台工单情况").Range("B3").Value = "更新中"
    
    '设置工作表对象
    Set ws1 = ThisWorkbook.Worksheets("综合查询")
    Set ws2 = ThisWorkbook.Worksheets("新平台工单情况")
    
    Application.ScreenUpdating = False '关闭屏幕更新以加快执行速度
    
    '获取日期阈值
    dateThreshold = Int(ws2.Range("A3").Value) ' 只保留日期部分
    
    '获取最后一行行号
    lastRow1 = ws1.Cells(ws1.Rows.Count, "Y").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    '循环遍历综合查询工作表的第三行到最后一行
    For i = lastRow1 To 3 Step -1
        '如果Y列的日期大于等于阈值，则删除这一行
        If Int(ws1.Cells(i, "Y").Value) > dateThreshold Then ' 只比较日期部分
            ws1.Rows(i).Delete
        End If
    Next i
    
    Application.ScreenUpdating = True '恢复屏幕更新
    Worksheets("新平台工单情况").Range("B3").Formula = "=SUMPRODUCT(--(TEXT('综合查询'!Y:Y,""yyyy-mm-dd"")=TEXT(A3,""yyyy-mm-dd"")))"
End Sub

