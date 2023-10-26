Attribute VB_Name = "AutoUpdate"
Sub Calling()
    Worksheets("��ƽ̨�������").Range("B3").Value = "������"
    DeleteRowsAndColumns
    ImportData
    UnHideRows
    SortData
    HideRows
    setNowTime
    RemoveBackgroundColor
    ChangeBackgroundColor
    Worksheets("��ƽ̨�������").Range("B3").Formula = "=SUMPRODUCT(--(TEXT('�ۺϲ�ѯ'!Y:Y,""yyyy-mm-dd"")=TEXT(A3,""yyyy-mm-dd"")))"
End Sub



Sub CopyValues()
    'ѡ����
    Worksheets("��ƽ̨�������").Activate
    
    '����A3��F3��ֵ��A4��F4
    Range("A3:F3").Copy
    Range("A4:F4").PasteSpecial xlPasteValues
End Sub





Sub ����ɾ���ۺϲ�ѯ����()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("�ۺϲ�ѯ") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    '��յڶ��е�A��AG��
    With ws
        .Range("A3:AG3").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 3 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub ����ɾ�������ۻ����ݱ�()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("�����ۻ�����") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:K2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub ����ɾ�������������ݱ�()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("������������") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:K2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub �����ۻ�����()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '�����ļ�ѡ�񴰿ڣ����û�ѡ��Ҫ�����Excel�ļ�
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '����û�δѡ���ļ����˳����ӹ���
    If selectedFile = False Then Exit Sub

    
    '�򿪹�����
    Workbooks.Open selectedFile
    
    '��ȡ�򿪵Ĺ������ĵ�һ��������
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '��ȡ���һ�е��к�
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '���ڶ��е����һ�е�A��BG�и��Ƶ���ǰ��������Ϊ�����ݡ�����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("�����ۻ�����")
    sourceSheet.Range("A2:K" & LastRow).Copy targetSheet.Range("A2")
    
    '�رմ򿪵Ĺ�����
    ActiveWorkbook.Close False
    
    '��ȡ���һ�е��к�
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub


Sub ������������()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '�����ļ�ѡ�񴰿ڣ����û�ѡ��Ҫ�����Excel�ļ�
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '����û�δѡ���ļ����˳����ӹ���
    If selectedFile = False Then Exit Sub

    
    '�򿪹�����
    Workbooks.Open selectedFile
    
    '��ȡ�򿪵Ĺ������ĵ�һ��������
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '��ȡ���һ�е��к�
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '���ڶ��е����һ�е�A��BG�и��Ƶ���ǰ��������Ϊ�����ݡ�����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("������������")
    sourceSheet.Range("A2:K" & LastRow).Copy targetSheet.Range("A2")
    
    '�رմ򿪵Ĺ�����
    ActiveWorkbook.Close False
    
    '��ȡ���һ�е��к�
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub

Sub �����ۺϲ�ѯ()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    Dim pw As String
    
    '�����ļ�ѡ�񴰿ڣ����û�ѡ��Ҫ�����Excel�ļ�
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '����û�δѡ���ļ����˳����ӹ���
    If selectedFile = False Then Exit Sub
    
    '����������������û���������
    pw = InputBox("���������룺")
    
    '����û�δ�������룬�˳����ӹ���
    If pw = "" Then Exit Sub
    
    '����ʹ������������ѡ�����ļ���������벻��ȷ��������ʾ���˳����ӹ���
    On Error Resume Next
    Workbooks.Open selectedFile, , , , pw
    If Err.Number <> 0 Then
        MsgBox "���벻��ȷ��"
        Exit Sub
    End If
    On Error GoTo 0
    
    '��ȡ�򿪵Ĺ������ĵ�һ��������
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '��ȡ���һ�е��к�
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '���ڶ��е����һ�е�A��BG�и��Ƶ���ǰ��������Ϊ�����ݡ�����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("�ۺϲ�ѯ")
    sourceSheet.Range("A3:BG" & LastRow).Copy targetSheet.Range("A3")
    
    '�رմ򿪵Ĺ�����
    ActiveWorkbook.Close False
    
    '��ȡ���һ�е��к�
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub




Sub SortDataWithP()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("������������")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("E3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A3:E27")
        
        'Ӧ������
        .Header = xlNo '������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
    
    Set ws = ThisWorkbook.Worksheets("��λ��������")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("F3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A3:F11")
        
        'Ӧ������
        .Header = xlNo '��������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
End Sub



Sub ��������������()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("������������")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("C4"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A4:C28")
        
        'Ӧ������
        .Header = xlNo '������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
    
    Set ws = ThisWorkbook.Worksheets("��λ��������")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("D4"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A4:D14")
        
        'Ӧ������
        .Header = xlNo '��������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
End Sub


Sub setNowTime()
    Worksheets("��ƽ̨�������").Range("A3").Value = Date - 1
End Sub

Sub HideRowsWithP()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("��λ��������")
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У�������б����أ���ȡ������
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У����C��Ϊ0�������ظ���
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 And ws.Cells(i, "E").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("������������")
    
    
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' ѭ��ÿһ�У����B��Ϊ0�������ظ���
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
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("��λ��������")
    
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У�������б����أ���ȡ������
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    Set ws = ThisWorkbook.Sheets("������������")
    
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
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("��λ��������")
    
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У�������б����أ���ȡ������
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У����C��Ϊ0�������ظ���
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("������������")
    
    
    For i = 1 To LastRow
        If ws.Rows(i).Hidden Then
            ws.Rows(i).Hidden = False
        End If
    Next i
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row
    
    ' ѭ��ÿһ�У����B��Ϊ0�������ظ���
    For i = 3 To LastRow
        If ws.Cells(i, "B").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
    
End Sub

Sub ChangeCellContentForWeek()
    Sheets("��ƽ̨�������").Range("B2").Value = "����������������"
    Sheets("��ƽ̨�������").Range("C2").Value = "ʣ�����ᣨ����"
    Sheets("��λ��������").Range("B1").Value = "ͼ��������ʣ�����Ṥ��"
    Sheets("��λ��������").Range("C2").Value = "ʣ�����Ṥ����(����"
    Sheets("������������").Range("A1").Value = "ͼ��������ʣ�����Ṥ��"
    Sheets("������������").Range("B2").Value = "ʣ�����Ṥ����(����"
End Sub



Sub ChangeCellContentForDay()
    Sheets("��ƽ̨�������").Range("B2").Value = "������������"
    Sheets("��ƽ̨�������").Range("C2").Value = "����ᣨ����"
    Sheets("��λ��������").Range("B1").Value = "ͼ�������Ŵ���Ṥ��"
    Sheets("��λ��������").Range("C2").Value = "����Ṥ����(����"
    Sheets("������������").Range("A1").Value = "ͼ������������Ṥ��"
    Sheets("������������").Range("B2").Value = "����Ṥ����(����"
End Sub

Sub UnmergeAndCenter()
    With Worksheets("��λ��������").Range("B1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("��λ��������").Range("B1:D1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Worksheets("������������").Range("A1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("������������").Range("A1:C1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub UnmergeAndCenterWithP()
    With Worksheets("��λ��������").Range("B1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("��λ��������").Range("B1:F1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Worksheets("������������").Range("E1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .UnMerge
    End With
    
    With Worksheets("������������").Range("A1:E1")
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Sub RemoveBackgroundColor()
    Worksheets("��λ��������").Range("C3:F17").Interior.ColorIndex = xlNone
    Worksheets("������������").Range("B3:E27").Interior.ColorIndex = xlNone
End Sub

Sub ChangeBackgroundColor()
    Worksheets("��λ��������").Range("C11:F12").Interior.Color = RGB(255, 153, 204)
    Worksheets("������������").Range("B27:E27").Interior.Color = RGB(255, 153, 204)
End Sub






Sub DeleteRowsAndCopyData()
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim LastRow As Long, i As Long, j As Long
    
    '��ȡ��ִ�����ۺ�ƥ�䡱�͡��ۺϲ�ѯ������������
    Set ws1 = ThisWorkbook.Sheets("ִ�����ۺ�ƥ��")
    Set ws2 = ThisWorkbook.Sheets("�ۺϲ�ѯ")
    
    'ɾ����ִ�����ۺ�ƥ�䡱�ڶ��е����һ��
    LastRow = ws1.Cells(ws1.Rows.Count, "A").End(xlUp).Row
    If LastRow > 1 Then
        ws1.Range("A2:A" & LastRow).EntireRow.Delete
    End If
    
    '���ҡ��ۺϲ�ѯ���а������ۺ�����ִ���족���У�����C�к�AD�и��Ƶ���ִ�����ۺ�ƥ�䡱��
    LastRow = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    If LastRow > 2 Then
        LastRow = ws2.Cells(LastRow, "A").End(xlUp).Row
    ElseIf LastRow <= 2 Then
        LastRow = ws2.Cells(3, "A").End(xlDown).Row
    End If
    For i = 3 To LastRow
        If InStr(ws2.Range("AD" & i).Value, "�ۺ�����ִ����") > 0 Then
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
    
    Worksheets("��ƽ̨�������").Range("B3").Value = "������"
    
    '���ù��������
    Set ws1 = ThisWorkbook.Worksheets("�ۺϲ�ѯ")
    Set ws2 = ThisWorkbook.Worksheets("��ƽ̨�������")
    
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    
    '��ȡ������ֵ
    dateThreshold = Int(ws2.Range("A3").Value) ' ֻ�������ڲ���
    
    '��ȡ���һ���к�
    lastRow1 = ws1.Cells(ws1.Rows.Count, "Y").End(xlUp).Row
    lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
    
    'ѭ�������ۺϲ�ѯ������ĵ����е����һ��
    For i = lastRow1 To 3 Step -1
        '���Y�е����ڴ��ڵ�����ֵ����ɾ����һ��
        If Int(ws1.Cells(i, "Y").Value) > dateThreshold Then ' ֻ�Ƚ����ڲ���
            ws1.Rows(i).Delete
        End If
    Next i
    
    Application.ScreenUpdating = True '�ָ���Ļ����
    Worksheets("��ƽ̨�������").Range("B3").Formula = "=SUMPRODUCT(--(TEXT('�ۺϲ�ѯ'!Y:Y,""yyyy-mm-dd"")=TEXT(A3,""yyyy-mm-dd"")))"
End Sub

