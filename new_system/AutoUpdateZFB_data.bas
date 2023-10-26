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
    Set ws = ThisWorkbook.Worksheets("ִ����") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:H2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("2:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub

Sub DeleteRowsAndColumnsZFB2()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ִ����ƥ��") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:H2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("2:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    Application.ScreenUpdating = True '�ָ���Ļ����
End Sub


Sub ImportDataZFB()
    Dim selectedFile As Variant
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim LastRow As Long
    
    '�����ļ�ѡ�񴰿ڣ����û�ѡ��Ҫ�����Excel�ļ�
    selectedFile = Application.GetOpenFilename("Excel Files (*.xlsx),*.xlsx", , "Select a File to Import")
    
    '����û�δѡ���ļ����˳����ӹ���
    If selectedFile = False Then Exit Sub
    
    '��ѡ�����ļ�
    Workbooks.Open selectedFile
    
    '��ȡ�򿪵Ĺ������ĵ�һ��������
    Set sourceSheet = ActiveWorkbook.Sheets(1)
    
    '��ȡ���һ�е��к�
    If IsEmpty(sourceSheet.Range("B2")) Then
        LastRow = 1
    Else
        LastRow = sourceSheet.Cells(Rows.Count, 2).End(xlUp).Row
    End If
    
    '���ڶ��е����һ�е�B��H�и��Ƶ���ǰ��������Ϊ��ִ���족����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("ִ����")
    sourceSheet.Range("B2:H" & LastRow).Copy targetSheet.Range("B1")
    
    '�رմ򿪵Ĺ�����
    ActiveWorkbook.Close False
    
    ' ��ȡ���һ�е��к�
    LastRow = targetSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub


Sub NumberRowsZFB()

    ' ���á�ִ���족���
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ִ����")
    
    ' �������
    Dim i As Integer
    Dim LastRow As Long
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' ѭ������ÿһ��
    For i = 2 To LastRow
        ' ������е� B �������ݣ����� A ���������
        If ws.Cells(i, "B").Value <> "" Then
            ws.Cells(i, "A").Value = i - 1
        End If
    Next i
    
End Sub



Sub MatchDataZFB()
    Application.ScreenUpdating = False '��ͣˢ����Ļ���ӿ��ٶ�
    
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim matchSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("�ۺϲ�ѯ")
    Set targetSheet = ThisWorkbook.Sheets("ִ����ƥ��")
    Set matchSheet = ThisWorkbook.Sheets("ִ����")
    
    '��һ�������Ҳ���������
    Dim sourceLastRow As Long
    sourceLastRow = sourceSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Dim targetLastRow As Long
    targetLastRow = targetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    Dim i As Long, j As Long
    For i = 3 To sourceLastRow
        If InStr(1, sourceSheet.Range("AH" & i).Value, "�ۺ�����ִ��") > 0 Or InStr(1, sourceSheet.Range("AH" & i).Value, "���й����") > 0 Then
            targetLastRow = targetLastRow + 1
            targetSheet.Range("A" & targetLastRow).Value = sourceSheet.Range("C" & i).Value
            targetSheet.Range("B" & targetLastRow).Value = sourceSheet.Range("AH" & i).Value
        End If
    Next i
    
    '�ڶ��������Ҳ���������
    Dim matchLastRow As Long
    matchLastRow = matchSheet.Cells(Rows.Count, "B").End(xlUp).Row
    targetLastRow = targetSheet.Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To targetLastRow
        For j = 2 To matchLastRow
            If targetSheet.Range("A" & i).Value = matchSheet.Range("B" & j).Value Then
                targetSheet.Range("C" & i).Value = matchSheet.Range("E" & j).Value
                Exit For '�ҵ��˾��˳�ѭ�����ӿ��ٶ�
            End If
        Next j
    Next i
    
    
    Application.ScreenUpdating = True '�ָ�ˢ����Ļ
    MsgBox "ִ����ƥ����ɣ�"
End Sub


Sub AutoFillZFB()
    Dim LastRow As Long
    LastRow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To LastRow
        If Range("B" & i) <> "" And Range("C" & i) = "" Then
            If InStr(Range("B" & i), "ִ����") > 0 Then
                Range("C" & i) = "δ�ֲ�"
            ElseIf InStr(Range("B" & i), "���й����") > 0 Then
                Range("C" & i) = "���й����"
            End If
        End If
    Next i
    
    MsgBox "�Զ����ɹ���"
End Sub


Sub SortDataZFB()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ִ���Ӵ�������")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("C3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A3:D12")
        
        'Ӧ������
        .Header = xlNo '������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
    
        MsgBox "�Զ�����ɹ���"
End Sub

Sub UnHideRowsZFB()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim i As Long
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("ִ���Ӵ�������")
    
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    ' ѭ��ÿһ�У�������б����أ���ȡ������
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
    
    ' ���幤����
    Set ws = ThisWorkbook.Sheets("ִ���Ӵ�������")
    
    
    ' ��ȡ���һ��
    LastRow = ws.Cells(Rows.Count, "C").End(xlUp).Row
    
    
    ' ѭ��ÿһ�У����C��Ϊ0�������ظ���
    For i = 3 To LastRow
        If ws.Cells(i, "C").Value = 0 Then
            ws.Rows(i).Hidden = True
        End If
    Next i
End Sub

