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
    
    Set ws = ThisWorkbook.Sheets("���ݷ���")
    
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

    Set ws = ThisWorkbook.Sheets("���ݷ�������ƪ")
    
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

    Set ws = ThisWorkbook.Sheets("���ݷ�����λƪ")
    
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
    Set ws = ThisWorkbook.Worksheets("����") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:AG2").ClearContents
        If .Cells(.Rows.Count, 1).End(xlUp).Row > 2 Then '����Ƿ��г���2�е�����
            'ɾ�������м�������������
            .Range("3:" & .Cells(.Rows.Count, 1).End(xlUp).Row).Delete
        End If
    End With
    
    Application.ScreenUpdating = True '�ָ���Ļ����
    
End Sub

Sub ImportData()
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
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '���ڶ��е����һ�е�A��AG�и��Ƶ���ǰ��������Ϊ�����ݡ�����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("����")
    sourceSheet.Range("A2:AG" & LastRow).Copy targetSheet.Range("A2")
    
    '�رմ򿪵Ĺ�����
    ActiveWorkbook.Close False


    
    ' ��ȡ���һ�е��к�
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' �� AH2 ��ʽ��䵽���� AH ��
    Range("AH2").AutoFill Destination:=Range("AH2:AH" & LastRow), Type:=xlFillDefault
    Range("AI2").AutoFill Destination:=Range("AI2:AI" & LastRow), Type:=xlFillDefault
    Range("AJ2").AutoFill Destination:=Range("AJ2:AJ" & LastRow), Type:=xlFillDefault
    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub

Sub Add7Days()
    Dim rng As Range
    Set rng = ThisWorkbook.Sheets("���ݷ�������ƪ").Range("K2:Q2")
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = cell.Value + 7
        End If
    Next cell
    Set rng = ThisWorkbook.Sheets("���ݷ�����λƪ").Range("K2:Q2")
    For Each cell In rng
        If IsDate(cell.Value) Then
            cell.Value = cell.Value + 7
        End If
    Next cell
End Sub

Sub SortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("���ݷ�������ƪ")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("D3"), _
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
    
    Set ws = ThisWorkbook.Worksheets("���ݷ�����λƪ")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("D3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlDescending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A3:D9")
        
        'Ӧ������
        .Header = xlNo '������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
End Sub







