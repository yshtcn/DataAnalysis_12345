Attribute VB_Name = "ģ��1"
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
    
    Set ws = ThisWorkbook.Sheets("���ݷ���")
    
    For i = 2 To 36
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next i


    Set ws = ThisWorkbook.Sheets("���ݷ�������ƪ")
    
     For i = 3 To 27
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next

     ws.Range("H" & 36).Value = ws.Range("B" & 36).Value
    
    Set ws = ThisWorkbook.Sheets("���ݷ�����λƪ")
    
     For i = 28 To 38
        ws.Range("H" & i).Value = ws.Range("B" & i).Value
        ws.Range("I" & i).Value = ws.Range("C" & i).Value
        ws.Range("J" & i).Value = ws.Range("D" & i).Value
        ws.Range("K" & i).Value = ws.Range("E" & i).Value
    Next i
     
End Sub

Sub DeleteRowsAndColumns()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("����") '�������ݡ�����Ϊ��Ҫɾ���еĹ���������
    
    Application.ScreenUpdating = False '�ر���Ļ�����Լӿ�ִ���ٶ�
    
    '��յڶ��е�A��AG��
    With ws
        .Range("A2:AO2").ClearContents
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
    
  
    '��ȡ�򿪵Ĺ������ĵ�һ��������
    Set sourceSheet = ThisWorkbook.Sheets(5)
    
    '��ȡ���һ�е��к�
    LastRow = sourceSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    '���ڶ��е����һ�е�A��AG�и��Ƶ���ǰ��������Ϊ�����ݡ�����ͬλ��
    Set targetSheet = ThisWorkbook.Sheets("����")
    sourceSheet.Range("A2:AO" & LastRow).Copy targetSheet.Range("A2")
    
    
    ' ��ȡ���һ�е��к�
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' �� AH2 ��ʽ��䵽���� AH ��
    Range("AP2").AutoFill Destination:=Range("AP2:AP" & LastRow), Type:=xlFillDefault
    Range("AQ2").AutoFill Destination:=Range("AQ2:AQ" & LastRow), Type:=xlFillDefault

    
    '�����û����������
    MsgBox "�����ѳɹ����롣"
End Sub

Sub SortData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("���ݷ�������ƪ")
    
    With ws.Sort
        .SortFields.Clear '��������ֶ�
        
        '��������ֶ�
        .SortFields.Add Key:=ws.Range("D3"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A3:K27")
        
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
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        
        '��������Χ
        .SetRange ws.Range("A28:K34")
        
        'Ӧ������
        .Header = xlNo '��������ͷ
        .MatchCase = False '�����ִ�Сд
        .Orientation = xlTopToBottom '�����򣺴��ϵ���
        .SortMethod = xlPinYin '��ƴ������
        .Apply
    End With
End Sub



