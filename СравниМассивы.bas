Attribute VB_Name = "Module5"
Sub �������������()
    Dim firstRange As Range
    Dim secondRange As Range
    Dim response As VbMsgBoxResult
    Dim isEqual As Boolean
    Dim i As Long, j As Long
    
    ' ������ �� ��������� ������� �������
    On Error Resume Next
    Set firstRange = Application.InputBox("�������� ������ ������", "����� �������", Type:=8)
    On Error GoTo 0
    
    ' ���� ������������ ����� "������", ������� �� �������
    If firstRange Is Nothing Then
        MsgBox "������ ������ �� ������. ������ ��������.", vbExclamation
        Exit Sub
    End If
    
    ' ������ �� ��������� ������� �������
    On Error Resume Next
    Set secondRange = Application.InputBox("�������� ������ ������", "����� �������", Type:=8)
    On Error GoTo 0
    
    ' ���� ������������ ����� "������", ������� �� �������
    If secondRange Is Nothing Then
        MsgBox "������ ������ �� ������. ������ ��������.", vbExclamation
        Exit Sub
    End If
    
    ' �������� �� ��������� �������� ��������
    If firstRange.Rows.count <> secondRange.Rows.count Or firstRange.Columns.count <> secondRange.Columns.count Then
        response = MsgBox("���������� ������� �� �����. ����������?", vbYesNo + vbExclamation, "��������������")
        If response = vbNo Then
            Exit Sub
        End If
    End If
    
    ' ��������� �������� � ��������
    isEqual = True
    For i = 1 To firstRange.Rows.count
        For j = 1 To firstRange.Columns.count
            If firstRange.Cells(i, j).Value <> secondRange.Cells(i, j).Value Then
                firstRange.Cells(i, j).Interior.Color = RGB(222, 180, 180) ' ������������ ������
                isEqual = False
            Else
                firstRange.Cells(i, j).Interior.ColorIndex = xlNone ' ���������� ����, ���� �������� �����
            End If
        Next j
    Next i
    
    ' ��������� � ����������� ���������
    If isEqual Then
        MsgBox "������� ���������.", vbInformation
    Else
        MsgBox "������� �������� � ��������.", vbExclamation
    End If

End Sub

