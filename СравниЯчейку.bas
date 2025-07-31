Attribute VB_Name = "Module6"
Sub ������������()
    Dim range1 As Range, range2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim data1 As String, data2 As String
    Dim words1() As String, words2() As String
    Dim i As Long, j As Long
    Dim hasDifferences As Boolean
    
    ' ������ ������ ������� ������� �����
    On Error Resume Next
    Set range1 = Application.InputBox("�������� ������ ������ ����� ��� ���������", Type:=8)
    On Error GoTo 0
    If range1 Is Nothing Then Exit Sub ' ���� ������������ ������� �����
    
    ' ������ ������ ������� ������� �����
    On Error Resume Next
    Set range2 = Application.InputBox("�������� ������ ������ ����� ��� ���������", Type:=8)
    On Error GoTo 0
    If range2 Is Nothing Then Exit Sub ' ���� ������������ ������� �����
    
    ' ���������, ��� ������� ����� ����� ���������� ������
    If range1.Cells.count <> range2.Cells.count Then
        MsgBox "������� ����� ����� ������ ������. ��������� ����������.", vbExclamation
        Exit Sub
    End If
    
    ' �������������� ����
    hasDifferences = False
    
    ' ���������� ������ ������ � ��������
    For i = 1 To range1.Cells.count
        Set cell1 = range1.Cells(i)
        Set cell2 = range2.Cells(i)
        
        ' �������� ������ �� ����� (� ������ ����������� �����)
        data1 = cell1.MergeArea.Cells(1, 1).Value
        data2 = cell2.MergeArea.Cells(1, 1).Value
        
        ' ���������� �������������� ������
        cell1.MergeArea.Cells(1, 1).Font.ColorIndex = xlAutomatic
        cell2.MergeArea.Cells(1, 1).Font.ColorIndex = xlAutomatic
        
        ' ��������� ������ �� ����� (�� ��������)
        words1 = Split(data1, " ")
        words2 = Split(data2, " ")
        
        ' ���������� �����
        For j = LBound(words1) To UBound(words1)
            Dim found As Boolean
            found = False
            
            ' ���� ������� ����� �� ������ ������
            Dim k As Long
            For k = LBound(words2) To UBound(words2)
                If AreWordsSimilar(words1(j), words2(k)) Then
                    found = True
                    ' ���������� ������� � ������� ������
                    CompareAndHighlightWords cell1, words1(j), cell2, words2(k), data1, data2
                    Exit For
                End If
            Next k
            
            ' ���� ����� �� �������, �������� ��� �������
            If Not found Then
                HighlightText cell1, words1(j), data1
                hasDifferences = True
            End If
        Next j
    Next i
    
    ' ������� ��������� � ����������� �� ������� �����������
    If hasDifferences Then
        MsgBox "����������� ������� � ���������� �������.", vbInformation
    Else
        MsgBox "����������� �����������.", vbInformation
    End If
End Sub

' ������� ��� ��������� ���� (���������� ����������)
Function AreWordsSimilar(word1 As String, word2 As String) As Boolean
    ' ���� ����� ��������� �� 70% ��� �����, ������� �� ��������
    Dim similarityThreshold As Double
    similarityThreshold = 0.7
    
    ' ��������� ������� ����������
    Dim matchCount As Long
    matchCount = 0
    Dim i As Long
    For i = 1 To Len(word1)
        If i <= Len(word2) And Mid(word1, i, 1) = Mid(word2, i, 1) Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' ���������� ���������
    AreWordsSimilar = (matchCount / Len(word1)) >= similarityThreshold
End Function

' ������� ��� ��������� � ��������� �������� � ������� ������
Sub CompareAndHighlightWords(cell1 As Range, word1 As String, cell2 As Range, word2 As String, fullText1 As String, fullText2 As String)
    Dim i As Long
    For i = 1 To Len(word1)
        If i > Len(word2) Or Mid(word1, i, 1) <> Mid(word2, i, 1) Then
            ' ���� ������� �� ���������, �������� ������ ������ � ������ �����
            HighlightText cell1, Mid(word1, i, 1), fullText1
        End If
    Next i
End Sub

' ��������������� ������� ��� ��������� ������
Sub HighlightText(cell As Range, textToHighlight As String, fullText As String)
    Dim startPos As Long
    startPos = InStr(fullText, textToHighlight)
    If startPos > 0 Then
        With cell.MergeArea.Cells(1, 1).Characters(startPos, Len(textToHighlight)).Font
            .Color = RGB(255, 0, 0) ' ������� ����
        End With
    End If
End Sub
