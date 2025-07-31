Attribute VB_Name = "Module3"
Sub ���������������������()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    
    ' ���������� ����������� ��� ��������� ������������������
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' �������� ������ ��� ����������� �� ������ E8 ����� "�������"
    Dim mainSheet As Worksheet
    Set mainSheet = Vb.Sheets("�������")
    Dim rowRanges As String
    rowRanges = mainSheet.Range("E8").Value
    
    ' ���������, ������� �� ������
    If Trim(rowRanges) = "" Then
        MsgBox "������� ������ ��� ������ � ������ E8!", vbExclamation
        Exit Sub
    End If
    
    ' �������� ������������ ������� �����
    If Not ValidateRowRanges(rowRanges) Then
        MsgBox "���� ����� ������������ � ������ E8 � ���� �������� ����� ���� (�����) ��� ��� ������������ ������ ����� �������", vbExclamation
        Exit Sub
    End If
    
    ' ��������� ���������� ���� ��� ������ ������
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog5(Vb.Path)
    
    ' ���������, ��� �� ������ ���� �� ���� ����
    If filePaths.count = 0 Then
        MsgBox "���� �� ������!", vbExclamation
        Exit Sub
    End If
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetSheetName As String
    Dim suffix As Integer
    Dim maxSheetNameLength As Integer
    maxSheetNameLength = 31 ' ������������ ����� ����� ����� � Excel
    
    ' ���������� ��� �������� �����
    targetSheetName = "����"
    suffix = 1
    
    ' ������� ��������� ��������� ����� ��� �����
    Do While SheetExists5(targetSheetName & suffix, Vb)
        suffix = suffix + 1
    Loop
    
    ' ������� ����� ���� � ���������� ������
    Dim targetSheet As Worksheet
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName & suffix
    
    ' ��������� ������� ���������� �����
    Dim lastRow As Long
    lastRow = 1 ' ��������� ������ ��� ������ ������
    
    For Each filePath In filePaths
        ' ��������� �����
        On Error Resume Next
        Set wb = Workbooks.Open(filePath)
        On Error GoTo 0
        
        If Not wb Is Nothing Then
            ' ���������� ��� ����� � �������� �����
            For Each ws In wb.Sheets
                ' �������� ������ ����� ��� �����������
                Dim rowsToCopy As Variant
                rowsToCopy = GetRowsToCopy(rowRanges)
                
                ' �������� ��������� ������
                Dim row As Variant
                For Each row In rowsToCopy
                    If row <= ws.UsedRange.Rows.count Then
                        ws.Rows(row).Copy
                        targetSheet.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteAll
                        lastRow = lastRow + 1
                    End If
                Next row
            Next ws
            
            ' ��������� �������� ����� ��� ���������� ���������
            wb.Close SaveChanges:=False
        Else
            MsgBox "�� ������� ������� ����: " & filePath, vbExclamation
        End If
    Next filePath
    
    ' ��������������� ��������� Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    MsgBox "������ ������� �������!", vbInformation
End Sub

' ������� ��� �������� ������������� ����� � ��������� �����
Function SheetExists5(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists5 = Not ws Is Nothing
    On Error GoTo 0
End Function

' ������� ��� ����������� ����������� ���� ������ ������ (������ ���������� OpenFileDialog3)
Function OpenFileDialog5(initialPath As String) As Collection
    Dim fileDialog As fileDialog
    Dim selectedFiles As New Collection
    Dim selectedFile As Variant
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "�������� ����� ��� �������"
        .InitialFileName = initialPath
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            For Each selectedFile In .SelectedItems
                selectedFiles.Add selectedFile
            Next selectedFile
        End If
    End With
    
    Set OpenFileDialog5 = selectedFiles
End Function

' ������� ��� ��������� ������� ����� �� ������ ����������
Function GetRowsToCopy(rowRanges As String) As Variant
    Dim result As Collection
    Set result = New Collection
    
    ' ��������� ������ �� �������
    Dim ranges As Variant
    ranges = Split(rowRanges, ",")
    
    Dim r As Variant
    For Each r In ranges
        ' ������� �������
        r = Trim(r)
        
        ' ���������, �������� �� �������� �����/����
        If InStr(r, "-") > 0 Or InStr(r, "�") > 0 Then
            ' ��������� �������� �� ��������� � �������� ������
            Dim startRow As Long, endRow As Long
            Dim parts As Variant
            parts = Split(r, "-")
            If UBound(parts) = 1 Then
                startRow = CLng(parts(0))
                endRow = CLng(parts(1))
                
                ' ��������� ��� ������ � ���������
                Dim i As Long
                For i = startRow To endRow
                    result.Add i
                Next i
            End If
        Else
            ' ��������� ��������� ������
            result.Add CLng(r)
        End If
    Next r
    
    ' ����������� ��������� � ������
    Dim output() As Long
    ReDim output(result.count - 1)
    Dim j As Long
    For j = 1 To result.count
        output(j - 1) = result(j)
    Next j
    
    GetRowsToCopy = output
End Function

' ������� ��� ��������� �����
Function ValidateRowRanges(rowRanges As String) As Boolean
    Dim ranges As Variant
    ranges = Split(rowRanges, ",")
    
    Dim r As Variant
    For Each r In ranges
        r = Trim(r)
        
        ' ���� �������� �������� �����/����
        If InStr(r, "-") > 0 Or InStr(r, "�") > 0 Then
            Dim parts As Variant
            parts = Split(r, "-")
            If UBound(parts) <> 1 Then
                ValidateRowRanges = False
                Exit Function
            End If
            
            ' ���������, ��� ��� ����� �����
            If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
                ValidateRowRanges = False
                Exit Function
            End If
        ElseIf Not IsNumeric(r) Then
            ' ���� ��� �� �����
            ValidateRowRanges = False
            Exit Function
        End If
    Next r
    
    ValidateRowRanges = True
End Function
