Attribute VB_Name = "Module2"
Sub ���������������()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    
    ' ���������� ����������� ��� ��������� ������������������
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' ��������� ���������� ���� ��� ������ ������
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog4(Vb.Path)
    
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
    
    ' �������� �� �������� ���� �������� �����
    Do While SheetExists4(targetSheetName, Vb)
        If Len(targetSheetName) + Len(CStr(suffix)) > maxSheetNameLength Then
            targetSheetName = Left(targetSheetName, maxSheetNameLength - Len(CStr(suffix)) - 1)
        End If
        targetSheetName = targetSheetName & suffix
        suffix = suffix + 1
    Loop
    
    ' ������� ����� ���� � ���������� ������
    Dim targetSheet As Worksheet
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName
    
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
                ' �������� ������ �� ���� ������ � ������� ����
                If ws.UsedRange.Rows.count > 0 And ws.UsedRange.Columns.count > 0 Then
                    ws.UsedRange.Copy
                    targetSheet.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteAll
                    
                    ' ��������� �������� ��������� ������
                    lastRow = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row + 1
                End If
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
Function SheetExists4(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists4 = Not ws Is Nothing
    On Error GoTo 0
End Function

' ������� ��� ����������� ����������� ���� ������ ������ (������ ���������� OpenFileDialog3)
Function OpenFileDialog4(initialPath As String) As Collection
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
    
    Set OpenFileDialog4 = selectedFiles
End Function
