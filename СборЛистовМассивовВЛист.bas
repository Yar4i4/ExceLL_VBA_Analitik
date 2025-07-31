Attribute VB_Name = "Module4"
Option Explicit

Sub �����������������������()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook

    ' ���������� ����������� ��� ��������� ������������������
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler ' �������� ����� ��������� ������

    ' �������� ������ �� ����� "�������"
    Dim mainSheet As Worksheet
    Set mainSheet = Nothing ' ���������� ����� ����������
    On Error Resume Next ' �������� ���������� ������, ���� ���� �� ������
    Set mainSheet = Vb.Sheets("�������")
    On Error GoTo ErrorHandler ' ��������������� ��������� ������
    If mainSheet Is Nothing Then
        MsgBox "���� '�������' �� ������ � ���� �����.", vbCritical
        GoTo Cleanup
    End If

    ' ��������� ������������ ��������� �������� � ��������� ����
    Dim colStart As String, colEnd As String
    Dim searchStart As String, searchEnd As String

    ' �������� ������� ��� ������� ������
    colStart = CStr(mainSheet.Range("B13").Value) ' ���������� CStr ��� ����������
    If Not IsValidColumn(colStart) Then
        MsgBox "� ������ B13 ����� '�������' ������ ������������ ������������� ������� (��������� A-ZZZ).", vbExclamation, "������ �����"
        GoTo Cleanup
    End If

    ' �������� ���������� ����� ��� ������� ������
    searchStart = Trim(CStr(mainSheet.Range("C13").Value))
    If searchStart = "" Then
        MsgBox "� ������ C13 ����� '�������' �� ������� ����� ��� ������ ������ �������.", vbExclamation, "������ �����"
        GoTo Cleanup
    End If

    ' �������� ������� ��� ������ ������
    colEnd = CStr(mainSheet.Range("D13").Value)
    If Not IsValidColumn(colEnd) Then
        MsgBox "� ������ D13 ����� '�������' ������ ������������ ������������� ������� (��������� A-ZZZ).", vbExclamation, "������ �����"
        GoTo Cleanup
    End If

    ' �������� ���������� ����� ��� ������ ������
    searchEnd = Trim(CStr(mainSheet.Range("E13").Value))
    If searchEnd = "" Then
        MsgBox "� ������ E13 ����� '�������' �� ������� ����� ��� ������ ����� �������.", vbExclamation, "������ �����"
        GoTo Cleanup
    End If

    ' ��������� ���������� ���� ��� ������ ������
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog5(Vb.Path)

    ' ���������, ��� �� ������ ���� �� ���� ����
    If filePaths.count = 0 Then
        MsgBox "����� �� �������. �������� ��������.", vbInformation, "������"
        GoTo Cleanup ' ���������� GoTo Cleanup ��� �������������� ��������
    End If

    ' --- �������� ����� ��� �������� ---
    Dim collisionSheet As Worksheet
    Dim collisionSheetName As String
    Dim suffix As Long ' ���������� Long ��� ��������
    collisionSheetName = "��������"
    suffix = 1

    ' ������� ���������� ��� ��� ����� ��������
    Do While SheetExists6(collisionSheetName, Vb)
        collisionSheetName = "��������" & suffix
        suffix = suffix + 1
    Loop

    ' ������� ���� ��� �������� � ����������� ���������
    Set collisionSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    collisionSheet.Name = collisionSheetName
    With collisionSheet.Rows(1)
        .Cells(1, 1).Value = "�����"
        .Cells(1, 2).Value = "����"
        .Cells(1, 3).Value = "������� ������"
        .Cells(1, 4).Value = "������� ����� (�� �������)"
        .RowHeight = 45
        .WrapText = True
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    collisionSheet.Columns("A:D").ColumnWidth = 20 ' ������� ����

    Dim collisionRow As Long
    collisionRow = 2 ' ��������� ������ ��� ������ ��������
    Dim collisionFound As Boolean
    collisionFound = False

    ' --- �������� �������� ����� ��� ������ ---
    Dim targetSheet As Worksheet
    Dim targetSheetName As String
    targetSheetName = "����"
    suffix = 1

    ' ������� ��������� ��������� ����� ��� �����
    Do While SheetExists6(targetSheetName & suffix, Vb)
        suffix = suffix + 1
    Loop

    ' ������� ����� ���� � ���������� ������
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName & suffix

    ' --- ��������� ������� ���������� ����� ---
    Dim lastRowTarget As Long ' ������������� ��� �������
    lastRowTarget = 1 ' ��������� ������ ��� ������ ������ �� ������� ����

    Dim filePath As Variant ' Variant ��� �������� �� ���������
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim startRows As Collection
    Dim endRows As Collection
    Dim startRowVariant As Variant
    Dim endRowVariant As Variant
    Dim currentStartRow As Long
    Dim foundEndRow As Long
    Dim blockCopied As Boolean

    For Each filePath In filePaths
        Set wb = Nothing ' ���������� ����� ��������� ������ �����
        On Error Resume Next ' ���������� ������, ���� ���� �� ������� �������
        Set wb = Workbooks.Open(filePath, ReadOnly:=True) ' ��������� ������ ��� ������
        On Error GoTo ErrorHandler ' ��������������� ����������� ��������� ������

        If Not wb Is Nothing Then
            For Each ws In wb.Sheets
                If ws Is Nothing Then GoTo NextSheet ' ���������� ���������� ������� ������

                ' ���� ��� ������ � ���������� � ��������� ���������
                Set startRows = FindAllRowsBySearchWord(ws, colStart, searchStart)
                Set endRows = FindAllRowsBySearchWord(ws, colEnd, searchEnd)

                blockCopied = False ' ����, ��� ���� �� ���� ���� ���������� � ����� �����

                ' ���������, ������� �� ������ ������� �� �����
                If startRows.count = 0 Then
                    Call LogCollision(collisionSheet, collisionRow, wb.Name, ws.Name, colStart, searchStart)
                    collisionFound = True
                End If
                 If endRows.count = 0 Then
                    Call LogCollision(collisionSheet, collisionRow, wb.Name, ws.Name, colEnd, searchEnd)
                    collisionFound = True
                 End If

                ' ���� ������� � ���������, � �������� �������, ���� ����
                If startRows.count > 0 And endRows.count > 0 Then
                    ' ��������� ��������� ����� �� ����������� (Find ���������� � ������� ����������)
                    ' ��� ���������� ����� ���� �� ����������� ����������, �� Find ������ ������� ������ ����.

                    For Each startRowVariant In startRows
                        currentStartRow = CLng(startRowVariant)
                        foundEndRow = 0 ' ���������� ����� �������� ������ ��� ������ ���������

                        ' ���� ������ ���������� �������� ������ ����� ������� ���������
                        For Each endRowVariant In endRows
                            If CLng(endRowVariant) >= currentStartRow Then
                                foundEndRow = CLng(endRowVariant)
                                Exit For ' ����� ��������� ���������� �������� ������
                            End If
                        Next endRowVariant

                        ' ���� ����� ���������� ���� (������ <= �����)
                        If foundEndRow > 0 Then
                            On Error Resume Next ' ��������� ��������� ������ �����������/�������
                            ws.Rows(currentStartRow & ":" & foundEndRow).Copy
                            If Err.Number = 0 Then
                                targetSheet.Cells(lastRowTarget, 1).PasteSpecial Paste:=xlPasteAll
                                Application.CutCopyMode = False ' ������� ����� ������
                                If Err.Number = 0 Then
                                    lastRowTarget = lastRowTarget + (foundEndRow - currentStartRow + 1)
                                    blockCopied = True ' ��������, ��� ����������� ����
                                Else
                                     ' ������ ������� - ����� �������� �����������
                                    Err.Clear
                                End If
                            Else
                                ' ������ ����������� - ����� �������� �����������
                                Err.Clear
                                Application.CutCopyMode = False
                            End If
                             On Error GoTo ErrorHandler ' ��������������� ����� ���������
                        End If
                        ' ��������� � ��������� ��������� ������
                    Next startRowVariant
                End If

                ' ���� �� ������ ����� �� ���� �����������, � ������� ������ ���� ���� (�.�. ��������� �� �����),
                ' �� �� ������� ��� start <= end, ����� �������� ���. ������ �������� �����, ���� �����.
                ' ������� ������ ������������ ������ ������ ���������� ��������.

NextSheet:
                Set startRows = Nothing ' ����������� ������
                Set endRows = Nothing
            Next ws

            wb.Close SaveChanges:=False ' ��������� �������� �����
        Else
            ' ���������� ��������, ���� �� ������� ������� ����
             Call LogCollision(collisionSheet, collisionRow, CStr(filePath), "N/A", "����", "�� ������� �������")
             collisionFound = True
            ' MsgBox "�� ������� ������� ����: " & filePath, vbExclamation ' ��������� � �����
        End If
    Next filePath

    ' --- ���������� ---
    ' ������� ���� ��������, ���� �������� �� ����
    If Not collisionFound Then
        Application.DisplayAlerts = False ' �������� ��������� �������������� �� ��������
        collisionSheet.Delete
        Application.DisplayAlerts = True
    Else
        collisionSheet.Columns.AutoFit ' ��������� ������ �������� �� ����� ��������
    End If

    ' ������� ������� ����, ���� �� ���� (������ �� �����������)
    On Error Resume Next
    If IsSheetEmpty(targetSheet) Then ' ���������, ���� �� ����
        Application.DisplayAlerts = False
        targetSheet.Delete
        Application.DisplayAlerts = True
        MsgBox "������ ��� ����������� �� ������� �� �������� ���������.", vbInformation, "���������"
    Else
        targetSheet.Columns.AutoFit ' ��������� ������ �������� �� ������� �����
        targetSheet.Cells(1, 1).Select ' ���������� ������ ������
        If collisionFound Then
            MsgBox "���� ������ ��������." & vbCrLf & "���������� ��������� ��������. ��������� ���� '" & collisionSheetName & "'!", vbExclamation, "��������� � ����������"
            collisionSheet.Activate ' ���������� ���� � ����������
        Else
            MsgBox "������ ������� ������� �� ���� '" & targetSheet.Name & "'!", vbInformation, "�����"
        End If
    End If
    On Error GoTo ErrorHandler

Cleanup:
    ' ��������������� ��������� Excel � ����� ������ (�����, ������, ������)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False ' ������� ������ ������

    Set Vb = Nothing
    Set mainSheet = Nothing
    Set filePaths = Nothing
    Set collisionSheet = Nothing
    Set targetSheet = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set startRows = Nothing
    Set endRows = Nothing

    Exit Sub ' ���������� �����

ErrorHandler:
    ' ���������� ������
    MsgBox "��������� �������������� ������:" & vbCrLf & vbCrLf & _
           "����� ������: " & Err.Number & vbCrLf & _
           "��������: " & Err.Description & vbCrLf & _
           "��������: " & Err.Source, vbCritical, "������ ����������"
    Resume Cleanup ' ��������� � ����� ������� ��� �������������� �������� Excel

End Sub

' --- ��������������� ������� ---

Function SheetExists6(sheetName As String, wb As Workbook) As Boolean
' ��������� ������������� ����� � ��������� ����� (��� ��������� ������)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists6 = Not ws Is Nothing
    On Error GoTo 0 ' ������������ ����������� ��������� ������
    Set ws = Nothing
End Function

Function OpenFileDialog5(initialPath As String) As Collection
' ���������� ���������� ���� ������ ������ Excel (������������� �����)
    Dim fileDialog As fileDialog
    Dim selectedFiles As New Collection
    Dim selectedFile As Variant

    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "�������� ���� ��� ��������� ������ Excel ��� ���������"
        If Right(initialPath, 1) <> "\" Then initialPath = initialPath & "\"
        .InitialFileName = initialPath ' ������������� ��������� �����
        .Filters.Clear
        .Filters.Add "����� Excel", "*.xls; *.xlsx; *.xlsm", 1 ' ��������� ������ � ������ ��� �� ���������
        .AllowMultiSelect = True ' ��������� ����� ���������� ������

        If .Show = -1 Then ' -1 ��������, ��� ������������ ����� OK
            For Each selectedFile In .SelectedItems
                selectedFiles.Add selectedFile
            Next selectedFile
        End If ' ���� ������������ ����� ������, ��������� ��������� ������
    End With

    Set OpenFileDialog5 = selectedFiles
    Set fileDialog = Nothing ' ����������� ������
End Function

Function IsValidColumn(col As String) As Boolean
' ���������, �������� �� ������ ���������� ��������������� ������� (A-ZZZ)
' ���������� ����� ������� �������� ����� � ��������
    Dim i As Long
    Dim L As Long
    Dim charCode As Integer

    col = Trim(UCase(col)) ' ������� �������, ��������� � ������� �������
    L = Len(col)

    If L = 0 Or L > 3 Then ' ���������� ����� �� 1 �� 3
        IsValidColumn = False
        Exit Function
    End If

    For i = 1 To L
        charCode = Asc(Mid(col, i, 1))
        If charCode < 65 Or charCode > 90 Then ' 65='A', 90='Z'
            IsValidColumn = False
            Exit Function
        End If
    Next i

    ' �������������� �������� ��� 3 ���� (�� ������ ��������� "ZZZ")
    ' �� �������� ColumnLetterToNumber ������ 0, ���� ������� ����������,
    ' ������� ������� �������� ����� ����� ���� ���������, �� ������� ��� �������.
    ' �������� ColumnLetterToNumber(col) > 0 ����� �������.
    If ColumnLetterToNumber(col) > 18278 Then ' 18278 = ����� ������� ZZZ
         IsValidColumn = False
    Else
         IsValidColumn = True
    End If

End Function


Function FindAllRowsBySearchWord(ws As Worksheet, colLetter As String, searchWord As String) As Collection
' ������� ��� ������, ���������� ������ �������� searchWord � ��������� ������� colLetter
' ���������� ��������� ������� ����� ��� ������ ���������, ���� ������ �� �������
    Dim foundRows As New Collection
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim colNum As Long

    colNum = ColumnLetterToNumber(colLetter)
    If colNum = 0 Then ' ���������� �������
        Set FindAllRowsBySearchWord = foundRows ' ���������� ������ ���������
        Exit Function
    End If

    On Error Resume Next ' ���������� ������, ���� ������� ���� ��� �������
    Set searchRange = ws.Columns(colNum)
    If searchRange Is Nothing Then
         Set FindAllRowsBySearchWord = foundRows
         Exit Function
    End If
    On Error GoTo 0 ' ��������������� ��������� ������

    ' ���� ������ ���������
    Set foundCell = searchRange.Find(What:=searchWord, _
                                     After:=searchRange.Cells(searchRange.Cells.count), _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False) ' ������������ �������

    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Do
            foundRows.Add foundCell.row ' ��������� ����� ������ � ���������
            ' ���� ��������� ���������
            Set foundCell = searchRange.FindNext(foundCell)
            ' ���������, �� ��������� �� � ������ � �� �������� �� ������ ������
            If foundCell Is Nothing Then Exit Do
        Loop While foundCell.Address <> firstAddress
    End If

    Set FindAllRowsBySearchWord = foundRows ' ���������� ��������� (����� ���� ������)
    ' �������
    Set searchRange = Nothing
    Set foundCell = Nothing
End Function

Function ColumnLetterToNumber(colLetter As String) As Long
' ����������� ��������� ����������� ������� (A-ZZZ) � ��������
    Dim colNum As Long
    Dim i As Long
    Dim L As Long
    Dim charCode As Integer

    colLetter = Trim(UCase(colLetter))
    L = Len(colLetter)

    If L = 0 Or L > 3 Then GoTo InvalidInput

    For i = 1 To L
        charCode = Asc(Mid(colLetter, i, 1))
        If charCode < 65 Or charCode > 90 Then GoTo InvalidInput ' �� ����� A-Z
        colNum = colNum * 26 + (charCode - 64)
    Next i

    ' �������� ������������� �������� ��� Excel (16384 = XFD)
    If colNum > 16384 Then GoTo InvalidInput

    ColumnLetterToNumber = colNum
    Exit Function

InvalidInput:
    ColumnLetterToNumber = 0 ' ���������� 0 ��� ������
End Function

Sub LogCollision(sheet As Worksheet, ByRef rowNum As Long, bookName As String, sheetName As String, searchCol As String, missingWord As String)
' ���������� ���������� � �������� �� ���� ��������
    On Error Resume Next ' ���������� ������ ��� ������ �� ���� ��������
    With sheet
        .Cells(rowNum, 1).Value = bookName
        .Cells(rowNum, 2).Value = sheetName
        .Cells(rowNum, 3).Value = searchCol
        .Cells(rowNum, 4).Value = missingWord
    End With
    If Err.Number = 0 Then
        rowNum = rowNum + 1 ' ����������� ������� ����� ������ ��� �������� ������
    Else
        Err.Clear ' ������� ������, ���� ������ �� �������
    End If
    On Error GoTo 0 ' ��������������� ��������� ������ (���� � ���� ������������ ��� �� ����� ����� Resume Next)
End Sub

Function IsSheetEmpty(sht As Worksheet) As Boolean
' ���������, �������� �� ���� �����-���� ������
    On Error Resume Next
    IsSheetEmpty = sht.UsedRange.Address = sht.Range("A1").Address And IsEmpty(sht.Range("A1").Value)
    If Err.Number <> 0 Then ' ���� UsedRange �������� ������ (��������, ����� ������ �������?), ������� ������
        IsSheetEmpty = True
        Err.Clear
    End If
    On Error GoTo 0
End Function
