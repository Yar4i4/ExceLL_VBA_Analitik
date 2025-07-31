Attribute VB_Name = "Module2"
Sub СборЛистовВЛист()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    
    ' Отключение оптимизаций для улучшения производительности
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' Открываем диалоговое окно для выбора файлов
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog4(Vb.Path)
    
    ' Проверяем, был ли выбран хотя бы один файл
    If filePaths.count = 0 Then
        MsgBox "Файл не выбран!", vbExclamation
        Exit Sub
    End If
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetSheetName As String
    Dim suffix As Integer
    Dim maxSheetNameLength As Integer
    maxSheetNameLength = 31 ' Максимальная длина имени листа в Excel
    
    ' Определяем имя целевого листа
    targetSheetName = "Сбор"
    suffix = 1
    
    ' Проверка на конфликт имен целевого листа
    Do While SheetExists4(targetSheetName, Vb)
        If Len(targetSheetName) + Len(CStr(suffix)) > maxSheetNameLength Then
            targetSheetName = Left(targetSheetName, maxSheetNameLength - Len(CStr(suffix)) - 1)
        End If
        targetSheetName = targetSheetName & suffix
        suffix = suffix + 1
    Loop
    
    ' Создаем новый лист с уникальным именем
    Dim targetSheet As Worksheet
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName
    
    ' Обработка каждого выбранного файла
    Dim lastRow As Long
    lastRow = 1 ' Начальная строка для записи данных
    
    For Each filePath In filePaths
        ' Открываем книгу
        On Error Resume Next
        Set wb = Workbooks.Open(filePath)
        On Error GoTo 0
        
        If Not wb Is Nothing Then
            ' Перебираем все листы в открытой книге
            For Each ws In wb.Sheets
                ' Копируем данные со всех листов в целевой лист
                If ws.UsedRange.Rows.count > 0 And ws.UsedRange.Columns.count > 0 Then
                    ws.UsedRange.Copy
                    targetSheet.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteAll
                    
                    ' Обновляем значение последней строки
                    lastRow = targetSheet.Cells(targetSheet.Rows.count, 1).End(xlUp).row + 1
                End If
            Next ws
            
            ' Закрываем исходную книгу без сохранения изменений
            wb.Close SaveChanges:=False
        Else
            MsgBox "Не удалось открыть файл: " & filePath, vbExclamation
        End If
    Next filePath
    
    ' Восстанавливаем настройки Excel
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    
    MsgBox "Данные успешно собраны!", vbInformation
End Sub

' Функция для проверки существования листа в указанной книге
Function SheetExists4(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists4 = Not ws Is Nothing
    On Error GoTo 0
End Function

' Функция для отображения диалогового окна выбора файлов (пример реализации OpenFileDialog3)
Function OpenFileDialog4(initialPath As String) As Collection
    Dim fileDialog As fileDialog
    Dim selectedFiles As New Collection
    Dim selectedFile As Variant
    
    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Выберите файлы для импорта"
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
