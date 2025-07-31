Attribute VB_Name = "Module3"
Sub СборЛистовСтрокиВЛист()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook
    
    ' Отключение оптимизаций для улучшения производительности
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    
    ' Получаем строки для копирования из ячейки E8 листа "Главный"
    Dim mainSheet As Worksheet
    Set mainSheet = Vb.Sheets("Главный")
    Dim rowRanges As String
    rowRanges = mainSheet.Range("E8").Value
    
    ' Проверяем, указаны ли строки
    If Trim(rowRanges) = "" Then
        MsgBox "Укажите строку или строки в ячейке E8!", vbExclamation
        Exit Sub
    End If
    
    ' Проверка корректности формата ввода
    If Not ValidateRowRanges(rowRanges) Then
        MsgBox "Ввод строк производится в ячейке E8 в виде массивов через тире (дефис) или как обособленные строки через запятую", vbExclamation
        Exit Sub
    End If
    
    ' Открываем диалоговое окно для выбора файлов
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog5(Vb.Path)
    
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
    
    ' Находим следующий доступный номер для листа
    Do While SheetExists5(targetSheetName & suffix, Vb)
        suffix = suffix + 1
    Loop
    
    ' Создаем новый лист с уникальным именем
    Dim targetSheet As Worksheet
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName & suffix
    
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
                ' Получаем массив строк для копирования
                Dim rowsToCopy As Variant
                rowsToCopy = GetRowsToCopy(rowRanges)
                
                ' Копируем указанные строки
                Dim row As Variant
                For Each row In rowsToCopy
                    If row <= ws.UsedRange.Rows.count Then
                        ws.Rows(row).Copy
                        targetSheet.Cells(lastRow, 1).PasteSpecial Paste:=xlPasteAll
                        lastRow = lastRow + 1
                    End If
                Next row
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
Function SheetExists5(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists5 = Not ws Is Nothing
    On Error GoTo 0
End Function

' Функция для отображения диалогового окна выбора файлов (пример реализации OpenFileDialog3)
Function OpenFileDialog5(initialPath As String) As Collection
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
    
    Set OpenFileDialog5 = selectedFiles
End Function

' Функция для получения массива строк из строки диапазонов
Function GetRowsToCopy(rowRanges As String) As Variant
    Dim result As Collection
    Set result = New Collection
    
    ' Разделяем строку по запятой
    Dim ranges As Variant
    ranges = Split(rowRanges, ",")
    
    Dim r As Variant
    For Each r In ranges
        ' Убираем пробелы
        r = Trim(r)
        
        ' Проверяем, содержит ли диапазон дефис/тире
        If InStr(r, "-") > 0 Or InStr(r, "–") > 0 Then
            ' Разделяем диапазон на начальную и конечную строки
            Dim startRow As Long, endRow As Long
            Dim parts As Variant
            parts = Split(r, "-")
            If UBound(parts) = 1 Then
                startRow = CLng(parts(0))
                endRow = CLng(parts(1))
                
                ' Добавляем все строки в диапазоне
                Dim i As Long
                For i = startRow To endRow
                    result.Add i
                Next i
            End If
        Else
            ' Добавляем одиночную строку
            result.Add CLng(r)
        End If
    Next r
    
    ' Преобразуем коллекцию в массив
    Dim output() As Long
    ReDim output(result.count - 1)
    Dim j As Long
    For j = 1 To result.count
        output(j - 1) = result(j)
    Next j
    
    GetRowsToCopy = output
End Function

' Функция для валидации строк
Function ValidateRowRanges(rowRanges As String) As Boolean
    Dim ranges As Variant
    ranges = Split(rowRanges, ",")
    
    Dim r As Variant
    For Each r In ranges
        r = Trim(r)
        
        ' Если диапазон содержит дефис/тире
        If InStr(r, "-") > 0 Or InStr(r, "–") > 0 Then
            Dim parts As Variant
            parts = Split(r, "-")
            If UBound(parts) <> 1 Then
                ValidateRowRanges = False
                Exit Function
            End If
            
            ' Проверяем, что обе части числа
            If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
                ValidateRowRanges = False
                Exit Function
            End If
        ElseIf Not IsNumeric(r) Then
            ' Если это не число
            ValidateRowRanges = False
            Exit Function
        End If
    Next r
    
    ValidateRowRanges = True
End Function
