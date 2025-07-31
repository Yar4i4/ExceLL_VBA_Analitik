Attribute VB_Name = "Module1"
Sub СборЛистовВКнигу()
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
    Set filePaths = OpenFileDialog3(Vb.Path)
    
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
    
    ' Обработка каждого выбранного файла
    For Each filePath In filePaths
        ' Открываем книгу
        On Error Resume Next
        Set wb = Workbooks.Open(filePath)
        On Error GoTo 0
        
        If Not wb Is Nothing Then
            ' Перебираем все листы в открытой книге
            For Each ws In wb.Sheets
                ' Определяем имя нового листа с учетом возможных конфликтов
                targetSheetName = ws.Name
                suffix = 1
                
                ' Проверка на конфликт имен
                Do While SheetExists(targetSheetName, Vb)
                    ' Если имя слишком длинное, удаляем последние 3 символа
                    If Len(targetSheetName) + Len(CStr(suffix)) > maxSheetNameLength Then
                        targetSheetName = Left(targetSheetName, maxSheetNameLength - Len(CStr(suffix)) - 1)
                    End If
                    
                    ' Добавляем суффикс
                    targetSheetName = targetSheetName & suffix
                    suffix = suffix + 1
                Loop
                
                ' Копируем лист в целевую книгу
                ws.Copy After:=Vb.Sheets(Vb.Sheets.count)
                Vb.Sheets(Vb.Sheets.count).Name = targetSheetName
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
    
    MsgBox "Листы успешно собраны!", vbInformation
End Sub

' Функция для проверки существования листа в указанной книге
Function SheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' Функция для отображения диалогового окна выбора файлов (пример реализации OpenFileDialog3)
Function OpenFileDialog3(initialPath As String) As Collection
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
    
    Set OpenFileDialog3 = selectedFiles
End Function
