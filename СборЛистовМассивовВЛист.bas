Attribute VB_Name = "Module4"
Option Explicit

Sub СборЛистовМассивовВЛист()
    Dim Vb As Workbook
    Set Vb = ThisWorkbook

    ' Отключение оптимизаций для улучшения производительности
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    On Error GoTo ErrorHandler ' Включаем общую обработку ошибок

    ' Получаем данные из листа "Главный"
    Dim mainSheet As Worksheet
    Set mainSheet = Nothing ' Сбрасываем перед установкой
    On Error Resume Next ' Временно игнорируем ошибку, если лист не найден
    Set mainSheet = Vb.Sheets("Главный")
    On Error GoTo ErrorHandler ' Восстанавливаем обработку ошибок
    If mainSheet Is Nothing Then
        MsgBox "Лист 'Главный' не найден в этой книге.", vbCritical
        GoTo Cleanup
    End If

    ' Проверяем корректность указанных столбцов и поисковых слов
    Dim colStart As String, colEnd As String
    Dim searchStart As String, searchEnd As String

    ' Проверка столбца для верхней строки
    colStart = CStr(mainSheet.Range("B13").Value) ' Используем CStr для надежности
    If Not IsValidColumn(colStart) Then
        MsgBox "В ячейке B13 листа 'Главный' указан некорректный идентификатор столбца (допустимы A-ZZZ).", vbExclamation, "Ошибка ввода"
        GoTo Cleanup
    End If

    ' Проверка поискового слова для верхней строки
    searchStart = Trim(CStr(mainSheet.Range("C13").Value))
    If searchStart = "" Then
        MsgBox "В ячейке C13 листа 'Главный' не указано слово для поиска начала массива.", vbExclamation, "Ошибка ввода"
        GoTo Cleanup
    End If

    ' Проверка столбца для нижней строки
    colEnd = CStr(mainSheet.Range("D13").Value)
    If Not IsValidColumn(colEnd) Then
        MsgBox "В ячейке D13 листа 'Главный' указан некорректный идентификатор столбца (допустимы A-ZZZ).", vbExclamation, "Ошибка ввода"
        GoTo Cleanup
    End If

    ' Проверка поискового слова для нижней строки
    searchEnd = Trim(CStr(mainSheet.Range("E13").Value))
    If searchEnd = "" Then
        MsgBox "В ячейке E13 листа 'Главный' не указано слово для поиска конца массива.", vbExclamation, "Ошибка ввода"
        GoTo Cleanup
    End If

    ' Открываем диалоговое окно для выбора файлов
    Dim filePaths As Collection
    Set filePaths = OpenFileDialog5(Vb.Path)

    ' Проверяем, был ли выбран хотя бы один файл
    If filePaths.count = 0 Then
        MsgBox "Файлы не выбраны. Операция прервана.", vbInformation, "Отмена"
        GoTo Cleanup ' Используем GoTo Cleanup для восстановления настроек
    End If

    ' --- Создание листа для коллизий ---
    Dim collisionSheet As Worksheet
    Dim collisionSheetName As String
    Dim suffix As Long ' Используем Long для суффикса
    collisionSheetName = "Коллизии"
    suffix = 1

    ' Находим уникальное имя для листа коллизий
    Do While SheetExists6(collisionSheetName, Vb)
        collisionSheetName = "Коллизии" & suffix
        suffix = suffix + 1
    Loop

    ' Создаем лист для коллизий и форматируем заголовки
    Set collisionSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    collisionSheet.Name = collisionSheetName
    With collisionSheet.Rows(1)
        .Cells(1, 1).Value = "Книга"
        .Cells(1, 2).Value = "Лист"
        .Cells(1, 3).Value = "Столбец поиска"
        .Cells(1, 4).Value = "Искомое слово (не найдено)"
        .RowHeight = 45
        .WrapText = True
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlVAlignCenter
    End With
    collisionSheet.Columns("A:D").ColumnWidth = 20 ' Немного шире

    Dim collisionRow As Long
    collisionRow = 2 ' Начальная строка для записи коллизий
    Dim collisionFound As Boolean
    collisionFound = False

    ' --- Создание целевого листа для данных ---
    Dim targetSheet As Worksheet
    Dim targetSheetName As String
    targetSheetName = "Сбор"
    suffix = 1

    ' Находим следующий доступный номер для листа
    Do While SheetExists6(targetSheetName & suffix, Vb)
        suffix = suffix + 1
    Loop

    ' Создаем новый лист с уникальным именем
    Set targetSheet = Vb.Sheets.Add(After:=Vb.Sheets(Vb.Sheets.count))
    targetSheet.Name = targetSheetName & suffix

    ' --- Обработка каждого выбранного файла ---
    Dim lastRowTarget As Long ' Переименовано для ясности
    lastRowTarget = 1 ' Начальная строка для записи данных на целевой лист

    Dim filePath As Variant ' Variant для итерации по коллекции
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
        Set wb = Nothing ' Сбрасываем перед открытием нового файла
        On Error Resume Next ' Игнорируем ошибку, если файл не удается открыть
        Set wb = Workbooks.Open(filePath, ReadOnly:=True) ' Открываем только для чтения
        On Error GoTo ErrorHandler ' Восстанавливаем стандартную обработку ошибок

        If Not wb Is Nothing Then
            For Each ws In wb.Sheets
                If ws Is Nothing Then GoTo NextSheet ' Пропускаем невалидные объекты листов

                ' Ищем ВСЕ строки с начальными и конечными маркерами
                Set startRows = FindAllRowsBySearchWord(ws, colStart, searchStart)
                Set endRows = FindAllRowsBySearchWord(ws, colEnd, searchEnd)

                blockCopied = False ' Флаг, что хотя бы один блок скопирован с этого листа

                ' Проверяем, найдены ли вообще маркеры на листе
                If startRows.count = 0 Then
                    Call LogCollision(collisionSheet, collisionRow, wb.Name, ws.Name, colStart, searchStart)
                    collisionFound = True
                End If
                 If endRows.count = 0 Then
                    Call LogCollision(collisionSheet, collisionRow, wb.Name, ws.Name, colEnd, searchEnd)
                    collisionFound = True
                 End If

                ' Если найдены и начальные, и конечные маркеры, ищем пары
                If startRows.count > 0 And endRows.count > 0 Then
                    ' Сортируем коллекции строк по возрастанию (Find возвращает в порядке нахождения)
                    ' Для надежности можно было бы реализовать сортировку, но Find обычно находит сверху вниз.

                    For Each startRowVariant In startRows
                        currentStartRow = CLng(startRowVariant)
                        foundEndRow = 0 ' Сбрасываем поиск конечной строки для каждой начальной

                        ' Ищем ПЕРВУЮ подходящую конечную строку ПОСЛЕ текущей начальной
                        For Each endRowVariant In endRows
                            If CLng(endRowVariant) >= currentStartRow Then
                                foundEndRow = CLng(endRowVariant)
                                Exit For ' Нашли ближайшую подходящую конечную строку
                            End If
                        Next endRowVariant

                        ' Если нашли подходящую пару (начало <= конец)
                        If foundEndRow > 0 Then
                            On Error Resume Next ' Локальная обработка ошибок копирования/вставки
                            ws.Rows(currentStartRow & ":" & foundEndRow).Copy
                            If Err.Number = 0 Then
                                targetSheet.Cells(lastRowTarget, 1).PasteSpecial Paste:=xlPasteAll
                                Application.CutCopyMode = False ' Очищаем буфер обмена
                                If Err.Number = 0 Then
                                    lastRowTarget = lastRowTarget + (foundEndRow - currentStartRow + 1)
                                    blockCopied = True ' Отмечаем, что скопировали блок
                                Else
                                     ' Ошибка вставки - можно добавить логирование
                                    Err.Clear
                                End If
                            Else
                                ' Ошибка копирования - можно добавить логирование
                                Err.Clear
                                Application.CutCopyMode = False
                            End If
                             On Error GoTo ErrorHandler ' Восстанавливаем общую обработку
                        End If
                        ' Переходим к следующей начальной строке
                    Next startRowVariant
                End If

                ' Если ни одного блока не было скопировано, а маркеры должны были быть (т.е. коллекции не пусты),
                ' но не нашлось пар start <= end, можно добавить доп. логику коллизий здесь, если нужно.
                ' Текущая логика регистрирует только полное отсутствие маркеров.

NextSheet:
                Set startRows = Nothing ' Освобождаем память
                Set endRows = Nothing
            Next ws

            wb.Close SaveChanges:=False ' Закрываем исходную книгу
        Else
            ' Записываем коллизию, если не удалось открыть файл
             Call LogCollision(collisionSheet, collisionRow, CStr(filePath), "N/A", "Файл", "Не удалось открыть")
             collisionFound = True
            ' MsgBox "Не удалось открыть файл: " & filePath, vbExclamation ' Сообщение в конце
        End If
    Next filePath

    ' --- Завершение ---
    ' Удаляем лист коллизий, если коллизий не было
    If Not collisionFound Then
        Application.DisplayAlerts = False ' Временно отключаем предупреждения об удалении
        collisionSheet.Delete
        Application.DisplayAlerts = True
    Else
        collisionSheet.Columns.AutoFit ' Подгоняем ширину столбцов на листе коллизий
    End If

    ' Удаляем целевой лист, если он пуст (ничего не скопировано)
    On Error Resume Next
    If IsSheetEmpty(targetSheet) Then ' Проверяем, пуст ли лист
        Application.DisplayAlerts = False
        targetSheet.Delete
        Application.DisplayAlerts = True
        MsgBox "Данные для копирования не найдены по заданным критериям.", vbInformation, "Результат"
    Else
        targetSheet.Columns.AutoFit ' Подгоняем ширину столбцов на целевом листе
        targetSheet.Cells(1, 1).Select ' Активируем первую ячейку
        If collisionFound Then
            MsgBox "Сбор данных завершен." & vbCrLf & "Обнаружены некоторые проблемы. Проверьте лист '" & collisionSheetName & "'!", vbExclamation, "Завершено с коллизиями"
            collisionSheet.Activate ' Показываем лист с коллизиями
        Else
            MsgBox "Данные успешно собраны на лист '" & targetSheet.Name & "'!", vbInformation, "Успех"
        End If
    End If
    On Error GoTo ErrorHandler

Cleanup:
    ' Восстанавливаем настройки Excel в любом случае (успех, ошибка, отмена)
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.CutCopyMode = False ' Очистка буфера обмена

    Set Vb = Nothing
    Set mainSheet = Nothing
    Set filePaths = Nothing
    Set collisionSheet = Nothing
    Set targetSheet = Nothing
    Set wb = Nothing
    Set ws = Nothing
    Set startRows = Nothing
    Set endRows = Nothing

    Exit Sub ' Нормальный выход

ErrorHandler:
    ' Обработчик ошибок
    MsgBox "Произошла непредвиденная ошибка:" & vbCrLf & vbCrLf & _
           "Номер ошибки: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description & vbCrLf & _
           "Источник: " & Err.Source, vbCritical, "Ошибка выполнения"
    Resume Cleanup ' Переходим к блоку очистки для восстановления настроек Excel

End Sub

' --- Вспомогательные функции ---

Function SheetExists6(sheetName As String, wb As Workbook) As Boolean
' Проверяет существование листа в указанной книге (без активации ошибок)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    SheetExists6 = Not ws Is Nothing
    On Error GoTo 0 ' Восстановить стандартную обработку ошибок
    Set ws = Nothing
End Function

Function OpenFileDialog5(initialPath As String) As Collection
' Отображает диалоговое окно выбора файлов Excel (множественный выбор)
    Dim fileDialog As fileDialog
    Dim selectedFiles As New Collection
    Dim selectedFile As Variant

    Set fileDialog = Application.fileDialog(msoFileDialogFilePicker)
    With fileDialog
        .Title = "Выберите ОДИН или НЕСКОЛЬКО файлов Excel для обработки"
        If Right(initialPath, 1) <> "\" Then initialPath = initialPath & "\"
        .InitialFileName = initialPath ' Устанавливаем начальную папку
        .Filters.Clear
        .Filters.Add "Файлы Excel", "*.xls; *.xlsx; *.xlsm", 1 ' Добавляем фильтр и делаем его по умолчанию
        .AllowMultiSelect = True ' Разрешаем выбор нескольких файлов

        If .Show = -1 Then ' -1 означает, что пользователь нажал OK
            For Each selectedFile In .SelectedItems
                selectedFiles.Add selectedFile
            Next selectedFile
        End If ' Если пользователь нажал Отмена, коллекция останется пустой
    End With

    Set OpenFileDialog5 = selectedFiles
    Set fileDialog = Nothing ' Освобождаем объект
End Function

Function IsValidColumn(col As String) As Boolean
' Проверяет, является ли строка допустимым идентификатором столбца (A-ZZZ)
' Использует более простую проверку длины и символов
    Dim i As Long
    Dim L As Long
    Dim charCode As Integer

    col = Trim(UCase(col)) ' Убираем пробелы, переводим в верхний регистр
    L = Len(col)

    If L = 0 Or L > 3 Then ' Допустимая длина от 1 до 3
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

    ' Дополнительная проверка для 3 букв (не должно превышать "ZZZ")
    ' На практике ColumnLetterToNumber вернет 0, если столбец невалидный,
    ' поэтому строгая проверка здесь может быть избыточна, но оставим для полноты.
    ' Проверка ColumnLetterToNumber(col) > 0 более надежна.
    If ColumnLetterToNumber(col) > 18278 Then ' 18278 = номер столбца ZZZ
         IsValidColumn = False
    Else
         IsValidColumn = True
    End If

End Function


Function FindAllRowsBySearchWord(ws As Worksheet, colLetter As String, searchWord As String) As Collection
' Находит ВСЕ строки, содержащие ТОЧНОЕ значение searchWord в указанном столбце colLetter
' Возвращает коллекцию номеров строк или пустую коллекцию, если ничего не найдено
    Dim foundRows As New Collection
    Dim searchRange As Range
    Dim foundCell As Range
    Dim firstAddress As String
    Dim colNum As Long

    colNum = ColumnLetterToNumber(colLetter)
    If colNum = 0 Then ' Невалидный столбец
        Set FindAllRowsBySearchWord = foundRows ' Возвращаем пустую коллекцию
        Exit Function
    End If

    On Error Resume Next ' Игнорируем ошибку, если столбец пуст или защищен
    Set searchRange = ws.Columns(colNum)
    If searchRange Is Nothing Then
         Set FindAllRowsBySearchWord = foundRows
         Exit Function
    End If
    On Error GoTo 0 ' Восстанавливаем обработку ошибок

    ' Ищем первое вхождение
    Set foundCell = searchRange.Find(What:=searchWord, _
                                     After:=searchRange.Cells(searchRange.Cells.count), _
                                     LookIn:=xlValues, _
                                     LookAt:=xlWhole, _
                                     SearchOrder:=xlByRows, _
                                     SearchDirection:=xlNext, _
                                     MatchCase:=False) ' Игнорировать регистр

    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Do
            foundRows.Add foundCell.row ' Добавляем номер строки в коллекцию
            ' Ищем следующее вхождение
            Set foundCell = searchRange.FindNext(foundCell)
            ' Проверяем, не вернулись ли к началу и не является ли объект пустым
            If foundCell Is Nothing Then Exit Do
        Loop While foundCell.Address <> firstAddress
    End If

    Set FindAllRowsBySearchWord = foundRows ' Возвращаем коллекцию (может быть пустой)
    ' Очистка
    Set searchRange = Nothing
    Set foundCell = Nothing
End Function

Function ColumnLetterToNumber(colLetter As String) As Long
' Преобразует буквенное обозначение столбца (A-ZZZ) в числовое
    Dim colNum As Long
    Dim i As Long
    Dim L As Long
    Dim charCode As Integer

    colLetter = Trim(UCase(colLetter))
    L = Len(colLetter)

    If L = 0 Or L > 3 Then GoTo InvalidInput

    For i = 1 To L
        charCode = Asc(Mid(colLetter, i, 1))
        If charCode < 65 Or charCode > 90 Then GoTo InvalidInput ' Не буква A-Z
        colNum = colNum * 26 + (charCode - 64)
    Next i

    ' Проверка максимального значения для Excel (16384 = XFD)
    If colNum > 16384 Then GoTo InvalidInput

    ColumnLetterToNumber = colNum
    Exit Function

InvalidInput:
    ColumnLetterToNumber = 0 ' Возвращаем 0 при ошибке
End Function

Sub LogCollision(sheet As Worksheet, ByRef rowNum As Long, bookName As String, sheetName As String, searchCol As String, missingWord As String)
' Записывает информацию о коллизии на лист коллизий
    On Error Resume Next ' Игнорируем ошибки при записи на лист коллизий
    With sheet
        .Cells(rowNum, 1).Value = bookName
        .Cells(rowNum, 2).Value = sheetName
        .Cells(rowNum, 3).Value = searchCol
        .Cells(rowNum, 4).Value = missingWord
    End With
    If Err.Number = 0 Then
        rowNum = rowNum + 1 ' Увеличиваем счетчик строк только при успешной записи
    Else
        Err.Clear ' Очищаем ошибку, если запись не удалась
    End If
    On Error GoTo 0 ' Восстанавливаем обработку ошибок (хотя в этой подпрограмме она не нужна после Resume Next)
End Sub

Function IsSheetEmpty(sht As Worksheet) As Boolean
' Проверяет, содержит ли лист какие-либо данные
    On Error Resume Next
    IsSheetEmpty = sht.UsedRange.Address = sht.Range("A1").Address And IsEmpty(sht.Range("A1").Value)
    If Err.Number <> 0 Then ' Если UsedRange вызывает ошибку (например, очень старые форматы?), считаем пустым
        IsSheetEmpty = True
        Err.Clear
    End If
    On Error GoTo 0
End Function
