Attribute VB_Name = "Module7"
Sub ВПР()

    ' --- Объявление переменных ---
    Dim lookup_ranges As Range
    Dim search_ranges As Range
    Dim source_ranges As Range
    Dim destination_ranges As Range

    Dim lookup_area As Range
    Dim search_area As Range
    Dim source_cell As Range
    Dim dest_cell As Range
    Dim dest_row_range As Range
    Dim source_row_range As Range ' Не используется для копирования, но нужна для GetRowRangeByOverallIndex
    Dim first_match_source_row_range As Range

    Dim s_row As Long
    Dim s_col As Long
    Dim k As Long
    Dim m As Long
    Dim first_match_overall_search_row As Long

    Dim match_found_for_this_lookup As Boolean
    Dim multiple_matches_found As Boolean
    Dim current_row_matches As Boolean
    Dim dest_is_empty As Boolean
    Dim source_is_not_empty As Boolean

    Dim lookup_val As String
    Dim search_val As String

    ' --- ИСПРАВЛЕНИЕ: Объявляем цвет как переменную Dim ---
    Dim HIGHLIGHT_COLOR As Long
    ' --- Присваиваем значение переменной в начале кода ---
    HIGHLIGHT_COLOR = RGB(222, 180, 180)

    ' --- Получение диапазонов от пользователя ---
    On Error Resume Next
    Set lookup_ranges = Application.InputBox("1. Выберите КЛЮЧЕВЫЕ ячейки/столбцы для поиска (можно несколько через Ctrl)", "Выбор ключей", Type:=8)
    If lookup_ranges Is Nothing Then
        MsgBox "Операция отменена пользователем (Выбор ключей).", vbInformation
        Exit Sub
    End If

    Set search_ranges = Application.InputBox("2. Выберите ячейки/столбцы, ГДЕ искать ключи (можно несколько через Ctrl)", "Выбор области поиска", Type:=8)
    If search_ranges Is Nothing Then
        MsgBox "Операция отменена пользователем (Выбор области поиска).", vbInformation
        Exit Sub
    End If

    Set source_ranges = Application.InputBox("3. Выберите ячейки/столбцы, ОТКУДА копировать данные при совпадении (можно несколько через Ctrl)", "Выбор источника данных", Type:=8)
    If source_ranges Is Nothing Then
        MsgBox "Операция отменена пользователем (Выбор источника данных).", vbInformation
        Exit Sub
    End If

    Set destination_ranges = Application.InputBox("4. Выберите ячейки/столбцы, КУДА вставлять скопированные данные (можно несколько через Ctrl)", "Выбор назначения", Type:=8)
    If destination_ranges Is Nothing Then
        MsgBox "Операция отменена пользователем (Выбор назначения).", vbInformation
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' --- Настройки производительности ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- Валидация (Проверка) выбранных диапазонов ---
    Dim num_lookup_cols As Long
    Dim num_search_cols As Long
    Dim num_source_cols As Long
    Dim num_dest_cols As Long
    Dim total_lookup_rows As Long
    Dim total_search_rows As Long
    Dim total_dest_rows As Long
    Dim total_source_rows As Long

    num_lookup_cols = GetTotalColumns(lookup_ranges)
    num_search_cols = GetTotalColumns(search_ranges)
    num_source_cols = GetTotalColumns(source_ranges)
    num_dest_cols = GetTotalColumns(destination_ranges)

    total_lookup_rows = GetTotalRows(lookup_ranges)
    total_search_rows = GetTotalRows(search_ranges)
    total_dest_rows = GetTotalRows(destination_ranges)
    total_source_rows = GetTotalRows(source_ranges)

    ' Проверка 1: Несовпадение кол-ва столбцов (Ключи vs Поиск)
    If num_lookup_cols <> num_search_cols Then
        MsgBox "Ошибка: Количество столбцов в диапазоне КЛЮЧЕЙ (" & num_lookup_cols & ") не совпадает с количеством столбцов в диапазоне ПОИСКА (" & num_search_cols & "). Операция прервана.", vbCritical
        GoTo CleanUpAndExit
    End If

    ' Проверка 2: Несовпадение кол-ва столбцов (Источник vs Назначение)
    If num_source_cols <> num_dest_cols Then
        MsgBox "Ошибка: Количество столбцов в диапазоне ИСТОЧНИКА (" & num_source_cols & ") не совпадает с количеством столбцов в диапазоне НАЗНАЧЕНИЯ (" & num_dest_cols & "). Операция прервана.", vbCritical
        GoTo CleanUpAndExit
    End If

    ' Проверка 3: Несовпадение общего кол-ва строк (Ключи vs Назначение)
    If total_lookup_rows <> total_dest_rows Then
        MsgBox "Ошибка: Общее количество строк в диапазоне КЛЮЧЕЙ (" & total_lookup_rows & ") не совпадает с общим количеством строк в диапазоне НАЗНАЧЕНИЯ (" & total_dest_rows & "). Операция прервана.", vbCritical
        GoTo CleanUpAndExit
    End If

    ' Проверка 4: Несовпадение общего кол-ва строк (Поиск vs Источник)
     If total_search_rows <> total_source_rows Then
        MsgBox "Ошибка: Общее количество строк в диапазоне ПОИСКА (" & total_search_rows & ") не совпадает с общим количеством строк в диапазоне ИСТОЧНИКА (" & total_source_rows & "). Операция прервана.", vbCritical
        GoTo CleanUpAndExit
    End If

    ' --- Основной цикл обработки ---
    k = 0
    For Each lookup_area In lookup_ranges.Areas
        Dim r_lookup As Long
        For r_lookup = 1 To lookup_area.Rows.count
            k = k + 1

            match_found_for_this_lookup = False
            multiple_matches_found = False
            first_match_overall_search_row = -1
            Set first_match_source_row_range = Nothing

            m = 0
            For Each search_area In search_ranges.Areas
                For s_row = 1 To search_area.Rows.count
                    m = m + 1
                    current_row_matches = True

                    For s_col = 1 To num_lookup_cols
                        Dim lookup_cell As Range
                        Dim search_cell As Range

                        On Error Resume Next
                        Set lookup_cell = lookup_area.Cells(r_lookup, s_col)
                        Set search_cell = search_area.Cells(s_row, s_col)
                        On Error GoTo ErrorHandler

                        If lookup_cell Is Nothing Or search_cell Is Nothing Then
                            current_row_matches = False
                            Debug.Print "Предупреждение: Не удалось получить lookup_cell или search_cell для k=" & k & ", m=" & m & ", s_col=" & s_col
                            Exit For
                        End If

                        lookup_val = GetComparableString(lookup_cell)
                        search_val = GetComparableString(search_cell)

                        If lookup_val <> search_val Then
                            current_row_matches = False
                            Exit For
                        End If
                    Next s_col

                    If current_row_matches Then
                        If Not match_found_for_this_lookup Then
                            match_found_for_this_lookup = True
                            first_match_overall_search_row = m
                            Set first_match_source_row_range = GetRowRangeByOverallIndex(source_ranges, first_match_overall_search_row, num_source_cols)
                        Else
                            multiple_matches_found = True
                        End If
                    End If
                Next s_row
            Next search_area

            If match_found_for_this_lookup Then
                Set dest_row_range = GetRowRangeByOverallIndex(destination_ranges, k, num_dest_cols)

                If Not dest_row_range Is Nothing And Not first_match_source_row_range Is Nothing Then
                    For s_col = 1 To num_source_cols
                        On Error Resume Next
                        Set dest_cell = dest_row_range.Cells(1, s_col)
                        Set source_cell = first_match_source_row_range.Cells(1, s_col)
                        On Error GoTo ErrorHandler

                         If dest_cell Is Nothing Or source_cell Is Nothing Then
                             Debug.Print "Предупреждение: Не удалось получить dest_cell или source_cell при копировании для k=" & k & ", s_col=" & s_col
                         Else
                            dest_is_empty = (Len(Trim(CStr(dest_cell.Value2))) = 0)
                            source_is_not_empty = (Len(Trim(CStr(GetComparableString(source_cell)))) > 0)

                            If dest_is_empty And source_is_not_empty Then
                                dest_cell.Value = source_cell.Value
                            End If
                        End If
                    Next s_col

                    If multiple_matches_found Then
                        dest_row_range.Interior.Color = HIGHLIGHT_COLOR ' Используем переменную
                    Else
                         ' dest_row_range.Interior.ColorIndex = xlNone
                    End If
                Else
                    Debug.Print "Предупреждение: Не удалось получить dest_row_range (k=" & k & ") или first_match_source_row_range (m=" & first_match_overall_search_row & ")"
                End If
            End If
        Next r_lookup
    Next lookup_area

    MsgBox "Операция ВПР() успешно завершена.", vbInformation

CleanUpAndExit:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Set lookup_ranges = Nothing
    Set search_ranges = Nothing
    Set source_ranges = Nothing
    Set destination_ranges = Nothing
    Set lookup_area = Nothing
    Set search_area = Nothing
    Set source_cell = Nothing
    Set dest_cell = Nothing
    Set dest_row_range = Nothing
    Set source_row_range = Nothing
    Set first_match_source_row_range = Nothing
    On Error GoTo 0

    Exit Sub

ErrorHandler:
    MsgBox "Произошла ошибка выполнения:" & vbCrLf & vbCrLf & _
           "Номер ошибки: " & Err.Number & vbCrLf & _
           "Описание: " & Err.Description, vbCritical, "Ошибка в макросе ВПР()"
    Debug.Print "Ошибка #" & Err.Number & ": " & Err.Description & " в процедуре ВПР()"
    Resume CleanUpAndExit

End Sub



Private Function GetTotalRows(rng As Range) As Long
    Dim area As Range
    Dim totalRows As Long
    If rng Is Nothing Then
        GetTotalRows = 0
        Exit Function
    End If
    totalRows = 0
    For Each area In rng.Areas
        totalRows = totalRows + area.Rows.count
    Next area
    GetTotalRows = totalRows
End Function

Private Function GetTotalColumns(rng As Range) As Long
    If rng Is Nothing Then
        GetTotalColumns = 0
        Exit Function
    End If
    GetTotalColumns = rng.Columns.count
End Function

Private Function GetRowRangeByOverallIndex(baseRange As Range, overallRowIndex_1Based As Long, numCols As Long) As Range
    Dim area As Range
    Dim cumulativeRows As Long
    Dim relativeRowIndex As Long

    Set GetRowRangeByOverallIndex = Nothing
    If baseRange Is Nothing Or overallRowIndex_1Based <= 0 Then Exit Function

    cumulativeRows = 0

    For Each area In baseRange.Areas
        If overallRowIndex_1Based <= cumulativeRows + area.Rows.count Then
            relativeRowIndex = overallRowIndex_1Based - cumulativeRows
            On Error Resume Next
            Set GetRowRangeByOverallIndex = area.Cells(relativeRowIndex, 1).Resize(1, area.Columns.count)
            On Error GoTo 0
            Exit Function
        End If
        cumulativeRows = cumulativeRows + area.Rows.count
    Next area
End Function

Private Function GetComparableString(cell As Range) As String
    If cell Is Nothing Then
        GetComparableString = ""
    ElseIf IsError(cell.Value) Then
        GetComparableString = ""
    Else
        GetComparableString = LCase(Trim(CStr(cell.Value2)))
    End If
End Function


