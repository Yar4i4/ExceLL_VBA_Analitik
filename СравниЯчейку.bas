Attribute VB_Name = "Module6"
Sub СравниЯчейку()
    Dim range1 As Range, range2 As Range
    Dim cell1 As Range, cell2 As Range
    Dim data1 As String, data2 As String
    Dim words1() As String, words2() As String
    Dim i As Long, j As Long
    Dim hasDifferences As Boolean
    
    ' Запрос выбора первого массива ячеек
    On Error Resume Next
    Set range1 = Application.InputBox("Выберите первый массив ячеек для сравнения", Type:=8)
    On Error GoTo 0
    If range1 Is Nothing Then Exit Sub ' Если пользователь отменил выбор
    
    ' Запрос выбора второго массива ячеек
    On Error Resume Next
    Set range2 = Application.InputBox("Выберите второй массив ячеек для сравнения", Type:=8)
    On Error GoTo 0
    If range2 Is Nothing Then Exit Sub ' Если пользователь отменил выбор
    
    ' Проверяем, что массивы ячеек имеют одинаковый размер
    If range1.Cells.count <> range2.Cells.count Then
        MsgBox "Массивы ячеек имеют разный размер. Сравнение невозможно.", vbExclamation
        Exit Sub
    End If
    
    ' Инициализируем флаг
    hasDifferences = False
    
    ' Сравниваем каждую ячейку в массивах
    For i = 1 To range1.Cells.count
        Set cell1 = range1.Cells(i)
        Set cell2 = range2.Cells(i)
        
        ' Получаем данные из ячеек (с учётом объединённых ячеек)
        data1 = cell1.MergeArea.Cells(1, 1).Value
        data2 = cell2.MergeArea.Cells(1, 1).Value
        
        ' Сбрасываем форматирование шрифта
        cell1.MergeArea.Cells(1, 1).Font.ColorIndex = xlAutomatic
        cell2.MergeArea.Cells(1, 1).Font.ColorIndex = xlAutomatic
        
        ' Разделяем тексты на слова (по пробелам)
        words1 = Split(data1, " ")
        words2 = Split(data2, " ")
        
        ' Сравниваем слова
        For j = LBound(words1) To UBound(words1)
            Dim found As Boolean
            found = False
            
            ' Ищем похожее слово во второй ячейке
            Dim k As Long
            For k = LBound(words2) To UBound(words2)
                If AreWordsSimilar(words1(j), words2(k)) Then
                    found = True
                    ' Сравниваем символы в похожих словах
                    CompareAndHighlightWords cell1, words1(j), cell2, words2(k), data1, data2
                    Exit For
                End If
            Next k
            
            ' Если слово не найдено, выделяем его целиком
            If Not found Then
                HighlightText cell1, words1(j), data1
                hasDifferences = True
            End If
        Next j
    Next i
    
    ' Выводим сообщение в зависимости от наличия расхождений
    If hasDifferences Then
        MsgBox "Расхождения найдены и подсвечены красным.", vbInformation
    Else
        MsgBox "Расхождения отсутствуют.", vbInformation
    End If
End Sub

' Функция для сравнения слов (простейшая реализация)
Function AreWordsSimilar(word1 As String, word2 As String) As Boolean
    ' Если слова совпадают на 70% или более, считаем их похожими
    Dim similarityThreshold As Double
    similarityThreshold = 0.7
    
    ' Вычисляем процент совпадения
    Dim matchCount As Long
    matchCount = 0
    Dim i As Long
    For i = 1 To Len(word1)
        If i <= Len(word2) And Mid(word1, i, 1) = Mid(word2, i, 1) Then
            matchCount = matchCount + 1
        End If
    Next i
    
    ' Возвращаем результат
    AreWordsSimilar = (matchCount / Len(word1)) >= similarityThreshold
End Function

' Функция для сравнения и выделения символов в похожих словах
Sub CompareAndHighlightWords(cell1 As Range, word1 As String, cell2 As Range, word2 As String, fullText1 As String, fullText2 As String)
    Dim i As Long
    For i = 1 To Len(word1)
        If i > Len(word2) Or Mid(word1, i, 1) <> Mid(word2, i, 1) Then
            ' Если символы не совпадают, выделяем лишний символ в первом слове
            HighlightText cell1, Mid(word1, i, 1), fullText1
        End If
    Next i
End Sub

' Вспомогательная функция для выделения текста
Sub HighlightText(cell As Range, textToHighlight As String, fullText As String)
    Dim startPos As Long
    startPos = InStr(fullText, textToHighlight)
    If startPos > 0 Then
        With cell.MergeArea.Cells(1, 1).Characters(startPos, Len(textToHighlight)).Font
            .Color = RGB(255, 0, 0) ' Красный цвет
        End With
    End If
End Sub
