Attribute VB_Name = "Module5"
Sub СравниМассивы()
    Dim firstRange As Range
    Dim secondRange As Range
    Dim response As VbMsgBoxResult
    Dim isEqual As Boolean
    Dim i As Long, j As Long
    
    ' Запрос на выделение первого массива
    On Error Resume Next
    Set firstRange = Application.InputBox("Выделите первый массив", "Выбор массива", Type:=8)
    On Error GoTo 0
    
    ' Если пользователь нажал "Отмена", выходим из макроса
    If firstRange Is Nothing Then
        MsgBox "Первый массив не выбран. Макрос завершен.", vbExclamation
        Exit Sub
    End If
    
    ' Запрос на выделение второго массива
    On Error Resume Next
    Set secondRange = Application.InputBox("Выделите второй массив", "Выбор массива", Type:=8)
    On Error GoTo 0
    
    ' Если пользователь нажал "Отмена", выходим из макроса
    If secondRange Is Nothing Then
        MsgBox "Второй массив не выбран. Макрос завершен.", vbExclamation
        Exit Sub
    End If
    
    ' Проверка на равенство размеров массивов
    If firstRange.Rows.count <> secondRange.Rows.count Or firstRange.Columns.count <> secondRange.Columns.count Then
        response = MsgBox("Выделенные массивы не равны. Продолжить?", vbYesNo + vbExclamation, "Предупреждение")
        If response = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Сравнение значений в массивах
    isEqual = True
    For i = 1 To firstRange.Rows.count
        For j = 1 To firstRange.Columns.count
            If firstRange.Cells(i, j).Value <> secondRange.Cells(i, j).Value Then
                firstRange.Cells(i, j).Interior.Color = RGB(222, 180, 180) ' Подкрашиваем ячейку
                isEqual = False
            Else
                firstRange.Cells(i, j).Interior.ColorIndex = xlNone ' Сбрасываем цвет, если значения равны
            End If
        Next j
    Next i
    
    ' Сообщение о результатах сравнения
    If isEqual Then
        MsgBox "Массивы идентичны.", vbInformation
    Else
        MsgBox "Найдены различия в массивах.", vbExclamation
    End If

End Sub

