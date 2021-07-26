Attribute VB_Name = "Functions"
Option Explicit
Function FindEmptyRowInColumn(cell As Range) As Integer
    Dim i As Integer
    
    i = 0
    Do While Sheets(cell.Worksheet.Name).Cells(cell.Row + i, cell.Column).Value <> ""
        i = i + 1
    Loop
    
    FindEmptyRowInColumn = cell.Row + i
End Function
Function FindDoubleEmptyRowInColumn(cell As Range) As Integer
    Dim i As Integer
    
    i = 0
    Do While Sheets(cell.Worksheet.Name).Cells(cell.Row + i, cell.Column).Value <> "" _
    Or Sheets(cell.Worksheet.Name).Cells(cell.Row + i + 1, cell.Column).Value <> ""
        i = i + 1
    Loop
    
    FindDoubleEmptyRowInColumn = cell.Row + i
End Function
Function CountMeasures(cell As Range, ByRef LastRow As Integer) As Integer
    Dim i%, s%
    
    i = 0
    s = 0
    Do While Sheets(cell.Worksheet.Name).Cells(cell.Row + i, cell.Column).Value <> ""
        i = i + 1
        If Sheets(cell.Worksheet.Name).Cells(cell.Row + i, cell.Column + 1).Value <> "" Then
            s = s + 1
        End If
    Loop
    
    CountMeasures = s
    LastRow = i + cell.Row
    
End Function
Function SetApplianceColumn(SheetName As String, Optional StartCell As Range, _
Optional FomratConditionsOnly As Boolean = False, Optional FormatConditions_1 As String = "Применимо", Optional FormatConditions_2 As String = "Неприменимо", _
Optional DropToDefault As Integer = 0)
    Dim i, LastRow As Integer
    Dim temp As Range
    
    If Not StartCell Is Nothing Then
        Set temp = StartCell
    Else
        Set temp = Sheets(SheetName).Range("2:2").Find(What:="ПРИМЕНИМОСТЬ", Lookat:=xlPart)
    End If
    
    If Not temp Is Nothing Then
        If temp.Column > 1 Then
            LastRow = FindEmptyRowInColumn(Sheets(SheetName).Cells(temp.Row + 2, temp.Column - 1))
        Else
            LastRow = FindEmptyRowInColumn(Sheets(SheetName).Cells(temp.Row + 2, temp.Column + 1))
        End If
        For i = 1 To LastRow - 4
            With Sheets(SheetName).Cells(temp.Row + 1 + i, temp.Column)
                If Not FomratConditionsOnly Then
                    With .Validation
                        .Delete
                        .Add Type:=xlValidateList, Formula1:=CStr(FormatConditions_1 + "," + FormatConditions_2)
                        .ErrorTitle = "Ошибка"
                        .InputMessage = CStr(FormatConditions_1 + "/" + FormatConditions_2)
                        .ErrorMessage = "Неверный ввод"
                    End With
                End If
                With .FormatConditions
                    .Delete
                    .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:=FormatConditions_1
                    .Item(1).Interior.Color = RGB(150, 255, 150)
                End With
                With .FormatConditions
                    .Add Type:=xlCellValue, Operator:=xlEqual, Formula1:=FormatConditions_2
                    .Item(2).Interior.Color = RGB(255, 150, 150)
                End With
                If DropToDefault <> 0 Then
                    If DropToDefault = 1 Then
                        .Value = FormatConditions_1
                    ElseIf DropToDefault = 2 Then
                        .Value = FormatConditions_2
                    End If
                End If
            End With
        Next i
    Else
        MsgBox "Не найдено слово ПРИМЕНИМОСТЬ на листе " + SheetName, , "ERROR: Find"
    End If

End Function
Function FindIDorName(Name As String, MyDitcionary As Variant, Optional FromIDtoName As Boolean = False) As String
    
    Dim i As Variant
    
    'По имени ищём ID
    If FromIDtoName = False Then
        For Each i In MyDitcionary.keys
            If Name = MyDitcionary.Item(i) Then
                FindIDorName = i
            End If
        Next i
    'По ID ищем имя
    Else
        For Each i In MyDitcionary.keys
            If Name = i Then
                FindIDorName = MyDitcionary.Item(i)
            End If
        Next i
    End If
End Function
Function WriteDictionary(SheetName As String, ByRef Result As Variant, _
Optional DictionaryName As String, Optional Questionary As Boolean = False)
    Dim i, j As Integer
    Dim temp As Range
    
    Set temp = Sheets(SheetName).Range("2:2").Find(What:="ID", LookIn:=xlFormulas, Lookat:=xlPart)
    Set Result = CreateObject("Scripting.Dictionary")
    Result.RemoveAll
        
    i = 2
    Do While Sheets(SheetName).Cells(temp.Row + i, temp.Column) <> "" Or _
    Sheets(SheetName).Cells(temp.Row + i + 1, temp.Column) <> ""
        'Если это опросник, то выписываем только нужные объекты
        If Questionary = True Then
            If Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1) = "Применимо" Then
                If Not Result.Exists(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value) Then
                    Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, Sheets(SheetName).Cells(temp.Row + i, temp.Column - 1).Value
                Else
                    MsgBox ("В словаре " + SheetName + " дублируется ID " + CStr(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value))
                End If
            End If
        'Если это просто словарь, то все
        ElseIf Sheets(SheetName).Cells(temp.Row + i, temp.Column) <> "" Then
            If Not Result.Exists(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value) Then
                Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, Sheets(SheetName).Cells(temp.Row + i, temp.Column - 1).Value
            Else
                MsgBox ("В словаре " + SheetName + " дублируется ID " + CStr(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value))
            End If
        End If
        i = i + 1
    Loop
    
    If Result.Count = 0 Then
        MsgBox "Cправочник " + DictionaryName + " пустой!", , "ERROR: " + SheetName
    End If
End Function
Function WriteBookOfReferenceFromDefault(SheetName As String, ByRef Result As Variant, _
Optional DictionaryName As String, Optional Reverse As Boolean = False)
    Dim i, j As Integer
    Dim IDs() As String
    Dim ReverseIDs() As String
    Dim ID_key As Variant
    Dim temp As Range
    
    Set temp = Sheets(SheetName).Range("2:2").Find(What:="ID", LookIn:=xlFormulas, Lookat:=xlPart)
    Set Result = CreateObject("Scripting.Dictionary")
    Result.RemoveAll
        
    i = 2
    'Если это прямой справочник
    If Reverse = False Then
'        Индексирует по столбцу 2
        Do While Sheets(SheetName).Cells(temp.Row + i, temp.Column) <> "" Or _
        Sheets(SheetName).Cells(temp.Row + i + 1, temp.Column) <> ""
            If Not Result.Exists(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value) Then
                IDs = Split(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, "|")
                Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, IDs
            Else
                MsgBox ("В справочнике " + SheetName + " дублируется ID " + CStr(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value))
            End If
            i = i + 1
        Loop
    'Если это обратный справочник
    Else
'        Индексирует по столбцу 3
        Do While Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1) <> "" Or _
        Sheets(SheetName).Cells(temp.Row + i + 1, temp.Column + 1) <> ""
'            Строка разбивается на ID
            IDs = Split(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, "|")
'            Проходит по каждму ID
            For Each ID_key In IDs
                If Not Result.Exists(ID_key) Then
'                    Если ID нет, то задается массив с одним значением, в котром лежит ID соответствующего элемента
                    ReDim ReverseIDs(0)
                    ReverseIDs(0) = Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value
                    Result.Add ID_key, ReverseIDs
                Else
'                    Если ID уже есть в словаре, то создаётся массив, который содержит все значения массива соответствующего ключа
                    Erase ReverseIDs
                    ReverseIDs() = Result.Item(ID_key)
'                    Затем этот массив расширается и дописывается новым элементом
                    ReDim Preserve ReverseIDs(UBound(ReverseIDs) + 1)
                    ReverseIDs(UBound(ReverseIDs)) = Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value
'                    Т.к. переобозначить Item для ключа нельзя, то ключ удаляется и записывается заново с дополненным массивом
                    Result.Remove ID_key
                    Result.Add ID_key, ReverseIDs
                End If
            Next ID_key
            i = i + 1
        Loop
    End If
    
    If Result.Count = 0 Then
        MsgBox "Cправочник " + DictionaryName + " пустой!", , "ERROR: " + SheetName
    End If
End Function
Function WriteBookOfReferenceFromAuto(SheetName As String, ByRef Result As Variant, _
Optional DictionaryName As String, Optional Reverse As Boolean = False, Optional Questionary As Boolean = False, _
Optional StartCell As Range)

    Dim i, j As Integer
    Dim ReverseIDs() As String
    Dim temp As Range
    
    If Not StartCell Is Nothing Then
        Set temp = StartCell
    Else
        Set temp = Sheets(SheetName).Range("2:2").Find(What:="ID", LookIn:=xlFormulas, Lookat:=xlPart)
    End If
    
    Set Result = CreateObject("Scripting.Dictionary")
    Result.RemoveAll
    
    i = 2
    'Если это прямой справочник
    If Reverse = False Then
'        Индексирует по столбцу 2
        Do While Sheets(SheetName).Cells(temp.Row + i, temp.Column) <> "" Or _
        Sheets(SheetName).Cells(temp.Row + i + 1, temp.Column) <> ""
            If Not Questionary Or Sheets(SheetName).Cells(temp.Row + i, temp.Column + 2).Value = "Применимо" Then
                If Not Result.Exists(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value) Then
                    ReDim ReverseIDs(0)
                    ReverseIDs(0) = Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value
                    Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, ReverseIDs
                ElseIf Not CheckReferences(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, Result) Then
    '               Если ID уже есть в словаре, то создаётся массив, который содержит все значения массива соответствующего ключа
                    Erase ReverseIDs
                    ReverseIDs() = Result.Item(Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value)
    '               Затем этот массив расширается и дописывается новым элементом
                    ReDim Preserve ReverseIDs(UBound(ReverseIDs) + 1)
                    ReverseIDs(UBound(ReverseIDs)) = Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value
    '               Т.к. переобозначить Item для ключа нельзя, то ключ удаляется и записывается заново с дополненным массивом
                    Result.Remove Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value
                    Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, ReverseIDs
                End If
            End If
            i = i + 1
        Loop
    'Если это обратный справочник
    Else
'        Индексирует по столбцу 3
        Do While Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1) <> "" Or _
        Sheets(SheetName).Cells(temp.Row + i + 1, temp.Column + 1) <> ""
            If Not Questionary Or Sheets(SheetName).Cells(temp.Row + i, temp.Column + 2).Value = "Применимо" Then
                If Not Result.Exists(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value) Then
                    ReDim ReverseIDs(0)
                    ReverseIDs(0) = Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value
                    Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, ReverseIDs
                ElseIf Not CheckReferences(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value, Result) Then
    '               Если ID уже есть в словаре, то создаётся массив, который содержит все значения массива соответствующего ключа
                    Erase ReverseIDs
                    ReverseIDs() = Result.Item(Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value)
    '               Затем этот массив расширается и дописывается новым элементом
                    ReDim Preserve ReverseIDs(UBound(ReverseIDs) + 1)
                    ReverseIDs(UBound(ReverseIDs)) = Sheets(SheetName).Cells(temp.Row + i, temp.Column).Value
    '               Т.к. переобозначить Item для ключа нельзя, то ключ удаляется и записывается заново с дополненным массивом
                    Result.Remove Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value
                    Result.Add Sheets(SheetName).Cells(temp.Row + i, temp.Column + 1).Value, ReverseIDs
                End If
            End If
            i = i + 1
        Loop
    End If
    
    If Result.Count = 0 Then
        MsgBox "Cправочник " + DictionaryName + " пустой!", , "ERROR: " + SheetName
    End If
End Function
Function CheckReferences(CheckingID As Variant, KeyID As Variant, InputDict As Variant)
    Dim ID As Variant
    CheckReferences = False
    If InputDict.Exists(KeyID) Then
        For Each ID In InputDict.Item(KeyID)
            If CheckingID = ID Then
                CheckReferences = True
            End If
        Next ID
    End If
End Function
Function CheckCategory(ID1 As Variant, InputDict1 As Variant, ID2 As Variant, InputDict2 As Variant)
    Dim Index1 As Variant
    Dim Index2 As Variant
    CheckCategory = False
    If InputDict1.Exists(ID1) And InputDict2.Exists(ID2) Then
        For Each Index1 In InputDict1.Item(ID1)
            For Each Index2 In InputDict2.Item(ID2)
                If CStr(Index1) = CStr(Index2) Then
                    CheckCategory = True
                    Exit For
                End If
            Next Index2
        Next Index1
    End If
End Function
Function DisplayDictionary(ByRef ShowingDictionary As Variant, Optional NameOfDictionary As String)
    Dim i As Integer
    Dim s() As String
    Dim Key As Variant
    Dim str As Variant
    Dim ItemStr As String
    
    Debug.Print "Start: DisplayDictionary=========="
    Debug.Print "Name: " + NameOfDictionary + "---------"

    For Each Key In ShowingDictionary.keys
        If TypeName(ShowingDictionary.Item(Key)) <> "String" Then
            Erase s
            s = ShowingDictionary.Item(Key)
            ItemStr = ""
            For Each str In s
                ItemStr = ItemStr + "|" + CStr(str)
            Next str
            Debug.Print CStr(Key) + "=" + ItemStr
        Else
            Debug.Print CStr(Key) + "=" + CStr(ShowingDictionary.Item(Key))
        End If
    Next Key
    Debug.Print "==========End: DisplayDictionary=========="
End Function
Function CategoryOutput(Key As Variant, ByRef InputDictionary As Variant) As String
    Dim i As Integer
    Dim Buffer() As String
    Dim Element As Variant
    Dim Output As String
    
    If TypeName(InputDictionary.Item(Key)) <> "String" Then
        Erase Buffer
        Buffer = InputDictionary.Item(Key)
        Output = ""
        For Each Element In Buffer
            Output = Output + CStr(Element) + "|"
        Next Element
        CategoryOutput = Left(Output, Len(Output) - 1)
    Else
        CategoryOutput = CStr(InputDictionary.Item(Key))
    End If
    
End Function
Function AddItemToKey(Item As Variant, Key As Variant, ByRef InputDictionary As Variant)
    Dim i As Integer
    Dim Buffer() As String
    
    If Not CheckReferences(Item, Key, InputDictionary) Then
        Buffer() = InputDictionary.Item(Key)
        'Затем этот массив расширается и дописывается новым элементом
        ReDim Preserve Buffer(UBound(Buffer) + 1)
        Buffer(UBound(Buffer)) = Item
        'Т.к. переобозначить Item для ключа нельзя, то ключ удаляется и записывается заново с дополненным массивом
        InputDictionary.Remove Key
        InputDictionary.Add Key, Buffer
    End If

End Function
Function MakeStep(CurrentSheet As String, TargetSheet As String)
    ActiveWorkbook.Worksheets(TargetSheet).Visible = xlSheetVisible
    ActiveWorkbook.Worksheets(TargetSheet).Activate
    If ChoosingListForm.HideSheet.Value <> True Then
        ActiveWorkbook.Worksheets(CurrentSheet).Visible = xlSheetHidden
    End If
End Function
Function DisplayDictionaryOnList(DictionaryForDisplay As Variant, ID_Dictionary1 As Variant, ID_Dictionary2 As Variant)
    Dim i As Integer
    Dim Numeration As Integer
    Dim ID1 As Variant
    Dim ID2 As Variant

    i = FindEmptyRowInColumn(ActiveSheet.Cells(3, 9))
    ActiveSheet.Range("F4:K" + CStr(i)).ClearContents
    ActiveSheet.Range("F4:K" + CStr(i)).ClearFormats
           
    'Индекс для последствий
    Numeration = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого последствия
    For Each ID1 In ID_Dictionary1.keys
        For Each ID2 In ID_Dictionary2.keys
            Numeration = Numeration + 1
            ActiveSheet.Cells(i, 6).Value = Numeration
            ActiveSheet.Cells(i, 7).Value = ID_Dictionary1.Item(ID1)
            ActiveSheet.Cells(i, 8).Value = ID_Dictionary2.Item(ID2)
            ActiveSheet.Cells(i, 9).Value = ID1
            ActiveSheet.Cells(i, 10).Value = ID2
            If Functions.CheckReferences(ID2, ID1, DictionaryForDisplay) Then
                ActiveSheet.Cells(i, 11).Value = "Применимо"
            Else
                ActiveSheet.Cells(i, 11).Value = "Неприменимо"
            End If
            i = i + 1
        Next ID2
    Next ID1
    Call SetApplianceColumn(ActiveSheet.Name)
    
    i = FindEmptyRowInColumn(ActiveSheet.Cells(3, 9)) - 1
    'Настраивается внешний вид
    With ActiveSheet.Range("F4:K" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
End Function
Function RewriteDictionary(InputDictionary As Variant)
    Dim i As Integer
    Dim Numeration As Integer
    Dim Key As Variant
    Dim Element As Variant
    Dim ID_Item$, Output$
    Dim Buffer() As String
    
    i = FindEmptyRowInColumn(ActiveSheet.Cells(3, 3))
    ActiveSheet.Range("B4:D" + CStr(i)).ClearContents
    ActiveSheet.Range("B4:D" + CStr(i)).ClearFormats
           
    'Индекс для последствий
    Numeration = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого последствия
    For Each Key In InputDictionary.keys
        Numeration = Numeration + 1
        ActiveSheet.Cells(i, 2).Value = Numeration
        ActiveSheet.Cells(i, 3).Value = Key
        If TypeName(InputDictionary.Item(Key)) <> "String" Then
            Erase Buffer
            Buffer = InputDictionary.Item(Key)
            Output = ""
            For Each Element In Buffer
                Output = Output + CStr(Element) + "|"
            Next Element
            ActiveSheet.Cells(i, 4).Value = Left(Output, Len(Output) - 1)
        Else
            ActiveSheet.Cells(i, 4).Value = CStr(InputDictionary.Item(Key))
        End If
        i = i + 1
    Next Key

    
    i = FindEmptyRowInColumn(ActiveSheet.Cells(3, 3)) - 1
    'Настраивается внешний вид
    With ActiveSheet.Range("B4:D" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
End Function
Function CheckExistense(CheckingVar As Variant, InputMassive() As Variant) As Boolean
    CheckExistense = False
    Dim Element As Variant
    For Each Element In InputMassive
        If Element = CheckingVar Then
            CheckExistense = True
            Exit For
        End If
    Next Element
End Function
Function CreateOutputStringForBDU(SheetName As String, sRow As Integer, sColumn As Integer, Optional Description As String) As String
    CreateOutputStringForBDU = ""
    Dim i%
    
    i = 2
    Do While Sheets(SheetName).Cells(sRow + i, sColumn).Value <> ""
        If Sheets(SheetName).Cells(sRow + i, sColumn + 1).Value = "Применимо" Then
            CreateOutputStringForBDU = CreateOutputStringForBDU + Sheets(SheetName).Cells(sRow + i, sColumn).Value + "|"
        End If
        i = i + 1
    Loop
    If CreateOutputStringForBDU = "" Then
        MsgBox Description + " все неприменимы? Мне бы твою уверенность...", , "WARNING: EMPTY FIELDS"
    Else
        CreateOutputStringForBDU = Left(CreateOutputStringForBDU, Len(CreateOutputStringForBDU) - 1)
    End If
End Function
Function SortDict(ByRef InputDictionary As Object)
    Dim temp() As String
    Dim Element As Variant
    Dim itX As Integer
    Dim itY As Integer
    Dim TempDict As Object
    Set TempDict = CreateObject("Scripting.Dictionary")
    
    For Each Element In InputDictionary.keys
        TempDict.Add Element, InputDictionary.Item(Element)
    Next Element

    'Only sort if more than one item in the dict
    If InputDictionary.Count > 1 Then

        'Populate the array
        ReDim temp(InputDictionary.Count)
        itX = 0
        For Each Element In InputDictionary
            temp(itX) = Element
            itX = itX + 1
        Next

        'Do the sort in the array
        For itX = 0 To (InputDictionary.Count - 2)
            For itY = (itX + 1) To (InputDictionary.Count - 1)
                If StrComp(temp(itX), temp(itY), 0) = 1 Then
                    Element = temp(itY)
                    temp(itY) = temp(itX)
                    temp(itX) = Element
                End If
            Next
        Next

        'Create the new dictionary
        Set InputDictionary = CreateObject("Scripting.Dictionary")
        For itX = 0 To (TempDict.Count - 1)
            InputDictionary.Add temp(itX), TempDict.Item(temp(itX))
        Next
    End If
End Function
Function DisplayDictionaryOnDesk(InputDictionary As Variant, StartCell As Range, SheetName As String, DatabaseSheet As String, ID As Integer, DatabaseColumn As Integer)
    Dim Element As Variant
    Dim Element_1 As Variant
    Dim Num%
    Dim temp() As String
    
    Num = 0
    For Each Element In InputDictionary.keys
        Num = Num + 1
        Sheets(SheetName).Cells(StartCell.Row + Num + 1, StartCell.Column).Value = Num
        Sheets(SheetName).Cells(StartCell.Row + Num + 1, StartCell.Column + 1).Value = InputDictionary.Item(Element)
        Sheets(SheetName).Cells(StartCell.Row + Num + 1, StartCell.Column + 2).Value = Element
        Sheets(SheetName).Cells(StartCell.Row + Num + 1, StartCell.Column + 3).Value = "Неприменимо"
        If Sheets(DatabaseSheet).Cells(ID + 3, DatabaseColumn).Value <> "" Then
            'Строка разбивается на индексы
            temp = Split(Sheets(DatabaseSheet).Cells(ID + 3, DatabaseColumn).Value, "|")
            For Each Element_1 In temp
                If Element_1 = Element Then
                    Sheets(SheetName).Cells(StartCell.Row + Num + 1, StartCell.Column + 3).Value = "Применимо"
                    Exit For
                End If
            Next Element_1
        End If
    Next Element
End Function
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
Function WriteMeasuresDictionary(ByRef SheetWithMeasures As Worksheet, ByRef SheetWithMeasuresForBIOS As Worksheet, _
ByRef Result As Variant, _
MeasuresColumn As Integer, IndexColumn As Integer, ByVal RowIndex As Integer, _
Optional DictionaryName As String)
    Dim i%
    Dim ID$
        
    Set Result = CreateObject("Scripting.Dictionary")
    Result.RemoveAll
       
    'Мер нет?
    If SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Value = "Меры для нейтрализации данной угрозы прорабатываются в рамках реализации ИТ-проекта." Then

        Result.Add "EMPTY", True
        RowIndex = RowIndex + 1
        'Если (строка с номером пустая ИЛИ строка с номером содержит номер УБИ ИДЕНТИЧНЫЙ тому, который мы уже имеем)
        'И
        '(строка с номером НЕ пустая и объединенная)
        Do While (SheetWithMeasures.Cells(RowIndex, IndexColumn).Value = "" Or _
        ID = SheetWithMeasures.Cells(RowIndex, IndexColumn).Value) And _
        (SheetWithMeasures.Cells(RowIndex, IndexColumn).Value <> "" And _
        SheetWithMeasures.Cells(RowIndex, IndexColumn).MergeCells)
            RowIndex = RowIndex + 1
        Loop
'    Меры надо искать на другой странице?
    ElseIf InStr(SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text, "УБИ к BIOS") <> 0 Then
        i = 3
        Do While SheetWithMeasuresForBIOS.Cells(i, 1).Value <> ""
            If Not Result.Exists(Left(SheetWithMeasuresForBIOS.Cells(i, 1).Text, InStr(SheetWithMeasuresForBIOS.Cells(i, 1).Text, " ") - 1)) Then
                Result.Add Left(SheetWithMeasuresForBIOS.Cells(i, 1).Text, InStr(SheetWithMeasuresForBIOS.Cells(i, 1).Text, " ") - 1), True
            End If
            i = i + 1
        Loop
        
        i = 3
        Do While SheetWithMeasuresForBIOS.Cells(i, 2).Value <> ""
            If Not Result.Exists(Left(SheetWithMeasuresForBIOS.Cells(i, 2).Text, InStr(SheetWithMeasuresForBIOS.Cells(i, 1).Text, " ") - 1)) Then
                Result.Add Left(SheetWithMeasuresForBIOS.Cells(i, 2).Text, InStr(SheetWithMeasuresForBIOS.Cells(i, 1).Text, " ") - 1), True
            End If
            i = i + 1
        Loop
'    Имеется нормальный перечень мер?
    Else
        
        'Если (строка с номером пустая ИЛИ строка с номером содержит номер УБИ ИДЕНТИЧНЫЙ тому, который мы уже имеем)
        'И
        'ячейка имеет границ больше 1 (там все ячейки имеет от 2 до 4, кроме тех, которые идут после всего списка)
        ID = SheetWithMeasures.Cells(RowIndex, IndexColumn).Value
        Do While (ID = SheetWithMeasures.Cells(RowIndex, IndexColumn).Value Or _
        SheetWithMeasures.Cells(RowIndex, IndexColumn).Value = "") And _
        SheetWithMeasures.Cells(RowIndex, IndexColumn).Borders.Count > 1
            'Если текущая мера уже есть в словаре, то писать её не надо
            If SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text <> "" Then
                If Not Result.Exists(Left(SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text, InStr(SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text, " ") - 1)) Then
                    Result.Add Left(SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text, InStr(SheetWithMeasures.Cells(RowIndex, MeasuresColumn).Text, " ") - 1), True
                End If
            End If
            RowIndex = RowIndex + 1
        Loop
    End If
End Function
Function ExtractMeasuresFromDictionary(ByRef InputDictionary As Variant) As String
    Dim Element As Variant
    ExtractMeasuresFromDictionary = ""
    For Each Element In InputDictionary
        ExtractMeasuresFromDictionary = ExtractMeasuresFromDictionary + CStr(Element) + "|"
    Next Element
    ExtractMeasuresFromDictionary = Left(ExtractMeasuresFromDictionary, Len(ExtractMeasuresFromDictionary) - 1)
End Function
Function RewriteOrderList(SheetName As String, ChosenBook As Excel.Workbook, InputDictionary As Variant)
    Dim i%
    Dim letter$
    Dim InputDictionaryElement As Variant
    
    Select Case SheetName:
        Case "Order239"
            letter = ":G"
        Case "Order31"
            letter = ":G"
        Case "Order21"
            letter = ":H"
        Case "Order17"
            letter = ":G"
    End Select
    
    i = 4
    Do While ChosenBook.Sheets(SheetName).Cells(i, 2).Value <> "Внутренние меры защиты"
        i = i + 1
    Loop
    Set temp = ChosenBook.Sheets(SheetName).Cells(i, 2)
    i = FindDoubleEmptyRowInColumn(temp) - 1
'    Очистка старых мер на всякий случай
    ChosenBook.Sheets(SheetName).Range("B" + CStr(temp.Row + 1) + letter + CStr(i)).ClearContents
    ChosenBook.Sheets(SheetName).Range("B" + CStr(temp.Row + 1) + letter + CStr(i)).ClearFormats
'    Вписываем меры из подгруженного файла
    i = 1 + temp.Row
    For Each InputDictionaryElement In InputDictionary
        ChosenBook.Sheets(SheetName).Cells(i, 2).Value = InputDictionaryElement
        ChosenBook.Sheets(SheetName).Cells(i, 3).Value = InputDictionary.Item(InputDictionaryElement)
        i = i + 1
    Next InputDictionaryElement
    

    
    With Sheets(SheetName).Range("B" + CStr(temp.Row + 1) + letter + CStr(i - 1))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
        .Font.Name = "Times New Roman"
        .Font.Size = 10
    End With
End Function
Function ConvertCollectionToMassive(InputCollection As Collection, ByRef OutputMassive() As String)
    ReDim OutputMassive(InputCollection.Count - 1, UBound(InputCollection.Item(1)) + 1)
    Dim Element As Variant
    
    Dim i%, j%
    i = 0
    For Each Element In InputCollection
        OutputMassive(i, 0) = i + 1
        For j = 1 To UBound(InputCollection.Item(1)) + 1
            OutputMassive(i, j) = Element(j - 1)
        Next j
        i = i + 1
    Next Element
End Function
Function ChangeElementInArrayOfCollection(ByRef InputCollection As Collection, ItemIndex As Integer, ElementIndex As Integer, NewValue As Variant)
    Dim temp() As Variant
    Dim i%
    
    ReDim temp(UBound(InputCollection.Item(ItemIndex)))
    
    For i = 0 To UBound(InputCollection.Item(ItemIndex))
        temp(i) = InputCollection.Item(ItemIndex)(i)
    Next i
    
    If ElementIndex <= UBound(temp) Then
        temp(ElementIndex) = NewValue
    Else
        ReDim Preserve temp(UBound(temp) + 1)
        temp(UBound(temp)) = NewValue
    End If
    
    If ItemIndex = InputCollection.Count Then
        InputCollection.Remove ItemIndex
        InputCollection.Add temp
    Else
        InputCollection.Remove ItemIndex
        InputCollection.Add temp, , ItemIndex
    End If
End Function
