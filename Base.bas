Attribute VB_Name = "Base"
Option Explicit

'Переменные
Public CategoryOfSystem As Integer
Public ClearQuestionaryOfMeasures As Integer
Public RegulatoryDocumentPage As String
Public ShadowStart As Boolean
Public TypeOfWorkmode As String
Public ThreatsDeskBusy As Boolean
Public ContinueExraction As Boolean
Public OutputTable As New Collection

'Словари
Public Consequences As Object
Public GoalsOfIntruder As Object
Public Interfaces As Object
Public Intruder As Object
Public Things As Object
Public TypesOfImpact As Object
Public WaysOfRealization As Object
Public Threats() As New Threat
Public Techniques() As New Technique
Public Measures() As New Measure
Public LevelsOfIntruder As Object
Public CategoriesOfIntruder As Object

'Справочники
Public RefConsequencesToThings As Object
Public RefGoalsOfIntruderToConsequences As Object
Public RefInterfacesToCategory As Object
Public RefInterfacesToLvl As Object
Public RefIntrudersToCategory As Object
Public RefIntrudersToConsequences As Object
Public RefIntrudersToGoals As Object
Public RefIntrudersToLvl As Object
Public RefIntrudersToThings As Object
Public RefThingsToConsequences As Object
Public RefThingsToInterfaces As Object
Public RefTypesOfImpactToConsequences As Object
Public RefTypesOfImpactToThings As Object
Public RefWaysOfRealizationToCategory As Object
Public RefWaysOfRealizationToInterfaces As Object
Public RefWaysOfRealizationToLvl As Object
Public RefWaysOfRealizationToThings As Object
Public RefWaysOfRealizationToTypesOfImpact As Object

Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Sub Start()
    ShadowStart = True
    ShadowStartForm.Show
End Sub
Sub QNC_UpdateRefs()
    Call WriteDictionary("QNC", Consequences, "Consequences", True)
    Call WriteDictionary("DNoI", Intruder, "Intruder")
    Call WriteDictionary("DGoI", GoalsOfIntruder, "GoalsOfIntruder")
    Call WriteDictionary("DToI", TypesOfImpact, "TypesOfImpact")
    Call WriteBookOfReferenceFromDefault("RefCToI", RefTypesOfImpactToConsequences, "RefTypesOfImpactToConsequences", True)
    Call WriteBookOfReferenceFromDefault("RefTTToI", RefTypesOfImpactToThings, "RefTypesOfImpactToThings", True)
End Sub
Sub QTT_UpdateDict()
    Call WriteDictionary("QTT", Things, "Things", True)
End Sub
Sub QTTToI_Write()
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim NumOfCons As Integer
    Dim temp As Range
    Dim ID_Consequences As Variant
    Dim ID_TypesOfImpact As Variant
    Dim ID_Things As Variant
       
    'поиск столбца "ВИД ВОЗДЕЙСТВИЯ" для очистки
    Set temp = Sheets("QTTToI").Range("2:2").Find(What:="ВИД ВОЗДЕЙСТВИЯ", Lookat:=xlWhole)

    Call Functions.WriteBookOfReferenceFromDefault("RefTNC", RefThingsToConsequences, "RefThingsToConsequences")

    i = FindEmptyRowInColumn(temp)
    If Not (ShadowStartForm.ClearAppliances.Value = False _
    And ShadowStart = True) Then
        Sheets("QTTToI").Range("B4:H" + CStr(i)).ClearContents
        Sheets("QTTToI").Range("B4:H" + CStr(i)).ClearFormats
    Else
        Sheets("QTTToI").Range("B4:G" + CStr(i)).ClearContents
        Sheets("QTTToI").Range("B4:G" + CStr(i)).ClearFormats
    End If
    
    'Индекс для последствий
    NumOfCons = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого последствия
    For Each ID_Consequences In Consequences.keys
        'Выводятся все актуальные объекты
        For Each ID_Things In Things.keys
            Sheets("QTTToI").Cells(i, 4).Value = Things.Item(ID_Things)
            For Each ID_TypesOfImpact In TypesOfImpact.keys
'                Если этот ID воздействия есть в обоих словарях в качестве ключа
                If RefTypesOfImpactToConsequences.Exists(ID_TypesOfImpact) _
                And RefTypesOfImpactToThings.Exists(ID_TypesOfImpact) Then
                    'И если этому ID есть последсвтия и угрозы в соответствующих словарях, то мы выводим это воздействие
                    'Этот кусок *№";%?"* не способен передать массив из словаря в функцию
                    If CheckReferences(CStr(ID_Consequences), CStr(ID_TypesOfImpact), RefTypesOfImpactToConsequences) _
                    And CheckReferences(CStr(ID_Things), CStr(ID_TypesOfImpact), RefTypesOfImpactToThings) Then
                        NumOfCons = NumOfCons + 1
                        Sheets("QTTToI").Cells(i, 2).Value = NumOfCons
                        Sheets("QTTToI").Cells(i, 3).Value = Consequences.Item(ID_Consequences)
                        Sheets("QTTToI").Cells(i, 4).Value = Things.Item(ID_Things)
                        Sheets("QTTToI").Cells(i, 5).Value = TypesOfImpact.Item(ID_TypesOfImpact)
                        Sheets("QTTToI").Cells(i, 6).Value = ID_Things
                        Sheets("QTTToI").Cells(i, 7).Value = ID_Consequences
                        'Авторасстановка применимости по предгенеренному словарю
                        'Если не в теневом режиме
                        If ShadowStart = False Then
                            If Functions.CheckReferences(ID_Consequences, ID_Things, RefThingsToConsequences) Then
                                Sheets("QTTToI").Cells(i, 8).Value = "Применимо"
                            Else
                                Sheets("QTTToI").Cells(i, 8).Value = "Неприменимо"
                            End If
                        End If
                        i = i + 1
                    End If
                End If
            Next ID_TypesOfImpact
'            Если для объекта не нашлось воздействий, то он стирается
            If Sheets("QTTToI").Cells(i, 5).Value = "" _
            And Sheets("QTTToI").Cells(i, 4).Value <> "" Then
                Sheets("QTTToI").Cells(i, 2).Value = ""
                Sheets("QTTToI").Cells(i, 3).Value = ""
                Sheets("QTTToI").Cells(i, 4).Value = ""
                Sheets("QTTToI").Cells(i, 6).Value = ""
                Sheets("QTTToI").Cells(i, 7).Value = ""
            End If
        Next ID_Things
    Next ID_Consequences
    
    Call SetApplianceColumn("QTTToI")
    
    i = FindEmptyRowInColumn(temp) - 1
    'Настраивается внешний вид
    With Sheets("QTTToI").Range("B4:H" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
End Sub
Sub QTTToI_UpdateRefs()

    Dim i As Integer
    Dim NumOfCons As Integer
    Dim temp As Range
    Dim Name_TypesOfImpact$, Name_Things$, Name_Consequences$
    Dim ID_TypesOfImpact As Variant
    Dim ID_Things As Variant
    Dim ID_Consequences As Variant
    Dim IsUsing As Boolean
       
    'поиск столбца "ВИД ВОЗДЕЙСТВИЯ" для индексации
    Set temp = Sheets("QTTToI").Range("2:2").Find(What:="ВИД ВОЗДЕЙСТВИЯ", Lookat:=xlWhole)
 
    'Для каждого воздействия выполняется проверка наличия актуальных объектов
    For Each ID_TypesOfImpact In RefTypesOfImpactToThings.keys
        i = 2
        IsUsing = False
        'Ищет имя способа воздейтсвия по ID
        Name_TypesOfImpact = FindIDorName(CStr(ID_TypesOfImpact), TypesOfImpact, True)
        
        Do While Sheets("QTTToI").Cells(temp.Row + i, temp.Column) <> "" Or _
        Sheets("QTTToI").Cells(temp.Row + i + 1, temp.Column) <> ""
            'Если значение ячейки и имя совпадает, а также применимость = true, то этот вид воздействия останется
            If Sheets("QTTToI").Cells(temp.Row + i, temp.Column) = Name_TypesOfImpact _
            And Sheets("QTTToI").Cells(temp.Row + i, temp.Column + 3) = "Применимо" Then
                IsUsing = True
                'Перебирать дальше смысла нет
                Exit Do
            End If
            i = i + 1
        Loop
        'Если все воздействия признали неактуальными, то его можно из словаря удалять
        If Not IsUsing Then
            RefTypesOfImpactToThings.Remove (ID_TypesOfImpact)
        End If
    Next ID_TypesOfImpact
    
    'Актуализация словаря объектов
    For Each ID_Things In Things.keys
        i = 2
        IsUsing = False
        'Ищет имя объекта по ID
        Name_Things = FindIDorName(CStr(ID_Things), Things, True)
        
        'Продолжает, пока не натыкается на две пустых ячейки подряд в столбце воздействий
        'Или пока не подтвердится актуальность объекта
        Do While (Sheets("QTTToI").Cells(temp.Row + i, temp.Column) <> "" _
        Or Sheets("QTTToI").Cells(temp.Row + i + 1, temp.Column) <> "") _
        And Not IsUsing
            'Если значение ячейки и имя совпадает, то производится проверка всех ячеек ниже, на случай, если одна из них применима
            Do While Sheets("QTTToI").Cells(temp.Row + i, temp.Column - 1) = Name_Things _
            Or Sheets("QTTToI").Cells(temp.Row + i, temp.Column - 1) = ""
                If Sheets("QTTToI").Cells(temp.Row + i, temp.Column + 3).Value = "Применимо" Then
                    IsUsing = True
                    'Перебирать дальше смысла нет
                    Exit Do
                End If
                'Если иденкс уже вышел за пределы таблицы
                If Sheets("QTTToI").Cells(temp.Row + i, temp.Column) = "" _
                Or Sheets("QTTToI").Cells(temp.Row + i + 1, temp.Column) = "" Then
                    Exit Do
                End If
                i = i + 1
            Loop
            i = i + 1
        Loop
        'Если все воздействия признали неактуальными, то его можно из словаря удалять
        If Not IsUsing Then
            Things.Remove (ID_Things)
        End If
    Next ID_Things
    
    'Аткуализация словаря последствий
    For Each ID_Consequences In Consequences.keys
        i = 2
        IsUsing = False
        'Ищет имя объекта по ID
        Name_Consequences = FindIDorName(CStr(ID_Consequences), Consequences, True)
        
        'Продолжает, пока не натыкается на две пустых ячейки подряд в столбце воздействий
        'Или пока не подтвердится актуальность объекта
        Do While (Sheets("QTTToI").Cells(temp.Row + i, temp.Column) <> "" _
        Or Sheets("QTTToI").Cells(temp.Row + i + 1, temp.Column) <> "") _
        And Not IsUsing
            'Если значение ячейки и имя совпадает, то производится проверка всех ячеек ниже, на случай, если одна из них применима
            Do While Sheets("QTTToI").Cells(temp.Row + i, temp.Column - 2) = Name_Consequences _
            Or Sheets("QTTToI").Cells(temp.Row + i, temp.Column - 2) = ""
                If Sheets("QTTToI").Cells(temp.Row + i, temp.Column + 3).Value = "Применимо" Then
                    IsUsing = True
                    'Перебирать дальше смысла нет
                    Exit Do
                End If
                'Если иденкс уже вышел за пределы таблицы
                If Sheets("QTTToI").Cells(temp.Row + i, temp.Column) = "" _
                Or Sheets("QTTToI").Cells(temp.Row + i + 1, temp.Column) = "" Then
                    Exit Do
                End If
                i = i + 1
            Loop
            i = i + 1
        Loop
        'Если все воздействия признали неактуальными, то его можно из словаря удалять
        If Not IsUsing Then
            Consequences.Remove (ID_Consequences)
        End If
    Next ID_Consequences
    
    Call Functions.WriteBookOfReferenceFromAuto("QTTToI", RefThingsToConsequences, "RefThingsToConsequences", , True)
    
End Sub
Sub QNCGoI_Write()
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim NumOfCons As Integer
    Dim temp As Range
    Dim ID_Consequences As Variant
    Dim ID_Goals As Variant
       
    'поиск столбца "ВОЗМОЖНЫЕ ЦЕЛИ РЕАЛИЗАЦИИ УГРОЗ БЕЗОПАСНОСТИ ИНФОРМАЦИИ" для очистки
    Set temp = Sheets("QNCGoI").Range("3:3").Find(What:="3", Lookat:=xlWhole)

    i = FindEmptyRowInColumn(temp)
    If Not (ShadowStartForm.ClearAppliances.Value = False _
    And ShadowStart = True) Then
        Sheets("QNCGoI").Range("B4:G" + CStr(i)).ClearContents
        Sheets("QNCGoI").Range("B4:G" + CStr(i)).ClearFormats
    Else
        Sheets("QNCGoI").Range("B4:F" + CStr(i)).ClearContents
        Sheets("QNCGoI").Range("B4:F" + CStr(i)).ClearFormats
    End If

    Call WriteBookOfReferenceFromDefault("RefCGOI", RefGoalsOfIntruderToConsequences, "RefGoalsOfIntruderToConsequences", True)
    
    'Индекс для последствий
    NumOfCons = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого последствия
    For Each ID_Consequences In Consequences.keys
        'Перебираются все цели
        For Each ID_Goals In GoalsOfIntruder.keys
            NumOfCons = NumOfCons + 1
            Sheets("QNCGoI").Cells(i, 2).Value = NumOfCons
            Sheets("QNCGoI").Cells(i, 3).Value = Consequences.Item(ID_Consequences)
            Sheets("QNCGoI").Cells(i, 4).Value = GoalsOfIntruder.Item(ID_Goals)
            Sheets("QNCGoI").Cells(i, 5).Value = ID_Consequences
            Sheets("QNCGoI").Cells(i, 6).Value = ID_Goals
            If Functions.CheckReferences(ID_Consequences, ID_Goals, RefGoalsOfIntruderToConsequences) Then
                Sheets("QNCGoI").Cells(i, 7).Value = "Применимо"
            Else
                Sheets("QNCGoI").Cells(i, 7).Value = "Неприменимо"
            End If
            i = i + 1
        Next ID_Goals
    Next ID_Consequences
    Call SetApplianceColumn("QNCGoI")
    
    i = FindEmptyRowInColumn(temp) - 1
    'Настраивается внешний вид
    With Sheets("QNCGoI").Range("B4:G" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
End Sub
Sub QNCGoI_UpdateRefs()

    Call WriteBookOfReferenceFromAuto("QNCGoI", RefGoalsOfIntruderToConsequences, "RefGoalsOfIntruderToConsequences", True, True)
    Call WriteBookOfReferenceFromDefault("RefNoIGoI", RefIntrudersToGoals, "RefIntrudersToGoals")
    
End Sub
Sub TNCGoINoI_Write()
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim NumOfIntruder As Integer
    Dim temp As Range
    Dim IDs() As String
    Dim ID_Consequences As Variant
    Dim ID_Goals As Variant
    Dim ID_Intruder As Variant
    Dim ID_ConsInRefs As Variant
    Dim ID_Things As Variant

    'поиск столбца для очистки
    Set temp = Sheets("TNCGoINoI").Range("2:2").Find(What:="ID", LookIn:=xlFormulas, Lookat:=xlPart)

    i = FindEmptyRowInColumn(temp)
    Sheets("TNCGoINoI").Range("B4:G" + CStr(i)).ClearContents
    Sheets("TNCGoINoI").Range("B4:G" + CStr(i)).ClearFormats
    
    'Индекс для последствий
    NumOfIntruder = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого нарушителя
    For Each ID_Intruder In Intruder.keys
        For Each ID_Goals In GoalsOfIntruder.keys
            'Проверка наличия цели для текущего нарушителя в справочнике Вид нарушителя/Цели"
            'Проверка наличия для цели негативных последствий
            If CheckReferences(CStr(ID_Goals), CStr(ID_Intruder), RefIntrudersToGoals) _
            And RefGoalsOfIntruderToConsequences.Exists(ID_Goals) Then
                'По соответствующей цели перебираются все найденные актуальные для неё ID последствий
                For Each ID_Consequences In RefGoalsOfIntruderToConsequences.Item(ID_Goals)
                    NumOfIntruder = NumOfIntruder + 1
                    Sheets("TNCGoINoI").Cells(i, 2).Value = NumOfIntruder
                    Sheets("TNCGoINoI").Cells(i, 3).Value = Intruder.Item(ID_Intruder)
                    Sheets("TNCGoINoI").Cells(i, 4).Value = Functions.CategoryOutput(ID_Intruder, RefIntrudersToCategory)
                    Sheets("TNCGoINoI").Cells(i, 5).Value = RefIntrudersToLvl.Item(ID_Intruder)
                    'Записывается цель
                    Sheets("TNCGoINoI").Cells(i, 6).Value = GoalsOfIntruder.Item(ID_Goals)
                    'Записывается последствие
                    Sheets("TNCGoINoI").Cells(i, 7).Value = Consequences.Item(ID_Consequences)
                    'Записывается ID пары "Нарушитель/Последствие"
                    Sheets("TNCGoINoI").Cells(i, 8).Value = ID_Intruder
                    Sheets("TNCGoINoI").Cells(i, 9).Value = ID_Consequences
                    i = i + 1
                Next ID_Consequences
            End If
        Next ID_Goals
    Next ID_Intruder
    
    Call WriteBookOfReferenceFromAuto("TNCGoINoI", RefIntrudersToConsequences, "RefIntrudersToConsequences")
    Set RefIntrudersToThings = CreateObject("Scripting.Dictionary")
    
    If RefIntrudersToConsequences.Count <> 0 Then
        For Each ID_Intruder In Intruder.keys
            If Not RefIntrudersToConsequences.Exists(ID_Intruder) Then
                Intruder.Remove (ID_Intruder)
            Else
                For Each ID_Things In Things.keys
                    For Each ID_ConsInRefs In RefIntrudersToConsequences.Item(ID_Intruder)
                        If CheckReferences(CStr(ID_ConsInRefs), CStr(ID_Things), RefThingsToConsequences) _
                        And CheckReferences(CStr(ID_ConsInRefs), CStr(ID_Intruder), RefIntrudersToConsequences) Then
                            If Not RefIntrudersToThings.Exists(ID_Intruder) Then
                                ReDim IDs(0)
                                IDs(0) = CStr(ID_Things)
                                RefIntrudersToThings.Add ID_Intruder, IDs
                            ElseIf Not CheckReferences(CStr(ID_Things), CStr(ID_Intruder), RefIntrudersToThings) Then
                '               Если ID уже есть в словаре, то создаётся массив, который содержит все значения массива соответствующего ключа
                                Erase IDs
                                IDs() = RefIntrudersToThings.Item(ID_Intruder)
                '               Затем этот массив расширается и дописывается новым элементом
                                ReDim Preserve IDs(UBound(IDs) + 1)
                                IDs(UBound(IDs)) = CStr(ID_Things)
                '               Т.к. переобозначить Item для ключа нельзя, то ключ удаляется и записывается заново с дополненным массивом
                                RefIntrudersToThings.Remove ID_Intruder
                                RefIntrudersToThings.Add ID_Intruder, IDs
                            End If
                        End If
                    Next
                Next ID_Things
            End If
        Next ID_Intruder
    End If
    
    i = FindEmptyRowInColumn(temp) - 1
    'Настраивается внешний вид
    With Sheets("TNCGoINoI").Range("B4:G" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
End Sub
Sub QCollusion_UpdateRef()
    Dim i As Integer
    Dim IDs() As String
    Dim Category As Variant
    Dim ID_Intruder As Variant
    
    Call Functions.WriteBookOfReferenceFromAuto("RefNoIC", RefIntrudersToCategory, "RefIntrudersToCategory")
    Call Functions.WriteBookOfReferenceFromAuto("RefNoILoC", RefIntrudersToLvl, "RefIntrudersToLvl")
    
    If Sheets("QCollusion").Cells(4, 4).Value = "Да" Then
        Call Functions.AddItemToKey("Внутренний", "NoI_1", RefIntrudersToCategory)
'        RefIntrudersToLvl.Item("NoI_5") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_6") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_7") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_8") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_9") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_10") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_11") = RefIntrudersToLvl.Item("NoI_1")
'        RefIntrudersToLvl.Item("NoI_12") = RefIntrudersToLvl.Item("NoI_1")
    End If
    If Sheets("QCollusion").Cells(5, 4).Value = "Да" Then
        Call Functions.AddItemToKey("Внутренний", "NoI_2", RefIntrudersToCategory)
'        If RefIntrudersToLvl.Item("NoI_10")(0) < RefIntrudersToLvl.Item("NoI_2")(0) Then
'            RefIntrudersToLvl.Item("NoI_10") = RefIntrudersToLvl.Item("NoI_2")
'            RefIntrudersToLvl.Item("NoI_11") = RefIntrudersToLvl.Item("NoI_2")
'            RefIntrudersToLvl.Item("NoI_12") = RefIntrudersToLvl.Item("NoI_2")
'        End If
    End If
    If Sheets("QCollusion").Cells(6, 4).Value = "Да" Then
        Call Functions.AddItemToKey("Внутренний", "NoI_3", RefIntrudersToCategory)
'        If RefIntrudersToLvl.Item("NoI_10")(0) < RefIntrudersToLvl.Item("NoI_3")(0) Then
'            RefIntrudersToLvl.Item("NoI_10") = RefIntrudersToLvl.Item("NoI_3")
'            RefIntrudersToLvl.Item("NoI_11") = RefIntrudersToLvl.Item("NoI_3")
'            RefIntrudersToLvl.Item("NoI_12") = RefIntrudersToLvl.Item("NoI_3")
'        End If
    End If

End Sub
Sub QIntOfTT_Write()
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim NumOfObj As Integer
    Dim temp As Range
    Dim ID_Thing As Variant
    Dim ID_Interface As Variant
    Dim ID_Intruder As Variant
    Dim MaxLvl$, Internal$, External$
    Dim ItemsBuffer() As String
       
    'поиск столбца "ПРИМЕНИМОСТЬ" для очистки
    Set temp = Sheets("QIntOfTT").Range("2:2").Find(What:="ID", LookIn:=xlFormulas, Lookat:=xlPart)
    
    'Заполняем необходимые словари
    Call Functions.WriteBookOfReferenceFromDefault("RefTTInt", RefThingsToInterfaces, "RefThingsToInterfaces")
    Call Functions.WriteBookOfReferenceFromDefault("RefIntLoC", RefInterfacesToLvl, "RefInterfacesToLvl")
    Call Functions.WriteBookOfReferenceFromDefault("RefIntCat", RefInterfacesToCategory, "RefInterfacesToCategory")
    Call Functions.WriteDictionary("DInt", Interfaces, "Interfaces")
    
    'Определяем максимальный уровень нарушителя и категории
    MaxLvl = "Н1"
    For Each ID_Intruder In Intruder.keys
        'Если элемент категории один
        If RefIntrudersToCategory.Item(ID_Intruder)(0) = "Внешний" Then
            External = "Внешний"
        ElseIf RefIntrudersToCategory.Item(ID_Intruder)(0) = "Внутренний" Then
            Internal = "Внутренний"
        End If
        'Если элементов стало больше одного
        If UBound(RefIntrudersToCategory.Item(ID_Intruder)) Then
            If RefIntrudersToCategory.Item(ID_Intruder)(1) = "Внешний" Then
                External = "Внешний"
            ElseIf RefIntrudersToCategory.Item(ID_Intruder)(1) = "Внутренний" Then
                Internal = "Внутренний"
            End If
        End If
        'Проверка максимального уровня
        If MaxLvl < RefIntrudersToLvl.Item(ID_Intruder)(0) Then
            MaxLvl = RefIntrudersToLvl.Item(ID_Intruder)(0)
        End If
    Next ID_Intruder
    

    i = FindEmptyRowInColumn(temp)
    If Not (ShadowStartForm.ClearAppliances.Value = False _
    And ShadowStart = True) Then
        Sheets("QIntOfTT").Range("B4:G" + CStr(i)).ClearContents
        Sheets("QIntOfTT").Range("B4:G" + CStr(i)).ClearFormats
    Else
        Sheets("QIntOfTT").Range("B4:F" + CStr(i)).ClearContents
        Sheets("QIntOfTT").Range("B4:F" + CStr(i)).ClearFormats
    End If
    
    'Индекс для объекта
    NumOfObj = 0
    'Индексация по листу опросника
    i = 4
    'Для каждого объекта
    For Each ID_Thing In Things.keys
        For Each ID_Interface In Interfaces.keys
            'Если максимальный уровень больше или равен необходимому для интерфейса
            'И (категория является внешней или внутренней, необходимой для интерфейса)
            If MaxLvl >= RefInterfacesToLvl.Item(ID_Interface)(0) _
            And (Functions.CheckReferences(External, ID_Interface, RefInterfacesToCategory) _
            Or Functions.CheckReferences(Internal, ID_Interface, RefInterfacesToCategory)) Then
                NumOfObj = NumOfObj + 1
                Sheets("QIntOfTT").Cells(i, 2).Value = NumOfObj
                Sheets("QIntOfTT").Cells(i, 3).Value = Things.Item(ID_Thing)
                Sheets("QIntOfTT").Cells(i, 4).Value = Interfaces.Item(ID_Interface)
                Sheets("QIntOfTT").Cells(i, 5).Value = ID_Thing
                Sheets("QIntOfTT").Cells(i, 6).Value = ID_Interface
                If RefThingsToInterfaces.Exists(ID_Thing) Then
                    If Functions.CheckReferences(ID_Interface, ID_Thing, RefThingsToInterfaces) Then
                        Sheets("QIntOfTT").Cells(i, 7).Value = "Применимо"
                    Else
                        Sheets("QIntOfTT").Cells(i, 7).Value = "Неприменимо"
                    End If
                Else
                    Sheets("QIntOfTT").Cells(i, 7).Value = "Применимо"
                End If
                i = i + 1
            End If
        Next ID_Interface
    Next ID_Thing
    Call SetApplianceColumn("QIntOfTT")
    
    i = FindEmptyRowInColumn(temp) - 1
    'Настраивается внешний вид
    With Sheets("QIntOfTT").Range("B4:G" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
End Sub
Sub QAoWoR_Write()
    Application.ScreenUpdating = False
    Dim i As Integer
    Dim Num As Integer
    Dim temp As Range
    Dim ID_Thing As Variant
    Dim ID_Interface As Variant
    Dim ID_Intruder As Variant
    Dim ID_TypeOfImpact As Variant
    Dim ID_WayOfRealization As Variant
    Dim OutputString() As String
       
    'поиск столбца "ПРИМЕНИМОСТЬ" для очистки
    Set temp = Sheets("QAoWoR").Range("2:2").Find(What:="СПОСОБ РЕАЛИЗАЦИИ", Lookat:=xlPart)
    Set OutputTable = Nothing
    'Заполняются необходимые словари
    Call Functions.WriteBookOfReferenceFromDefault("RefWoRC", RefWaysOfRealizationToCategory, "RefWaysOfRealizationToCategory")
    Call Functions.WriteBookOfReferenceFromDefault("RefWoRLoC", RefWaysOfRealizationToLvl, "RefWaysOfRealizationToLvl")
    Call Functions.WriteBookOfReferenceFromDefault("RefWoRInt", RefWaysOfRealizationToInterfaces, "RefWaysOfRealizationToInterfaces")
    Call Functions.WriteBookOfReferenceFromDefault("RefWoRTT", RefWaysOfRealizationToThings, "RefWaysOfRealizationToThings")
    Call Functions.WriteBookOfReferenceFromDefault("RefWoRToI", RefWaysOfRealizationToTypesOfImpact, "RefWaysOfRealizationToTypesOfImpact")
    Call Functions.WriteDictionary("DWoR", WaysOfRealization, "WaysOfRealization")

    i = FindEmptyRowInColumn(temp)
    If Not (ShadowStartForm.ClearAppliances.Value = False _
    And ShadowStart = True) Then
        Sheets("QAoWoR").Range("B4:J" + CStr(i)).ClearContents
        Sheets("QAoWoR").Range("B4:J" + CStr(i)).ClearFormats
    Else
        Sheets("QAoWoR").Range("B4:I" + CStr(i)).ClearContents
        Sheets("QAoWoR").Range("B4:I" + CStr(i)).ClearFormats
    End If
    
    'Индекс для нарушителя
    Num = 0
    'Индексация по листу опросника
    i = 4
    
    'Для каждого нарушителя
    ReDim OutputString(7)
    For Each ID_Intruder In Intruder.keys
        For Each ID_Thing In Things.keys
            'Если объект соответствует нарушителю (нарушитель связывается с объектом через последствие в процедуре TNCGoINoI_Write())
            If Functions.CheckReferences(ID_Thing, ID_Intruder, RefIntrudersToThings) Then
                'Для каждого интерфейса
                For Each ID_Interface In Interfaces.keys
                    'Если интерфейс хотя бы одну схожую категорию с нарушителем
                    'Если интерфейс подходит по уровню возможностей
                    'То он вписывается
                    If Functions.CheckCategory(ID_Interface, RefInterfacesToCategory, ID_Intruder, RefIntrudersToCategory) _
                    And RefInterfacesToLvl.Item(ID_Interface)(0) <= RefIntrudersToLvl.Item(ID_Intruder)(0) Then
                        'Выписываются все виды воздействий в каждый интерфейс
                        For Each ID_TypeOfImpact In TypesOfImpact.keys
                            'Если для объекта этот вид воздействий актуален, то мы его вписываем
                            If Functions.CheckReferences(ID_Thing, ID_TypeOfImpact, RefTypesOfImpactToThings) Then
                                'Для каждого вида воздействия выписываются все способы реализации
                                For Each ID_WayOfRealization In WaysOfRealization.keys
                                    'Если способ реализации подходит по категории
                                    'Если способ реализации подходит по уровню возможностей
                                    'Если способ реализации имеет текущий интерфейс в справочнике "Способ реализации>Интерфейсы"
                                    'Если способ реализации имеет текущий вид воздействия в справочнике "Способ реализации>Виды воздействия"
                                    'Если способ реализации имеет текущий объект в справочнике "Способ реализации>Объекты"
                                    If Functions.CheckCategory(ID_WayOfRealization, RefWaysOfRealizationToCategory, ID_Intruder, RefIntrudersToCategory) _
                                    And RefWaysOfRealizationToLvl.Item(ID_WayOfRealization)(0) <= RefIntrudersToLvl.Item(ID_Intruder)(0) _
                                    And Functions.CheckReferences(ID_Interface, ID_WayOfRealization, RefWaysOfRealizationToInterfaces) _
                                    And Functions.CheckReferences(ID_TypeOfImpact, ID_WayOfRealization, RefWaysOfRealizationToTypesOfImpact) _
                                    And Functions.CheckReferences(ID_Thing, ID_WayOfRealization, RefWaysOfRealizationToThings) Then
                                        'Данные о нарушителе
                                        OutputString(0) = Intruder.Item(ID_Intruder)
                                        OutputString(1) = Functions.CategoryOutput(ID_Intruder, RefIntrudersToCategory)
                                        OutputString(2) = Join(RefIntrudersToLvl.Item(ID_Intruder))
                                        'Название объекта
                                        OutputString(3) = Things.Item(ID_Thing)
                                        'Название интерфейса
                                        OutputString(4) = Interfaces.Item(ID_Interface)
                                        'Вид воздействия
                                        OutputString(5) = TypesOfImpact.Item(ID_TypeOfImpact)
                                        'Способ реализации
                                        OutputString(6) = WaysOfRealization.Item(ID_WayOfRealization)
                                        OutputString(7) = "Применимо"
                                        OutputTable.Add OutputString
                                        i = i + 1
                                    End If
                                Next ID_WayOfRealization
                            End If
                        Next ID_TypeOfImpact
                    End If
                Next ID_Interface
            End If
        Next ID_Thing
    Next ID_Intruder
    
    
    Call Functions.ConvertCollectionToMassive(OutputTable, OutputString)
    Sheets("QAoWoR").Range("B4:J" + CStr(OutputTable.Count + 3)) = OutputString
    Sheets("QAoWoR").ScrollArea = "A1:J" + CStr(OutputTable.Count + 5)
    
    Call SetApplianceColumn("QAoWoR")

    i = FindEmptyRowInColumn(temp) - 1
    'Настраивается внешний вид
    With Sheets("QAoWoR").Range("B4:J" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
End Sub
Sub TofThreats_Write()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    Dim i%, i_BDU%, i_TofThreats%, j%, k%
    Dim Num%
    Dim temp As Range
    Dim ID_Thing$, ID_TypeOfImpact$, ID_WayOfRealization$, RelativeThreats$, RelativeThreatsVisible$
    Dim IntruderStats() As String
    Dim stemp() As String
    Dim temp_WoR() As String
    Dim temp_ToI() As String
    Dim temp_Objects() As String
    Dim ThreatsElement As Variant
    Dim Element As Variant
    Dim RowInOutputTable As Variant
    Dim IntruderIsExists As Boolean
    Dim y As Long
    
    Set temp = Sheets("TofThreats").Range("2:2").Find(What:="№", Lookat:=xlPart)
    
    Base.DeclareThreats
    
    Call WriteDictionary("DToI", TypesOfImpact, "TypesOfImpact")
    Call WriteDictionary("QTT", Things, "Things", True)
    Call WriteDictionary("DWoR", WaysOfRealization, "WaysOfRealization")
    
    i = FindEmptyRowInColumn(temp)
    Sheets("TofThreats").Range("B4:K" + CStr(i)).ClearContents
    Sheets("TofThreats").Range("B4:K" + CStr(i)).ClearFormats
    

    'Индексация по листу опросника QAoWoR
    i = 4
    'Индексация по листу опросника TofThreats
    i_TofThreats = 4
    'Нумерация
    Num = 1
    
    'Обновление применимостей способов реализации
    Do While Sheets("QAoWoR").Cells(i, 9) <> "" Or _
    Sheets("QAoWoR").Cells(i + 1, 9) <> ""
        If Sheets("QAoWoR").Cells(i, 10) <> "Применимо" Then
            OutputTable.Remove Num
        End If
        i = i + 1
    Loop
    
    'Индексация по коллекции
    i = 1
    
    For Each RowInOutputTable In OutputTable
        'Находятся ID по их текстовым описаниям
        ID_Thing = Functions.FindIDorName(CStr(RowInOutputTable(3)), Base.Things)
        ID_TypeOfImpact = Functions.FindIDorName(CStr(RowInOutputTable(5)), Base.TypesOfImpact)
        ID_WayOfRealization = Functions.FindIDorName(CStr(RowInOutputTable(6)), Base.WaysOfRealization)

        stemp = Split(RowInOutputTable(1), "|")
        If TypeName(stemp) <> "String" Then
            ReDim IntruderStats(UBound(stemp), 1)
        Else
            ReDim IntruderStats(0, 1)
        End If
        j = 0
        For Each Element In stemp
            'Объект IntruderStats - двумерный массив, где 0 элмент - категория нарушителя, а 1 - его потенциал
            If Element = "Внешний" Then
                IntruderStats(j, 0) = "Внешний"
            ElseIf Element = "Внутренний" Then
                IntruderStats(j, 0) = "Внутренний"
            End If
            IntruderStats(j, 1) = RowInOutputTable(2)
            j = j + 1
        Next Element
        'Перебираются все угрозы
        RelativeThreats = ""
        RelativeThreatsVisible = ""
        For Each ThreatsElement In Threats
            'Если для угрозы способ реализации присутствует
            'И присутствует вид воздействия
            'И присутствует объект
            If ThreatsElement.CheckWoR(ID_WayOfRealization) And _
            ThreatsElement.CheckToI(ID_TypeOfImpact) And _
            ThreatsElement.CheckObjects(ID_Thing) Then
                'Выполняется сверка с нарушителем
                IntruderIsExists = False
                For j = 0 To UBound(IntruderStats)
                    For k = 0 To ThreatsElement.NumberOfIntruder
                        'Если категория совпадает
                        'А уровень возможностей, необходимый для угрозы, <= заявленного
                        If IntruderStats(j, 0) = ThreatsElement.Intruder(k, 0) And _
                        IntruderStats(j, 1) >= ThreatsElement.Intruder(k, 1) Then
                            RelativeThreatsVisible = RelativeThreatsVisible + CStr(ThreatsElement.ID) _
                            + ". " + ThreatsElement.Name + vbNewLine
                            RelativeThreats = RelativeThreats + ThreatsElement.ID + "|"
                            IntruderIsExists = True
                            Exit For
                        End If
                    Next k
                    If IntruderIsExists Then
                        Exit For
                    End If
                Next j
            End If
        Next ThreatsElement
        'Если угрозы нашлись, то можно продолжать индексацию
        If RelativeThreats <> "" Then
            Call Functions.ChangeElementInArrayOfCollection(OutputTable, i, 7, Left(RelativeThreats, Len(RelativeThreats) - 1))
            Call Functions.ChangeElementInArrayOfCollection(OutputTable, i, 8, RelativeThreatsVisible)
            i = i + 1
        Else
            OutputTable.Remove i
        End If
    Next RowInOutputTable
    
    Call Functions.ConvertCollectionToMassive(OutputTable, stemp)
    Sheets("TofThreats").Range("B4:K" + CStr(OutputTable.Count + 3)) = stemp
    Sheets("TofThreats").ScrollArea = "A1:K" + CStr(OutputTable.Count + 5)

    'Настраивается внешний вид
    With Sheets("TofThreats").Range("B4:K" + CStr(OutputTable.Count + 3))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
End Sub
Sub TofTechniques_Write()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    
    Dim i%, i_TofThreats%, i_TofTechniques%, j%, k%
    Dim Num%, NumTech%
    Dim temp As Range
    Dim ID_Thing$, ID_TypeOfImpact$, ID_WayOfRealization$, RelativeTechniques$, RelativeTechniquesVisible$
    Dim IntruderStats() As String
    Dim stemp() As String
    Dim temp_WoR() As String
    Dim temp_ToI() As String
    Dim temp_Objects() As String
    Dim TechniquesElement As Variant
    Dim Element As Variant
    Dim RowInOutputTable As Variant
    Dim IntruderIsExists As Boolean
    Dim y As Long
    
    Base.DeclareTechniques
    Base.DeclareThreats
    
    i = FindEmptyRowInColumn(Sheets("TofTechniques").Cells(2, 2))
    Sheets("TofTechniques").Range("B4:M" + CStr(i)).ClearContents
    Sheets("TofTechniques").Range("B4:M" + CStr(i)).ClearFormats
    
    i = FindEmptyRowInColumn(Sheets("BDU").Cells(2, 3))
    Sheets("BDU").Range("A4:A" + CStr(i)).Value = "Неактуальна"
    
    i = FindEmptyRowInColumn(Sheets("RefTactics").Cells(2, 3))
    Sheets("RefTactics").Range("A4:A" + CStr(i)).Value = "Неактуальна"
    
    
    Call WriteDictionary("DToI", TypesOfImpact, "TypesOfImpact")
    Call WriteDictionary("QTT", Things, "Things", True)
    Call Functions.WriteDictionary("DWoR", WaysOfRealization, "WaysOfRealization")
    

    'Индексация по коллекции
    i = 1
    For Each RowInOutputTable In OutputTable
        'Находятся ID по их текстовым описаниям
        ID_Thing = Functions.FindIDorName(CStr(RowInOutputTable(3)), Base.Things)
        ID_TypeOfImpact = Functions.FindIDorName(CStr(RowInOutputTable(5)), Base.TypesOfImpact)
        ID_WayOfRealization = Functions.FindIDorName(CStr(RowInOutputTable(6)), Base.WaysOfRealization)
        
        stemp = Split(RowInOutputTable(1), "|")
        If TypeName(stemp) <> "String" Then
            ReDim IntruderStats(UBound(stemp), 1)
        Else
            ReDim IntruderStats(0, 1)
        End If
        j = 0
        For Each Element In stemp
            'Объект IntruderStats - двумерный массив, где 0 элмент - категория нарушителя, а 1 - его потенциал
            If Element = "Внешний" Then
                IntruderStats(j, 0) = "Внешний"
            ElseIf Element = "Внутренний" Then
                IntruderStats(j, 0) = "Внутренний"
            End If
            IntruderStats(j, 1) = RowInOutputTable(2)
            j = j + 1
        Next Element
        'Перебираются все техники
        RelativeTechniques = ""
        RelativeTechniquesVisible = ""
        For Each TechniquesElement In Techniques
            'Если для угрозы способ реализации присутствует
            'И присутствует вид воздействия
            'И присутствует объект
            If TechniquesElement.CheckWoR(ID_WayOfRealization) And _
            TechniquesElement.CheckToI(ID_TypeOfImpact) And _
            TechniquesElement.CheckObjects(ID_Thing) Then
                'Выполняется сверка с нарушителем
                IntruderIsExists = False
                For j = 0 To UBound(IntruderStats)
                    For k = 0 To TechniquesElement.NumberOfIntruder
                        'Если категория совпадает
                        'А уровень возможностей, необходимый для техники, <= заявленного
                        If IntruderStats(j, 0) = TechniquesElement.Intruder(k, 0) And _
                        IntruderStats(j, 1) >= TechniquesElement.Intruder(k, 1) Then
                            RelativeTechniquesVisible = RelativeTechniquesVisible + CStr(TechniquesElement.ID_Technique) _
                                + ". " + TechniquesElement.Description_Technique + vbNewLine
                            RelativeTechniques = RelativeTechniques + TechniquesElement.ID_Technique + "|"
                            IntruderIsExists = True
                            Exit For
                        End If
                    Next k
                    If IntruderIsExists Then
                        Exit For
                    End If
                Next j
            End If
        Next TechniquesElement
        'Если техники нашлись, то можно продолжать индексацию
        If RelativeTechniques <> "" Then
            RelativeTechniques = Left(RelativeTechniques, Len(RelativeTechniques) - 1)
            'Расстановка актуальности угроз на листе BDU
            If InStr(RowInOutputTable(7), "|") Then
                For Each Element In Split(RowInOutputTable(7), "|")
                    Threats(CInt(Element) - 1).Relevance = True
                    Sheets("BDU").Cells(CInt(Element) + 3, 1).Value = "Актуальна"
                Next Element
            ElseIf RowInOutputTable(7) <> "" Then
                Threats(CInt(RowInOutputTable(7)) - 1).Relevance = True
                Sheets("BDU").Cells(CInt(RowInOutputTable(7)) + 3, 1).Value = "Актуальна"
            End If
            'Указание актуальных техник
            Call Functions.ChangeElementInArrayOfCollection(OutputTable, i, 9, RelativeTechniques)
            Call Functions.ChangeElementInArrayOfCollection(OutputTable, i, 10, RelativeTechniquesVisible)
            'Расстановка актуальности техник на листе RefTactics
            If InStr(RelativeTechniques, "|") Then
                For Each Element In Split(RelativeTechniques, "|")
                    For Each TechniquesElement In Techniques
                        If Element = TechniquesElement.ID_Technique Then
                            TechniquesElement.Relevance = True
                            Sheets("RefTactics").Cells(TechniquesElement.NumTact + 3, 1).Value = "Актуальна"
                            Exit For
                        End If
                    Next TechniquesElement
                Next Element
            ElseIf RelativeTechniques <> "" Then
                For Each TechniquesElement In Techniques
                    If RelativeTechniques = TechniquesElement.ID_Technique Then
                        TechniquesElement.Relevance = True
                        Sheets("RefTactics").Cells(TechniquesElement.NumTact + 3, 1).Value = "Актуальна"
                        Exit For
                    End If
                Next TechniquesElement
            End If
            i = i + 1
        Else
            OutputTable.Remove i
        End If
    Next RowInOutputTable
    
    Call Functions.ConvertCollectionToMassive(OutputTable, stemp)
    Sheets("TofTechniques").Range("B4:M" + CStr(OutputTable.Count + 3)) = stemp
    Sheets("TofTechniques").ScrollArea = "A1:M" + CStr(OutputTable.Count + 5)
    
    'Настраивается внешний вид
    With Sheets("TofTechniques").Range("B4:M" + CStr(OutputTable.Count + 3))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
End Sub
Sub QIntOfTT_UpdateRefs()
    Call WriteBookOfReferenceFromAuto("QIntOfTT", RefThingsToInterfaces, "RefThingsToInterfaces", , True)
End Sub
Sub DeclareThreats(Optional PurposeOfLaunch As String)
    Dim NumberOfThreats%, i_BDU%, i%
    Dim Element As Variant
    'Определение числа угроз
    NumberOfThreats = FindEmptyRowInColumn(Sheets("BDU").Cells(2, 2)) - 4
    'Переопределение массива угроз
    ReDim Threats(NumberOfThreats - 1)
    'Индексация по массиву
    If PurposeOfLaunch = "Для определения мер защиты" Then
        For i_BDU = 0 To NumberOfThreats - 1
            'Заполнение полей
            Threats(i_BDU).Fill i_BDU + 4, PurposeOfLaunch
        Next i_BDU
    Else
        For i_BDU = 0 To NumberOfThreats - 1
            'Заполнение полей
            Threats(i_BDU).Fill i_BDU + 4
        Next i_BDU
    End If
    Call Functions.SetApplianceColumn("BDU", Sheets("BDU").Cells(2, 1), True, "Актуальна", "Неактуальна")
End Sub
Sub DeclareTechniques()
    Dim NumberOfTechniques%, i_Techniques%, i%
    Dim Element As Variant
    'Определение числа техник
    NumberOfTechniques = FindEmptyRowInColumn(Sheets("RefTactics").Cells(2, 2)) - 4
    'Переопределение массива техник
    ReDim Techniques(NumberOfTechniques - 1)
    'Индексация по массиву
    For i_Techniques = 0 To NumberOfTechniques - 1
        'Заполнение полей
        'Выписывается ID, название и описание
        Techniques(i_Techniques).Fill (i_Techniques + 4)
    Next i_Techniques
    Call Functions.SetApplianceColumn("RefTactics", Sheets("RefTactics").Cells(2, 1), True, "Актуальна", "Неактуальна")
End Sub
Sub ThreatsDesk_Write()
    Application.ScreenUpdating = False
    Call Functions.WriteDictionary("DWoR", WaysOfRealization, "WaysOfRealization")
    Call Functions.WriteDictionary("DToI", TypesOfImpact, "TypesOfImpact")
    Call Functions.WriteDictionary("QTT", Things, "Things")
    ThreatsDeskBusy = True
    
    
    Dim Element As Variant
    Dim Num%, i%
    Dim ID_Threat%
    Dim temp As Range
    
    ID_Threat = CInt(Sheets("ThreatsDesk").Cells(4, 1).Value)
    If ID_Threat < 1 Then
        MsgBox "Номер угрозы слишком мал", , "ERROR: ID_Threat"
    ElseIf ID_Threat > (FindEmptyRowInColumn(Sheets("BDU").Cells(2, 2)) - 4) Then
        MsgBox "Номер угрозы слишком велик", , "ERROR: ID_Threat"
    End If
    
    i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(12, 1)) - 1
    If i < FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(12, 6)) - 1 Then
        i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(12, 6)) - 1
    End If
    If i < FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(12, 11)) - 1 Then
        i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(12, 11)) - 1
    End If
    
    Sheets("ThreatsDesk").Range("A14:N" + CStr(i)).ClearContents
    Sheets("ThreatsDesk").Range("A14:N" + CStr(i)).ClearFormats
    

    Call Functions.DisplayDictionaryOnDesk(WaysOfRealization, Sheets("ThreatsDesk").Cells(12, 1), "ThreatsDesk", "BDU", ID_Threat, 12)
    Call Functions.DisplayDictionaryOnDesk(TypesOfImpact, Sheets("ThreatsDesk").Cells(12, 6), "ThreatsDesk", "BDU", ID_Threat, 13)
    Call Functions.DisplayDictionaryOnDesk(Things, Sheets("ThreatsDesk").Cells(12, 11), "ThreatsDesk", "BDU", ID_Threat, 14)

    
    Call Functions.SetApplianceColumn("ThreatsDesk", Sheets("ThreatsDesk").Cells(12, 4))
    Call Functions.SetApplianceColumn("ThreatsDesk", Sheets("ThreatsDesk").Cells(12, 9))
    Call Functions.SetApplianceColumn("ThreatsDesk", Sheets("ThreatsDesk").Cells(12, 14))
    
    i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(14, 4)) - 1
    'Настраивается внешний вид Способы реализации угроз
    With Sheets("ThreatsDesk").Range("A14:D" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(14, 9)) - 1
    'Настраивается внешний вид Виды воздействия
    With Sheets("ThreatsDesk").Range("F14:I" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    i = FindEmptyRowInColumn(Sheets("ThreatsDesk").Cells(14, 14)) - 1
    'Настраивается внешний вид Объекты
    With Sheets("ThreatsDesk").Range("K14:N" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    ThreatsDeskBusy = False
    Application.ScreenUpdating = True
End Sub
Sub TechniquesDesk_Write()
    Application.ScreenUpdating = False
    Call Functions.WriteDictionary("DWoR", WaysOfRealization, "WaysOfRealization")
    Call Functions.WriteDictionary("DToI", TypesOfImpact, "TypesOfImpact")
    Call Functions.WriteDictionary("QTT", Things, "Things")
    Call Functions.WriteDictionary("DLoC", LevelsOfIntruder, "LevelsOfIntruder")
    Call Functions.WriteDictionary("DCat", CategoriesOfIntruder, "CategoriesOfIntruder")
    
    Dim Element As Variant
    Dim Num%, i%
    Dim ID_Technique%
    Dim temp As Range
    
    ID_Technique = CInt(Sheets("TechniquesDesk").Cells(4, 1).Value)
    If ID_Technique < 1 Then
        MsgBox "Номер тактики слишком мал", , "ERROR: ID_Technique"
    ElseIf ID_Technique > (FindEmptyRowInColumn(Sheets("RefTactics").Cells(2, 2)) - 4) Then
        MsgBox "Номер тактики слишком велик", , "ERROR: ID_Technique"
    End If
    
    i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(12, 1)) - 1
    If i < FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(12, 6)) - 1 Then
        i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(12, 6)) - 1
    End If
    If i < FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(12, 11)) - 1 Then
        i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(12, 11)) - 1
    End If
    
    Sheets("TechniquesDesk").Range("A14:N" + CStr(i)).ClearContents
    Sheets("TechniquesDesk").Range("A14:N" + CStr(i)).ClearFormats
    
    Call Functions.DisplayDictionaryOnDesk(Things, Sheets("TechniquesDesk").Cells(12, 11), "TechniquesDesk", "RefTactics", ID_Technique, 9)
    Call Functions.DisplayDictionaryOnDesk(TypesOfImpact, Sheets("TechniquesDesk").Cells(12, 6), "TechniquesDesk", "RefTactics", ID_Technique, 10)
    Call Functions.DisplayDictionaryOnDesk(WaysOfRealization, Sheets("TechniquesDesk").Cells(12, 1), "TechniquesDesk", "RefTactics", ID_Technique, 11)
    
    With Sheets("TechniquesDesk").Cells(14, 16)
        With .Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Внутренний,Внешний,Внутренний|Внешний"
            .ErrorTitle = "Ошибка"
            .ErrorMessage = "Неверный ввод"
        End With
    End With
    Sheets("TechniquesDesk").Cells(14, 16).Value = Sheets("RefTactics").Cells(ID_Technique + 3, 7).Value
    
    With Sheets("TechniquesDesk").Cells(14, 17)
        With .Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:="Н1,Н2,Н3,Н4"
            .ErrorTitle = "Ошибка"
            .ErrorMessage = "Неверный ввод"
        End With
    End With
    Sheets("TechniquesDesk").Cells(14, 17).Value = Sheets("RefTactics").Cells(ID_Technique + 3, 8).Value
    
    
    Call Functions.SetApplianceColumn("TechniquesDesk", Sheets("TechniquesDesk").Cells(12, 4))
    Call Functions.SetApplianceColumn("TechniquesDesk", Sheets("TechniquesDesk").Cells(12, 9))
    Call Functions.SetApplianceColumn("TechniquesDesk", Sheets("TechniquesDesk").Cells(12, 14))
    
    i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(14, 4)) - 1
    'Настраивается внешний вид Способы реализации угроз
    With Sheets("TechniquesDesk").Range("A14:D" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(14, 9)) - 1
    'Настраивается внешний вид Виды воздействия
    With Sheets("TechniquesDesk").Range("F14:I" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
    
    i = FindEmptyRowInColumn(Sheets("TechniquesDesk").Cells(14, 14)) - 1
    'Настраивается внешний вид Объекты
    With Sheets("TechniquesDesk").Range("K14:N" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
Application.ScreenUpdating = True
End Sub
Sub AThreats_Write()
    Application.ScreenUpdating = False
    
    Dim Element As Variant
    Dim Num%, i%, i_BDU%
    Dim temp As Range
    
    
    i = FindEmptyRowInColumn(Sheets("AThreats").Cells(2, 2))
    
    Sheets("AThreats").Range("B4:G" + CStr(i)).ClearContents
    Sheets("AThreats").Range("B4:G" + CStr(i)).ClearFormats
    
    i = 4
    i_BDU = 4
    Num = 0
    
    Do While Sheets("BDU").Cells(i_BDU, 2) <> "" Or _
    Sheets("BDU").Cells(i_BDU + 1, 2) <> ""
        If Sheets("BDU").Cells(i_BDU + 1, 1) = "Актуальна" Then
             Num = Num + 1
             Sheets("AThreats").Cells(i, 2).Value = Num
             Sheets("AThreats").Cells(i, 3) = Sheets("BDU").Cells(i_BDU + 1, 2)
             Sheets("AThreats").Cells(i, 4) = Sheets("BDU").Cells(i_BDU + 1, 3)
             Sheets("AThreats").Cells(i, 5) = Sheets("BDU").Cells(i_BDU + 1, 4)
             Sheets("AThreats").Cells(i, 6) = Sheets("BDU").Cells(i_BDU + 1, 5)
             Sheets("AThreats").Cells(i, 7) = Sheets("BDU").Cells(i_BDU + 1, 6)
             i = i + 1
        End If
        i_BDU = i_BDU + 1
    Loop
    
    i = FindEmptyRowInColumn(Sheets("AThreats").Cells(2, 2))
    'Настраивается внешний вид
    With Sheets("AThreats").Range("B4:G" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
Application.ScreenUpdating = True
End Sub
Sub ShowAll()
    Dim temp As Boolean
    Dim i%
    For i = 1 To ActiveWorkbook.Sheets.Count
        ChoosingListForm.ListOfLists.AddItem (ActiveWorkbook.Worksheets(i).Cells(1, 1))
        If Not ActiveWorkbook.Worksheets(i).Name = ActiveSheet.Name Then
            ActiveWorkbook.Worksheets(i).Visible = xlSheetVisible
        End If
    Next i
End Sub
Sub CreateOutput()
    Dim FileType$, NewFileName$
    Dim i%, j%
    Dim Element As Variant
    Dim NewBook As Workbook

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each Element In ThisWorkbook.Sheets
        Select Case Element.Name
            Case "QNC"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 4)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 3"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:F" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.Sheets(1).Columns(4).Delete
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "QTT"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 4"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:E" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.Sheets(1).Columns(3).Delete
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "QIntOfTT"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 5"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:G" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "QTTToI"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 6"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:H" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "QCollusion"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 7"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:D" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "TNCGoINoI"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 8"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                Element.Range("B2:G" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "QAoWoR"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 9"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                For j = 0 To 8
                    NewBook.Sheets(1).Columns(1 + j).ColumnWidth = Element.Columns(2 + j).ColumnWidth
                Next j
                Element.Range("B2:J" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "TofThreats"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 10"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                For j = 0 To 10
                    NewBook.Sheets(1).Columns(1 + j).ColumnWidth = Element.Columns(2 + j).ColumnWidth
                Next j
                Element.Range("B2:K" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "TofTechniques"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 11"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                For j = 0 To 12
                    NewBook.Sheets(1).Columns(1 + j).ColumnWidth = Element.Columns(2 + j).ColumnWidth
                Next j
                Element.Range("B2:M" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
            Case "AThreats"
                i = Functions.FindEmptyRowInColumn(Element.Cells(2, 2)) - 1
                Set NewBook = Workbooks.Add
                NewBook.Sheets(1).Name = "Приложение 12"
                If CInt(Left(Application.Version, InStr(1, Application.Version, ".") - 1)) <= 10 Then
                    NewBook.Sheets(2).Delete
                    NewBook.Sheets(2).Delete
                End If
                For j = 0 To 6
                    NewBook.Sheets(1).Columns(1 + j).ColumnWidth = Element.Columns(2 + j).ColumnWidth
                Next j
                Element.Range("B2:G" + CStr(i)).Copy
                NewBook.Sheets(1).Cells(1, 1).PasteSpecial Paste:=xlPasteAll
                NewBook.SaveAs Filename:=ThisWorkbook.Path + "\" + Element.Cells(1, 1).Text, FileFormat:=xlWorkbookNormal, CreateBackup:=False
                NewBook.Close
        End Select
    Next Element
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Beep 900, 150
    Beep 900, 150
    Beep 600, 300
    MsgBox "Выгрузка листов произведена успешно.", , "Сообщение"
    
End Sub
Sub QBasic_UpdateRefs()
    Dim temp As Range
    Dim i%
    
    If Sheets("QBasic").Cells(3, 4).Value = "Да" Then
        Base.ClearQuestionaryOfMeasures = 2
    ElseIf Sheets("QBasic").Cells(3, 4).Value = "Нет" Then
        Base.ClearQuestionaryOfMeasures = 0
    End If
    
    Base.CategoryOfSystem = CInt(Sheets("QBasic").Cells(3, 2).Value)
    Select Case Sheets("QBasic").Cells(3, 3).Value
    Case "Приказ ФСТЭК №239"
        Base.RegulatoryDocumentPage = "Order239"
    Case "Приказ ФСТЭК №31"
        Base.RegulatoryDocumentPage = "Order31"
    Case "Приказ ФСТЭК №21"
        Base.RegulatoryDocumentPage = "Order21"
    Case "Приказ ФСТЭК №17"
        Base.RegulatoryDocumentPage = "Order17"
    End Select
    
    Base.DeclareMeasures
    Call Functions.SetApplianceColumn("QoMfD", Sheets("QoMfD").Cells(2, 6), False, "Да", "Нет", ClearQuestionaryOfMeasures)
    
    i = FindEmptyRowInColumn(Sheets("QoMfD").Cells(2, 2)) - 1
    'Настраивается внешний вид Опросника удаляемых мер
    With Sheets("QoMfD").Range("B4:F" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
    End With
End Sub
Sub DeclareMeasures()
    Dim NumberOfMeasures%, i_Order%, i%, MLastRow%
    Dim Element As Variant
    'Определение числа угроз
    NumberOfMeasures = Functions.CountMeasures(Sheets(RegulatoryDocumentPage).Cells(4, 2), MLastRow)
    'Переопределение массива угроз
    ReDim Measures(NumberOfMeasures - 1)
    
    'Индексация по массиву
    i = 0
    'Индексация по листу (отличается от массива из-за пустых стрк с названиями разделов)
    For i_Order = 4 To MLastRow - 1
        'Если ячейка справа от индексируемой ячейки пуста, то это название раздела мер защиты
        If Sheets(RegulatoryDocumentPage).Cells(i_Order, 3).Value <> "" Then
            Call Measures(i).Fill(RegulatoryDocumentPage, i_Order)
            i = i + 1
        End If
    Next i_Order
End Sub
Sub QoMfD_UpdateRefs()

    Dim i%
    Dim temp As Range
    Dim TempArray() As String
    Dim Element As Variant
    Dim MeasuresElement As Variant
    Dim ThreatsElement As Variant
    Dim Status$
    Dim AddedMeasures As Object
    
    Set temp = Sheets("QoMfA").Range("2:2").Find(What:="№", LookIn:=xlFormulas, Lookat:=xlPart)
    i = 4
    
    Set AddedMeasures = CreateObject("Scripting.Dictionary")
    AddedMeasures.RemoveAll
    
    Base.DeclareThreats ("Для определения мер защиты")
    i = FindDoubleEmptyRowInColumn(Sheets("QoMfA").Cells(3, 2))
    Sheets("QoMfA").Range("B4:E" + CStr(i)).ClearContents
    Sheets("QoMfA").Range("B4:E" + CStr(i)).ClearFormats
    
    i = 4
    For Each ThreatsElement In Threats
        If ThreatsElement.Relevance Then
            If ThreatsElement.MeasuresMoreThanZero(Base.RegulatoryDocumentPage) Then
                For Each MeasuresElement In Measures
                    Status = MeasuresElement.Status(Base.CategoryOfSystem)
                    If Status = "Неактуальна" _
                    And ThreatsElement.CheckMeasure(MeasuresElement.ShortName, Base.RegulatoryDocumentPage) _
                    And Not AddedMeasures.Exists(MeasuresElement.ShortName) Then
                        Sheets("QoMfA").Cells(i, 2).Value = i - 3
                        Sheets("QoMfA").Cells(i, 3).Value = MeasuresElement.ShortName
                        Sheets("QoMfA").Cells(i, 4).Value = MeasuresElement.Description
                        If Not AddedMeasures.Exists(MeasuresElement.ShortName) Then
                            AddedMeasures.Add MeasuresElement.ShortName, 1
                        End If
                        i = i + 1
                    End If
                Next MeasuresElement
            End If
        End If
    Next ThreatsElement
    
    Call Functions.SetApplianceColumn("QoMfA", Sheets("QoMfA").Cells(2, 5), False, "Да", "Нет", Base.ClearQuestionaryOfMeasures)
    
    i = FindDoubleEmptyRowInColumn(Sheets("QoMfA").Cells(2, 2)) - 1
    'Настраивается внешний вид Опросника удаляемых мер
    With Sheets("QoMfA").Range("B4:E" + CStr(i))
        .Borders.LineStyle = True
        .WrapText = True
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignTop
        .Font.Name = "Times New Roman"
        .Font.Size = 10
    End With
End Sub
Sub QoMfA_UpdateRefs()

    Dim DeletedMeasures As Object
    Set DeletedMeasures = CreateObject("Scripting.Dictionary")
    DeletedMeasures.RemoveAll
    
    Dim AddedMeasures As Object
    Set AddedMeasures = CreateObject("Scripting.Dictionary")
    AddedMeasures.RemoveAll
        
    Dim i%, IDM%, iAM%, i_m%, z_m%, CategoryColumn%
    Dim ZeroMeasures As Boolean
    Dim temp As Range
    Dim TempArray() As String
    Dim Element As Variant
    Dim MeasuresElement As Variant
    Dim ThreatsElement As Variant
    Dim MeasuresWithConflicts$, Status$, letter$
    
    Set temp = Sheets("QoMfD").Range("2:2").Find(What:="№", LookIn:=xlFormulas, Lookat:=xlPart)
    i = 4
    
    Do While Sheets("QoMfD").Cells(i, temp.Column + 1).Value <> ""
        If Sheets("QoMfD").Cells(i, temp.Column + 4).Value = "Да" Then
            TempArray = Split(Sheets("QoMfD").Cells(i, temp.Column + 1), "|")
            For Each Element In TempArray
                If Not DeletedMeasures.Exists(CStr(Element)) Then
                    DeletedMeasures.Add CStr(Element), Sheets("QoMfD").Cells(i, temp.Column + 2).Value
                End If
            Next Element
        End If
        i = i + 1
    Loop
    
    Set temp = Sheets("QoMfA").Range("2:2").Find(What:="№", LookIn:=xlFormulas, Lookat:=xlPart)
    i = 4
    
    Do While Sheets("QoMfA").Cells(i, temp.Column + 1).Value <> ""
        If Sheets("QoMfA").Cells(i, temp.Column + 3).Value = "Да" Then
            TempArray = Split(Sheets("QoMfA").Cells(i, temp.Column + 1), "|")
            For Each Element In TempArray
                If Not AddedMeasures.Exists(CStr(Element)) Then
                    If Not DeletedMeasures.Exists(CStr(Element)) Then
                        AddedMeasures.Add CStr(Element), 1
                    Else
                        MeasuresWithConflicts = MeasuresWithConflicts + CStr(Element) + ";"
                    End If
                End If
            Next Element
        End If
        i = i + 1
    Loop
    
    
    
    If MeasuresWithConflicts <> "" Then
        MsgBox "Следующие меры одновременно удалены и добавлены: [" + Left(MeasuresWithConflicts, Len(MeasuresWithConflicts) - 1) + "]" + Chr(10) _
        + "Рекомендуется вернуться назад и проверить введённые данные. Дальнейшая работа логически невозможна.", , "ERROR"
        Buttons.QoMfA_Back
    Else
        IDM = 1
        iAM = 1

        i = FindDoubleEmptyRowInColumn(Sheets("DMeasures").Cells(3, 2))
        Sheets("DMeasures").Range("B4:D" + CStr(i)).ClearContents
        Sheets("DMeasures").Range("B4:D" + CStr(i)).ClearFormats
        i = FindDoubleEmptyRowInColumn(Sheets("AMeasures").Cells(3, 2))
        Sheets("AMeasures").Range("B4:C" + CStr(i)).ClearContents
        Sheets("AMeasures").Range("B4:C" + CStr(i)).ClearFormats
        
        For Each MeasuresElement In Measures
            If DeletedMeasures.Exists(MeasuresElement.ShortName) Then
                MeasuresElement.mIsDeleted = True
                MeasuresElement.mIsAdded = False
                If MeasuresElement.Status(Base.CategoryOfSystem) = "Удалена" Then
                    Sheets("DMeasures").Cells(3 + IDM, 2).Value = IDM
                    Sheets("DMeasures").Cells(3 + IDM, 3).Value = MeasuresElement.ShortName + " — " + MeasuresElement.Description
                    Sheets("DMeasures").Cells(3 + IDM, 4).Value = DeletedMeasures.Item(MeasuresElement.ShortName)
                    IDM = IDM + 1
                End If
            ElseIf AddedMeasures.Exists(MeasuresElement.ShortName) Then
                MeasuresElement.mIsAdded = True
                MeasuresElement.mIsDeleted = False
                If MeasuresElement.Status(Base.CategoryOfSystem) = "Добавлена" Then
                    Sheets("AMeasures").Cells(3 + iAM, 2).Value = iAM
                    Sheets("AMeasures").Cells(3 + iAM, 3).Value = MeasuresElement.ShortName + " — " + MeasuresElement.Description
                    iAM = iAM + 1
                End If
            End If
        
        Next MeasuresElement
        i = FindDoubleEmptyRowInColumn(Sheets("DMeasures").Cells(3, 2)) - 1
        Sheets("DMeasures").Range("B4:D" + CStr(i)).Borders.LineStyle = xlContinuous
        i = FindDoubleEmptyRowInColumn(Sheets("AMeasures").Cells(3, 2)) - 1
        Sheets("AMeasures").Range("B4:C" + CStr(i)).Borders.LineStyle = xlContinuous
        
        'Заполнение базовых мер
        i = FindDoubleEmptyRowInColumn(Sheets("BasicMeasures").Cells(3, 2))
        Sheets("BasicMeasures").Range("B4:D" + CStr(i)).ClearContents
        Sheets("BasicMeasures").Range("B4:D" + CStr(i)).ClearFormats
        
'        Индексация по листу приказа ФСТЭК
        i = 4
'        Индексацая по листу с мерами
        i_m = 4
'        Номер меры
        z_m = 0
        ZeroMeasures = True
        
        'Если приказ 21, то там 4 категории
        If RegulatoryDocumentPage = "Order21" Then
            CategoryColumn = 8
        Else
            CategoryColumn = 7
        End If
        
        Do
            'Проверка на то, есть ли мера в классе или это название раздела мер
            If Sheets(Base.RegulatoryDocumentPage).Cells(i, CategoryColumn - Base.CategoryOfSystem).Value = "+" _
            Or IsEmpty(Sheets(Base.RegulatoryDocumentPage).Cells(i, 3)) Then
                If IsEmpty(Sheets(Base.RegulatoryDocumentPage).Cells(i, 3)) Then
                    If ZeroMeasures Then 'Нужно для удаления пустого раздела мер
                        Sheets("BasicMeasures").Cells(i_m, 2).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value
                        i_m = i_m + 1
                        z_m = z_m + 1
                        ZeroMeasures = False
                    Else
                        Sheets("BasicMeasures").Cells(i_m - 1, 2).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value
                    End If
                Else
                    Sheets("BasicMeasures").Cells(i_m, 2).Value = CStr(i_m - 3 - z_m)
                    Sheets("BasicMeasures").Cells(i_m, 3).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value
                    Sheets("BasicMeasures").Cells(i_m, 4).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 3).Value
                    i_m = i_m + 1
                    ZeroMeasures = True
                End If
            End If
            i = i + 1
        Loop While Not IsEmpty(Sheets(Base.RegulatoryDocumentPage).Cells(i, 2))
        
        i = FindDoubleEmptyRowInColumn(Sheets("BasicMeasures").Cells(3, 2)) - 1
        Sheets("BasicMeasures").Range("B4:D" + CStr(i)).Borders.LineStyle = xlContinuous
        
        i = FindDoubleEmptyRowInColumn(Sheets("ResultMeasures").Cells(3, 2))
        Sheets("ResultMeasures").Range("B4:D" + CStr(i)).ClearContents
        Sheets("ResultMeasures").Range("B4:D" + CStr(i)).ClearFormats
        
        
        Select Case RegulatoryDocumentPage:
            Case "Order239"
                letter = ":G"
            Case "Order31"
                letter = ":G"
            Case "Order21"
                letter = ":H"
            Case "Order17"
                letter = ":G"
        End Select
        
        i = FindDoubleEmptyRowInColumn(Sheets(Base.RegulatoryDocumentPage).Cells(4, 2))
        Sheets(Base.RegulatoryDocumentPage).Range("B4" + letter + CStr(i)).Interior.ColorIndex = 2
        
        
'        Индексация по листу приказа ФСТЭК
        i = 4
'        Индексацая по листу с мерами
        i_m = 4
'        Номер меры
        z_m = 0
        ZeroMeasures = True
        Do
            If IsEmpty(Sheets(Base.RegulatoryDocumentPage).Cells(i, 3)) Then
                If ZeroMeasures Then 'Нужно для удаления пустого раздела мер (см 3 раздел мер)
                    Sheets("ResultMeasures").Cells(i_m, 2).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value
                    i_m = i_m + 1
                    z_m = z_m + 1
                    ZeroMeasures = False
                Else
                    Sheets("ResultMeasures").Cells(i_m - 1, 2).Value = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value
                End If
            Else
                For Each MeasuresElement In Measures
                    If MeasuresElement.ShortName = Sheets(Base.RegulatoryDocumentPage).Cells(i, 2).Value Then
                        Status = MeasuresElement.Status(Base.CategoryOfSystem)
                        If Status = "Актуальна" Then
                            Sheets("ResultMeasures").Cells(i_m, 2).Value = CStr(i_m - 3 - z_m)
                            Sheets("ResultMeasures").Cells(i_m, 3).Value = MeasuresElement.ShortName
                            Sheets("ResultMeasures").Cells(i_m, 4).Value = MeasuresElement.Description
                            i_m = i_m + 1
                            ZeroMeasures = True
                            Exit For
                        ElseIf Status = "Добавлена" Then
                            Sheets("ResultMeasures").Cells(i_m, 2).Value = CStr(i_m - 3 - z_m)
                            Sheets("ResultMeasures").Cells(i_m, 3).Value = MeasuresElement.ShortName
                            Sheets("ResultMeasures").Cells(i_m, 4).Value = MeasuresElement.Description
                            Sheets(Base.RegulatoryDocumentPage).Range("B" + CStr(i) + letter + CStr(i)).Interior.ColorIndex = 37
                            i_m = i_m + 1
                            ZeroMeasures = True
                            Exit For
                        ElseIf Status = "Удалена" Then
                            Sheets(Base.RegulatoryDocumentPage).Range("B" + CStr(i) + letter + CStr(i)).Interior.ColorIndex = 22
                            Exit For
                        End If
                    End If
                Next MeasuresElement
            End If
            i = i + 1
        Loop While Not IsEmpty(Sheets(Base.RegulatoryDocumentPage).Cells(i, 2))
        
        i = FindDoubleEmptyRowInColumn(Sheets("ResultMeasures").Cells(3, 2)) - 1
        Sheets("ResultMeasures").Range("B4:D" + CStr(i)).Borders.LineStyle = xlContinuous
        
        i = FindDoubleEmptyRowInColumn(Sheets("LoTaM").Cells(3, 5))
        Sheets("LoTaM").Range("B4:E" + CStr(i)).ClearContents
        Sheets("LoTaM").Range("B4:E" + CStr(i)).ClearFormats
        
        
'        Индексация по листу угроз и мер
        i = 4
'        Номер угрозы
        i_m = 1
        For Each ThreatsElement In Threats
            If ThreatsElement.Relevance Then
                Sheets("LoTaM").Cells(i, 2).Value = i_m
                Sheets("LoTaM").Cells(i, 3).Value = ThreatsElement.ID
                Sheets("LoTaM").Cells(i, 4).Value = ThreatsElement.Name
                If ThreatsElement.MeasuresMoreThanZero(Base.RegulatoryDocumentPage) Then
                    ZeroMeasures = True
                    For Each MeasuresElement In Measures
                        Status = MeasuresElement.Status(Base.CategoryOfSystem)
                        If (Status = "Актуальна" Or Status = "Добавлена") _
                        And ThreatsElement.CheckMeasure(MeasuresElement.ShortName, Base.RegulatoryDocumentPage) Then
                            Sheets("LoTaM").Cells(i, 5).Value = MeasuresElement.ShortName + " — " + MeasuresElement.Description
                            If Status = "Добавлена" Then
                                Sheets("LoTaM").Cells(i, 5).Font.Bold = True
                            End If
                            i = i + 1
                            ZeroMeasures = False
                        End If
                    Next MeasuresElement
                    If ZeroMeasures = True Then
                        Sheets("LoTaM").Cells(i, 5).Value = "Среди указанных компенсирующих мер нет мер, присутствующих в выбранном приказе ФСТЭК"
                        Sheets("LoTaM").Cells(i, 5).Font.Bold = True
                        Sheets("LoTaM").Cells(i, 5).Font.Color = RGB(255, 0, 0)
                    End If
                Else
                    Sheets("LoTaM").Cells(i, 5).Value = "EMPTY"
                    i = i + 1
                End If
                i_m = i_m + 1
            End If
        Next ThreatsElement
        i = FindDoubleEmptyRowInColumn(Sheets("LoTaM").Cells(3, 5)) - 1
        If Sheets("LoTaM").Cells(i + 1, 4) <> "" Then
            i = i + 1
        End If
        Sheets("LoTaM").Range("B4:E" + CStr(i)).Borders.LineStyle = xlContinuous
    End If
    
    
End Sub
Sub ExtractMeasuresForThreats()
    Dim launch As Boolean
    Dim i As Integer
    Dim myExcel As New Excel.Application
    Dim MeasureBook As Excel.Workbook
    Dim CurrentBook As Excel.Workbook
    Dim ImpotingFileName$, ThreatID$
    Dim Measures31 As Object
    Set Measures31 = CreateObject("Scripting.Dictionary")
    Measures31.RemoveAll
    Dim Measures239 As Object
    Set Measures239 = CreateObject("Scripting.Dictionary")
    Measures239.RemoveAll
    Dim MeasuresIS As Object
    Set MeasuresIS = CreateObject("Scripting.Dictionary")
    MeasuresIS.RemoveAll
    Dim IDM As Object
    Set IDM = CreateObject("Scripting.Dictionary")
    IDM.RemoveAll
    Dim IDMElement As Variant
    Dim fDialog As FileDialog
    Dim temp As Range
    
    Application.ScreenUpdating = False
    
SelectingFile:
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    fDialog.Title = "Выберите опросник для импорта модели угроз"
    fDialog.InitialFileName = ActiveWorkbook.Path
    fDialog.Filters.Clear
    fDialog.Filters.Add "Книга Excell с поддержкой макросов(*.xlsm)", "*.xlsm, *.xlsx"
    
    If fDialog.Show = -1 Then
        ImpotingFileName = fDialog.SelectedItems(1) 'The full path to the file selected by the user
    Else
        MsgBox ("Ошибка в пути к Перечню мер ИБ для БДУ")
        GoTo SelectingFile
    End If

    Set CurrentBook = ActiveWorkbook
    Set MeasureBook = myExcel.Workbooks.Open(ImpotingFileName)
    
    'Проверка наличия листа с моделью угроз
    If Not WorksheetExists("Меры защиты", MeasureBook) Then
        MsgBox "Выбран неверный файл для импорта (нет страницы Меры защиты). Проверься на IQ довен.", , "ERROR"
        FileChoosingForm.Show
        If Base.ContinueExraction Then
            GoTo SelectingFile
        Else
            GoTo StopExtracting
        End If
    End If
    
    i = 2
    'Проходит по листу, пока не появитс пустая ячейка (или необъединеная)
    Do While MeasureBook.Worksheets("Меры защиты").Cells(i, 2).Value <> "" _
    Or MeasureBook.Worksheets("Меры защиты").Cells(i, 2).MergeCells
        'Обновляет значения Индекса меры только по непустым ячейкам
        If MeasureBook.Worksheets("Меры защиты").Cells(i, 2).Text <> "" And _
        ThreatID <> MeasureBook.Worksheets("Меры защиты").Cells(i, 2).Text Then
            ThreatID = MeasureBook.Worksheets("Меры защиты").Cells(i, 2).Text
            Set temp = CurrentBook.Worksheets("BDU").Range("B:B").Find(What:=CStr(ThreatID), LookIn:=xlFormulas, Lookat:=xlWhole)
            Call Functions.WriteMeasuresDictionary(MeasureBook.Worksheets("Меры защиты"), MeasureBook.Worksheets("УБИ к BIOS"), Measures239, 10, 2, i)
            CurrentBook.Worksheets("BDU").Cells(temp.Row, 16).Value = Functions.ExtractMeasuresFromDictionary(Measures239)
            Call Functions.WriteMeasuresDictionary(MeasureBook.Worksheets("Меры защиты"), MeasureBook.Worksheets("УБИ к BIOS"), Measures31, 11, 2, i)
            CurrentBook.Worksheets("BDU").Cells(temp.Row, 15).Value = Functions.ExtractMeasuresFromDictionary(Measures31)
            Call Functions.WriteMeasuresDictionary(MeasureBook.Worksheets("Меры защиты"), MeasureBook.Worksheets("УБИ к BIOS"), MeasuresIS, 12, 2, i)
            CurrentBook.Worksheets("BDU").Cells(temp.Row, 17).Value = Functions.ExtractMeasuresFromDictionary(MeasuresIS)
        End If
        i = i + 1
    Loop
    
    'Вписываются в словарь меры из Внутренних Мер Защиты
    i = 3
    Do While MeasureBook.Worksheets(" (ВМЗ.х) Внутренние меры защиты").Cells(i, 1).Value <> ""
        If Not IDM.Exists(MeasureBook.Worksheets(" (ВМЗ.х) Внутренние меры защиты").Cells(i, 1).Value) Then
            IDM.Add MeasureBook.Worksheets(" (ВМЗ.х) Внутренние меры защиты").Cells(i, 1).Value, MeasureBook.Worksheets(" (ВМЗ.х) Внутренние меры защиты").Cells(i, 2).Value
        End If
        i = i + 1
    Loop
    
    Call Functions.RewriteOrderList("Order239", CurrentBook, IDM)
    Call Functions.RewriteOrderList("Order31", CurrentBook, IDM)
    Call Functions.RewriteOrderList("Order21", CurrentBook, IDM)
    Call Functions.RewriteOrderList("Order17", CurrentBook, IDM)
    
StopExtracting:
    MeasureBook.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox ("Заполнение завершено")
End Sub
Sub CreateThreatsForAct()
    Dim i%
    Dim ActualThreats$, NonactualThreats$
    
    i = 4
    Do While ActiveWorkbook.Worksheets("BDU").Cells(i, 2).Text <> ""
        If ActiveWorkbook.Worksheets("BDU").Cells(i, 1).Text = "Актуальна" Then
            ActualThreats = ActualThreats + "УБИ." + ActiveWorkbook.Worksheets("BDU").Cells(i, 2).Text _
            + " " + ActiveWorkbook.Worksheets("BDU").Cells(i, 3).Text + "," + vbNewLine
        Else
            NonactualThreats = NonactualThreats + "УБИ." + ActiveWorkbook.Worksheets("BDU").Cells(i, 2).Text _
            + " " + ActiveWorkbook.Worksheets("BDU").Cells(i, 3).Text + "," + vbNewLine
        End If
        i = i + 1
    Loop
    
    ActiveWorkbook.Worksheets("ThreatsForAct").Cells(4, 2).Value = ActualThreats
    ActiveWorkbook.Worksheets("ThreatsForAct").Cells(4, 3).Value = NonactualThreats
    
End Sub
Sub Measures_Show()
    Dim i%
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Not ActiveWorkbook.Worksheets(i).Name = "QoMfA" _
        And ChoosingListForm.HideSheet.Value <> True Then
            ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
        End If
    Next i

    ActiveWorkbook.Worksheets("DMeasures").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("AMeasures").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("ResultMeasures").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("BasicMeasures").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("LoTaM").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets(Base.RegulatoryDocumentPage).Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("QoMfA").Visible = xlSheetHidden
    
End Sub
