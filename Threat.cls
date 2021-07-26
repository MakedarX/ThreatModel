VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Threat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tID As Integer
Public Name$, Description$
Private tIntruder() As String
Public Conftidentiality As Boolean, Integrity As Boolean, Availability As Boolean
Public CreationDate As Date
Public UpdateDate As Date
Private tWoR() As String
Private tToI() As String
Private tObjects() As String
Private tMeasures31() As String
Private tMeasures239() As String
Private tMeasuresIS() As String
Public Relevance As Boolean

Public Property Get CheckWoR(CheckingVar As Variant) As Boolean
    CheckWoR = False
    Dim Element As Variant
    For Each Element In tWoR
        If Element = CheckingVar Then
            CheckWoR = True
            Exit For
        End If
    Next Element
End Property
Public Property Get CheckToI(CheckingVar As Variant) As Boolean
    CheckToI = False
    Dim Element As Variant
    For Each Element In tToI
        If Element = CheckingVar Then
            CheckToI = True
            Exit For
        End If
    Next Element
End Property
Public Property Get CheckObjects(CheckingVar As Variant) As Boolean
    CheckObjects = False
    Dim Element As Variant
    For Each Element In tObjects
        If Element = CheckingVar Then
            CheckObjects = True
            Exit For
        End If
    Next Element
End Property
Public Property Get CheckMeasure(CheckingVar As Variant, RegulatoryDocument As String) As Boolean
    CheckMeasure = False
    Dim Element As Variant
    Dim tMeasures() As String
    Select Case RegulatoryDocument:
        Case "Order239"
            tMeasures = tMeasures239
        Case "Order31"
            tMeasures = tMeasures31
        Case "Order21"
            tMeasures = tMeasuresIS
        Case "Order17"
            tMeasures = tMeasuresIS
        Case Else
            tMeasures = tMeasuresIS
    End Select
    For Each Element In tMeasures
        If Element = CheckingVar Then
            CheckMeasure = True
            Exit For
        End If
    Next Element
End Property
Public Property Get MeasuresMoreThanZero(RegulatoryDocument As String) As Boolean

    Dim tMeasures() As String
    Select Case RegulatoryDocument:
        Case "Order239"
            tMeasures = tMeasures239
        Case "Order31"
            tMeasures = tMeasures31
        Case "Order21"
            tMeasures = tMeasuresIS
        Case "Order17"
            tMeasures = tMeasuresIS
        Case Else
            tMeasures = tMeasuresIS
    End Select
    If UBound(tMeasures) > 1 Then
        MeasuresMoreThanZero = True
    ElseIf tMeasures(0) <> "EMPTY" Then
        MeasuresMoreThanZero = True
    Else
        MeasuresMoreThanZero = False
    End If
End Property
Public Property Get NumberOfIntruder() As Integer
    NumberOfIntruder = UBound(tIntruder, 1)
End Property
Public Property Get Intruder(ByVal i As Integer, ByVal j As Integer) As String
    Intruder = tIntruder(i, j)
End Property
Public Property Get ID() As String
    ID = tID
End Property
Sub Fill(ByVal i As Integer, Optional ByVal FillingType As String)
    Dim temp() As String
    Dim Element As Variant
    Dim j%
    Relevance = False
    
    With Sheets("BDU")
        'Если это не запуск для заполнения модели угроз, то мы можем использовать значения актуальность в качестве актуальных референсов
        If FillingType = "Для определения мер защиты" Then
            If .Cells(i, 1).Value = "Неактуальна" Then
                Relevance = False
            ElseIf .Cells(i, 1).Value = "Актуальна" Then
                Relevance = True
            End If
        'В противном случае обнуляем актуальности
        Else
            .Cells(i, 1).Value = "Неактуальна"
        End If
        
        'Выписывается tID, название и описание
        tID = CInt(.Cells(i, 2).Value)
        Name = .Cells(i, 3).Value
        Description = .Cells(i, 4).Value
        'Строка с источником разбивается
        temp = Split(.Cells(i, 5).Value, ";")
        'Переопределяется массив нарушителей (неизвестно сколько нарушителей будет)
        ReDim tIntruder(UBound(temp), 1)
        j = 0
        For Each Element In temp
            'Объект tIntruder - двумерный массив, где 0 элмент - категория нарушителя, а 1 - его потенциал
            If InStr(Element, "Внешний") <> 0 Then
                tIntruder(j, 0) = "Внешний"
            ElseIf InStr(Element, "Внутренний") <> 0 Then
                tIntruder(j, 0) = "Внутренний"
            End If
            If InStr(Element, "высоким") <> 0 Then
                tIntruder(j, 1) = "Н4"
            ElseIf InStr(Element, "средним") <> 0 Then
                tIntruder(j, 1) = "Н3"
            ElseIf InStr(Element, "низким") <> 0 Then
                tIntruder(j, 1) = "Н1"
            End If
            j = j + 1
        Next Element
        'Проверка последствий
        If .Cells(i, 7).Value = 1 Then
            Conftidentiality = True
        Else
            Conftidentiality = False
        End If
        If .Cells(i, 8).Value = 1 Then
            Integrity = True
        Else
            Integrity = False
        End If
            If .Cells(i, 9).Value = 1 Then
            Availability = True
        Else
            Availability = False
        End If
        'Записываются даты
        CreationDate = CDate(.Cells(i, 10).Value)
        UpdateDate = CDate(.Cells(i, 11).Value)
        'Строка со способами реализации разбивается в массив
        tWoR = Split(.Cells(i, 12).Value, "|")
        'Строка с видами воздействий разбивается в массив
        tToI = Split(.Cells(i, 13).Value, "|")
        'Строка с объектами разбивается в массив
        tObjects = Split(.Cells(i, 14).Value, "|")
        'Строка с мерами по 31 приказу разбивается в массив
        tMeasures31 = Split(.Cells(i, 15).Value, "|")
        'Строка с мерами по 239 приказу разбивается в массив
        tMeasures239 = Split(.Cells(i, 16).Value, "|")
        'Строка с мерами для ИС разбивается в массив
        tMeasuresIS = Split(.Cells(i, 17).Value, "|")
    End With
End Sub
