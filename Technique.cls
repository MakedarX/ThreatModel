VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Technique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public NumTact As Integer
Public ID_Technique$, Description_Technique$, ID_Tactic$, Description_Tactic$
Private tIntruder() As String
Private tWoR() As String
Private tToI() As String
Private tObjects() As String
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
Public Property Get NumberOfIntruder() As Integer
    NumberOfIntruder = UBound(tIntruder, 1)
End Property
Public Property Get Intruder(ByVal i As Integer, ByVal j As Integer) As String
    Intruder = tIntruder(i, j)
End Property
Sub Fill(ByVal i As Integer)
    Dim temp() As String
    Dim Element As Variant
    Dim j%
    Relevance = False
    
    With Sheets("RefTactics")
        .Cells(i, 1).Value = "Неактуальна"
        'Выписывается NumTact, ID техники, её описание, ID тактики, её описание
        NumTact = CInt(.Cells(i, 2).Value)
        ID_Technique = .Cells(i, 3).Value
        Description_Technique = .Cells(i, 4).Value
        ID_Tactic = .Cells(i, 5).Value
        Description_Tactic = .Cells(i, 6).Value
        If .Cells(i, 7).Value <> "" Then
            'Строка с категорией разбивается (уровень будет один, а вот категорий может быть и две)
            temp = Split(.Cells(i, 7).Value, "|")
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
                tIntruder(j, 1) = .Cells(i, 8).Value
                j = j + 1
            Next Element
        Else
            ReDim tIntruder(0, 1)
            tIntruder(0, 0) = "Не задано"
            tIntruder(0, 1) = "Не задано"
        End If
        If .Cells(i, 9).Value <> "" Then
            'Строка с объектами разбивается в массив
            tObjects = Split(.Cells(i, 9).Value, "|")
        Else
            ReDim tObjects(0)
            tObjects(0) = "Не задано"
        End If
        If .Cells(i, 10).Value <> "" Then
            'Строка с видами воздействий разбивается в массив
            tToI = Split(.Cells(i, 10).Value, "|")
        Else
            ReDim tToI(0)
            tToI(0) = "Не задано"
        End If
        If .Cells(i, 11).Value <> "" Then
            'Строка со способами реализации разбивается в массив
            tWoR = Split(.Cells(i, 11).Value, "|")
        Else
            ReDim tWoR(0)
            tWoR(0) = "Не задано"
        End If
    End With
End Sub

