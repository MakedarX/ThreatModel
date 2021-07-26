VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Measure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mShortName$, mDescription$, mStatus$, mDeleteReason$
Public LowestCategory%
Public mIsUsing, mIsDeleted, mIsAdded As Boolean

Public Property Get IsApplicably(CheckingVar As Integer) As Boolean
    IsApplicably = False
    If CheckingVar <= LowestCategory Then
        IsApplicably = True
    End If
End Property
Public Property Get ShortName() As String
    ShortName = mShortName
End Property
Public Property Get Description() As String
    Description = mDescription
End Property
Public Property Get Status(CurrentCategory As Integer) As String
    If CurrentCategory <= LowestCategory Then
        Status = "Актуальна"
        If mIsDeleted Then
            Status = "Удалена"
        End If
    ElseIf mIsAdded Then
        Status = "Добавлена"
    Else
        Status = "Неактуальна"
    End If
End Property
Sub Fill(SheetName As String, i As Integer, _
Optional FillingType As String, Optional DebugNeeded As Boolean = False)
    Dim temp() As String
    
    If SheetName <> "Order21" Then
        With Sheets(SheetName)
            'Выписывается название и описание
            mShortName = .Cells(i, 2).Value
            mDescription = .Cells(i, 3).Value
            'Определяется минимальная категория
            If .Cells(i, 4).Value = "+" Then
                LowestCategory = 3
            ElseIf .Cells(i, 5).Value = "+" Then
                LowestCategory = 2
            ElseIf .Cells(i, 6).Value = "+" Then
                LowestCategory = 1
            Else
                LowestCategory = 0
                If DebugNeeded = True Then
                    MsgBox "Мера " + ShortName + " не имеет категории", , "ERROR"
                End If
            End If
            
            If .Cells(i, 7).Value = "+" Then
                mIsUsing = True
            Else
                mIsUsing = False
            End If
        End With
    Else
        With Sheets(SheetName)
            'Выписывается название и описание
            mShortName = .Cells(i, 2).Value
            mDescription = .Cells(i, 3).Value
            'Определяется минимальная категория
            If .Cells(i, 4).Value = "+" Then
                LowestCategory = 4
            ElseIf .Cells(i, 5).Value = "+" Then
                LowestCategory = 3
            ElseIf .Cells(i, 6).Value = "+" Then
                LowestCategory = 2
            ElseIf .Cells(i, 7).Value = "+" Then
                LowestCategory = 1
            Else
                LowestCategory = 0
                If DebugNeeded = True Then
                    MsgBox "Мера " + ShortName + " не имеет категории", , "ERROR"
                End If
            End If
            
            If .Cells(i, 8).Value = "+" Then
                mIsUsing = True
            Else
                mIsUsing = False
            End If
        End With
    End If
End Sub

