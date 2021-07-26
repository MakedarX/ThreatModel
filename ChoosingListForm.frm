VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChoosingListForm 
   Caption         =   "Выбор листа"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8130
   OleObjectBlob   =   "ChoosingListForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChoosingListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lists As Object
Private Sub ChooseList_Click()
    Dim i%
    i% = Lists.Item(ChoosingListForm.ListOfLists.Value)
    ActiveWorkbook.Worksheets(i).Visible = xlSheetVisible
    ActiveWorkbook.Worksheets(i).Activate
    ChoosingListForm.Hide
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Not ActiveWorkbook.Worksheets(i).Name = ActiveSheet.Name _
        And ChoosingListForm.HideSheet.Value <> True Then
            ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
        End If
    Next i
    ChoosingListForm.ListOfLists.Enabled = False
    ChooseList.Enabled = False
End Sub

Private Sub ShowMeasuresLists_Click()
    Dim i As Integer
    Dim Element As Variant
    Set Lists = CreateObject("Scripting.Dictionary")

    ReDim temp(ActiveWorkbook.Sheets.Count)

    For i = 1 To ActiveWorkbook.Sheets.Count
        If InStr(1, ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, "»") <> 0 Then
            If Lists.Exists(ActiveWorkbook.Worksheets(i).Cells(1, 1).Value) Then
                Lists.Add CStr(i) + "x" + ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, i
            Else
                Lists.Add ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, i
            End If
        End If
    Next i

    Call Functions.SortDict(Lists)
    
    ChoosingListForm.ListOfLists.Clear
    
    For Each Element In Lists
        ChoosingListForm.ListOfLists.AddItem (Element)
    Next Element
    
    ChoosingListForm.ListOfLists.Enabled = True
    ChooseList.Enabled = True
End Sub

Private Sub ShowMULists_Click()
    Dim i As Integer
    Dim Element As Variant
    Set Lists = CreateObject("Scripting.Dictionary")

    ReDim temp(ActiveWorkbook.Sheets.Count)

    For i = 1 To ActiveWorkbook.Sheets.Count
        If InStr(1, ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, "»") = 0 Then
            If Lists.Exists(ActiveWorkbook.Worksheets(i).Cells(1, 1).Value) Then
                Lists.Add CStr(i) + "x" + ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, i
            Else
                Lists.Add ActiveWorkbook.Worksheets(i).Cells(1, 1).Value, i
            End If
        End If
    Next i

    Call Functions.SortDict(Lists)
    
    ChoosingListForm.ListOfLists.Clear
    
    For Each Element In Lists
        ChoosingListForm.ListOfLists.AddItem (Element)
    Next Element
    ChoosingListForm.ListOfLists.Enabled = True
    ChooseList.Enabled = True
End Sub
