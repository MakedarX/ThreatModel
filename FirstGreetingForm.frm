VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FirstGreetingForm 
   Caption         =   "Окно приветствия"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5040
   OleObjectBlob   =   "FirstGreetingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FirstGreetingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DeveloperMode_Click()
    FirstGreetingForm.Hide
    TypeOfWorkmode = "Разработчик"
End Sub
Private Sub StartupType_Change()
    Select Case FirstGreetingForm.StartupType.Value:
        Case "Построение модели угроз"
            FirstGreetingForm.UserMode.Enabled = True
        Case "Построение мер защиты по модели угроз"
            FirstGreetingForm.UserMode.Enabled = True
        Case "Отображение результата построения мер защиты"
            FirstGreetingForm.UserMode.Enabled = True
        Case "Отображение приложений модели угроз"
            FirstGreetingForm.UserMode.Enabled = True
        Case Else
            FirstGreetingForm.UserMode.Enabled = False
    End Select
End Sub

Private Sub UserMode_Click()
    Application.ScreenUpdating = False
    TypeOfWorkmode = "Пользователь"
    Dim i%
    Select Case FirstGreetingForm.StartupType.Value:
        Case "Построение модели угроз"
            ActiveWorkbook.Worksheets("QNC").Visible = xlSheetVisible
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Worksheets(i).Name <> "QNC" Then
                    ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
                End If
            Next i
            FirstGreetingForm.Hide
        Case "Построение мер защиты по модели угроз"
            ActiveWorkbook.Worksheets("QBasic").Visible = xlSheetVisible
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Worksheets(i).Name <> "QBasic" Then
                    ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
                End If
            Next i
            FirstGreetingForm.Hide
            
            TypeOfWorkmode = "Построение мер защиты"
            
            With ActiveWorkbook.Worksheets("QBasic").Cells(3, 3)
                With .Validation
                    .Delete
                    .Add Type:=xlValidateList, Formula1:="Приказ ФСТЭК №17,Приказ ФСТЭК №21,Приказ ФСТЭК №31,Приказ ФСТЭК №239"
                    .ErrorTitle = "Ошибка"
                    .ErrorMessage = "Неверный ввод"
                End With
            End With
            With ActiveWorkbook.Worksheets("QBasic").Cells(3, 2)
                With .Validation
                    .Delete
                    .Add Type:=xlValidateList, Formula1:="1,2,3,4"
                    .ErrorTitle = "Ошибка"
                    .ErrorMessage = "Неверный ввод"
                End With
            End With
                With ActiveWorkbook.Worksheets("QBasic").Cells(3, 4)
                With .Validation
                    .Delete
                    .Add Type:=xlValidateList, Formula1:="Да,Нет"
                    .ErrorTitle = "Ошибка"
                    .ErrorMessage = "Неверный ввод"
                End With
            End With
        Case "Отображение результата построения мер защиты"
            ActiveWorkbook.Worksheets("QoMfA").Visible = xlSheetVisible
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Worksheets(i).Name <> "QoMfA" _
                And ChoosingListForm.HideSheet.Value <> True Then
                    ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
                End If
            Next i
        
            ActiveWorkbook.Worksheets("DMeasures").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("AMeasures").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("ResultMeasures").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("BasicMeasures").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("LoTaM").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("Order239").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("Order31").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("Order21").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("Order17").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QoMfA").Visible = xlSheetHidden
            TypeOfWorkmode = "Отображение"
        Case "Отображение приложений модели угроз"
            
            ActiveWorkbook.Worksheets("QNC").Visible = xlSheetVisible
            For i = 1 To ActiveWorkbook.Sheets.Count
                If ActiveWorkbook.Worksheets(i).Name <> "QNC" Then
                    ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
                End If
            Next i
            ActiveWorkbook.Worksheets("QTT").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QTTToI").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QNCGoI").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("TNCGoINoI").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QCollusion").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QIntOfTT").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("QAoWoR").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("TofThreats").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("TofTechniques").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("AThreats").Visible = xlSheetVisible
            ActiveWorkbook.Worksheets("ThreatsForAct").Visible = xlSheetVisible
            
            TypeOfWorkmode = "Отображение"
        Case Else
            MsgBox "Не знаю, как ты это провернул, но поздравляю. Ты молодец.", , "ERROR"
    End Select
    FirstGreetingForm.UserMode.Enabled = False
    Application.ScreenUpdating = True
    FirstGreetingForm.Hide
End Sub
