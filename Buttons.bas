Attribute VB_Name = "Buttons"
Option Explicit
Sub QNC_Next()
    '1-2
    Base.QNC_UpdateRefs
    ActiveWorkbook.Worksheets("QTT").Visible = xlSheetVisible
    ActiveWorkbook.Worksheets("QNC").Visible = xlSheetHidden
End Sub
Sub QTT_Next()
    '2-3
    If TypeOfWorkmode <> "Пользователь" Then
        ShadowStart = True
        Base.QNC_UpdateRefs
        ShadowStart = False
    End If
    Base.QTT_UpdateDict
    Base.QTTToI_Write
    
    Call Functions.MakeStep("QTT", "QTTToI")
End Sub
Sub QTT_Back()
    '2-1
    Call Functions.MakeStep("QTT", "QNC")
End Sub
Sub QTTToI_Next()
    '3-4
    If TypeOfWorkmode <> "Пользователь" Then
        ShadowStart = True
        Base.QNC_UpdateRefs
        Base.QTT_UpdateDict
        Base.QTTToI_Write
        ShadowStart = False
    End If
    Base.QTTToI_UpdateRefs
    Base.QNCGoI_Write
    
    Call Functions.MakeStep("QTTToI", "QNCGoI")
End Sub
Sub QTTToI_Back()
    '3-2
    Call Functions.MakeStep("QTTToI", "QTT")
End Sub
Sub QNCGoI_Next()
    '4-5
    If TypeOfWorkmode <> "Пользователь" Then
        ShadowStart = True
        Base.QNC_UpdateRefs
        Base.QTT_UpdateDict
        Base.QTTToI_Write
        Base.QTTToI_UpdateRefs
        Base.QNCGoI_Write
        ShadowStart = False
    End If
    Base.QNCGoI_UpdateRefs
       
    Call Functions.MakeStep("QNCGoI", "QCollusion")
End Sub
Sub QNCGoI_Back()
    '4-3
    Call Functions.MakeStep("QNCGoI", "QTTToI")
End Sub
Sub QCollusion_Next()
    '5-6
    If TypeOfWorkmode <> "Пользователь" Then
        ShadowStart = True
        Base.QNC_UpdateRefs
        Base.QTT_UpdateDict
        Base.QTTToI_Write
        Base.QTTToI_UpdateRefs
        Base.QNCGoI_Write
        Base.QNCGoI_UpdateRefs
        ShadowStart = False
    End If

    Base.QCollusion_UpdateRef
    Base.QIntOfTT_Write
    Base.TNCGoINoI_Write
    
    Call Functions.MakeStep("QCollusion", "QIntOfTT")
End Sub
Sub QCollusion_Back()
    '5-4
    Call Functions.MakeStep("QCollusion", "QTTToI")
End Sub
Sub QIntOfTT_Next()
    '6-7
    If TypeOfWorkmode <> "Пользователь" Then
        ShadowStart = True
        Base.QNC_UpdateRefs
        Base.QTT_UpdateDict
        Base.QTTToI_Write
        Base.QTTToI_UpdateRefs
        Base.QNCGoI_Write
        Base.QNCGoI_UpdateRefs
        Base.TNCGoINoI_Write
        Base.QCollusion_UpdateRef
        Base.QIntOfTT_Write
        ShadowStart = False
    End If
    
    Base.QIntOfTT_UpdateRefs
    Base.QAoWoR_Write
    
    Call Functions.MakeStep("QIntOfTT", "QAoWoR")
    
    Beep 900, 150
    Beep 900, 150
    Beep 600, 300
    
End Sub
Sub QIntOfTT_Back()
    '6-5
    Call Functions.MakeStep("QIntOfTT", "QCollusion")
End Sub
Sub QAoWoR_Next()
    '7-8
    If Sheets("QAoWoR").Cells(2, 4).Value <> "" Then
        Base.TofThreats_Write
        Call Functions.MakeStep("QAoWoR", "TofThreats")
        Beep 900, 150
        Beep 900, 150
        Beep 600, 300
    Else
        MsgBox "Лист пустой. Невозможно определить угрозы", , "ERROR: QAoWoR"
    End If
    
    
End Sub
Sub QAoWoR_Back()
    '7-6
    Call Functions.MakeStep("QAoWoR", "QIntOfTT")
End Sub
Sub TofThreats_Next()
    '8-9
    If Sheets("TofThreats").Cells(2, 4).Value <> "" Then
        Base.TofTechniques_Write
        Base.AThreats_Write
        Call Functions.MakeStep("TofThreats", "TofTechniques")
        Beep 900, 150
        Beep 900, 150
        Beep 600, 300
    Else
        MsgBox "Лист пустой. Невозможно определить техники", , "ERROR: TofThreats"
    End If
End Sub
Sub TofThreats_Back()
    '8-7
    Call Functions.MakeStep("TofThreats", "QAoWoR")
End Sub
Sub TofTofTechniques_Next()
    '9-10
    Base.CreateThreatsForAct
    Call Functions.MakeStep("TofTechniques", "ThreatsForAct")
End Sub
Sub TofTofTechniques_Back()
    '9-8
    Call Functions.MakeStep("TofTechniques", "TofThreats")
End Sub
Sub CreateQuestionary()
    Dim RefDictionary As Object
    Dim IDDictionaryKeys As Object
    Dim IDDictionaryItems As Object
    Dim SheetID1$, SheetID2$
    
    Select Case ActiveSheet.Name:
    Case "RefWoRInt"
        SheetID1 = "DWoR"
        SheetID2 = "DInt"
    Case "RefIntLoC"
        SheetID1 = "DInt"
        SheetID2 = "DLoC"
    Case "RefNoIC"
        SheetID1 = "DNoI"
        SheetID2 = "DCat"
    Case "RefNoIGoI"
        SheetID1 = "DNoI"
        SheetID2 = "DGoI"
    Case "RefTTToI"
        SheetID1 = "QTT"
        SheetID2 = "DToI"
    Case "RefWoRToI"
        SheetID1 = "DWoR"
        SheetID2 = "DToI"
    Case "RefTTInt"
        SheetID1 = "QTT"
        SheetID2 = "DInt"
    Case "RefWoRC"
        SheetID1 = "DWoR"
        SheetID2 = "DCat"
    Case "RefWoRTT"
        SheetID1 = "DWoR"
        SheetID2 = "QTT"
    Case "RefCGoI"
        SheetID1 = "QNC"
        SheetID2 = "DGoI"
        Call Functions.WriteBookOfReferenceFromAuto(SheetID1, IDDictionaryKeys, "IDDictionaryKeys", True, , Sheets(SheetID1).Cells(2, 4))
    Case "RefNoILoC"
        SheetID1 = "DNoI"
        SheetID2 = "DLoC"
    Case "RefIntCat"
        SheetID1 = "DInt"
        SheetID2 = "DCat"
    Case "RefWoRLoC"
        SheetID1 = "DWoR"
        SheetID2 = "DLoC"
    Case "RefTNC"
        SheetID1 = "QTT"
        SheetID2 = "QNC"
        Call Functions.WriteBookOfReferenceFromAuto(SheetID2, IDDictionaryItems, "IDDictionaryItems", True, , Sheets(SheetID2).Cells(2, 4))
    Case "RefCToI"
        SheetID1 = "QNC"
        SheetID2 = "DToI"
        Call Functions.WriteBookOfReferenceFromAuto(SheetID1, IDDictionaryKeys, "IDDictionaryKeys", True, , Sheets(SheetID1).Cells(2, 4))
    Case Else
        MsgBox "Oups! Case Else", , "CreateQuestionary"
    End Select
    
    If SheetID1 <> "" And SheetID2 <> "" Then
        Call Functions.WriteBookOfReferenceFromDefault(ActiveSheet.Name, RefDictionary, "RefDictionary")
        If IDDictionaryKeys Is Nothing Then
            Call Functions.WriteDictionary(SheetID1, IDDictionaryKeys, "IDDictionaryKeys")
        End If
        If IDDictionaryItems Is Nothing Then
            Call Functions.WriteDictionary(SheetID2, IDDictionaryItems, "IDDictionaryItems")
        End If
        Call Functions.DisplayDictionaryOnList(RefDictionary, IDDictionaryKeys, IDDictionaryItems)
    End If
    
End Sub
Sub ThreatsDesk_Plus()
    Sheets("ThreatsDesk").Cells(4, 1).Value = Sheets("ThreatsDesk").Cells(4, 1).Value + 1
End Sub
Sub ThreatsDesk_Minus()
    Sheets("ThreatsDesk").Cells(4, 1).Value = Sheets("ThreatsDesk").Cells(4, 1).Value - 1
    If Sheets("ThreatsDesk").Cells(4, 1).Value < 1 Then
        Sheets("ThreatsDesk").Cells(4, 1).Value = 1
    End If
End Sub
Sub ThreatsDesk_RewriteBDU()
    ThreatsDeskBusy = True
    
    Dim Num%, i%, ID_Threat%
    
    ID_Threat = CInt(Sheets("ThreatsDesk").Cells(4, 1).Value)
    If ID_Threat < 1 Then
        MsgBox "Номер угрозы слишком мал", , "ERROR: ID_Threat"
    ElseIf ID_Threat > (FindEmptyRowInColumn(Sheets("BDU").Cells(2, 2)) - 4) Then
        MsgBox "Номер угрозы слишком велик", , "ERROR: ID_Threat"
    End If
    
    Sheets("BDU").Cells(ID_Threat + 3, 12).Value = Functions.CreateOutputStringForBDU("ThreatsDesk", 12, 3, "Способы реализации угроз")
    Sheets("BDU").Cells(ID_Threat + 3, 13).Value = Functions.CreateOutputStringForBDU("ThreatsDesk", 12, 8, "Виды воздействия")
    Sheets("BDU").Cells(ID_Threat + 3, 14).Value = Functions.CreateOutputStringForBDU("ThreatsDesk", 12, 13, "Объекты")
    ThreatsDeskBusy = False
End Sub
Sub TechniquesDesk_Plus()
    Sheets("TechniquesDesk").Cells(4, 1).Value = Sheets("TechniquesDesk").Cells(4, 1).Value + 1
End Sub
Sub TechniquesDesk_Minus()
    Sheets("TechniquesDesk").Cells(4, 1).Value = Sheets("TechniquesDesk").Cells(4, 1).Value - 1
    If Sheets("TechniquesDesk").Cells(4, 1).Value < 1 Then
        Sheets("TechniquesDesk").Cells(4, 1).Value = 1
    End If
End Sub
Sub TechniquesDesk_RewriteTactic()
    Dim Num%, i%, ID_Tactic%
    
    ID_Tactic = CInt(Sheets("TechniquesDesk").Cells(4, 1).Value)
    If ID_Tactic < 1 Then
        MsgBox "Номер тактики слишком мал", , "ERROR: ID_Tactic"
    ElseIf ID_Tactic > (FindEmptyRowInColumn(Sheets("RefTactics").Cells(2, 2)) - 4) Then
        MsgBox "Номер тактики слишком велик", , "ERROR: ID_Tactic"
    End If
    
    Sheets("RefTactics").Cells(ID_Tactic + 3, 11).Value = Functions.CreateOutputStringForBDU("TechniquesDesk", 12, 3, "Способы реализации угроз")
    Sheets("RefTactics").Cells(ID_Tactic + 3, 10).Value = Functions.CreateOutputStringForBDU("TechniquesDesk", 12, 8, "Виды воздействия")
    Sheets("RefTactics").Cells(ID_Tactic + 3, 9).Value = Functions.CreateOutputStringForBDU("TechniquesDesk", 12, 13, "Объекты")
    If Sheets("TechniquesDesk").Cells(14, 17).Value = "" Then
        MsgBox "Не заполнен уровень нарушителя", , "WARNING"
    Else
        Sheets("RefTactics").Cells(ID_Tactic + 3, 8).Value = Sheets("TechniquesDesk").Cells(14, 17).Value
    End If
    If Sheets("TechniquesDesk").Cells(14, 16).Value = "" Then
        MsgBox "Не заполнена категория нарушителя", , "WARNING"
    Else
        Sheets("RefTactics").Cells(ID_Tactic + 3, 7).Value = Sheets("TechniquesDesk").Cells(14, 16).Value
    End If
End Sub
Sub SelectList()
    If TypeOfWorkmode = "Разработчик" Then
        ChoosingListForm.ListOfLists.Clear
        ChoosingListForm.ListOfLists.ListRows = ActiveWorkbook.Sheets.Count
        ChoosingListForm.Show
    ElseIf TypeOfWorkmode = "" Then
        ChoosingListForm.ListOfLists.Clear
        ChoosingListForm.ListOfLists.ListRows = ActiveWorkbook.Sheets.Count
        ChoosingListForm.Show
    Else
        MsgBox "В данный момент используется режим [" + TypeOfWorkmode + "] и переключение между листами невозможно", , "Предупреждение"
    End If
End Sub
Sub QBasic_Next()
    '1-2
    Base.QBasic_UpdateRefs
    Call Functions.MakeStep("QBasic", "QoMfD")
End Sub
Sub QoMfD_Next()
    '2-3
    Base.QoMfD_UpdateRefs
    Call Functions.MakeStep("QoMfD", "QoMfA")
End Sub
Sub QoMfD_Back()
    '2-1
    Call Functions.MakeStep("QoMfD", "QBasic")
End Sub
Sub QoMfA_Next()
    '3-4
    Base.QoMfA_UpdateRefs
    Base.Measures_Show
End Sub
Sub QoMfA_Back()
    '3-2
    Call Functions.MakeStep("QoMfA", "QoMfD")
End Sub
Sub Restart_Creation()
    Dim i%
    '4-1
    Call Functions.MakeStep(ActiveWorkbook.ActiveSheet.Name, "QBasic")
    
    
    For i = 1 To ActiveWorkbook.Sheets.Count
        If Not ActiveWorkbook.Worksheets(i).Name = "QBasic" _
        And ChoosingListForm.HideSheet.Value <> True Then
            ActiveWorkbook.Worksheets(i).Visible = xlSheetHidden
        End If
    Next i
    
End Sub
