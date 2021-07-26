VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ShadowStartForm 
   Caption         =   "Закадровый прогон рассчетов"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "ShadowStartForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ShadowStartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BeginShadowStart_Click()
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
    Base.QIntOfTT_UpdateRefs
    Base.QAoWoR_Write
    Base.TofThreats_Write
    ShadowStart = False
    ShadowStartForm.Hide
End Sub
