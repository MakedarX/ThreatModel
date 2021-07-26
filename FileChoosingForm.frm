VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FileChoosingForm 
   Caption         =   "Ошибка имени файла"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3750
   OleObjectBlob   =   "FileChoosingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FileChoosingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ExitButton_Click()
    Base.ContinueExraction = False
    FileChoosingForm.Hide
End Sub

Private Sub TryAgain_Click()
    Base.ContinueExraction = True
    FileChoosingForm.Hide
End Sub
