VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObservacao 
   Caption         =   "Observação para Sienge"
   ClientHeight    =   2970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   OleObjectBlob   =   "frmObservacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmObservacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CarregarTexto(ByVal texto As String)
    Me.txtObservacao.Value = texto
End Sub



Private Sub txtObservacao_Change()

End Sub

Private Sub UserForm_Initialize()
    Me.Width = 360
    Me.Height = 250
    
    Me.txtObservacao.Left = 12
    Me.txtObservacao.Top = 18
    Me.txtObservacao.Width = 250
    Me.txtObservacao.Height = 150
    
   
End Sub


