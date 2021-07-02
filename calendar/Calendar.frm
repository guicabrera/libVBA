VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendar 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Calendar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
Call calendIni(Me)
'btDataCalendClass = "" ---------> deve ser colocado o valor do textbox ao qual é desejado a atribuição do valor de data
End Sub
Private Sub lblHoje_Click()
Call calendChangeMonth(Me)
End Sub
Private Sub sb_Change()

Call Atualizar(DateSerial(sb \ 12, sb Mod 12, 1), Me)
End Sub
