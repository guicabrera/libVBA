VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarioClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CalendarioClass
Public WithEvents lblDtSel As MSForms.Label
Attribute lblDtSel.VB_VarHelpID = -1

Private Sub lblDtSel_Click()

Dim vSel    As String


If Left(lblDtSel.Name, 1) = "l" Then

    vSel = lblDtSel.Tag
    meObj.Controls(btDataCalendClass).Text = vSel
    meObj.frCalend.Visible = False
    
End If

End Sub


'
''definicação do bloco que da o efeito de atribuir a data para o campo necessário
''---------------------------------------------------------------------------------------------
'Dim BtAtualCalendClass() As CalendarioClass '----> precisa sempre ser declarado no inicio do bloco do userform
'Dim ObjetoBtCalendClass    As Object
'Dim btSelecCalendClass     As Long
'
'ReDim BtAtualCalendClass(1 To Me.Controls.Count)
'
'For Each ObjetoBtCalendClass In Me.Controls
'If TypeName(ObjetoBtCalendClass) = "Label" Then
'    btSelecCalendClass = btSelecCalendClass + 1
'    Set BtAtualCalendClass(btSelecCalendClass) = New CalendarioClass
'    Set BtAtualCalendClass(btSelecCalendClass).lblDtSel = ObjetoBtCalendClass
'End If
'
'Next ObjetoBtCalendClass
'Set ObjetoBtCalendClass = Nothing
'
'ReDim Preserve BtAtualCalendClass(1 To btSelecCalendClass)
'
''---------------------------------------------------------------------------------------------


















'bloco de código teste ---> tentativa de deixar em função do Me setado através da meobj
'---------------------------------------------------------------------------------------------------
'Public Sub effectCalendClass()
'
'
'Dim ObjetoBtCalendClass    As Object
'Dim btSelecCalendClass     As Long
'
'ReDim BtAtualCalendClass(1 To meobj.Controls.Count)
'
'For Each ObjetoBtCalendClass In meobj.Controls
'If TypeName(ObjetoBtCalendClass) = "Label" Then
'    btSelecCalendClass = btSelecCalendClass + 1
'    Set BtAtualCalendClass(btSelecCalendClass) = New CalendarioClass
'    Set BtAtualCalendClass(btSelecCalendClass).lblDtSel = ObjetoBtCalendClass
'End If
'
'Next ObjetoBtCalendClass
'Set ObjetoBtCalendClass = Nothing
'
'ReDim Preserve BtAtualCalendClass(1 To btSelecCalendClass)
'
'End Sub
'---------------------------------------------------------------------------------------------------
