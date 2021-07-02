Attribute VB_Name = "zEfeitos"
Option Explicit
Private meObj As Object
Public BtAtual() As ClasseEfeito

'sub para adicionar efeito toda vez que passar o mouse
'deve ser adicionado no initialize do userform
'não está funcionando ainda --> deve ser colocado o código da sub trocando o obj pelo o me do userform

Sub addEfeito(ByVal obj As Object)
'efeito hover
'--------------------------------------------------
Dim ObjetoBt    As Object
Dim BtSelec     As Long

ReDim BtAtual(1 To Me.Controls.Count)

For Each ObjetoBt In Me.Controls
If TypeName(ObjetoBt) = "Label" Then
    BtSelec = BtSelec + 1
    Set BtAtual(BtSelec) = New ClasseEfeito
    Set BtAtual(BtSelec).aplicamod = ObjetoBt
End If

Next ObjetoBt
Set ObjetoBt = Nothing
'--------------------------------------------------

End Sub


'sub para retirar o efeito quando retirar o mouse
'deve ser adicionado em:
    'Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'Call TiraEfeitos(Me)
    'End Sub
'na seção do userform
Sub TiraEfeitos(ByVal obj As Object)
Set meObj = obj
Dim contForm As control


For Each contForm In meObj.Controls
    
    If TypeName(contForm) = "Label" And Left(contForm.Name, 3) = "cmd" Then
        
        contForm.BackStyle = fmBackStyleTransparent
        contForm.BorderStyle = 1
        
    End If

Next contForm
Set contForm = Nothing

End Sub
