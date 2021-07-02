Attribute VB_Name = "subCalendIni"
Public Sub Atualizar(dt As Date, meObj As Object)
      
    Dim L As Long
    Dim C As Long
    Dim cInício As Long
    Dim dtDia As Date
    Dim ctrl As Control
    
    meObj.MesAno = Format(dt, "mmmm yyyy")
    
    For L = 1 To 6
        For C = 1 To 7
            Set ctrl = meObj.Controls("l" & L & "c" & C)
            
            dtDia = DateSerial(Year(dt), Month(dt), (L - 1) * 7 + C - Weekday(dt) + 1)
            ctrl.Caption = Format(Day(dtDia), "00")
            ctrl.Tag = dtDia
           
            If Month(dtDia) <> Month(dt) Then
                ctrl.ForeColor = &HC0C0&
            Else
                ctrl.ForeColor = RGB(8, 25, 48)
            End If
            
            If dtDia = Date Then
                ctrl.ForeColor = rgbRed
            End If
        Next C
    Next L

End Sub

'precisa ser iniciado no userform initialization
Sub calendIni(ByVal meObj As Object)
meObj.lblHoje = Format(Date, "dddd") & vbNewLine & _
Format(Date, "dd/mm/yyyy")
meObj.sb = Year(Date) * 12 + Month(Date)
End Sub

Sub calendChangeMonth(ByVal meObj As Object)
meObj.sb = Year(Date) * 12 + Month(Date)
End Sub


''comandos de click e mudança do calendar -----------> precisam ser colodos no escopo do userform
''-------------------------------------------------------------------
'Private Sub lblHoje_Click()
'Call calendChangeMonth(Me)
'End Sub
'Private Sub sb_Change()
'
'Call Atualizar(DateSerial(sb \ 12, sb Mod 12, 1), Me)
'End Sub
''-------------------------------------------------------------------
