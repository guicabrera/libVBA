Attribute VB_Name = "CalendarioClassDeclarations"
Public btDataCalendClass As String
Public meObj As Object '-------------> deve ser setado no userform initialize e activate



'faz com que a
Sub locationCalendar(meObj As Object, meFieldTop As Integer, meFieldLeft As Integer, meFieldHeight As Integer, meFieldName As String, Optional frameTop As Integer, Optional frameLeft As Integer)


    If meObj.frCalend.Visible Then
        meObj.frCalend.Visible = False
    Else
        meObj.frCalend.Visible = True
        btDataCalendClass = meFieldName
        meObj.frCalend.Top = meFieldTop + meFieldHeight + frameTop
        meObj.frCalend.Left = meFieldLeft + frameLeft
    End If
    
End Sub
