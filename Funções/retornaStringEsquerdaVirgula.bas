Attribute VB_Name = "trasStringLeftComma"
Function analistaAceitaCaso(ByVal valueCell As String)

For x = 1 To Len(valueCell)
    If InStr(1, Left(valueCell, x), ",") Then
    analistaAceitaCaso = Left(valueCell, x - 1)
    Exit For
    End If
    
    
Next x

End Function

