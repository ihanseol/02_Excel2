Sub 매크로3()
'
' 매크로3 매크로
'

'
    Range("M2").Select
    Selection.AutoFill Destination:=Range("M2:M17")
    Range("M2:M17").Select
End Sub
