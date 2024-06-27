' Class Module: Class_ReturnTrueFalse
Private mValue As Boolean

Private Sub Class_Initialize()
    ' Initialize default values
    mValue = False
End Sub

Public Property Let result(val As Boolean)
    mValue = val
End Property

Public Property Get result() As Boolean
    result = mValue
End Property
