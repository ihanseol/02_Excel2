' Class Module: Class_ReturnTrueFalse
Private mValue As Boolean

Private Sub Class_Initialize()
    ' Initialize default values
    mValue = False
End Sub

Public Property Let Result(val As Boolean)
    mValue = val
End Property

Public Property Get Result() As Boolean
    Result = mValue
End Property
