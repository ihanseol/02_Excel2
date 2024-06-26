VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_Boolean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
