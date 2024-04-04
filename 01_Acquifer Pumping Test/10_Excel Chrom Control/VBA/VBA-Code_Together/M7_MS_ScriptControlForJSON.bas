Option Explicit

Private ScriptEngine As scriptControl

Public Sub InitScriptEngine()
On Error Resume Next
    Set ScriptEngine = New scriptControl
    ScriptEngine.Language = "JScript"
On Error GoTo 0

    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal JSonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + JSonString + ")")
End Function

Public Function GetProperty(ByVal JsonObject As Object, ByVal propertyName As String) 'As Variant
    GetProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal JsonObject As Object, ByVal propertyName As String) 'As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", JsonObject, propertyName)
End Function

Public Function GetKeys(ByVal JsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", JsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function

Sub ConvertJsonToDictionary()
    Dim jsonStr As String
    Dim dict As Object
    
    ' Your JSON string
    jsonStr = "{""Name"": ""John"", ""Age"": 30, ""City"": ""New York""}"
    
    Call InitScriptEngine
    Debug.Print DecodeJsonString(jsonStr)
    
    
    
End Sub

