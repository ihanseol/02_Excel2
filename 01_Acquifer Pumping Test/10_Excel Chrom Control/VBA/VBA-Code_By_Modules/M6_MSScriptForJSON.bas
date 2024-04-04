Attribute VB_Name = "M6_MSScriptForJSON"
Sub ConvertJsonToDictionary()
    Dim jsonStr As String
    Dim dict As Object
    
    ' Your JSON string
    jsonStr = "{""Name"": ""John"", ""Age"": 30, ""City"": ""New York""}"
    
    ' Convert JSON string to dictionary
    Set dict = JsonToDictionary(jsonStr)
    
    ' Print values from dictionary
    Debug.Print "Name: " & dict("Name")
    Debug.Print "Age: " & dict("Age")
    Debug.Print "City: " & dict("City")
End Sub

Function JsonToDictionary(jsonStr As String) As Object
    Dim scriptControl As Object
    Dim JsonObject As Object
    
    ' Create a ScriptControl object
    Set scriptControl = CreateObject("ScriptControl")
    scriptControl.Language = "JScript"
    
    ' Evaluate the JSON string to create a JSON object
    Set JsonObject = scriptControl.Eval("(" + jsonStr + ")")
    
    ' Convert JSON object to dictionary
    Set JsonToDictionary = JsonToDictionaryRecursive(JsonObject)
End Function

Function JsonToDictionaryRecursive(JsonObject As Object) As Object
    Dim dict As Object
    Dim Key As Variant
    
    ' Create a dictionary object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Iterate through the JSON object
    For Each Key In JsonObject.Keys
        If IsObject(JsonObject(Key)) Then
            ' Recursively convert nested JSON objects
            dict.Add Key, JsonToDictionaryRecursive(JsonObject(Key))
        Else
            ' Add key-value pairs to the dictionary
            dict.Add Key, JsonObject(Key)
        End If
    Next Key
    
    ' Return the dictionary
    Set JsonToDictionaryRecursive = dict
End Function

