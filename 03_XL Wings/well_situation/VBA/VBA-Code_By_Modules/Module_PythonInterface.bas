Attribute VB_Name = "Module_PythonInterface"


Sub SampleCall()
'    mymodule = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
'    RunPython "import " & mymodule & ";" & mymodule & ".main()"

    RunPython "import demo; demo.main()"

End Sub

Sub SampleRemoteCall()
    RunRemotePython "http://127.0.0.1:8000/hello", apiKey:="DEVELOPMENT"
End Sub


Sub retTest()

    Dim ret As Variant
    Dim i As Variant
    
    ret = Application.Run("ret_test")
    
    For Each i In ret
        Debug.Print (i)
    Next i

End Sub

Sub get_wellinfo_test()

    Dim ret As Variant
    Dim i As Variant
    
    ret = Application.Run("get_wellinfo")
    
    For Each i In ret
        Debug.Print (i)
    Next i

End Sub


Function get_wellinfo_function(ByVal factor As Integer) As Variant
    Dim ret As Variant
    Dim yongdo As Variant
    Dim sebu As Variant
    Dim simdo As Variant
    Dim well_diameter As Variant
    Dim well_hp As Variant
    Dim well_q As Variant
    Dim well_tochul As Variant
    
    ret = Application.Run("get_wellinfo", factor)
    
    yongdo = ret(0)
    sebu = ret(1)
    simdo = ret(2)
    well_diameter = ret(3)
    well_hp = ret(4)
    well_q = ret(5)
    well_tochul = ret(6)
    

    get_wellinfo_function = Array(yongdo, sebu, simdo, well_diameter, well_hp, well_q, well_tochul)
End Function



Sub Test()
    Dim ret As Variant
    Dim r As Variant
    
    ret = get_wellinfo_function()
    
    For Each r In ret
        Debug.Print (r)
    Next r
    
    Dim yongdo As Variant
    Dim sebu As Variant
    Dim simdo As Variant
    Dim well_diameter As Variant
    Dim well_hp As Variant
    Dim well_q As Variant
    Dim well_tochul As Variant
    
    yongdo = ret(0)
    sebu = ret(1)
    simdo = ret(2)
    well_diameter = ret(3)
    well_hp = ret(4)
    well_q = ret(5)
    well_tochul = ret(6)
End Sub


