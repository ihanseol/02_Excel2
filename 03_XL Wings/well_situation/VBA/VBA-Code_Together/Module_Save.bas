'
'' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA
'Sub ToggleOX()
'    Dim activeCellColumn, activeCellRow As String
'    Dim row As Long
'    Dim col As Long
'    Dim lastRow As Long
'    Dim cp, fillRange As String
'
'
'
'    activeCellColumn = Split(ActiveCell.address, "$")(1)
'    activeCellRow = Split(ActiveCell.address, "$")(2)
'
'    row = ActiveCell.row
'    col = ActiveCell.Column
'
'    Debug.Print Alpha_Column(ActiveCell)
'
'    If activeCellColumn = "S" Then
'        If ActiveCell.Value = "O" Then
'            ActiveCell.Value = "X"
'        Else
'            ActiveCell.Value = "O"
'        End If
'    End If
'
'
'    If activeCellColumn = "B" Then
'        If ActiveCell.Value = "신고공" Then
'            ActiveCell.Value = "허가공"
'            With Selection.Font
'                .Color = -16776961
'                .TintAndShade = 0
'            End With
'            Selection.Font.Bold = True
'        Else
'            ActiveCell.Value = "신고공"
'             With Selection.Font
'                .ThemeColor = xlThemeColorLight1
'                .TintAndShade = 0
'            End With
'            Selection.Font.Bold = False
'        End If
'    End If
'
'    If activeCellColumn = "D" Then
'        cp = Replace(ActiveCell.address, "$", "")
'        lastRow = lastRowByKey(ActiveCell.address)
'
'        fillRange = "D" & Range(cp).row & ":D" & lastRow
'
'        Range(cp).Select
'        Selection.AutoFill Destination:=Range(fillRange)
'
'        Range(cp).Select
'    End If
'
'    If activeCellColumn = "C" Then
'        cp = Replace(ActiveCell.address, "$", "")
'        lastRow = lastRowByKey(ActiveCell.address)
'
'        fillRange = "C" & Range(cp).row & ":C" & lastRow
'
'        Range(cp).Select
'        Selection.AutoFill Destination:=Range(fillRange)
'
'        Range(cp).Select
'    End If
'
'
'    ' activeCellColumn = F, G, H, I, J
'    ' activeCellRow
'    If activeCellColumn = "F" Or activeCellColumn = "G" Or activeCellColumn = "H" Or activeCellColumn = "I" Or activeCellColumn = "J" Then
'        Dim ret, r As Variant
'        Dim i As Integer: i = 0
'
'        If Not IsEmpty(ActiveCell.Value) Then
'            GoTo FLAG_END
'        End If
'
'        ret = get_wellinfo_function()
'
'        Dim yongdo As Variant
'        Dim sebu As Variant
'        Dim simdo As Variant
'        Dim well_diameter As Variant
'        Dim well_hp As Variant
'        Dim well_q As Variant
'        Dim well_tochul As Variant
'        Dim yongdo_s As String
'
'        yongdo = ret(0)
'        sebu = ret(1)
'        simdo = ret(2)
'        well_diameter = ret(3)
'        well_hp = ret(4)
'        well_q = ret(5)
'        well_tochul = ret(6)
'
'        Select Case yongdo
'            Case "농업용"
'            Case "농어업용"
'                    yongdo_s = "aa"
'
'            Case "생활용"
'                    yongdo_s = "ss"
'
'            Case "공업용"
'                    yongdo_s = "ii"
'
'            Case Else
'                    yongdo_s = "ss"
'        End Select
'
'        If yongdo_s <> ActiveSheet.name Then
'            Debug.Print yongdo
'            Call BeepExample
'        Else
'            Cells(activeCellRow, "K").Value = sebu
'            Cells(activeCellRow, "F").Value = simdo
'            Cells(activeCellRow, "G").Value = well_diameter
'            Cells(activeCellRow, "H").Value = well_hp
'            Cells(activeCellRow, "I").Value = well_q
'            Cells(activeCellRow, "J").Value = well_tochul
'        End If
'
'    End If
'
'
'    If ActiveSheet.name = "ss" And activeCellColumn = "K" Then
'        UserForm_SS.Show
'    End If
'
'    If ActiveSheet.name = "aa" And activeCellColumn = "K" Then
'        UserForm_AA.Show
'    End If
'
'    If ActiveSheet.name = "ii" And activeCellColumn = "K" Then
'        UserForm_II.Show
'    End If
'
'FLAG_END:
'
'End Sub
'
