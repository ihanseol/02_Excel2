Option Explicit

Private Sub CommandButton_PressAll_Click()
    Call PressAll_Button
End Sub

Private Sub CommandButton1_Click()
' add well

    BaseData_ETC_02.TurnOffStuff
    Call AddWell_CopyOneSheet
    BaseData_ETC_02.TurnOnStuff
    
End Sub


'AggChart Button
Private Sub CommandButton10_Click()
    Sheets("AggChart").Visible = True
    Sheets("AggChart").Select
End Sub


'AggFx Button
Private Sub CommandButton11_Click()
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
End Sub

Private Sub CommandButton12_Click()
    Sheets("water").Visible = True
    Sheets("water").Select
End Sub

Private Sub CommandButton13_Click()
' 이것은, FX에서 양수일보 데이터를 가지고 오면,
' 각각의 관정, 1, 2, 3 번 식으로
' FX 에서 가지고 온다.
' Import All Well Spec

    Call ImportAll_EachWellSpec
End Sub


Private Sub CommandButton14_Click()
    'wSet, WellSpec Setting

    Dim nofwell, i As Integer

    nofwell = sheets_count()
    
    For i = 1 To nofwell
        Cells(i + 3, "E").formula = "=Recharge!$I$24"
        Cells(i + 3, "F").formula = "=All!$B$2"
        Cells(i + 3, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i
End Sub

Private Sub CommandButton15_Click()
' 2024/6/24 - dupl, duplicate basic well data ...
' 기본관정데이타 복사하는것
' 관정을 순회하면서, 거기에서 데이터를 가지고 오는데 …
' 와파 , 장축부, 단축부
' 유향, 거리, 관정높이, 지표수표고 이렇게 가지고 오면 될듯하다.

' k6 - 장축부 / long axis
' k7 - 단축부 / short axis
' k12 - degree of flow
' k13 - well distance
' k14 - well height
' k15 - surfacewater height

    Call DuplicateBasicWellData
  
End Sub

'AggSum Button
Private Sub CommandButton3_Click()
    Sheets("AggSum").Visible = True
    Sheets("AggSum").Select
End Sub



'Aggregate1 Button
Private Sub CommandButton4_Click()
    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
End Sub




Private Sub CommandButton5_Click()
' Aggregate2 Button
' 집계함수 2번

    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
End Sub


'AggWhpa Button
Private Sub CommandButton7_Click()
    Sheets("aggWhpa").Visible = True
    Sheets("aggWhpa").Select
End Sub


Private Sub CommandButton9_Click()
  Sheets("AggStep").Visible = True
  Sheets("AggStep").Select
End Sub


'Jojung Button
'add new feature - correct border frame ...
Private Sub CommandButton2_Click()
    Call JojungButton
End Sub



' delete last
Private Sub CommandButton8_Click()
    Call DeleteLast
End Sub

Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub

'one button / delete all well except for one ...

Private Sub CommandButton6_Click()
   Call Make_OneButton
End Sub





