Option Explicit


' AggSum_Intake : 취수계획량
' AggSum_Simdo : 굴착심도
' AggSum_MotorHP : 펌프마력
' AggSum_NaturalLevel : 자연수위
' AggSum_StableLevel : 안정수위
' AggSum_ToChool : 토출구경
' AggSum_MotorSimdo : 모터심도
'

' AggSum_ROI : radius of influence
' AggSum_DI : drastic index
' AggSum_ROI_Stat :
' AggSum_26_AC : 26, AquiferCharacterization
' AggSum_26_RightAC : 26, Right AquiferCharacterizationn

Const WELL_BUFFER = 30

Private Sub CommandButton1_Click()
    Sheets("AggSum").Visible = False
    Sheets("Well").Select
End Sub


' Summary Button
Private Sub CommandButton2_Click()
   Call Summary_Button
End Sub


' 2025-2-12, CheckBoxClick 추가해줌 ...
Private Sub CheckBox1_Click()
  Application.Run "Sheet_AggSum.Summary_Button"
End Sub


' 2025-2-12, CheckBoxClick 추가해줌 ...
Sub Summary_Button()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggSum" Then Sheets("AggSum").Select

    ' Summary, Aquifer Characterization  Appropriated Water Analysis
    BaseData_ETC_02.TurnOffStuff
    
    Call Write23_SummaryDevelopmentPotential
    Call Write26_AquiferCharacterization(nofwell)
    Call Write26_Right_AquiferCharacterization(nofwell)
    
    Call Write_RadiusOfInfluence(nofwell)
    Call Write_WaterIntake(nofwell)
    Call Check_DI
    
    Call Write_DiggingDepth(nofwell)
    Call Write_MotorPower(nofwell)
    Call Write_DrasticIndex(nofwell)
    
    Call Write_NaturalLevel(nofwell)
    Call Write_StableLevel(nofwell)
    
    
    Call Write_MotorTochool(nofwell)
    Call Write_MotorSimdo(nofwell)
    
    Call Write_WellDiameter(nofwell)
    Call Write_CasingDepth(nofwell)


    Range("D5").Select
    BaseData_ETC_02.TurnOnStuff
End Sub

