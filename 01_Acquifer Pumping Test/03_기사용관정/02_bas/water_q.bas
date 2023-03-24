Attribute VB_Name = "water_q"
Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double
Public ISIT_FIRST As Boolean

Public Enum SS_VALUE
    svGAJUNG = 1
    svILBAN = 2
    svSCHOOL = 3
    svGONGDONG = 4
    svMAEUL = 5
End Enum

Public Enum AA_VALUE
    avJEONJAK = 1
    avDAPJAK = 2
    avWONYE = 3
    avCOW = 4
    avPIG = 5
    avCHICKEN = 6
End Enum

Function CheckBoxFind(objNAME As String) As MSForms.CheckBox
    
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myCheckBox As MSForms.CheckBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myCheckBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.CheckBox Then
            If obj.NAME = objNAME Then
                Set myCheckBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myCheckBox Is Nothing) Then
        ' found
        Set CheckBoxFind = myCheckBox
    Else
        ' not found
        Set CheckBoxFind = Nothing
    End If

End Function

Function ComboBoxFind(objNAME As String) As MSForms.ComboBox
    
    Dim ws As Worksheet
    Dim obj As OLEObject
    Dim myComboBox As MSForms.ComboBox
    
    
    Set ws = ThisWorkbook.Worksheets("ss")
    Set myComboBox = Nothing
    
    For Each obj In ws.OLEObjects
        If TypeOf obj.Object Is MSForms.ComboBox Then
            If obj.NAME = objNAME Then
                Set myComboBox = obj.Object
                Exit For
            End If
        End If
    Next obj
    
    If Not (myComboBox Is Nothing) Then
        ' found
        Set ComboBoxFind = myComboBox
    Else
        ' not found
        Set ComboBoxFind = Nothing
    End If

End Function



Sub initialize()
    Dim checkboxJIYEOL As MSForms.CheckBox
    Dim comboboxAREA As MSForms.ComboBox


    Set checkboxJIYEOL = CheckBoxFind("chkboxJIYEOL")
    Set comboboxAREA = ComboBoxFind("comboAREA")
    
    Debug.Print "comboAREA", comboboxAREA.Value
    
    If checkboxJIYEOL.Value Then
        Call initialize_JIYEOL(comboboxAREA.Value)
    Else
        Call initialize_CNU(comboboxAREA.Value)
    End If

End Sub


Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


' �������
Sub ComputeQ()
    Dim i As Integer
    Dim lastRow As Long

    Call initialize
    
    Sheets("ss").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = ss_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
    
    Sheets("aa").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = aa_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
End Sub


Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double
    '���� �ó���
    If CheckSubstring(strPurpose, "��") Then
        ss_water = qhp * 0.01
        Exit Function
    End If
    
    ' �Ϲݿ�
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' ��Ÿ
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' ���Ȱ���
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' û�ҿ�
    If CheckSubstring(strPurpose, "û") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    '���̻����
    If CheckSubstring(strPurpose, "��") Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    ' �����
    If CheckSubstring(strPurpose, "����") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' �������ÿ�
    If CheckSubstring(strPurpose, "����") Then
        ss_water = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    ' �ι�����
    If CheckSubstring(strPurpose, "�ι�") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' �б���
    If CheckSubstring(strPurpose, "�б�") Then
        ss_water = Round(SS(svSCHOOL, 1) + npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
   ss_water = 900
      
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double

    'nhead - ������ �μ� ....

    ' ���ۿ�
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' ���ۿ�
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    ' ���Ȱ���
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' ������
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    '����
    If CheckSubstring(strPurpose, "��") Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
    ' ��Ÿ
    If CheckSubstring(strPurpose, "��Ÿ") Then
        aa_water = Round(AA(avDAPJAK, 1) + nhead * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
   aa_water = 900
      
End Function










