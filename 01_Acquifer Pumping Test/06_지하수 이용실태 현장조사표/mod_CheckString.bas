Attribute VB_Name = "mod_CheckString"
Option Explicit


Function CheckSubstring(str As String, chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function

Function SS_StringCheck(str As String) As String
    
    ' ������ - �缳
    If CheckSubstring(str, "����") Then SS_StringCheck = "g,"
    
    ' �Ϲݿ� - �缳
    If CheckSubstring(str, "�Ϲ�") Then SS_StringCheck = "h,"
    
    ' �б��� - ����
    If CheckSubstring(str, "�б�") Then SS_StringCheck = "i,"
        
    ' �ι����� - ����
    If CheckSubstring(str, "�ι�") Then SS_StringCheck = "j,"
    
    ' �������ÿ� - �缳
    If CheckSubstring(str, "����") Then SS_StringCheck = "k,"
    
    ' ���̻���� - ����
    If CheckSubstring(str, "����") Then SS_StringCheck = "l,"
    
    ' ���Ȱ��� - �缳
    If CheckSubstring(str, "���") Then SS_StringCheck = "m,"
    
    ' ��Ÿ - �缳
    If CheckSubstring(str, "��Ÿ") Then SS_StringCheck = "n,"

End Function

Function AA_StringCheck(str As String) As String
    
    ' ������� ���� �缳, ���� �㰡�� - ����
    If CheckSubstring(str, "����") Then AA_StringCheck = "v,"
    If CheckSubstring(str, "����") Then AA_StringCheck = "w,"
    If CheckSubstring(str, "����") Then AA_StringCheck = "x,"
    
    If CheckSubstring(str, "���") Then AA_StringCheck = "y,"
    If CheckSubstring(str, "���") Then AA_StringCheck = "z,"
    If CheckSubstring(str, "��Ÿ") Then AA_StringCheck = "aa,"
    
End Function


Function II_StringCheck(str As String) As String
    
    ' �ذ�, ����, ��� - ����
    If CheckSubstring(str, "����") Then II_StringCheck = "p,"
    If CheckSubstring(str, "����") Then II_StringCheck = "q,"
    If CheckSubstring(str, "���") Then II_StringCheck = "r,"
    
    ' ��������, ��Ÿ - �缳
    If CheckSubstring(str, "��������") Then II_StringCheck = "s,"
    If CheckSubstring(str, "��Ÿ") Then II_StringCheck = "t,"

End Function



Function SS_PublicCheck(str As String) As String
    
    ' ������ - �缳
    If CheckSubstring(str, "����") Then SS_PublicCheck = "ab,"
    
    ' �Ϲݿ� - �缳
    If CheckSubstring(str, "�Ϲ�") Then SS_PublicCheck = "ac,"
    
    ' �б��� - ����
    If CheckSubstring(str, "�б�") Then SS_PublicCheck = "ab,"
        
    ' �ι����� - ����
    If CheckSubstring(str, "�ι�") Then SS_PublicCheck = "ab,"
    
    ' �������ÿ� - �缳
    If CheckSubstring(str, "����") Then SS_PublicCheck = "ac,"
    
    ' ���̻���� - ����
    If CheckSubstring(str, "����") Then SS_PublicCheck = "ab,"
    
    ' ���Ȱ��� - �缳
    If CheckSubstring(str, "���") Then SS_PublicCheck = "ac,"
    
    ' ��Ÿ - �缳
    If CheckSubstring(str, "��Ÿ") Then SS_PublicCheck = "ac,"

End Function

Function AA_PublicCheck(str As String) As String
    
    ' ������� ���� �缳, ���� �㰡�� - ����
    If CheckSubstring(str, "����") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "����") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "����") Then AA_PublicCheck = "ac,"
    
    If CheckSubstring(str, "���") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "���") Then AA_PublicCheck = "ac,"
    If CheckSubstring(str, "��Ÿ") Then AA_PublicCheck = "ac,"
    
End Function


Function II_PublicCheck(str As String) As String
    
    ' �ذ�, ����, ��� - ����
    If CheckSubstring(str, "����") Then II_PublicCheck = "ab,"
    If CheckSubstring(str, "����") Then II_PublicCheck = "ab,"
    If CheckSubstring(str, "���") Then II_PublicCheck = "ab,"
    
    ' ��������, ��Ÿ - �缳
    If CheckSubstring(str, "��������") Then II_PublicCheck = "ac,"
    If CheckSubstring(str, "��Ÿ") Then II_PublicCheck = "ac,"

End Function


