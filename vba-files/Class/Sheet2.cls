VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub battle()
ReDim characterA(17), characterB(17)
ReDim attaker(17), defender(17), exchanger(17)
'���Ը�ֵ
characterA(1) = Sheets("Battle").Range("D6")
characterA(2) = Sheets("Battle").Range("F6")
characterA(3) = Sheets("Battle").Range("D7")
characterA(4) = Sheets("Battle").Range("F7")
characterA(5) = Sheets("Battle").Range("D8")
characterA(6) = Sheets("Battle").Range("F8")
characterA(7) = Sheets("Battle").Range("D9")
characterA(8) = Sheets("Battle").Range("F9")
characterA(9) = Sheets("Battle").Range("D10")
characterA(10) = Sheets("Battle").Range("F10")
characterA(11) = Sheets("Battle").Range("D11")
characterA(12) = Sheets("Battle").Range("F11")
characterA(13) = Sheets("Battle").Range("D12")
characterA(14) = Sheets("Battle").Range("F12")
characterA(15) = Sheets("Battle").Range("D13")
characterA(16) = Sheets("Battle").Range("F13")
characterB(1) = Sheets("Battle").Range("D19")
characterB(2) = Sheets("Battle").Range("F19")
characterB(3) = Sheets("Battle").Range("D20")
characterB(4) = Sheets("Battle").Range("F20")
characterB(5) = Sheets("Battle").Range("D21")
characterB(6) = Sheets("Battle").Range("F21")
characterB(7) = Sheets("Battle").Range("D22")
characterB(8) = Sheets("Battle").Range("F22")
characterB(9) = Sheets("Battle").Range("D23")
characterB(10) = Sheets("Battle").Range("F23")
characterB(11) = Sheets("Battle").Range("D24")
characterB(12) = Sheets("Battle").Range("F24")
characterB(13) = Sheets("Battle").Range("D25")
characterB(14) = Sheets("Battle").Range("F25")
characterB(15) = Sheets("Battle").Range("D26")
characterB(16) = Sheets("Battle").Range("F26")
characterA(0) = "A"
characterB(0) = "B"
characterA(17) = 0
characterB(17) = 0
'ս���м�ֵ
' ����  1
' ����  2
' ����  3
' ����  4
' ����  5
' �����ֿ�  6
' ����  7
' �����ֿ�  8
' ����  9
' �����ֿ�  10
' ����  11
' ���ֿܵ�  12
' ����  13
' ���εֿ�  14
' ��Ѫ  15
' ��Ѫ�ֿ�  16
' ���α�� 17
Dim damlog As String
Dim rhp As Single
Dim j As Integer
j = 6

'====ս������=====
Sheets("battle").Range("N:N").Clear
'��������ȷ������Ȩ,���幥����ɫ
If characterA(4) > characterB(4) Then
  
    attaker = characterA()
    defender = characterB()
    Cells(5, 14) = "A�ٶȴ���B,A�����ж�"
    
Else
    attaker = characterB()
    defender = characterA()
    Cells(5, 14) = "A�ٶȲ�����B,B�����ж�"
End If
rhp = defender(2)
For i = 1 To 16
    '�����涨�غ�˫����û��
    If i = 16 And rhp > 0 Then
        Cells(autorow(j), 14) = "---------ս������--------"
        Cells(autorow(j), 14) = "A��15�غ���δ����B,B��ʤ"
        '�غϽ���
    ElseIf rhp > 0 Then 'һ����������
        Cells(autorow(j), 14) = "---------��" & i & "�غ�--------"
      
        For e = 1 To 2
            rhp = defender(2)
            dam = damage(attaker(1), defender(3))
            If attribution(attaker(9), defender(10)) = True Then '�����춨
                dam = dam * 2
                damlog = dam & "(����!)"
            Else
                damlog = dam
            End If
            rhp = rhp - dam
            If rhp <= 0 Then
                rhp = 0
            End If
            If isdizzy(attaker(17)) = True Then '���������������������غ�
            Cells(autorow(j), 14) = attaker(0) & "(��ѣ),�޷��ж�"
            attaker(17) = attaker(17) - 1 '��ѣ״̬,value:1
        ElseIf attribution(attaker(13), defender(14)) = True Then '���μ춨
            defender(17) = 2 '������ѣЧ��,value:2
            Cells(autorow(j), 14) = attaker(0) & "��" & defender(0) & "���" & damlog & "�˺�," & defender(0) & "��ʣ��Ѫ��Ϊ:" & rhp & " " & defender(0) & "������"
        ElseIf defender(17) = 0 And attribution(defender(11), attaker(12)) = True Then '���ܼ춨
            Cells(autorow(j), 14) = attaker(0) & "��" & defender(0) & "���𹥻�, " & defender(0) & "����"
        ElseIf rhp = 0 Then
            Cells(autorow(j), 14) = attaker(0) & "��" & defender(0) & "���" & damlog & "�˺�" & defender(0) & "��ʣ��Ѫ��Ϊ:0"
            Exit For
        Else
            If defender(17) > 0 Then
                defender(17) = defender(17) - 1 '���ط��غ��Ƴ���ѣ,value:0
            End If
            Cells(autorow(j), 14) = attaker(0) & "��" & defender(0) & "���" & damlog & "�˺�," & defender(0) & "��ʣ��Ѫ��Ϊ:" & rhp
        End If
            defender(2) = rhp '���ط�Ѫ������
            '��������
            exchanger = attaker()
            attaker = defender()
            defender = exchanger()
            ReDim exchanger(17)
        Next
    Else
        Cells(autorow(j), 14) = death(attaker(2), defender(2))
        Cells(autorow(j), 14) = "---------ս������--------"
        Exit For
    End If
Next

roundcheck attaker(), defender()
Erase characterA(), characterB()
Erase attaker(), defender(), exchanger()
End Sub
'ս������
Sub roundcheck(a(), b())
    Dim x As Variant
    Dim y As Variant
    x = a(0)
    y = b(0)
    Cells(1, 14) = x
    Cells(2, 14) = y
End Sub
'�����춨
Function isdeath(hp As Variant) As Boolean
    If hp = 0 Then
        isdeath = True
    Else
        isdeath = False
    End If
End Function
'���ܼ춨
Function ismiss(def As Variant, atk As Variant) As Boolean
    If attribution(def, atk) = True Then
        ismiss = True
    Else
        ismiss = False
    End If
End Function
'�����춨
Function iscrit(atk As Variant, def As Variant) As Single
    If attribution(atk, def) = True Then
        iscrit = 2
    Else
        iscrit = 1
    End If
End Function
'���μ춨
Function isdizzy(a As Variant) As Boolean
    If a > 0 Then
        isdizzy = True
    Else
        isdizzy = False
    End If
End Function
'�����춨
Function iscombo(atk As Variant, def As Variant) As Boolean
    If attribution(atk, def) = True Then
        iscombo = True
    Else
        iscombo = False
    End If
End Function
'�����춨
Function iscounter(def As Variant, atk As Variant) As Boolean
    If attribution(def, atk) Then
        iscombo = True
    Else
        iscombo = False
    End If
End Function
'��Ѫ�춨
Function livesteal(atk As Variant, def As Variant) As Variant
    If attribution(atk, def) = True Then
        livesteal = 1 + atk - def
    Else
        livesteal = 1
    End If
End Function
'ս���˺�
Function damage(atk As Variant, def As Variant) As Single
    random = Application.RandBetween(95, 105)
    If (atk > def) Then
    damage = Round((atk - def) * random / 100)
    Else
    damage = Round(atk * random / 100 * 0.01)
    End If
End Function
'ս������-���ط��ط�Ѫ��
Function battleround(damg As Variant, rhpdef As Single) As Single
    rhpdef = rhpdef - damg
    If rhpdef <= 0 Then
        battleround = 0
    Else
        battleround = rhpdef
    End If
End Function
'�Ƿ����ս��
Function death(a As Variant, b As Variant) As String
    If a = 0 Then
        death = "A����"
    ElseIf b = 0 Then
        death = "B����"
    End If
End Function
'�߼����Լ춨
Function attribution(a As Variant, b As Variant) As Boolean
    random1 = Application.RandBetween(0, 9999)
    If (a - b) * 10000 >= random1 Then
        attribution = True '�춨ͨ��
    Else
        attribution = False '�춨ʧ��
    End If
End Function
'����
Function autorow(a As Integer) As Integer
    a = a + 1
    autorow = a
End Function