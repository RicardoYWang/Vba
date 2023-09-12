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
'属性赋值
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
'战斗中间值
' 力量  1
' 生命  2
' 体质  3
' 敏捷  4
' 连击  5
' 连击抵抗  6
' 反击  7
' 反击抵抗  8
' 暴击  9
' 暴击抵抗  10
' 闪避  11
' 闪避抵抗  12
' 击晕  13
' 击晕抵抗  14
' 吸血  15
' 吸血抵抗  16
' 击晕标记 17
Dim damlog As String, deflog As String
Dim rhp As Single, dam As Single
Dim j As Integer
j = 6

'====战斗过程=====

Sheets("battle").Range("m:N").Clear
'根据敏捷确定出手权,定义攻防角色
If characterA(4) > characterB(4) Then
  
    attaker = characterA()
    defender = characterB()
    Cells(5, 13) = "A速度大于B,A优先行动"
    
Else
    attaker = characterB()
    defender = characterA()
    Cells(5, 13) = "A速度不大于B,B优先行动"
End If
rhp = defender(2)
For i = 1 To 16
    '超过规定回合双方都没死
    If i = 16 Then
        Cells(autorow(j), 13) = "---------战斗结束--------"
        Cells(autorow(j), 13) = "A在15回合内未击败B,B获胜"
        '回合进行
    Else
            Cells(autorow(j), 13) = "-------------第" & i & "回合-------------"
            ' If defender(17) > 0 Then
            '     defender(17) = defender(17) - 1 '防守方回合移除晕眩,value:0
            ' End If
            For e = 1 To 2
            If rhp = 0 Then '一方死亡结束
                Exit For
            End If
            If attaker(17) > 0 Then '攻方被击晕则跳过攻方回合
                Cells(autorow(j), 14) = attaker(0) & "(晕眩),无法行动"
                attaker(17) = attaker(17) - 1 '晕眩状态,value:1
            Else
                If defender(17) = 0 And ismiss(defender(11), attaker(12)) = True Then '闪避检定
                    Cells(autorow(j), 14) = attaker(0) & "<攻击>" & defender(0) & "<闪避>"
                Else
                    If defender(17) > 0 Then
                        defender(17) = defender(17) - 1 '防守方回合移除晕眩,value:0
                    End If
                    rhp = defender(2)
                    dam = damage(attaker(1), defender(3))
                    damlog = "造成" & dam
                    deflog = "<攻击>" & defender(0)
                    If attribution(attaker(9), defender(10)) = True Then '暴击检定
                    dam = dam * 2
                    damlog = "造成" & dam & "(暴击!)"
                    End If
                    If attribution(attaker(13), defender(14)) = True Then '击晕检定
                        defender(17) = 2 '附加晕眩效果,value:2
                        deflog = "<击晕>" & defender(0)
                    End If
                        rhp = hp(dam, rhp)
                    Cells(autorow(j), 14) = attaker(0) & deflog & damlog & "伤害," & defender(0) & "剩余血量为:" & rhp & " "
                    defender(2) = rhp  '防守方血量更新
                    if defender(2)>0 and attribution(defender(7), attaker(8))=true then '反击检定前提为防守方存活
                        rhp = attaker(2)
                        dam = damage(defender(1), attaker(3))
                        damlog = "造成" & dam
                        deflog = "<反击>" & attaker(0)
                        If attribution(defender(9), attaker(10)) = True Then '反击暴击检定
                        dam = dam * 2
                        damlog = "造成" & dam & "(暴击!)"
                        End If
                        rhp = hp(dam, rhp)
                        Cells(autorow(j), 14) = defender(0) & deflog & damlog & "伤害," & attaker(0) & "剩余血量为:" & rhp & " "
                        attaker(2)=rhp
                        If rhp = 0 Then '一方死亡结束
                        Exit For
                        End If
                    end if
                End If
            End If
            rhp= defender(2)'血量参数更新为防守方,用于循环判断标记
            '攻防换边
            exchanger = attaker()
            attaker = defender()
            defender = exchanger()
            ReDim exchanger(17)

        Next
        If isdeath(attaker(2)) = True Or isdeath(defender(2)) = True Then   '一方死亡结束
            Cells(autorow(j), 14) = death(attaker(), defender())
            Cells(autorow(j), 13) = "---------战斗结束--------"
            Exit For
        End If

    End If
Next
End Sub
'攻击流程
Function atkloop(a(), d())

End Function
'死亡检定
Function isdeath(hp As Variant) As Boolean
    If hp = 0 Then
        isdeath = True
    Else
        isdeath = False
    End If
End Function
'闪避检定
Function ismiss(def As Variant, atk As Variant) As Boolean
    If attribution(def, atk) = True Then
        ismiss = True
    Else
        ismiss = False
    End If
End Function
'暴击检定
Function iscrit(atk As Variant, def As Variant) As Single
    If attribution(atk, def) = True Then
        iscrit = 2
    Else
        iscrit = 1
    End If
End Function
'击晕检定
Function isdizzy(a As Variant) As Boolean
    If a > 0 Then
        isdizzy = True
    Else
        isdizzy = False
    End If
End Function
'连击检定
Function iscombo(atk As Variant, def As Variant) As Boolean
    If attribution(atk, def) = True Then
        iscombo = True
    Else
        iscombo = False
    End If
End Function
'反击检定
Function iscounter(def As Variant, atk As Variant) As Boolean
    If attribution(def, atk) Then
        iscombo = True
    Else
        iscombo = False
    End If
End Function
'吸血检定
Function livesteal(atk As Variant, def As Variant) As Variant
    If attribution(atk, def) = True Then
        livesteal = 1 + atk - def
    Else
        livesteal = 1
    End If
End Function
'战斗伤害
Function damage(atk As Variant, def As Variant) As Single
    random = Application.RandBetween(95, 105)
    If (atk > def) Then
    damage = Round((atk - def) * random / 100)
    Else
    damage = Round(atk * random / 100 * 0.01)
    End If
End Function
'战斗过程-返回防守方血量
Function hp(damg As Single, rhpdef As Single) As Single
    rhpdef = rhpdef - damg
    If rhpdef <= 0 Then
        hp = 0
    Else
        hp = rhpdef
    End If
End Function
'是否结束战斗
Function death(a(), b()) As String
    If a(2) = 0 Then
        death = a(0) & "死亡"
    ElseIf b(2) = 0 Then
        death = b(0) & "死亡"
    End If
End Function
'高级属性检定
Function attribution(a As Variant, b As Variant) As Boolean
    random1 = Application.RandBetween(0, 9999)
    If (a - b) * 10000 >= random1 Then
        attribution = True '检定通过
    Else
        attribution = False '检定失败
    End If
End Function
'自增
Function autorow(a As Integer) As Integer
    a = a + 1
    autorow = a
End Function
