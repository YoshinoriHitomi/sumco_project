Attribute VB_Name = "s_cmmc001z"

'�T�v      :�ΐ͌v�Z�ɕK�v�Ȋe���v�d�ʎ��т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :CRYNUM        ,I  ,String    ,�����ԍ�
'          :wgtCharge     ,O  ,Long      ,�F���ʁi����`���[�W�ʁ|�O��܂ł̈��グ�d�ʁ|�O��܂ł�į�߶�ďd�ʁj
'          :wgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
'          :wgtTopCut     ,O  ,Double    ,�g�b�v�J�b�g�d�ʎ��ђl
'          :DM            ,O  ,Double    ,���a�P�`�R�̕���
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�P�{�����A�c�ʈ����ɂ��킹�Ď��уf�[�^���擾����
'����      :2001/8/29 �쐬  �쑺
Public Function GetCoeffParams(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset

    On Error GoTo Err
    GetCoeffParams = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    sql = "select decode(RONAI,null,CHARGE,RONAI) as RONAI, WGHTTOP, WGTOPCUT, (DM1+DM2+DM3)/3.0 as DM " & _
          "from TBCMH004 H004, " & _
          "  (select sum(CHARGE) - sum(UPWEIGHT) - sum(WGTOPCUT) as RONAI" & _
          "   From TBCMH004" & _
          "   where (CRYNUM<'" & CRYNUM & "')" & _
          "    and  (substr(CRYNUM,1,7)='" & left$(CRYNUM, 7) & "')" & _
          "  ) SUMDATA " & _
          "where (CRYNUM='" & CRYNUM & "')"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("RONAI")
        wgtTop = rs("WGHTTOP")
        wgtTopCut = rs("WGTOPCUT")
        DM = rs("DM")
    End If
    rs.Close
    
    GetCoeffParams = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function


'�T�v      :��R�l�ɑ΂���ʒu�𐄒肷��B
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :d             ,IO ,type_ResPosCal ,����v�Z�\����
'          :�߂�l        ,O  ,Double         ,����ʒu
'����      :
'����      :2001/06/23�@���� �M�Ɓ@�쐬
Public Function PosCalculation(d As type_ResPosCal) As Double
    Dim GS As Double        '��Top�ʒu���グ��
    Dim Ro As Double        '���R�l
    Dim Gx As Double
    
    On Error GoTo Err
    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    Ro = d.TOPRES * ((1 - GS) ^ (d.COEFFICIENT - 1))
    Gx = 1 - ((Ro / d.target) ^ (1 / (d.COEFFICIENT - 1)))
    
    PosCalculation = ((d.CHARGEWEIGHT - d.TOPWEIGHT) * Gx) / (d.DUNMENSEKI * HIJU_SILICONE)
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    PosCalculation = -9999
End Function

'�T�v      :�ʒu�ɑ΂����R�l�𐄒肷��B
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'          :d             ,IO ,type_ResPosCal ,����v�Z�\����
'          :�߂�l        ,O  ,Double         ,�����R�l
'����      :
'����      :2001/06/23�@���� �M�Ɓ@�쐬
Public Function ResCalculation(d As type_ResPosCal) As Double
    Dim GS As Double        '��Top�ʒu���グ��
    Dim Ro As Double        '���R�l
    Dim Gx As Double        '����Ώۈ��グ��

    On Error GoTo Err
    GS = (d.DUNMENSEKI * HIJU_SILICONE * d.TOPSMPLPOS) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    Ro = d.TOPRES * (1 - GS) ^ (d.COEFFICIENT - 1)
    Gx = d.DUNMENSEKI * d.target * HIJU_SILICONE / (d.CHARGEWEIGHT - d.TOPWEIGHT)

    ResCalculation = Ro / (1 - Gx) ^ (d.COEFFICIENT - 1)
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    ResCalculation = -9999
End Function

'�T�v      :�ΐ͌W�������߂�B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :d             ,IO ,type_Coefficient ,�ΐ͌W���v�Z�\����
'          :�߂�l        ,O  ,Double           ,�ΐ͌W��
'����      :
'����      :2001/06/23�@���� �M�Ɓ@�쐬
Public Function CoefficientCalculation(d As type_Coefficient) As Double
    Dim GT As Double
    Dim GB As Double
    
    On Error GoTo Err
    GT = (d.DUNMENSEKI * d.TOPSMPLPOS * HIJU_SILICONE) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    GB = (d.DUNMENSEKI * d.BOTSMPLPOS * HIJU_SILICONE) / (d.CHARGEWEIGHT - d.TOPWEIGHT)
    
    CoefficientCalculation = Log(d.BOTRES / (d.TOPRES * 1)) / Log((1 - GT) / (1 - GB)) + 1
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CoefficientCalculation = -9999
End Function


'�T�v      :�V���R���~���̏d�ʂ����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dblDiameter   ,I  ,Double    ,���a(mm)
'          :dblHeight     ,I  ,Double    ,����(mm)
'          :�߂�l        ,O  ,Double    ,�d��(g)
'����      :
'����      :2001/06/29 �쐬  �쑺
Public Function WeightOfCylinder(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCylinder = HIJU_SILICONE * cdblPI * (dblRadius ^ 2) * dblHeight
End Function


'�T�v      :�V���R���~���̏d�ʂ����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dblDiameter   ,I  ,Double    ,���a(mm)
'          :dblHeight     ,I  ,Double    ,����(mm)
'          :�߂�l        ,O  ,Double    ,�d��(g)
'����      :TOP�EBOT�d�ʂ̌v�Z�p
'����      :2001/06/29 �쐬  �쑺
Public Function WeightOfCone(ByVal dblDiameter As Double, ByVal dblHeight As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    WeightOfCone = HIJU_SILICONE * (cdblPI * (dblRadius ^ 2) * dblHeight) / 3#
End Function


'�T�v      :�~�̖ʐς����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dblDiameter   ,I  ,Double    ,���a(mm)
'          :�߂�l        ,O  ,Double    ,�ʐ�(mm2)
'����      :
'����      :2001/07/05 �쐬  �쑺
Public Function AreaOfCircle(ByVal dblDiameter As Double) As Double
Dim dblRadius As Double

    dblRadius = dblDiameter / 2#
    AreaOfCircle = cdblPI * (dblRadius ^ 2)
End Function


'�T�v      :�ΐ͌v�Z�ɕK�v�Ȋe���v�d�ʎ��т��擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :CRYNUM        ,I  ,String    ,�����ԍ�
'          :wgtCharge     ,O  ,Long      ,�F���ʁi����`���[�W�ʁ|�O��܂ł̈��グ�d�ʁ|�O��܂ł�į�߶�ďd�ʁj
'          :wgtTop        ,O  ,Double    ,�g�b�v�d�ʎ��ђl
'          :wgtTopCut     ,O  ,Double    ,�g�b�v�J�b�g�d�ʎ��ђl
'          :DM            ,O  ,Double    ,���a�P�`�R�̕���
'          :�߂�l        ,O  ,FUNCTION_RETURN,
'����      :�y�}���`����Ή��z �S�ʈ�����c�ʈ����RC�����ɂ��킹�Ď��уf�[�^���擾����
'����      :2008/4/21 �쐬  SETsw Nakada
Public Function GetCoeffParams_new(ByVal CRYNUM$, wgtCharge As Long, wgtTop As Double, wgtTopCut As Double, DM As Double) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset

    On Error GoTo Err
    GetCoeffParams_new = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    '' ����`���[�W�A�d�ʁiTOP�j�A�g�b�v�J�b�g�d�ʁA�������a�̕��ϒl �擾
    sql = " SELECT C1.SUICHARGE, C1.WGHTTOC1, C1.PUTCUTWC1, "
    sql = sql & " (C1.DIA1C1 + C1.DIA2C1 + C1.DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 C1 "
    sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("SUICHARGE")       ''����`���[�W
        wgtTop = rs("WGHTTOC1")           ''�d�ʁiTOP�j
        wgtTopCut = rs("PUTCUTWC1")       ''�g�b�v�J�b�g�d��
        DM = rs("DM")                     ''�������a(���ϒl)
    End If
    rs.Close
    
    GetCoeffParams_new = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function
