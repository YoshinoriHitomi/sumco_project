Attribute VB_Name = "s_cmzcjudg"
Option Explicit

'Public Enum FUNCTION_RETURN                 ''�֐��̖߂�l
'    FUNCTION_RETURN_SUCCESS = 0             '' ����
'    FUNCTION_RETURN_FAILURE = -1            '' �ُ�
'End Enum

Public Const EZJ00 = "EZJ00" ''����ΏۃR�[�h %s �ɂ͑Ή����Ă܂���B
Public Const ZJ001 = "ZJ001" ''�������@�R�[�h %s �́A�����ł��B
Public Const ZJ002 = "ZJ002" ''����_���ُ��BMD�ő�l�����߂��܂���B

Public Const JUDG_OK = True                 ''���茋�ʂ�OK�̏ꍇ�A���l
Public Const JUDG_NG = False                ''���茋�ʂ�NG�̏ꍇ�A���l

Public Const OI_JUDG = "�_�f�Z�x����"       ''���荀�ڕ�����(Oi)
Public Const CS_JUDG = "�Y�f�Z�x����"       ''���荀�ڕ�����(Cs)
Public Const RES_JUDG = "���R����"        ''���荀�ڕ�����(���R)
Public Const GFA_JUDG = "GFA����"           ''���荀�ڕ�����(GFA)
Public Const BMD_JUDG = "BMD����"           ''���荀�ڕ�����(BMD)
Public Const OSF_JUDG = "OSF����"           ''���荀�ڕ�����(OSF)
Public Const DEN_JUDG = "DEN����"           ''���荀�ڕ�����(Den)
Public Const LDL_JUDG = "L/DL����"          ''���荀�ڕ�����(L/DL)
Public Const DVD2_JUDG = "DVD2����"         ''���荀�ڕ�����(DVD2)
Public Const LT_JUDG = "���C�t�^�C������"   ''���荀�ڕ�����(���C�t�^�C��)
Public Const EPD_JUDG = "EPD����"           ''���荀�ڕ�����(EPD)
Public Const DOI_JUDG = "��Oi����"          ''���荀�ڕ�����(��Oi)
Public Const DZ_JUDG = "DZ����"             ''���荀�ڕ�����(DZ)
Public Const DSOD_JUDG = "DSOD����"         ''���荀�ڕ�����(DSOD)
Public Const SPV_JUDG = "SPV����"           ''���荀�ڕ�����(SPV)
Public Const AOI_JUDG = "AOi����"           ''���荀�ڕ�����(AOi)�@03/12/09 ooba

Public Const WFRES_JUDG = 1                 ''���莯�ʃt���O(RES)
Public Const WFOI_JUDG = 2                  ''���莯�ʃt���O(Oi)
Public Const WFDOI_JUDG = 3                 ''���莯�ʃt���O(��Oi)
Public Const WFOSF_JUDG = 4                 ''���莯�ʃt���O(OSF)
Public Const WFBMD_JUDG = 5                 ''���莯�ʃt���O(BMD)
Public Const WFDZ_JUDG = 6                  ''���莯�ʃt���O(DZ)
Public Const WFDSOD_JUDG = 7                ''���莯�ʃt���O(DSOD)
Public Const WFSPV_JUDG = 8                 ''���莯�ʃt���O(SPV)
Public Const WFAOI_JUDG = 9                 ''���莯�ʃt���O(AOi)�@03/12/09 ooba

Public Const ObjCode01 = "1"                ''���S����l
Public Const ObjCode02 = "2"                ''����l�̒����l
Public Const ObjCode03 = "3"                ''�S����_
Public Const ObjCode04 = "6"                ''R/2
Public Const ObjCode05 = "A"                ''�S�_�̕��ϒl
Public Const ObjCode06 = "B"                ''�S�_�̍ő�l
Public Const ObjCode07 = "C"                ''�S�_�̕��ϒl�ƍő�l
Public Const ObjCode08 = "D"                ''�S�_�̍ŏ��l
Public Const ObjCode09 = "E"                ''������2�_�A�O����2�_(5�_�����1,2,4,5)
Public Const ObjCode10 = "F"                ''MAX(2,4�_��)
Public Const ObjCode11 = "G"                ''MAX(2,3,4�_��)
Public Const ObjCode12 = "H"                ''���ۏ�
Public Const ObjCode13 = "N"                ''�_��
Public Const ObjCode14 = "Z"                ''�`�󑪒�(���R�x�A���Ԃ�AWARP)
Public Const ObjCode15 = " "                ''�K�i�Ȃ�
Public Const ObjCode16 = "K"                ''2001/09/19 S.Sano �S�_�̍ŏ��l�ƍő�l
Public Const ObjCode17 = "L"                ''AVE+MIN�@08/03/13 ooba
Public Const ObjCode18 = "7"                ''AVE+�O��1�_   '' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� ADD By Systech
Public Const ObjCodeGrp01 = "1236"          ''FTIR�AGFA�AWF���R�AWF�_�f�Z�x�A��Oi
'-----TEST2004/10
Public Const ObjCodeGrp05 = "1236N"          ''���R
'2001/09/19 S.SanoPublic Const ObjCodeGrp02 = "ABCDFG"        ''BMD�ADZ

'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� UPD By Systech Start
''Public Const ObjCodeGrp02 = "ABCDFGK"        ''2001/09/19 S.Sano BMD�ADZ
Public Const ObjCodeGrp02 = "ABCDFGK7"      ''BMD�ADZ
'' 2008/10/20 BMD�]��,�O��1�_�ۏ؋@�\�ǉ� UPD By Systech End

Public Const ObjCodeGrp03 = "ABC"           ''OSF
Public Const ObjCodeGrp04 = "3"             ''DSOD�ASPV
Public Const ObjCodeGrp06 = "ABCDFGK3"        ''2004/12/15 S.Sano BMD�ADZ

Public Const PosCode01 = "E"                ''
Public Const PosCode02 = "G"                ''
Public Const PosCode03 = "H"                ''
Public Const PosCode04 = "J"                ''
Public Const PosCode05 = "M"                ''
Public Const PosCode06 = "Q"                ''
Public Const PosCode07 = "N"                ''
Public Const PosCode08 = "P"                ''
Public Const PosCode09 = "R"                ''
Public Const PosCodeGrp01 = "EGHJMQNPRR"    ''

Public Const JudgCodeC01 = "H"              ''��������L��
Public Const JudgCodeC02 = "BS X"           ''�������薳���A�S��OK
Public Const JudgCodeW01 = "H"              ''WF����L��
Public Const JudgCodeW02 = "BS H"           ''WF���薳���A�S��OK
Public Const JudgCodeW03 = "XS"             ''����L��
'----- TEST2004/10
Public Const KSTAFF_J002 = "SHIJI"                    '

'Add Start 2011/01/26 SMPK Miyata
''Cu-deco ���i�d�l�p�^�[���敪
Public Const CudecoSpcPtnNR = "1"           '' �����O�����w��
Public Const CudecoSpcPtnND = "2"           '' �f�B�X�N�����w��
Public Const CudecoSpcPtnNP = "3"           '' �p�^�[�������w��
Public Const CudecoSpcPtnN = "4"            '' �s�� (�I���Ȃ�)
Public Const CudecoSpcPtnNB = "5"           '' �o���h�����w��
Public Const CudecoSpcPtnNPB = "6"          '' P�o���h�����w��
Public Const CudecoSpcPtnNBB = "7"          '' B�o���h�w�薳��
''Cu-deco ���уp�^�[���敪
Public Const CudecoJskPtnN = "0"            '' None
Public Const CudecoJskPtnR = "1"            '' Ring
Public Const CudecoJskPtnD = "2"            '' Disk
Public Const CudecoJskPtnDR = "3"           '' Disk & Ring
Public Const CudecoJskPtnPB_B = "5"         '' PB-band
Public Const CudecoJskPtnP_B = "6"          '' P-band
Public Const CudecoJskPtnB_B = "7"          '' B-band
'Add End   2011/01/26 SMPK Miyata


'�i���ۏ؏��\����
Public Type Guarantee
    cMeth  As String                        ''����ʒu_��
    cCount As String                        ''����ʒu_�_
    cPos   As String                        ''����ʒu_��(OSF�̏ꍇ ��)
    cObj   As String                        ''�ۏؕ��@_��
    cJudg  As String                        ''�ۏؕ��@_��
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga START   ---
'    cJudg2  As String                       ''�ۏؕ��@_��
'C�|OSF3����@�\�ǉ� 2007/04/23 M.Kaga END     ---
    cBunp  As String                        ''�ʓ����z�v�Z��    ' Res�COi �ʓ����z�v�Z���ǉ��˗�  No.030205  yakimura  2003.06.06
End Type
'�G���[���\����
Public Type ERROR_INFOMATION
    ErrCode     As Variant                  ''�G���[�R�[�h
    ErrStr(4)   As Variant                  ''�I�v�V����������
End Type

'�T�v      :�z����̍ŏ��l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function JudgMin(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    temp = d(0)
    For c0 = 1 To High
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMin = temp
End Function

'�T�v      :�z����̍ő�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function JudgMax(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    temp = d(0)
    For c0 = 1 To High
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMax = temp
End Function

'�T�v      :�z����f�[�^�̕��ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function JudgAve(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    temp = 0
    c1 = 0
    For c0 = 0 To High
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve = 0
    Else
        JudgAve = temp / c1
    End If
End Function

'�T�v      :�z����f�[�^�̕��ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :�z����f�[�^�����ׂ�NULL(-1)�̏ꍇ-1��Ԃ��B
'����      :�V�K�쐬 2005/06/22 ffc)tanabe
Public Function JudgAve3(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    temp = 0
    c1 = 0
    For c0 = 0 To High
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve3 = -1
    Else
        JudgAve3 = temp / c1
    End If
End Function

'�T�v      :�z����f�[�^�̒����l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�����l
'����      :�����l�Ƃ́A�ő�l�ƍŏ��l�̒��S�l�ɍł��߂��l�B
'����      :2001/06/06 ���� �M�� �쐬
Public Function JudgCenter(d() As Double) As Double
'�Z�F��
    Dim High As Integer
    Dim temp() As Double
    Dim c0 As Integer
    Dim c1 As Integer
    
    c1 = 0
    For c0 = 0 To UBound(d)
        If d(c0) <> -1 Then
            ReDim Preserve temp(c1) As Double
            temp(c1) = d(c0)
            c1 = c1 + 1
        End If
    Next
    
    If c1 <> 0 Then
        BubbleSort temp()
        
        High = UBound(temp)
        JudgCenter = temp(Int((High + 1) / 2))
    Else
        JudgCenter = -9999
    End If

'�O�H��
'    Dim c0 As Integer
'    Dim temp As double
'    Dim temp1 As double
'    Dim temp2 As double
'    Dim Center As double
'    Dim High As Integer
'
'    '' �z��̏�����擾���܂��B
'    High = UBound(d)
'
'    ''�ő�l�����߂�B
'    temp1 = JudgMax(d())
'    ''���ݍŏ��̐�Βl�����߂�����l�Ƃ���B
'    temp2 = temp1
'
'    ''���S�l�����߂�B
'    Center = (JudgMin(d()) + temp1) / 2
'
'    ''�ő�l�ƒ��S�l�̐�Βl�����߂�B
'    ''���ݍŏ��̐�Βl�Ƃ���B
'    temp1 = Abs(temp1)
'
'    ''�ő�l�ƍŏ��l�ɍł��߂�����l�����߂�B
'    For c0 = 0 To High
'        If d(c0) <> -1 Then
'            ''����l�ƒ��S�l�̐�Βl�����߂�B
'            temp = Abs(Center - d(c0))
'            ''�O�񋁂߂���Βl�ƍ��񋁂߂���Βl���r���A
'            ''���񋁂߂���Βl�������������ꍇ�B
'            If temp < temp1 Then
'                ''���񋁂߂���Βl�����ݍŏ��̐�Βl�Ƃ���B
'                temp1 = temp
'                ''���ݍŏ��̐�Βl�����߂�����l�Ƃ���B
'                temp2 = d(c0)
'            End If
'        End If
'    Next
'
'    ''���S�l�Ƃ̐�Βl���ł�����������l��Ԃ��B
'    JudgCenter = temp2
End Function

'�T�v      :�z��̃R�s�[���쐬����B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :d1()          ,O  ,double    ,����l�̃R�s�[
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Sub DataCopy(d() As Double, d1() As Double)
    Dim High As Integer
    Dim c0 As Integer
    
    '' �z��̏�����擾���܂��B
    High = UBound(d)
    
    ''��2�����ɑ�1�������R�s�[���܂��B
    For c0 = 0 To High
        d1(c0) = d(c0)
    Next
End Sub

'�T�v      :�o�u���\�[�g���s���܂��B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,IO ,double    ,����l�̃R�s�[
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Sub BubbleSort(d() As Double)
    Dim High As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    '' �z��̏�����擾���܂��B
    High = UBound(d)
    
    '' �z��̌X�̗v�f�ɑ΂��ČJ��Ԃ��������܂��B
    For c0 = 0 To High - 1
        '' �z��̌X�̗v�f�ɑ΂��ČJ��Ԃ��������܂��B
        For c1 = c0 + 1 To High
            '' �z��̑O���ɂ���l���A�z��̌���ɂ���l���
            '' �傫���ꍇ�ɂ́A�������������܂��B
            If d(c0) > d(c1) Then
                temp = d(c0)
                d(c0) = d(c1)
                d(c1) = temp
            End If
        Next
    Next
End Sub


'�T�v      :Side�f�[�^�̕��ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :
'����      :2003/06/06 yakimura �쐬
Public Function JudgSideAve(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim cnt As Integer
    Dim temp As Double
    Dim High As Integer
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    temp = 0
    cnt = 0
    For c0 = 1 To High
        If d(c0) <> -1 Then
            cnt = cnt + 1
            temp = temp + d(c0)
        End If
    Next
    If cnt = 0 Then
        JudgSideAve = 0
    Else
        JudgSideAve = temp / cnt
    End If
End Function

'�T�v      :�z���(0�`2)�̍ŏ��l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function JudgMin2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(0)
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMin2 = temp
End Function

'�T�v      :�z���(0�`2)�̍ő�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function JudgMax2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(0)
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgMax2 = temp
End Function

'�T�v      :�z���(0�`2)�f�[�^�̕��ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function JudgAve2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim c1 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = 0
    c1 = 0
    For c0 = 0 To 2
        If d(c0) <> -1 Then
            c1 = c1 + 1
            temp = temp + d(c0)
        End If
    Next
    If c1 = 0 Then
        JudgAve2 = 0
    Else
        JudgAve2 = temp / c1
    End If
End Function

'�T�v      :Side(1�`2)�f�[�^�̕��ϒl�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :
'����      :2003/06/06 yakimura �쐬
Public Function JudgSideAve2(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim cnt As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = 0
    cnt = 0
    For c0 = 1 To 2
        If d(c0) <> -1 Then
            cnt = cnt + 1
            temp = temp + d(c0)
        End If
    Next
    If cnt = 0 Then
        JudgSideAve2 = 0
    Else
        JudgSideAve2 = temp / cnt
    End If
End Function

'�T�v      :Side(1�`4)�̍ŏ��l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function JudgSideMin(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(1)
    For c0 = 2 To 4
        If d(c0) <> -1 Then
            If d(c0) < temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgSideMin = temp
End Function

'�T�v      :Side(1�`4)�̍ő�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,�ő�l
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function JudgSideMax(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim temp As Double
    Dim High As Integer
    
    temp = d(1)
    For c0 = 2 To 4
        If d(c0) <> -1 Then
            If d(c0) > temp Then
                temp = d(c0)
            End If
        End If
    Next
    
    JudgSideMax = temp
End Function


'�T�v      :�G���[���\���̂ɒl��������B
'���Ұ�    :�ϐ���        ,IO ,�^               ,����
'          :ErrInfo       ,O  ,ERROR_INFOMATION ,�G���[���\����
'          :[Code]        ,I  ,String           ,�I�v�V�����G���[�R�[�h
'          :[Str1]        ,I  ,String           ,�I�v�V����������P
'          :[Str2]        ,I  ,String           ,�I�v�V����������Q
'          :[Str3]        ,I  ,String           ,�I�v�V����������R
'          :[Str4]        ,I  ,String           ,�I�v�V����������S
'          :[Str5]        ,I  ,String           ,�I�v�V����������T
'          :�߂�l        ,O  ,FUNCTION_RETURN  ,FUNCTION_RETURN_FAILURE
'����      :�K���֐��ُ�I���R�[�h��Ԃ��B
'�@�@      :�I�v�V���������́A�ȗ������ꍇ�A""����������B
'�@�@      :�I�v�V�����������A�S�ďȗ������ꍇ�A�G���[���\���̂̏��������s����B
'�@�@      :�I�v�V�����������A�S�ďȗ������ꍇ�A�߂�l�́A��������B
'����      :2001/06/06 ���� �M�� �쐬
Public Function SetErrInfo(ErrInfo As ERROR_INFOMATION, Optional CODE, Optional Str1, Optional Str2, Optional Str3, Optional Str4, Optional Str5) As FUNCTION_RETURN
    ErrInfo.ErrCode = CODE
    ErrInfo.ErrStr(0) = Str1
    ErrInfo.ErrStr(1) = Str2
    ErrInfo.ErrStr(2) = Str3
    ErrInfo.ErrStr(3) = Str4
    ErrInfo.ErrStr(4) = Str5
    SetErrInfo = FUNCTION_RETURN_FAILURE
End Function

'�T�v      :����l�͈͔̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :JudgData      ,I  ,double    ,����l
'          :SpecMin       ,I  ,double    ,�����l
'          :SpecMax       ,I  ,double    ,����l
'          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function RangeDecision(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
End Function

'�T�v      :OSF�̃p�^�[���̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Kubun         ,I  ,string    ,�p�^�[���敪
'          :JudgData()    ,I  ,string    ,�p�^�[������
'          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'����      :
'����      :2003/05/17�@ooba
Public Function JudgPattern(KUBUN As String, JudgData() As String * 1) As Boolean
    Dim ct As Integer
    Dim RD As String
    
    '�y�p�^�[���敪�z�@1�F�����O�����@2�F�f�B�X�N�����@3�F�p�^�[�������@4�F�s��
    Select Case KUBUN
        Case "1"
            RD = "D "
        Case "2"
            RD = "R "
        Case "3"
            RD = " "
        Case "4", " "
            RD = "RD "
    End Select
    
    '�p�^�[���敪�ɊY�����镶����ƃp�^�[�����т��ׁA������s���B
    For ct = 0 To 2
        JudgPattern = (InStr(RD, JudgData(ct)) > 0)
        If JudgPattern = False Then
            Exit For
        End If
    Next
End Function

'�T�v      :DSOD������݂̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :Kubun         ,I  ,string    ,����݋敪
'          :JudgData()    ,I  ,string    ,����ݎ���
'          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'����      :
'����      :2004/07/28�@ooba
Public Function JudgDsodPattern(KUBUN As String, JudgData() As String * 3) As Boolean

    Dim iCnt As Integer
    Dim sPtn As String
    
    'DSOD����ݎ��ь��ʂɑ΂��āA�d�l�������݋敪���ނƔ�����s���B
    For iCnt = 0 To 1
    
        JudgDsodPattern = False
        sPtn = Trim(JudgData(iCnt))
        
        Select Case KUBUN
            Case "1"        '�ݸޖ���
                If sPtn = "" Or sPtn = "D" Then
                    JudgDsodPattern = True
                End If
            Case "2"        '�ި������
                If sPtn = "" Or sPtn = "R" Then
                    JudgDsodPattern = True
                End If
            Case "3"        '����ݖ���
                If sPtn = "" Then
                    JudgDsodPattern = True
                End If
            Case "4", " "   '�s��
                If sPtn = "" Or sPtn = "R" Or sPtn = "D" Or sPtn = "R,D" Then
                    JudgDsodPattern = True
                End If
        End Select
        If JudgDsodPattern = False Then
            Exit For
        End If
    Next
    
End Function

'�T�v      :BMD�̖ʓ����z�v�Z(�ʓ����z"P")���s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :�߂�l        ,O  ,double    ,���ϒl
'����      :
'����      :2003/05/21�@ooba
Public Function JudgBmdMBP(d() As Double, Optional iMax As Integer) As Double
    Dim c0 As Integer
    Dim side As Double
    Dim center As Double
    Dim High As Integer
    Dim deverrflag As Boolean
    
    deverrflag = False
    JudgBmdMBP = -1
    
    If iMax > 0 Then
        High = iMax - 1
    Else '' �z��̏�����擾���܂��B
        High = UBound(d)
    End If
    
    '�z��̍ŏ��̒l��side�ɃZ�b�g
    side = d(0)
    '�z��̍Ō�̒l��center�ɃZ�b�g
    For c0 = 1 To High
        If d(c0) <> -1 Then
            center = d(c0)
        End If
    Next
    
    'side��senter���ׁA�傫�����𕪎q��
    If side < center Then
        If side > 0 Then
            JudgBmdMBP = center / side * 100
        ElseIf side = 0 Then
            deverrflag = True
        End If
    Else
        If center > 0 Then
            JudgBmdMBP = side / center * 100
        ElseIf center = 0 Then
            deverrflag = True
        End If
    End If
    JudgBmdMBP = Round(JudgBmdMBP, 1)
    
    '�f�o�b�O�p�������R�����g�@2003/05/29 ooba
'    If deverrflag Then
'        WFCJudgDialog.WFCErrorMessage "���z�v�Z 0 ���Z�G���["
'    ElseIf JudgBmdMBP = -1 Then
'        WFCJudgDialog.WFCErrorMessage "����ʒu�A�Ώۃf�[�^�A���z�v�Z����"
'    End If
End Function
'�T�v      :RRG�AORG�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :d()           ,I  ,double    ,����l
'          :iMax          ,I  ,Integer   ,����_��
'          :�߂�l        ,O  ,double    ,RRG,ORG
'����      :
'����      :2001/06/06 ���� �M�� �쐬
Public Function RGCal(d() As Double, iMax As Integer) As Double
    Dim temp As Double
    Dim High As Integer
    
    '' �z��̏�����擾���܂��B
    High = UBound(d)
    If High < iMax - 1 Then
        RGCal = -1
        Exit Function
    End If
    
    temp = JudgMin(d(), iMax)
    If temp <> 0 Then
        temp = (JudgMax(d(), iMax) - temp) * 100 / temp
    End If
    
    RGCal = temp
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�T�v      :���R�C�_�f�Z�x�̖ʓ��v�Z�l�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,Mennai
'����      :
'����      :2003/06/06  yakimura  �쐬
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
'Public Function MENNAI_Cal(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double
Public Function MENNAI_Cal_NotRound(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double

Dim Min        As Double
Dim max        As Double
Dim AVE        As Double
Dim center     As Double
Dim side       As Double
Dim Side_Ave   As Double
Dim Mennai     As Double
Dim w_Max1     As Double
Dim w_Max2     As Double
Dim w_N1       As Double
Dim w_N2       As Double
Dim w_N3       As Double
Dim w_N4       As Double
Dim w_Nx       As Double
    
    Mennai = -1
    
    Select Case calcode

    Case "A"                '---> (max-min)/min�~100
        Select Case JudgFlag
           Case RES_JUDG ''���莯�ʃt���O(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''���莯�ʃt���O(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                Mennai = (max - Min) * 100 / Min
            Else
                Mennai = -1
            End If
        End If

    Case "B"                '---> (max-min)/max�~100
        Select Case JudgFlag
           Case RES_JUDG ''���莯�ʃt���O(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''���莯�ʃt���O(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (max <> 0) Then
                Mennai = (max - Min) * 100 / max
            Else
                Mennai = -1
            End If
        Else
            Mennai = -1
        End If

    Case "C"                '---> (max-min)/center�~100
        Select Case JudgFlag
           Case RES_JUDG ''���莯�ʃt���O(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''���莯�ʃt���O(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        center = R(0)
        If (Min <> -9999) And (max <> -9999) And (center <> -9999) Then
            If (center <> 0) Then
                Mennai = (max - Min) * 100 / center
            Else
                Mennai = -1
            End If
        Else
            Mennai = -1
        End If

    Case "D"                '---> |center-side|max/center�~100
        
        Select Case JudgFlag
        Case RES_JUDG ''���莯�ʃt���O(RES)
           
           center = R(0)
           Select Case G.cPos
           Case "1", "G", "H", "J", "M"

'����Ή��@2003.08.21 yakimura start
              If G.cCount = 2 Then
                 side = R(1)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
'����Ή��@2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 side = R(2)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
                  
              If G.cCount = 5 Then
                 side = R(3)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max1 = Abs(center - side) * 100 / center
                     Else
                         w_Max1 = -1
                     End If
                 Else
                     w_Max1 = -1
                 End If
              
                 side = R(4)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max2 = Abs(center - side) * 100 / center
                     Else
                         w_Max2 = -1
                     End If
                 Else
                     w_Max2 = -1
                 End If
              
                 Mennai = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
                
              End If
                  
           Case Else
              
              w_Max1 = SideMin_Cal(RES_JUDG, R(), G)
              w_Max1 = Abs(center - w_Max1)
           
              w_Max2 = SideMax_Cal(RES_JUDG, R(), G)
              w_Max2 = Abs(center - w_Max2)
           
              side = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
        
              If (side <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = side * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
           
           End Select
           
        Case OI_JUDG ''���莯�ʃt���O(Oi)
           
           center = R(0)
           Select Case G.cPos
           Case "1", "G", "H", "J", "M", "N", "Q"
           
'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 side = R(1)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If

'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 3 Then
                 side = R(2)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         Mennai = Abs(center - side) * 100 / center
                     Else
                         Mennai = -1
                     End If
                 Else
                     Mennai = -1
                 End If
              End If
                  
              If G.cCount = 5 Then
                 side = R(3)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max1 = Abs(center - side) * 100 / center
                     Else
                         w_Max1 = -1
                     End If
                 Else
                     w_Max1 = -1
                 End If
              
                 side = R(4)
                 If (side <> -1) And (center <> -1) Then
                     If (center <> 0) Then
                         w_Max2 = Abs(center - side) * 100 / center
                     Else
                         w_Max2 = -1
                     End If
                 Else
                     w_Max2 = -1
                 End If
              
                 Mennai = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
                
              End If
           
           Case Else
              
              w_Max1 = SideMin_Cal(OI_JUDG, R(), G)
              w_Max1 = Abs(center - w_Max1)
           
              w_Max2 = SideMax_Cal(OI_JUDG, R(), G)
              w_Max2 = Abs(center - w_Max2)
        
              side = IIf(w_Max1 >= w_Max2, w_Max1, w_Max2)
        
              If (side <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = side * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
           
           End Select
        
        End Select
    
    Case "E"                '---> |(centerave-sideave)|/centerave�~100
        
        Select Case G.cPos
           Case "1", "G", "H", "J", "M"

'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If
              End If

'����Ή��@2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If
              End If

' 2003.07.30 �đ򎖋Ə� �H�����̊m�F�ɂ��A����ʒu�l�̕��ς����߂Ă���ʓ��l���Z�o����  yakimura
              If G.cCount = 5 Then
'                 If R(0) <> 0 Then
'                    w_N3 = Abs(R(0) - R(3)) * 100 / R(0)
'                 Else
'                    w_N3 = -1
'                 End If
'                 If R(0) <> 0 Then
'                    w_N4 = Abs(R(0) - R(4)) * 100 / R(0)
'                 Else
'                    w_N4 = -1
'                 End If

'                 If w_N3 <> -1 and w_N4 <> -1 Then
'                    Mennai = (w_N3 + w_N4) / 2
'                 ElseIf w_N3 = -1 Then
'                    Mennai = w_N4
'                 ElseIf w_N4 = -1 Then
'                    Mennai = w_N3
'                 Else
'                    Mennai = -1
'                 End If

                 If R(3) <> -1 And R(4) <> -1 Then
                    w_Nx = (R(3) + R(4)) / 2
                 ElseIf R(3) = -1 Then
                    w_Nx = R(4)
                 ElseIf R(4) = -1 Then
                    w_Nx = R(3)
                 End If

                 If R(0) <> 0 Then
                    Mennai = Abs(R(0) - w_Nx) * 100 / R(0)
                 Else
                    Mennai = -1
                 End If

              End If

           Case Else

              center = R(0)
              Side_Ave = SideAve_Cal(RES_JUDG, R(), G)

              If (Side_Ave <> -9999) And (center <> -9999) Then
                  If (center <> 0) Then
                      Mennai = Abs(center - Side_Ave) * 100 / center
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If

        End Select

    Case "M"                '---> (max-min)/ave�~100
        Select Case JudgFlag
           Case RES_JUDG ''���莯�ʃt���O(RES)
              
              AVE = Ave_Cal(RES_JUDG, R(), G)
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
        
              If (Min <> -9999) And (max <> -9999) And (AVE <> -9999) Then
                 If (AVE <> 0) Then
                    Mennai = (max - Min) * 100 / AVE
                 Else
                    Mennai = -1
                 End If
              Else
                 Mennai = -1
              End If
           
           Case OI_JUDG ''���莯�ʃt���O(Oi)�@�@�@�@�@�@���ꏈ���@�uOi�v �́A"A" �Ōv�Z

              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
              If (Min <> -9999) And (max <> -9999) Then
                  If (Min <> 0) Then
                      Mennai = (max - Min) * 100 / Min
                  Else
                      Mennai = -1
                  End If
              Else
                  Mennai = -1
              End If
        
        End Select

    Case "N"                '---> |(center-side)/(center+side)|�~200
        
        Select Case JudgFlag
        Case RES_JUDG ''���莯�ʃt���O(RES)
           
           Select Case G.cPos
           Case "1", "G", "H", "J", "M"
                  
'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              End If

'����Ή��@2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 And R(2) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
              If G.cCount = 5 Then
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
              
                 Mennai = IIf(w_N3 >= w_N4, w_N3, w_N4)
              End If
              
           Case Else
              
'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              
              ElseIf G.cCount = 3 Then
'����Ή��@2003.08.21 yakimura end
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
              
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
              
              ElseIf G.cCount = 5 Then
              
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
           
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
                 Mennai = IIf(Mennai >= w_N3, Mennai, w_N3)
                 Mennai = IIf(Mennai >= w_N4, Mennai, w_N4)
              
              End If
           
           End Select
           
        Case OI_JUDG ''���莯�ʃt���O(Oi)
           
           Select Case G.cPos
           Case "1", "G", "H", "J", "M", "N", "Q"
                  
'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
'����Ή��@2003.08.21 yakimura end
              
              If G.cCount = 3 Then
                 If R(0) <> 0 And R(2) <> 0 Then
                    Mennai = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    Mennai = -1
                 End If
              End If
              
              If G.cCount = 5 Then
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
              
                 Mennai = IIf(w_N3 >= w_N4, w_N3, w_N4)
              End If
              
           Case Else
              
'����Ή��@2003.08.21 yakimura start
              
              If G.cCount = 2 Then
                 If R(0) <> 0 And R(1) <> 0 Then
                    Mennai = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    Mennai = -1
                 End If

              ElseIf G.cCount = 3 Then
'����Ή��@2003.08.21 yakimura end
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
              
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
              
              ElseIf G.cCount = 5 Then
              
                 If R(0) <> 0 And R(1) <> 0 Then
                    w_N1 = Abs(R(0) - R(1)) / (R(0) + R(1)) * 200
                 Else
                    w_N1 = -1
                 End If
                 If R(0) <> 0 And R(2) <> 0 Then
                    w_N2 = Abs(R(0) - R(2)) / (R(0) + R(2)) * 200
                 Else
                    w_N2 = -1
                 End If
                 If R(0) <> 0 And R(3) <> 0 Then
                    w_N3 = Abs(R(0) - R(3)) / (R(0) + R(3)) * 200
                 Else
                    w_N3 = -1
                 End If
                 If R(0) <> 0 And R(4) <> 0 Then
                    w_N4 = Abs(R(0) - R(4)) / (R(0) + R(4)) * 200
                 Else
                    w_N4 = -1
                 End If
           
                 Mennai = IIf(w_N1 >= w_N2, w_N1, w_N2)
                 Mennai = IIf(Mennai >= w_N3, Mennai, w_N3)
                 Mennai = IIf(Mennai >= w_N4, Mennai, w_N4)
              
              End If
           
           End Select
           
        End Select
           

    Case " ", ""
        
        ' �v�Z��������s��Ȃ�
        
    Case Else               '---> (max-min)/min�~100     ����P�[�X�@"A" �Ƃ��Čv�Z����
        Select Case JudgFlag
           Case RES_JUDG ''���莯�ʃt���O(RES)
              
              Min = Min_Cal(RES_JUDG, R(), G)
              max = Max_Cal(RES_JUDG, R(), G)
           
           Case OI_JUDG ''���莯�ʃt���O(Oi)
              
              Min = Min_Cal(OI_JUDG, R(), G)
              max = Max_Cal(OI_JUDG, R(), G)
        
        End Select
        
        If (Min <> -9999) And (max <> -9999) Then
            If (Min <> 0) Then
                Mennai = (max - Min) * 100 / Min
            Else
                Mennai = -1
            End If
        End If
    
    End Select

    'MENNAI_Cal = RoundUp(Mennai, 4)              '''�����̏����ł͂����Ȃ邪�c
    MENNAI_Cal_NotRound = Mennai        '�ۂ߂Ȃ��l��Ԃ��悤�ɕύX 2011/11/25 SETsw kubota

End Function
'�[���ۂ߂��s���֐��ƍs��Ȃ��֐��̓�ɕ��� 2011/11/25 SETsw kubota
Public Function MENNAI_Cal(JudgFlag As String, R() As Double, G As Guarantee, calcode As String) As Double
    '�[���ۂ߂��s��Ȃ��֐����Ăяo���A����4��(5���ڐ؂�グ)�ɂ��ĕԂ�
    MENNAI_Cal = RoundUp(MENNAI_Cal_NotRound(JudgFlag, R(), G, calcode), 4)
End Function



'�T�v      :����Ώ�MIN�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function Min_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim Min    As Double
    
    Min = -9999
    
    Select Case G.cCount
    Case "1"
            Min = d(0)
    Case "2"
            Min = IIf(d(0) <= d(1), d(0), d(1))
    Case "3"
            Min = JudgMin2(d())
    Case "5"
            Min = JudgMin(d())
    End Select
    
    Min_Cal = Min

End Function

'�T�v      :����Ώ�MAX�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function Max_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim max    As Double
    
    max = -9999
    
    Select Case G.cCount
    Case "1"
            max = d(0)
    Case "2"
            max = IIf(d(0) <= d(1), d(1), d(0))
    Case "3"
            max = JudgMax2(d())
    Case "5"
            max = JudgMax(d())
    End Select

    Max_Cal = max

End Function

'�T�v      :����Ώ�ave�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function Ave_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim AVE As Double
    
    AVE = -9999
    
    Select Case G.cCount
    Case "1"
            AVE = d(0)
    Case "2"
            If d(0) <> -1 And d(1) <> -1 Then
               AVE = (d(0) + d(1)) / 2
            ElseIf d(0) = -1 Then
               AVE = d(1)
            ElseIf d(1) = -1 Then
               AVE = d(0)
            End If
    Case "3"
            AVE = JudgAve2(d())
    Case "5"
            AVE = JudgAve(d())
    End Select

    Ave_Cal = AVE

End Function

'�T�v      :����Ώ�side_ave�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function SideAve_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim sideave As Double
    
    sideave = -9999
    
    Select Case G.cCount
    Case "2"
            sideave = d(1)
    Case "3"
            sideave = JudgSideAve2(d())
    Case "5"
            sideave = JudgSideAve(d())
    End Select

    SideAve_Cal = sideave

End Function


'�T�v      :����Ώ�side_min�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function SideMin_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim SideMin As Double
    
    SideMin = -9999
    
    Select Case G.cCount
    Case "2"
            SideMin = d(1)
    Case "3"
            If d(1) <> -1 And d(2) <> -1 Then
               SideMin = IIf(d(1) <= d(2), d(1), d(2))
            ElseIf d(1) = -1 Then
               SideMin = d(2)
            ElseIf d(2) = -1 Then
               SideMin = d(1)
            End If
    Case "5"
            SideMin = JudgSideMin(d())
    End Select

    SideMin_Cal = SideMin

End Function


'�T�v      :����Ώ�side_max�f�[�^�����߂�B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :�߂�l        ,O  ,double    ,RRG
'����      :
'����      :2003/06/06  yakimura  �쐬
Public Function SideMax_Cal(JudgFlag As String, d() As Double, G As Guarantee) As Double
Dim SideMax As Double
    
    SideMax = -9999
    
    Select Case G.cCount
    Case "2"
            SideMax = d(1)
    Case "3"
            If d(1) <> -1 And d(2) <> -1 Then
               SideMax = IIf(d(1) <= d(2), d(2), d(1))
            ElseIf d(1) = -1 Then
               SideMax = d(2)
            ElseIf d(2) = -1 Then
               SideMax = d(1)
            End If
    Case "5"
            SideMax = JudgSideMax(d())
    End Select

    SideMax_Cal = SideMax

End Function

'�T�v      :����l��NULL�������ꍇ�͈͔̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :JudgData      ,I  ,double    ,����l
'          :SpecMin       ,I  ,double    ,�����l
'          :SpecMax       ,I  ,double    ,����l
'          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'����      :
'����      :2003/12/11 �V�K�쐬 �V�X�e���u���C��
Public Function RangeDecision_nl(JudgData As Double, SpecMin As Double, SpecMax As Double) As Boolean
    RangeDecision_nl = False
    If (JudgData >= SpecMin) Or (SpecMin = -1) Then
        If (JudgData <= SpecMax) Or (SpecMax = -1) Then
            RangeDecision_nl = True
        End If
    End If
'    RangeDecision = ((JudgData >= SpecMin) And (JudgData <= SpecMax))
End Function

' �w�肳�ꂽ������R�f�[�^�Ɏd�l����ʒu��蒊�o������R�f�[�^���Z�b�g���Ȃ���
'sokuteiTensu  = ����_��(HSXRSPOT)
'sokuteiIchi   = ����ʒu(HSXRSPOI)
'typ_RS        = �����f�[�^(TBCMJ002)
'TEST2004/10
Public Function Set_Rs_Ichi(sokuteiTensu As String, sokuteiIchi As String, MEAS1 As Double, _
                            MEAS2 As Double, MEAS3 As Double, MEAS4 As Double, MEAS5 As Double) As FUNCTION_RETURN
Dim sTensu As String
Dim sName As String
Dim sMeas(1 To 5) As Double
Dim sMeas1(1 To 5) As Double
Dim i As Integer
Set_Rs_Ichi = FUNCTION_RETURN_FAILURE

''����1,2,3,5�_�̂ݑΉ�
If InStr("1235", sokuteiTensu) = 0 Then
    Exit Function
End If

''TBCMB005(1�_=info1,2�_=info2,3�_=info3,5�_=info5�j
sName = "info" & sokuteiTensu

sTensu = GetCodeField("SC", "30", sokuteiIchi, sName)
If sTensu = "" Then
    Exit Function
End If
''

sMeas(1) = MEAS1
sMeas(2) = MEAS2
sMeas(3) = MEAS3
sMeas(4) = MEAS4
sMeas(5) = MEAS5

For i = 1 To 5
    sMeas1(i) = -1
Next
For i = 1 To sokuteiTensu
    '�R�[�h�c�a = ex...1�_=1,3�_=133,5�_=13333
    sMeas1(i) = sMeas(Mid(sTensu, i, 1))
Next

MEAS1 = sMeas1(1)
MEAS2 = sMeas1(2)
MEAS3 = sMeas1(3)
MEAS4 = sMeas1(4)
MEAS5 = sMeas1(5)

Set_Rs_Ichi = FUNCTION_RETURN_SUCCESS

End Function

'Add Start 2011/01/26 SMPK Miyata
'�T�v      :Cudeco�̃p�^�[���̔�����s���B
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :SpcKbn        ,I  ,string    ,�p�^�[���敪
'          :JskKbn        ,I  ,string    ,�p�^�[������
'          :HanteiKbn     ,I  ,string    ,OK�p�^�[���敪(�P�����ڂ̓o�^�[���敪�A�Q�����ȍ~��OK�o�^�[���敪)
'          :�߂�l        ,O  ,Boolean   ,JUDG_OK or JUDG_NG
'����      :
'����      :
Public Function CudecoJudgPattern(SpcKbn As String, JskKbn As String, HanteiKbn() As String) As Boolean
    Dim ii As Integer
    
    CudecoJudgPattern = JUDG_NG
    For ii = 0 To UBound(HanteiKbn)
        If Mid(HanteiKbn(ii), 1, 1) = SpcKbn Then
            If InStr(2, HanteiKbn(ii), JskKbn) > 0 Then
                CudecoJudgPattern = JUDG_OK
                Exit For
            End If
        End If
    Next ii
    
End Function
'Add End   2011/01/26 SMPK Miyata

