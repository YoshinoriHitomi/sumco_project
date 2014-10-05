Attribute VB_Name = "s_cmmc001b"
''
'' ��R�ΐ͌v�Z��ʌv�Z���W���[��
''

'�T�v      :���̓p�����[�^�̍��v�l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,���v�l
'����      :
Public Function GetSum(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dWork   As Double

    On Error GoTo Err

    dWork = 0
    For Index = 0 To UBound(dParam)
        dWork = dWork + dParam(Index)
    Next Index

    GetSum = dWork
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetSum = 0
End Function


'�T�v      :���̓p�����[�^�̕��ϒl���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,���ϒl
'����      :
Public Function GetAve(dParam() As Double) As Double
    Dim dWork   As Double

    On Error GoTo Err

    GetAve = GetSum(dParam) / (UBound(dParam) + 1)

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetAve = 0
End Function


'�T�v      :���̓p�����[�^�̍ő�l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�ő�l
'����      :
Public Function GetMax(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMax    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMax = dParam(0): Exit Function
    
    dMax = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMax < dParam(Index) Then dMax = dParam(Index)
    Next Index

    GetMax = dMax

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMax = 0
End Function

'�T�v      :���̓p�����[�^�̍ŏ��l���擾����
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�ŏ��l
'����      :
Public Function GetMin(dParam() As Double) As Double
    Dim Index   As Integer
    Dim dMin    As Double

    On Error GoTo Err

    If UBound(dParam) = 0 Then GetMin = dParam(0): Exit Function
    
    dMin = dParam(0)
    For Index = 1 To UBound(dParam)
        If dMin > dParam(Index) Then dMin = dParam(Index)
    Next Index

    GetMin = dMin

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetMin = 0
End Function


'�T�v      :�ʓ����z�����߂�
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :dParam()      ,I   ,Double    ,�p�����[�^�l�z��
'          :�߂�l        ,O  ,Double    ,�ʓ����z�v�Z�l
'����      :
Public Function GetRG(dParam() As Double) As Double
    Dim dCalc1  As Double

    On Error GoTo Err

    dCalc1 = GetMin(dParam)
    If dCalc1 = 0 Then GetRG = 0: Exit Function

    GetRG = 100 * (GetMax(dParam) - GetMin(dParam)) / dCalc1

    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    GetRG = 0
End Function

