Attribute VB_Name = "s_cmbc001a_SQL"

'�T�v      :����w���ԍ��̘A�ԕ��ɒl��������
'���Ұ�    :�ϐ���        ,IO ,�^        ,����
'          :sijiNo        ,I  ,String    ,���̈���w���ԍ�
'          :addVal        ,I  ,Integer   ,���Z�l(�}�C�i�X����)
'          :�߂�l        ,O  ,String    ,���Z��̈���w���ԍ�
'����      :
'����      :2001/07/09 �쐬  �쑺 (2002/07 s_cmzcF_cmhc001d_SQL.bas���ړ�)
Public Function SijiNoAdd(sijiNo$, addVal%) As String
Dim seq As Integer
Dim newNo As String


    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function SijiNoAdd"

    seq = Val(Mid$(sijiNo, 5, 3))
    SijiNoAdd = Left$(sijiNo, 4) & Format$(seq + addVal, "000") & Mid$(sijiNo, 8)

proc_exit:
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    gErr.HandleError
    Resume proc_exit
End Function


