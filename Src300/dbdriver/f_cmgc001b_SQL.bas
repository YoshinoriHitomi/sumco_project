Attribute VB_Name = "f_cmgc001b_SQL"
Option Explicit
'
'' �������������
'Public Type typ_TBCMG001
'    MTRLNUM As String * 10      ' �����ԍ�
'    JDATE As Date               ' ���t
'    TRANCNT As Integer          ' ������
'    KRPROCCD As String * 5      ' �Ǘ��H���R�[�h
'    PROCCODE As String * 6      ' �H���R�[�h
'    MTRLTYPE As String * 3      ' �������
'    MAKERNO As String * 6       ' ���[�J�Ǘ�No
'    RVWEIGHT As Double          ' ����w���d��
'    CRYCOMMENT As String        ' �R�����g
'    TSTAFFID As String * 8      ' �o�^�Ј�ID
'    REGDATE As Date             ' �o�^���t
'    KSTAFFID As String * 8      ' �X�V�Ј��h�c
'    UPDDATE As Date             ' �X�V���t
'    SENDFLAG As String * 1      ' ���M�t���O
'    SENDDATE As Date            ' ���M���t
'End Type
'
'' �����݌ɊǗ�
'Public Type typ_TBCMG005
'    MTRLNUM As String * 10      ' �����ԍ�
'    USABLCLS As String * 1      ' �g�p�\�敪
'    WEIGHT As Integer           ' �d��
'    TSTAFFID As String * 8      ' �o�^�Ј�ID
'    REGDATE As Date             ' �o�^���t
'    KSTAFFID As String * 8      ' �X�V�Ј�ID
'    UPDDATE As Date             ' �X�V���t
'End Type

' f_cmgc001b_Exec
Public Type type_DBDRV_f_cmgc001b_Exec
    KRPROCCD As String * 5      ' �Ǘ��H���R�[�h
    PROCCODE As String * 6      ' �H���R�[�h
    TSTAFFID As String * 8      ' �o�^�Ј�ID
    MTRLTYPE As String * 3      ' �������
    MAKERNO As String * 6       ' ���[�J�Ǘ�No
    RVWEIGHT As Double          ' ����w���d��
    CRYCOMMENT As String        ' �R�����g
End Type

Public Function DBDRV_f_cmgc001b_Exec(DBDRV_f_cmgc001b_Exec As type_DBDRV_f_cmgc001b_Exec) As FUNCTION_RETURN

    f_cmgc001b_Exec = FUNCTION_RETURN_SUCCESS
End Function
