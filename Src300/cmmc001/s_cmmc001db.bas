Attribute VB_Name = "s_cmmc001db"
Option Explicit

'Type type_HinbanSyutoku   ''s_cmmc001db_sql�@��OUT�p
'    CRYNUM      As String * 12  ''�����ԍ�
'    INGOTPOS    As Integer      ''�g�b�v�T���v���ʒu
'    HINBAN      As String * 12  ''�{�g���T���v���ʒu
'    BLOCKID     As String * 12  ''�`���[�W��
'    LENGTH      As Integer      ''�g�b�v�d��
'End Type

'�T�v      :�@���グ�I�����ю擾�֐�
'���Ұ�    :�ϐ���        ,IO ,�^             ,����
'        :sCryNum        ,I   ,String         ,���͗p
'        :pTbcmh004()        ,O   ,typ_TBCMH004     ,���グ�I�����ю擾�p
'����      :
'����      :2001/06/28�@���с@�쐬
Public Function s_cmmc001db_sql(ByVal sCryNum As String, _
                pTbcmh004() As typ_TBCMH004) As Double
    Dim sql As String
    Dim ret As Integer
    
    sql = " where CRYNUM = '" & sCryNum & "' "

    ret = DBDRV_GetTBCMH004(pTbcmh004, sql, "order by CRYNUM")

End Function
