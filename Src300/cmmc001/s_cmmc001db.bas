Attribute VB_Name = "s_cmmc001db"
Option Explicit

'Type type_HinbanSyutoku   ''s_cmmc001db_sql　のOUT用
'    CRYNUM      As String * 12  ''結晶番号
'    INGOTPOS    As Integer      ''トップサンプル位置
'    HINBAN      As String * 12  ''ボトムサンプル位置
'    BLOCKID     As String * 12  ''チャージ量
'    LENGTH      As Integer      ''トップ重量
'End Type

'概要      :　引上げ終了実績取得関数
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'        :sCryNum        ,I   ,String         ,入力用
'        :pTbcmh004()        ,O   ,typ_TBCMH004     ,引上げ終了実績取得用
'説明      :
'履歴      :2001/06/28　小林　作成
Public Function s_cmmc001db_sql(ByVal sCryNum As String, _
                pTbcmh004() As typ_TBCMH004) As Double
    Dim sql As String
    Dim ret As Integer
    
    sql = " where CRYNUM = '" & sCryNum & "' "

    ret = DBDRV_GetTBCMH004(pTbcmh004, sql, "order by CRYNUM")

End Function
