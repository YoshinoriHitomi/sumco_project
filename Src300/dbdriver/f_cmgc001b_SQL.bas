Attribute VB_Name = "f_cmgc001b_SQL"
Option Explicit
'
'' 多結晶受入実績
'Public Type typ_TBCMG001
'    MTRLNUM As String * 10      ' 原料番号
'    JDATE As Date               ' 日付
'    TRANCNT As Integer          ' 処理回数
'    KRPROCCD As String * 5      ' 管理工程コード
'    PROCCODE As String * 6      ' 工程コード
'    MTRLTYPE As String * 3      ' 原料種類
'    MAKERNO As String * 6       ' メーカ管理No
'    RVWEIGHT As Double          ' 受入購入重量
'    CRYCOMMENT As String        ' コメント
'    TSTAFFID As String * 8      ' 登録社員ID
'    REGDATE As Date             ' 登録日付
'    KSTAFFID As String * 8      ' 更新社員ＩＤ
'    UPDDATE As Date             ' 更新日付
'    SENDFLAG As String * 1      ' 送信フラグ
'    SENDDATE As Date            ' 送信日付
'End Type
'
'' 原料在庫管理
'Public Type typ_TBCMG005
'    MTRLNUM As String * 10      ' 原料番号
'    USABLCLS As String * 1      ' 使用可能区分
'    WEIGHT As Integer           ' 重量
'    TSTAFFID As String * 8      ' 登録社員ID
'    REGDATE As Date             ' 登録日付
'    KSTAFFID As String * 8      ' 更新社員ID
'    UPDDATE As Date             ' 更新日付
'End Type

' f_cmgc001b_Exec
Public Type type_DBDRV_f_cmgc001b_Exec
    KRPROCCD As String * 5      ' 管理工程コード
    PROCCODE As String * 6      ' 工程コード
    TSTAFFID As String * 8      ' 登録社員ID
    MTRLTYPE As String * 3      ' 原料種類
    MAKERNO As String * 6       ' メーカ管理No
    RVWEIGHT As Double          ' 受入購入重量
    CRYCOMMENT As String        ' コメント
End Type

Public Function DBDRV_f_cmgc001b_Exec(DBDRV_f_cmgc001b_Exec As type_DBDRV_f_cmgc001b_Exec) As FUNCTION_RETURN

    f_cmgc001b_Exec = FUNCTION_RETURN_SUCCESS
End Function
