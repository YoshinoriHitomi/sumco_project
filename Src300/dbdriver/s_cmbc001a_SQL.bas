Attribute VB_Name = "s_cmbc001a_SQL"

'概要      :引上指示番号の連番部に値を加える
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :sijiNo        ,I  ,String    ,元の引上指示番号
'          :addVal        ,I  ,Integer   ,加算値(マイナスも可)
'          :戻り値        ,O  ,String    ,加算後の引上指示番号
'説明      :
'履歴      :2001/07/09 作成  野村 (2002/07 s_cmzcF_cmhc001d_SQL.basより移動)
Public Function SijiNoAdd(sijiNo$, addVal%) As String
Dim seq As Integer
Dim newNo As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmhc001d_SQL.bas -- Function SijiNoAdd"

    seq = Val(Mid$(sijiNo, 5, 3))
    SijiNoAdd = Left$(sijiNo, 4) & Format$(seq + addVal, "000") & Mid$(sijiNo, 8)

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function


