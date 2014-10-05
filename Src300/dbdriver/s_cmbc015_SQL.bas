Attribute VB_Name = "s_cmbc015_SQL"
Option Explicit

'================================================
'プロジェクトcmbc015用SQLbas
'2002/08 s_cmzcF_cmmc001a_SQL.basより移動
'================================================

'(2002/07 DBDRV_GetTBCME018より移動)
'フィールド名検索用
Dim fldNames() As String    '現rsに含まれるフィールド名保持配列
Dim fldCnt As Integer       '現rsに含まれるフィールド数


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME018」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME018 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME018_SQL.basより移動)
Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, MCNO, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME018"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME018 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
'NULL対応 ----- START ----- 2003/12/10
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .HINBAN = rs("HINBAN")           ' 品番
            .mnorevno = rs("MNOREVNO")       ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .HMGSTRRNO = rs("HMGSTRRNO")     ' 品管理仕様登録依頼番号
            .HMGSTFNO = rs("HMGSTFNO")       ' 品管理社員Ｎｏ
            .HMGSXSNO = rs("HMGSXSNO")       ' 品管理ＳＸ製品番号
            .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))       ' 品管理ＳＸ製品番号枝番
            .CONFLAG = rs("CONFLAG")         ' 確認フラグ
            .REINFLAG = rs("REINFLAG")       ' 再付与フラグ
            .HSXTRWKB = rs("HSXTRWKB")       ' 品ＳＸ統合可否区分
            .HSXTYPE = rs("HSXTYPE")         ' 品ＳＸタイプ
            .KSXTYPKW = rs("KSXTYPKW")       ' 品ＳＸタイプ検査方法
            .HSXDOP = rs("HSXDOP")           ' 品ＳＸドーパント
            .HSXRMIN = fncNullCheck(rs("HSXRMIN"))         ' 品ＳＸ比抵抗下限
            .HSXRMAX = fncNullCheck(rs("HSXRMAX"))         ' 品ＳＸ比抵抗上限
            .HSXRSPOH = rs("HSXRSPOH")       ' 品ＳＸ比抵抗測定位置＿方
            .HSXRSPOT = rs("HSXRSPOT")       ' 品ＳＸ比抵抗測定位置＿点
            .HSXRSPOI = rs("HSXRSPOI")       ' 品ＳＸ比抵抗測定位置＿位
            .HSXRHWYT = rs("HSXRHWYT")       ' 品ＳＸ比抵抗保証方法＿対
            .HSXRHWYS = rs("HSXRHWYS")       ' 品ＳＸ比抵抗保証方法＿処
            .HSXRKWAY = rs("HSXRKWAY")       ' 品ＳＸ比抵抗検査方法
            .HSXRKHNM = rs("HSXRKHNM")       ' 品ＳＸ比抵抗検査頻度＿枚
            .HSXRKHNI = rs("HSXRKHNI")       ' 品ＳＸ比抵抗検査頻度＿位
            .HSXRKHNH = rs("HSXRKHNH")       ' 品ＳＸ比抵抗検査頻度＿保
            .HSXRKHNS = rs("HSXRKHNS")       ' 品ＳＸ比抵抗検査頻度＿試
            .HSXRMCAL = rs("HSXRMCAL")       ' 品ＳＸ比抵抗面内計算
            .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))       ' 品ＳＸ比抵抗面内分布
            .HSXRMCL2 = rs("HSXRMCL2")       ' 品ＳＸ比抵抗面内計算２
            .HSXRMBP2 = fncNullCheck(rs("HSXRMBP2"))       ' 品ＳＸ比抵抗面内分布２
            .HSXRSDEV = fncNullCheck(rs("HSXRSDEV"))       ' 品ＳＸ比抵抗標準偏差
            .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))       ' 品ＳＸ比抵抗平均下限
            .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))       ' 品ＳＸ比抵抗平均上限
            .HSXFORM = rs("HSXFORM")         ' 品ＳＸ形状
            .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))       ' 品ＳＸ直径１中心
            .HSXD1MIN = fncNullCheck(rs("HSXD1MIN"))       ' 品ＳＸ直径１下限
            .HSXD1MAX = fncNullCheck(rs("HSXD1MAX"))       ' 品ＳＸ直径１上限
            .HSXD2CEN = fncNullCheck(rs("HSXD2CEN"))       ' 品ＳＸ直径２中心
            .HSXD2MIN = fncNullCheck(rs("HSXD2MIN"))       ' 品ＳＸ直径２下限
            .HSXD2MAX = fncNullCheck(rs("HSXD2MAX"))       ' 品ＳＸ直径２上限
            .HSXCDIR = rs("HSXCDIR")         ' 品ＳＸ結晶面方位
            .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))       ' 品ＳＸ結晶面傾中心
            .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))       ' 品ＳＸ結晶面傾下限
            .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))       ' 品ＳＸ結晶面傾上限
            .HSXCKWAY = rs("HSXCKWAY")       ' 品ＳＸ結晶面検査方法
            .HSXCKHNM = rs("HSXCKHNM")       ' 品ＳＸ結晶面検査頻度＿枚
            .HSXCKHNI = rs("HSXCKHNI")       ' 品ＳＸ結晶面検査頻度＿位
            .HSXCKHNH = rs("HSXCKHNH")       ' 品ＳＸ結晶面検査頻度＿保
            .HSXCKHNS = rs("HSXCKHNS")       ' 品ＳＸ結晶面検査頻度＿試
            .HSXCSDIR = rs("HSXCSDIR")       ' 品ＳＸ結晶面傾方位
            .HSXCSDIS = rs("HSXCSDIS")       ' 品ＳＸ結晶面傾方位指定
            .HSXCTDIR = rs("HSXCTDIR")       ' 品ＳＸ結晶面傾縦方位
            .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))       ' 品ＳＸ結晶面傾縦中心
            .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))       ' 品ＳＸ結晶面傾縦下限
            .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))       ' 品ＳＸ結晶面傾縦上限
            .HSXCYDIR = rs("HSXCYDIR")       ' 品ＳＸ結晶面傾横方位
            .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))       ' 品ＳＸ結晶面傾横中心
            .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))       ' 品ＳＸ結晶面傾横下限
            .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))       ' 品ＳＸ結晶面傾横上限
            .HSXOF1PD = rs("HSXOF1PD")       ' 品ＳＸＯＦ１位置方位
            .HSXOF1PN = fncNullCheck(rs("HSXOF1PN"))       ' 品ＳＸＯＦ１位置下限
            .HSXOF1PX = fncNullCheck(rs("HSXOF1PX"))       ' 品ＳＸＯＦ１位置上限
            .HSXOF1PW = rs("HSXOF1PW")       ' 品ＳＸＯＦ１位置検査方法
            .HSXOF1LC = fncNullCheck(rs("HSXOF1LC"))       ' 品ＳＸＯＦ１長中心
            .HSXOF1LN = fncNullCheck(rs("HSXOF1LN"))       ' 品ＳＸＯＦ１長下限
            .HSXOF1LX = fncNullCheck(rs("HSXOF1LX"))       ' 品ＳＸＯＦ１長上限
            .HSXOF1DC = fncNullCheck(rs("HSXOF1DC"))       ' 品ＳＸＯＦ１直径中心
            .HSXOF1DN = fncNullCheck(rs("HSXOF1DN"))       ' 品ＳＸＯＦ１直径下限
            .HSXOF1DX = fncNullCheck(rs("HSXOF1DX"))       ' 品ＳＸＯＦ１直径上限
            .HSXDFORM = rs("HSXDFORM")       ' 品ＳＸ溝形状
            .HSXDPDRC = rs("HSXDPDRC")       ' 品ＳＸ溝位置方向
            .HSXDPACN = fncNullCheck(rs("HSXDPACN"))       ' 品ＳＸ溝位置角度中心
            .HSXDPAMN = fncNullCheck(rs("HSXDPAMN"))       ' 品ＳＸ溝位置角度下限
            .HSXDPAMX = fncNullCheck(rs("HSXDPAMX"))       ' 品ＳＸ溝位置角度上限
            .HSXDPKWY = rs("HSXDPKWY")       ' 品ＳＸ溝位置検査方法
            .HSXDPDIR = rs("HSXDPDIR")       ' 品ＳＸ溝位置方位
            .HSXDPMIN = fncNullCheck(rs("HSXDPMIN"))       ' 品ＳＸ溝位置下限
            .HSXDPMAX = fncNullCheck(rs("HSXDPMAX"))       ' 品ＳＸ溝位置上限
            .HSXDWCEN = fncNullCheck(rs("HSXDWCEN"))       ' 品ＳＸ溝巾中心
            .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))       ' 品ＳＸ溝巾下限
            .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))       ' 品ＳＸ溝巾上限
            .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))       ' 品ＳＸ溝深中心
            .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))       ' 品ＳＸ溝深下限
            .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))       ' 品ＳＸ溝深上限
            .HSXDACEN = fncNullCheck(rs("HSXDACEN"))       ' 品ＳＸ溝角度中心
            .HSXDAMIN = fncNullCheck(rs("HSXDAMIN"))       ' 品ＳＸ溝角度下限
            .HSXDAMAX = fncNullCheck(rs("HSXDAMAX"))       ' 品ＳＸ溝角度上限
            .MCNO = rs("MCNO")               ' 結晶操業内製作条件
            .IFKBN = rs("IFKBN")             ' Ｉ／Ｆ区分
            .SYORIKBN = rs("SYORIKBN")       ' 処理区分
            .SPECRRNO = rs("SPECRRNO")       ' 仕様登録依頼番号
            .SXLMCNO = rs("SXLMCNO")         ' ＳＸＬ製作条件番号
            .WFMCNO = rs("WFMCNO")           ' ＷＦ製作条件番号
            .StaffID = rs("STAFFID")         ' 社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close
'NULL対応 -----  END  ----- 2003/12/10

    DBDRV_GetTBCME018 = FUNCTION_RETURN_SUCCESS
End Function




'概要      :引数のフィールド名がfldNames()配列に含まれているかどうかの判定。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :fldName       ,I  ,typ_TBCME018 ,抽出レコード
'          :戻り値        ,O  ,Boolean      ,True:在り／False：無し
'説明      :
'履歴      :2001/06/27作成　野村  (2002/07 DBDRV_GetTBCME018より移動)

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL全体
    Dim i As Integer                    'ﾙｰﾌﾟｶｳﾝﾄ


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc015_SQL.bas -- Function fldNameExist"

    fldNameExist = False                'ｴﾗｰｽﾃｰﾀｽ（初期値）ｾｯﾄ
    
    For i = 1 To fldCnt                 'ﾌｨｰﾙﾄﾞ数分ﾙｰﾌﾟ
        If fldName = fldNames(i) Then   '引数のﾌｨｰﾙﾄﾞ名と一致するものがあった場合
            fldNameExist = True         '正常ｽﾃｰﾀｽｾｯﾄ
            Exit For                    'ﾙｰﾌﾟを抜ける
        End If
    Next
    

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME037 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME037_SQL.basより移動)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' 結晶番号
            .DELCLS = rs("DELCLS")           ' 削除区分
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCD = rs("PROCCD")           ' 工程コード
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
            .RPFACT = rs("RPFACT")           ' ねらい品番工場
            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
            .PRODCOND = rs("PRODCOND")       ' 製作条件
            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .DIAMETER = rs("DIAMETER")       ' 直径
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMH004」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMH004 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMH004_SQL.basより移動)
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMH004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' 結晶番号
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .LENGTOP = rs("LENGTOP")         ' 長さ（TOP）
            .LENGTKDO = rs("LENGTKDO")       ' 長さ（直胴）
            .LENGTAIL = rs("LENGTAIL")       ' 長さ（TAIL）
            .LENGFREE = rs("LENGFREE")       ' フリー長さ
            .DM1 = rs("DM1")                 ' 直胴直径１
            .DM2 = rs("DM2")                 ' 直胴直径２
            .DM3 = rs("DM3")                 ' 直胴直径３
            .WGHTTOP = rs("WGHTTOP")         ' 重量（TOP）
            .WGHTTKDO = rs("WGHTTKDO")       ' 重量（直胴）
            .WGHTTAIL = rs("WGHTTAIL")       ' 重量（TAIL)
            .WGHTFREE = rs("WGHTFREE")       ' 重量（フリー長さ）
            .WGTOPCUT = rs("WGTOPCUT")       ' トップカット重量
            .UPWEIGHT = rs("UPWEIGHT")       ' 引上げ重量
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .STATCLS = rs("STATCLS")         ' BOT状況区分
            .JDGECODE = rs("JDGECODE")       ' 判定コード
            .PWTIME = rs("PWTIME")           ' パワー時間
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドーパント種類
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .ADDDPNAM = rs("ADDDPNAM")       ' 追加ドープ名
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMJ002」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ002 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMJ002_SQL.basより移動)
Public Function DBDRV_GetTBCMJ002(records() As typ_TBCMJ002, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
              " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID," & _
              " UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMJ002"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMJ002 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' 結晶番号
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .HINBAN = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .GOUKI = rs("GOUKI")             ' 号機
            .TYPE = rs("TYPE")               ' タイプ
            .MEAS1 = rs("MEAS1")             ' 測定値１
            .MEAS2 = rs("MEAS2")             ' 測定値２
            .MEAS3 = rs("MEAS3")             ' 測定値３
            .MEAS4 = rs("MEAS4")             ' 測定値４
            .MEAS5 = rs("MEAS5")             ' 測定値５
            .EFEHS = rs("EFEHS")             ' 実効偏析
            .RRG = rs("RRG")                 ' ＲＲＧ
            .JudgData = rs("JUDGDATA")       ' 検索対象値
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMJ002 = FUNCTION_RETURN_SUCCESS
End Function


