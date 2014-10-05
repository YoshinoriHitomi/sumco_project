Attribute VB_Name = "s_kensa2_SQL"
'------------------------------------------------
' DBアクセス関数
'------------------------------------------------
'フィールド名検索用
Dim fldNames() As String    '現rsに含まれるフィールド名保持配列
Dim fldCnt As Integer       '現rsに含まれるフィールド数

'概要      :テーブル「TBCME019」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME019 ,抽出レコード
'          :formID        ,I  ,String       ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban  ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2001/06/27作成　長野 (2002/07 s_cmzcF_TBCME019_SQL.basより移動)

Public Function DBDRV_GetTBCME019(records() As typ_TBCME019, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL全体
Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
Dim sqlWhere    As String           'SQLWhere部
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             'レコード数
Dim key         As String           '検索KEY
Dim i           As Long             'ﾙｰﾌﾟｶｳﾝﾄ


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME019_SQL.bas -- Function DBDRV_GetTBCME019"

 Select Case formID
        Case "f_cmbc021_1"           '「FTIR(Oi,Cs)実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc022_1"           '「GFA(Oi)実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc023_1"           '「抵抗実績入力」
           sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc024_1"           '「BMD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmec030_1"           '「BMD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc025_1"           '「OSF実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmec031_1"           '「OSF実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc027_1"           '「ライフタイム実績入力」
           sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
        
        Case "f_cmbc028_1i"           '「FPD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
                
        Case "f_cmbc029_1"           '「GFA校正情報設定」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXTMMAXN, HSXTMSPH, HSXTMSPT," & _
              " HSXTMSPR, HSXTMKHM, HSXTMKHI, HSXTMKHH, HSXTMKHS, HSXLTMIN, HSXLTMAX, HSXLTSPH, HSXLTSPT, HSXLTSPI, HSXLTHWT," & _
              " HSXLTHWS, HSXLTKWY, HSXLTNSW, HSXLTKHM, HSXLTKHI, HSXLTKHH, HSXLTKHS, HSXLTMBP, HSXLTMCL, HSXCNMIN, HSXCNMAX," & _
              " HSXCNSPH, HSXCNSPT, HSXCNSPI, HSXCNHWT, HSXCNHWS, HSXCNKWY, HSXCNKHM, HSXCNKHI, HSXCNKHH, HSXCNKHS, HSXONMIN," & _
              " HSXONMAX, HSXONSPH, HSXONSPT, HSXONSPI, HSXONHWT, HSXONHWS, HSXONKWY, HSXONKHM, HSXONKHI, HSXONKHH, HSXONKHS," & _
              " HSXONMBP, HSXONMCL, HSXONLTB, HSXONLTC, HSXONSDV, HSXONAMN, HSXONAMX, HSXOS1MN, HSXOS1MX, HSXOS1NS, HSXOS1SH," & _
              " HSXOS1ST, HSXOS1SI, HSXOS1HT, HSXOS1HS, HSXOS1HM, HSXOS1KI, HSXOS1KH, HSXOS1KS, HSXOS2MN, HSXOS2MX, HSXOS2NS," & _
              " HSXOS2SH, HSXOS2ST, HSXOS2SI, HSXOS2HT, HSXOS2HS, HSXOS2KM, HSXOS2KN, HSXOS2KH, HSXOS2KU, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG "
    
    End Select
    
    sqlBase = sqlBase & "From TBCME019"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME019 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''フィールド名を登録する
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
   
     ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
         With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")               ' 品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")         ' 製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")            ' 工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")            ' 操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")      ' 品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")         ' 品管理社員Ｎｏ
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")         ' 品管理ＳＸ製品番号
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))         ' 品管理ＳＸ製品番号枝番
            If fldNameExist("HSXTMMAXN") Then .HSXTMMAX = fncNullCheck(rs("HSXTMMAXN"))        ' 品ＳＸ転位密度上限    ＷＦサンプル処理変更 2003.05.20 yakimura
            If fldNameExist("HSXTMSPH") Then .HSXTMSPH = rs("HSXTMSPH")         ' 品ＳＸ転位密度測定位置＿方
            If fldNameExist("HSXTMSPT") Then .HSXTMSPT = rs("HSXTMSPT")         ' 品ＳＸ転位密度測定位置＿点
            If fldNameExist("HSXTMSPR") Then .HSXTMSPR = rs("HSXTMSPR")         ' 品ＳＸ転位密度測定位置＿領
            If fldNameExist("HSXTMKHM") Then .HSXTMKHM = rs("HSXTMKHM")         ' 品ＳＸ転位密度検査頻度＿枚
            If fldNameExist("HSXTMKHI") Then .HSXTMKHI = rs("HSXTMKHI")         ' 品ＳＸ転位密度検査頻度＿位
            If fldNameExist("HSXTMKHH") Then .HSXTMKHH = rs("HSXTMKHH")         ' 品ＳＸ転位密度検査頻度＿保
            If fldNameExist("HSXTMKHS") Then .HSXTMKHS = rs("HSXTMKHS")         ' 品ＳＸ転位密度検査頻度＿試
            If fldNameExist("HSXLTMIN") Then .HSXLTMIN = fncNullCheck(rs("HSXLTMIN"))         ' 品ＳＸＬタイム下限 'NULL対応
            If fldNameExist("HSXLTMAX") Then .HSXLTMAX = fncNullCheck(rs("HSXLTMAX"))         ' 品ＳＸＬタイム上限 'NULL対応
            If fldNameExist("HSXLTSPH") Then .HSXLTSPH = rs("HSXLTSPH")         ' 品ＳＸＬタイム測定位置＿方
            If fldNameExist("HSXLTSPT") Then .HSXLTSPT = rs("HSXLTSPT")         ' 品ＳＸＬタイム測定位置＿点
            If fldNameExist("HSXLTSPI") Then .HSXLTSPI = rs("HSXLTSPI")         ' 品ＳＸＬタイム測定位置＿位
            If fldNameExist("HSXLTHWT") Then .HSXLTHWT = rs("HSXLTHWT")         ' 品ＳＸＬタイム保証方法＿対
            If fldNameExist("HSXLTHWS") Then .HSXLTHWS = rs("HSXLTHWS")         ' 品ＳＸＬタイム保証方法＿処
            If fldNameExist("HSXLTKWY") Then .HSXLTKWY = rs("HSXLTKWY")         ' 品ＳＸＬタイム検査方法
            If fldNameExist("HSXLTNSW") Then .HSXLTNSW = rs("HSXLTNSW")         ' 品ＳＸＬタイム熱処理法
            If fldNameExist("HSXLTKHM") Then .HSXLTKHM = rs("HSXLTKHM")         ' 品ＳＸＬタイム検査頻度＿枚
            If fldNameExist("HSXLTKHI") Then .HSXLTKHI = rs("HSXLTKHI")         ' 品ＳＸＬタイム検査頻度＿位
            If fldNameExist("HSXLTKHH") Then .HSXLTKHH = rs("HSXLTKHH")         ' 品ＳＸＬタイム検査頻度＿保
            If fldNameExist("HSXLTKHS") Then .HSXLTKHS = rs("HSXLTKHS")         ' 品ＳＸＬタイム検査頻度＿試
            If fldNameExist("HSXLTMBP") Then .HSXLTMBP = fncNullCheck(rs("HSXLTMBP"))         ' 品ＳＸＬタイム面内分布
            If fldNameExist("HSXLTMCL") Then .HSXLTMCL = rs("HSXLTMCL")         ' 品ＳＸＬタイム面内計算
            If fldNameExist("HSXCNMIN") Then .HSXCNMIN = fncNullCheck(rs("HSXCNMIN"))         ' 品ＳＸ炭素濃度下限
            If fldNameExist("HSXCNMAX") Then .HSXCNMAX = fncNullCheck(rs("HSXCNMAX"))         ' 品ＳＸ炭素濃度上限
            If fldNameExist("HSXCNSPH") Then .HSXCNSPH = rs("HSXCNSPH")         ' 品ＳＸ炭素濃度測定位置＿方
            If fldNameExist("HSXCNSPT") Then .HSXCNSPT = rs("HSXCNSPT")         ' 品ＳＸ炭素濃度測定位置＿点
            If fldNameExist("HSXCNSPI") Then .HSXCNSPI = rs("HSXCNSPI")         ' 品ＳＸ炭素濃度測定位置＿位
            If fldNameExist("HSXCNHWT") Then .HSXCNHWT = rs("HSXCNHWT")         ' 品ＳＸ炭素濃度保証方法＿対
            If fldNameExist("HSXCNHWS") Then .HSXCNHWS = rs("HSXCNHWS")         ' 品ＳＸ炭素濃度保証方法＿処
            If fldNameExist("HSXCNKWY") Then .HSXCNKWY = rs("HSXCNKWY")         ' 品ＳＸ炭素濃度検査方法
            If fldNameExist("HSXCNKHM") Then .HSXCNKHM = rs("HSXCNKHM")         ' 品ＳＸ炭素濃度検査頻度＿枚
            If fldNameExist("HSXCNKHI") Then .HSXCNKHI = rs("HSXCNKHI")         ' 品ＳＸ炭素濃度検査頻度＿位
            If fldNameExist("HSXCNKHH") Then .HSXCNKHH = rs("HSXCNKHH")         ' 品ＳＸ炭素濃度検査頻度＿保
            If fldNameExist("HSXCNKHS") Then .HSXCNKHS = rs("HSXCNKHS")         ' 品ＳＸ炭素濃度検査頻度＿試
            If fldNameExist("HSXONMIN") Then .HSXONMIN = fncNullCheck(rs("HSXONMIN"))         ' 品ＳＸ酸素濃度下限
            If fldNameExist("HSXONMAX") Then .HSXONMAX = fncNullCheck(rs("HSXONMAX"))         ' 品ＳＸ酸素濃度上限
            If fldNameExist("HSXONSPH") Then .HSXONSPH = rs("HSXONSPH")         ' 品ＳＸ酸素濃度測定位置＿方
            If fldNameExist("HSXONSPT") Then .HSXONSPT = rs("HSXONSPT")         ' 品ＳＸ酸素濃度測定位置＿点
            If fldNameExist("HSXONSPI") Then .HSXONSPI = rs("HSXONSPI")         ' 品ＳＸ酸素濃度測定位置＿位
            If fldNameExist("HSXONHWT") Then .HSXONHWT = rs("HSXONHWT")         ' 品ＳＸ酸素濃度保証方法＿対
            If fldNameExist("HSXONHWS") Then .HSXONHWS = rs("HSXONHWS")         ' 品ＳＸ酸素濃度保証方法＿処
            If fldNameExist("HSXONKWY") Then .HSXONKWY = rs("HSXONKWY")         ' 品ＳＸ酸素濃度検査方法
            If fldNameExist("HSXONKHM") Then .HSXONKHM = rs("HSXONKHM")         ' 品ＳＸ酸素濃度検査頻度＿枚
            If fldNameExist("HSXONKHI") Then .HSXONKHI = rs("HSXONKHI")         ' 品ＳＸ酸素濃度検査頻度＿位
            If fldNameExist("HSXONKHH") Then .HSXONKHH = rs("HSXONKHH")         ' 品ＳＸ酸素濃度検査頻度＿保
            If fldNameExist("HSXONKHS") Then .HSXONKHS = rs("HSXONKHS")         ' 品ＳＸ酸素濃度検査頻度＿試
            If fldNameExist("HSXONMBP") Then .HSXONMBP = fncNullCheck(rs("HSXONMBP"))         ' 品ＳＸ酸素濃度面内分布
            If fldNameExist("HSXONMCL") Then .HSXONMCL = rs("HSXONMCL")         ' 品ＳＸ酸素濃度面内計算
            If fldNameExist("HSXONLTB") Then .HSXONLTB = fncNullCheck(rs("HSXONLTB"))         ' 品ＳＸ酸素濃度ＬＴ分布
            If fldNameExist("HSXONLTC") Then .HSXONLTC = rs("HSXONLTC")         ' 品ＳＸ酸素濃度ＬＴ計算
            If fldNameExist("HSXONSDV") Then .HSXONSDV = fncNullCheck(rs("HSXONSDV"))         ' 品ＳＸ酸素濃度標準偏差
            If fldNameExist("HSXONAMN") Then .HSXONAMN = fncNullCheck(rs("HSXONAMN"))         ' 品ＳＸ酸素濃度平均下限
            If fldNameExist("HSXONAMX") Then .HSXONAMX = fncNullCheck(rs("HSXONAMX"))         ' 品ＳＸ酸素濃度平均上限
            If fldNameExist("HSXOS1MN") Then .HSXOS1MN = fncNullCheck(rs("HSXOS1MN"))         ' 品ＳＸ酸素析出１下限
            If fldNameExist("HSXOS1MX") Then .HSXOS1MX = fncNullCheck(rs("HSXOS1MX"))         ' 品ＳＸ酸素析出１上限
            If fldNameExist("HSXOS1NS") Then .HSXOS1NS = rs("HSXOS1NS")         ' 品ＳＸ酸素析出１熱処理法
            If fldNameExist("HSXOS1SH") Then .HSXOS1SH = rs("HSXOS1SH")         ' 品ＳＸ酸素析出１測定位置＿方
            If fldNameExist("HSXOS1ST") Then .HSXOS1ST = rs("HSXOS1ST")         ' 品ＳＸ酸素析出１測定位置＿点
            If fldNameExist("HSXOS1SI") Then .HSXOS1SI = rs("HSXOS1SI")         ' 品ＳＸ酸素析出１測定位置＿位
            If fldNameExist("HSXOS1HT") Then .HSXOS1HT = rs("HSXOS1HT")         ' 品ＳＸ酸素析出１保証方法＿対
            If fldNameExist("HSXOS1HS") Then .HSXOS1HS = rs("HSXOS1HS")         ' 品ＳＸ酸素析出１保証方法＿処
            If fldNameExist("HSXOS1HM") Then .HSXOS1HM = rs("HSXOS1HM")         ' 品ＳＸ酸素析出１検査頻度＿枚
            If fldNameExist("HSXOS1KI") Then .HSXOS1KI = rs("HSXOS1KI")         ' 品ＳＸ酸素析出１検査頻度＿位
            If fldNameExist("HSXOS1KH") Then .HSXOS1KH = rs("HSXOS1KH")         ' 品ＳＸ酸素析出１検査頻度＿保
            If fldNameExist("HSXOS1KS") Then .HSXOS1KS = rs("HSXOS1KS")         ' 品ＳＸ酸素析出１検査頻度＿試
            If fldNameExist("HSXOS2MN") Then .HSXOS2MN = fncNullCheck(rs("HSXOS2MN"))         ' 品ＳＸ酸素析出２下限
            If fldNameExist("HSXOS2MX") Then .HSXOS2MX = fncNullCheck(rs("HSXOS2MX"))         ' 品ＳＸ酸素析出２上限
            If fldNameExist("HSXOS2NS") Then .HSXOS2NS = rs("HSXOS2NS")         ' 品ＳＸ酸素析出２熱処理法
            If fldNameExist("HSXOS2SH") Then .HSXOS2SH = rs("HSXOS2SH")         ' 品ＳＸ酸素析出２測定位置＿方
            If fldNameExist("HSXOS2ST") Then .HSXOS2ST = rs("HSXOS2ST")         ' 品ＳＸ酸素析出２測定位置＿点
            If fldNameExist("HSXOS2SI") Then .HSXOS2SI = rs("HSXOS2SI")         ' 品ＳＸ酸素析出２測定位置＿位
            If fldNameExist("HSXOS2HT") Then .HSXOS2HT = rs("HSXOS2HT")         ' 品ＳＸ酸素析出２保証方法＿対
            If fldNameExist("HSXOS2HS") Then .HSXOS2HS = rs("HSXOS2HS")         ' 品ＳＸ酸素析出２保証方法＿処
            If fldNameExist("HSXOS2KM") Then .HSXOS2KM = rs("HSXOS2KM")         ' 品ＳＸ酸素析出２検査頻度＿枚
            If fldNameExist("HSXOS2KN") Then .HSXOS2KN = rs("HSXOS2KN")         ' 品ＳＸ酸素析出２検査頻度＿位
            If fldNameExist("HSXOS2KH") Then .HSXOS2KH = rs("HSXOS2KH")         ' 品ＳＸ酸素析出２検査頻度＿保
            If fldNameExist("HSXOS2KU") Then .HSXOS2KU = rs("HSXOS2KU")         ' 品ＳＸ酸素析出２検査頻度＿試
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")                  ' Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")         ' 処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")         ' 仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")            ' ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")               ' ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")            ' 社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")            ' 登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")            ' 更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")         ' 送信フラグ
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME019 = FUNCTION_RETURN_SUCCESS
  

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

'概要      :テーブル「TBCME020」から条件にあったレコードを抽出する
'          :records()     ,O  ,typ_TBCME020 ,抽出レコード
'          :formID        ,I  ,String       ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban  ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2001/06/27作成　長野

Public Function DBDRV_GetTBCME020(records() As typ_TBCME020, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL全体
Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
Dim sqlWhere    As String           'SQLWhere部
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             'レコード数
Dim key         As String           '検索KEY
Dim i           As Long             'ﾙｰﾌﾟｶｳﾝﾄ
Dim j           As Long             'ﾙｰﾌﾟｶｳﾝﾄ2


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME020_SQL.bas -- Function DBDRV_GetTBCME020"

   Select Case formID
        Case "f_cmbc021_1"           '「FTIR(Oi,Cs)実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
              For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
              Next
              sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc022_1"           '「GFA(Oi)実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc023_1"           '「抵抗実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
              
        Case "f_cmbc024_1"           '「BMD実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP,"
' OSF，BMD項目追加対応  ↑　1行分　2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec030_1"           '「BMD実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXBMD1MBP, HSXBMD2MBP, HSXBMD3MBP,"
' OSF，BMD項目追加対応  ↑　1行分　2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc025_1"           '「OSF実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK,"
' OSF，BMD項目追加対応  ↑　1行分　2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec031_1"           '「OSF実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE," & _
              " HSXOSF1PTK, HSXOSF2PTK, HSXOSF3PTK, HSXOSF4PTK,"
' OSF，BMD項目追加対応  ↑　1行分　2002.04.02 yakimura
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc027_1"           '「ライフタイム実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc028_1"           '「FPD実績入力」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
             sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc029_1"           '「GFA校正情報設定」
             sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, HSXDENKU, HSXDENMX, HSXDENMN," & _
              " HSXDENHT, HSXDENHS, HSXDVDKU, HSXDVDMXN, HSXDVDMNN, HSXDVDHT, HSXDVDHS, HSXLDLKU, HSXLDLMX, HSXLDLMN, HSXLDLHT," & _
              " HSXLDLHS, HSXGDSZY, HSXGDSPH, HSXGDSPT, HSXGDSPR, HSXGDZAR, HSXGDKHM, HSXGDKHI, HSXGDKHH, HSXGDKHS, HSXDSOKE," & _
              " HSXDSOMX, HSXDSOMN, HSXDSOAX, HSXDSOAN, HSXDSOHT, HSXDSOHS, HSXDSOKM, HSXDSOKI, HSXDSOKH, HSXDSOKS, HSXLIFTW," & _
              " HSXSDSLP, HSXGKKNO, HSXCDOP, HSXCDOPN, HSXCDPNI, HSXGSFIN, HSXCLMIN, HSXCLMAX, HSXCLPMN, HSXCLPR, HSXWFWAR," & _
              " HSXOF1AX, HSXOF1MX, HSXOF1SH, HSXOF1ST, HSXOF1SR, HSXOF1HT, HSXOF1HS, HSXOF1SZ, HSXOF1KM, HSXOF1KI, HSXOF1KH," & _
              " HSXOF1KS, HSXOF1NS, HSXOF1ET, HSXOF2AX, HSXOF2MX, HSXOF2SH, HSXOF2ST, HSXOF2SR, HSXOF2HT, HSXOF2HS, HSXOF2SZ," & _
              " HSXOF2KM, HSXOF2KI, HSXOF2KH, HSXOF2KS, HSXOF2NS, HSXOF2ET, HSXOF3AX, HSXOF3MX, HSXOF3SH, HSXOF3ST, HSXOF3SR,"
            sqlBase = sqlBase & " HSXOF3HT, HSXOF3HS, HSXOF3SZ, HSXOF3KM, HSXOF3KI, HSXOF3KH, HSXOF3KS, HSXOF3NS, HSXOF3ET, HSXOF4AX, HSXOF4MX," & _
              " HSXOF4SH, HSXOF4ST, HSXOF4SR, HSXOF4HT, HSXOF4HS, HSXOF4SZ, HSXOF4KM, HSXOF4KI, HSXOF4KH, HSXOF4KS, HSXOF4NS," & _
              " HSXOF4ET, HSXBM1AN, HSXBM1AX, HSXBM1SH, HSXBM1ST, HSXBM1SR, HSXBM1HT, HSXBM1HS, HSXBM1SZ, HSXBM1KM, HSXBM1KI," & _
              " HSXBM1KH, HSXBM1KS, HSXBM1NS, HSXBM1ET, HSXBM2AN, HSXBM2AX, HSXBM2SH, HSXBM2ST, HSXBM2SR, HSXBM2HT, HSXBM2HS," & _
              " HSXBM2SZ, HSXBM2KM, HSXBM2KI, HSXBM2KH, HSXBM2KS, HSXBM2NS, HSXBM2ET, HSXBM3AN, HSXBM3AX, HSXBM3SH, HSXBM3ST," & _
              " HSXBM3SR, HSXBM3HT, HSXBM3HS, HSXBM3SZ, HSXBM3KM, HSXBM3KI, HSXBM3KH, HSXBM3KS, HSXBM3NS, HSXBM3ET, HSXNOTE,"
            For i = 1 To 10
                sqlBase = sqlBase & "HSXRS" & i & "N, "
                sqlBase = sqlBase & "HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    
    End Select
       
    sqlBase = sqlBase & "From TBCME020"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME020 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''フィールド名を登録する
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
     
    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")             ' 品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")       ' 製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")          ' 工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")          ' 操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")    ' 品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")       ' 品管理社員Ｎｏ
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")       ' 品管理ＳＸ製品番号
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))  ' 品管理ＳＸ製品番号枝番
            If fldNameExist("HSXDENKU") Then .HSXDENKU = rs("HSXDENKU")       ' 品ＳＸＤｅｎ検査有無
            If fldNameExist("HSXDENMX") Then .HSXDENMX = fncNullCheck(rs("HSXDENMX"))  ' 品ＳＸＤｅｎ上限
            If fldNameExist("HSXDENMN") Then .HSXDENMN = fncNullCheck(rs("HSXDENMN"))  ' 品ＳＸＤｅｎ下限
            If fldNameExist("HSXDENHT") Then .HSXDENHT = rs("HSXDENHT")       ' 品ＳＸＤｅｎ保証方法＿対
            If fldNameExist("HSXDENHS") Then .HSXDENHS = rs("HSXDENHS")       ' 品ＳＸＤｅｎ保証方法＿処
            If fldNameExist("HSXDVDKU") Then .HSXDVDKU = rs("HSXDVDKU")       ' 品ＳＸＤＶＤ２検査有無
            If fldNameExist("HSXDVDMXN") Then .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN")) ' 品ＳＸＤＶＤ２上限    ＷＦサンプル処理変更 2003.05.20 yakimura
            If fldNameExist("HSXDVDMNN") Then .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN")) ' 品ＳＸＤＶＤ２下限    ＷＦサンプル処理変更 2003.05.20 yakimura
            If fldNameExist("HSXDVDHT") Then .HSXDVDHT = rs("HSXDVDHT")       ' 品ＳＸＤＶＤ２保証方法＿対
            If fldNameExist("HSXDVDHS") Then .HSXDVDHS = rs("HSXDVDHS")       ' 品ＳＸＤＶＤ２保証方法＿処
            If fldNameExist("HSXLDLKU") Then .HSXLDLKU = rs("HSXLDLKU")       ' 品ＳＸＬ／ＤＬ検査有無
            If fldNameExist("HSXLDLMX") Then .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))   ' 品ＳＸＬ／ＤＬ上限
            If fldNameExist("HSXLDLMN") Then .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))   ' 品ＳＸＬ／ＤＬ下限
            If fldNameExist("HSXLDLHT") Then .HSXLDLHT = rs("HSXLDLHT")       ' 品ＳＸＬ／ＤＬ保証方法＿対
            If fldNameExist("HSXLDLHS") Then .HSXLDLHS = rs("HSXLDLHS")       ' 品ＳＸＬ／ＤＬ保証方法＿処
            If fldNameExist("HSXGDSZY") Then .HSXGDSZY = rs("HSXGDSZY")       ' 品ＳＸＧＤ測定条件
            If fldNameExist("HSXGDSPH") Then .HSXGDSPH = rs("HSXGDSPH")       ' 品ＳＸＧＤ測定位置＿方
            If fldNameExist("HSXGDSPT") Then .HSXGDSPT = rs("HSXGDSPT")       ' 品ＳＸＧＤ測定位置＿点
            If fldNameExist("HSXGDSPR") Then .HSXGDSPR = rs("HSXGDSPR")       ' 品ＳＸＧＤ測定位置＿領
            If fldNameExist("HSXGDZAR") Then .HSXGDZAR = fncNullCheck(rs("HSXGDZAR"))   ' 品ＳＸＧＤ除外領域
            If fldNameExist("HSXGDKHM") Then .HSXGDKHM = rs("HSXGDKHM")       ' 品ＳＸＧＤ検査頻度＿枚
            If fldNameExist("HSXGDKHI") Then .HSXGDKHI = rs("HSXGDKHI")       ' 品ＳＸＧＤ検査頻度＿位
            If fldNameExist("HSXGDKHH") Then .HSXGDKHH = rs("HSXGDKHH")       ' 品ＳＸＧＤ検査頻度＿保
            If fldNameExist("HSXGDKHS") Then .HSXGDKHS = rs("HSXGDKHS")       ' 品ＳＸＧＤ検査頻度＿試
            If fldNameExist("HSXDSOKE") Then .HSXDSOKE = rs("HSXDSOKE")       ' 品ＳＸＤＳＯＤ検査
            If fldNameExist("HSXDSOMX") Then .HSXDSOMX = fncNullCheck(rs("HSXDSOMX"))  ' 品ＳＸＤＳＯＤ上限
            If fldNameExist("HSXDSOMN") Then .HSXDSOMN = fncNullCheck(rs("HSXDSOMN"))  ' 品ＳＸＤＳＯＤ下限
            If fldNameExist("HSXDSOAX") Then .HSXDSOAX = fncNullCheck(rs("HSXDSOAX"))  ' 品ＳＸＤＳＯＤ領域上限
            If fldNameExist("HSXDSOAN") Then .HSXDSOAN = fncNullCheck(rs("HSXDSOAN"))  ' 品ＳＸＤＳＯＤ領域下限
            If fldNameExist("HSXDSOHT") Then .HSXDSOHT = rs("HSXDSOHT")       ' 品ＳＸＤＳＯＤ保証方法＿対
            If fldNameExist("HSXDSOHS") Then .HSXDSOHS = rs("HSXDSOHS")       ' 品ＳＸＤＳＯＤ保証方法＿処
            If fldNameExist("HSXDSOKM") Then .HSXDSOKM = rs("HSXDSOKM")       ' 品ＳＸＤＳＯＤ検査頻度＿枚
            If fldNameExist("HSXDSOKI") Then .HSXDSOKI = rs("HSXDSOKI")       ' 品ＳＸＤＳＯＤ検査頻度＿位
            If fldNameExist("HSXDSOKH") Then .HSXDSOKH = rs("HSXDSOKH")       ' 品ＳＸＤＳＯＤ検査頻度＿保
            If fldNameExist("HSXDSOKS") Then .HSXDSOKS = rs("HSXDSOKS")       ' 品ＳＸＤＳＯＤ検査頻度＿試
            If fldNameExist("HSXLIFTW") Then .HSXLIFTW = rs("HSXLIFTW")       ' 品ＳＸ引上方法
            If fldNameExist("HSXSDSLP") Then .HSXSDSLP = rs("HSXSDSLP")       ' 品ＳＸシード傾
            If fldNameExist("HSXGKKNO") Then .HSXGKKNO = rs("HSXGKKNO")       ' 品ＳＸ外観規格Ｎｏ
            If fldNameExist("HSXCDOP") Then .HSXCDOP = rs("HSXCDOP")         ' 品ＳＸ結晶ドープ
            If fldNameExist("HSXCDOPN") Then .HSXCDOPN = fncNullCheck(rs("HSXCDOPN"))       ' 品ＳＸ結晶ドープ濃度
            If fldNameExist("HSXCDPNI") Then .HSXCDPNI = rs("HSXCDPNI")       ' 品ＳＸ結晶ドープ濃度指数
            If fldNameExist("HSXGSFIN") Then .HSXGSFIN = rs("HSXGSFIN")       ' 品ＳＸ外周仕上げ
            If fldNameExist("HSXCLMIN") Then .HSXCLMIN = fncNullCheck(rs("HSXCLMIN"))  ' 品ＳＸ結晶長下限
            If fldNameExist("HSXCLMAX") Then .HSXCLMAX = fncNullCheck(rs("HSXCLMAX"))  ' 品ＳＸ結晶長上限
            If fldNameExist("HSXCLPMN") Then .HSXCLPMN = fncNullCheck(rs("HSXCLPMN"))  ' 品ＳＸ結晶長許容下限
            If fldNameExist("HSXCLPR") Then .HSXCLPR = fncNullCheck(rs("HSXCLPR"))     ' 品ＳＸ結晶長許容比率
            If fldNameExist("HSXWFWAR") Then .HSXWFWAR = rs("HSXWFWAR")       ' 品ＳＸＷＦＷａｒｐランク
#If False Then  'テーブルの型定義がs_cmzcTableDefs.basで異なるための対応
            For j = 1 To 4
                If fldNameExist("HSXOF" & j & "AX") Then .HSXOF_AX(j) = fncNullCheck(rs("HSXOF" & j & "AX"))  ' 品ＳＸＯＳＦ(n)平均上限
                If fldNameExist("HSXOF" & j & "MX") Then .HSXOF_MX(j) = fncNullCheck(rs("HSXOF" & j & "MX"))  ' 品ＳＸＯＳＦ(n)上限
                If fldNameExist("HSXOF" & j & "SH") Then .HSXOF_SH(j) = rs("HSXOF" & j & "SH")  ' 品ＳＸＯＳＦ(n)測定位置＿方
                If fldNameExist("HSXOF" & j & "ST") Then .HSXOF_ST(j) = rs("HSXOF" & j & "ST")  ' 品ＳＸＯＳＦ(n)測定位置＿点
                If fldNameExist("HSXOF" & j & "SR") Then .HSXOF_SR(j) = rs("HSXOF" & j & "SR")  ' 品ＳＸＯＳＦ(n)測定位置＿領
                If fldNameExist("HSXOF" & j & "HT") Then .HSXOF_HT(j) = rs("HSXOF" & j & "HT")  ' 品ＳＸＯＳＦ(n)保証方法＿対
                If fldNameExist("HSXOF" & j & "HS") Then .HSXOF_HS(j) = rs("HSXOF" & j & "HS")  ' 品ＳＸＯＳＦ(n)保証方法＿処
                If fldNameExist("HSXOF" & j & "SZ") Then .HSXOF_SZ(j) = rs("HSXOF" & j & "SZ")  ' 品ＳＸＯＳＦ(n)測定条件
                If fldNameExist("HSXOF" & j & "KM") Then .HSXOF_KM(j) = rs("HSXOF" & j & "KM")  ' 品ＳＸＯＳＦ(n)検査頻度＿枚
                If fldNameExist("HSXOF" & j & "KI") Then .HSXOF_KI(j) = rs("HSXOF" & j & "KI")  ' 品ＳＸＯＳＦ(n)検査頻度＿位
                If fldNameExist("HSXOF" & j & "KH") Then .HSXOF_KH(j) = rs("HSXOF" & j & "KH")  ' 品ＳＸＯＳＦ(n)検査頻度＿保
                If fldNameExist("HSXOF" & j & "KS") Then .HSXOF_KS(j) = rs("HSXOF" & j & "KS")  ' 品ＳＸＯＳＦ(n)検査頻度＿試
                If fldNameExist("HSXOF" & j & "NS") Then .HSXOF_NS(j) = rs("HSXOF" & j & "NS")  ' 品ＳＸＯＳＦ(n)熱処理法
                If fldNameExist("HSXOF" & j & "ET") Then .HSXOF_ET(j) = fncNullCheck(rs("HSXOF" & j & "ET"))  ' 品ＳＸＯＳＦ(n)選択ＥＴ代
                'NULL対応
                If fldNameExist("HSXOSF" & j & "PTK") Then                       ' 品ＳＸＯＳＦ(n)パタン区分
                   If IsNull(rs("HSXOSF" & j & "PTK")) = False Then .HSXOSF_PTK(j) = rs("HSXOSF" & j & "PTK")
                End If
            Next
            For j = 1 To 3
                If fldNameExist("HSXBM" & j & "AN") Then .HSXBM_AN(j) = fncNullCheck(rs("HSXBM" & j & "AN"))  ' 品ＳＸＢＭＤ(n)平均下限
                If fldNameExist("HSXBM" & j & "AX") Then .HSXBM_AX(j) = fncNullCheck(rs("HSXBM" & j & "AX"))  ' 品ＳＸＢＭＤ(n)平均上限
                If fldNameExist("HSXBM" & j & "SH") Then .HSXBM_SH(j) = rs("HSXBM" & j & "SH")  ' 品ＳＸＢＭＤ(n)測定位置＿方
                If fldNameExist("HSXBM" & j & "ST") Then .HSXBM_ST(j) = rs("HSXBM" & j & "ST")  ' 品ＳＸＢＭＤ(n)測定位置＿点
                If fldNameExist("HSXBM" & j & "SR") Then .HSXBM_SR(j) = rs("HSXBM" & j & "SR")  ' 品ＳＸＢＭＤ(n)測定位置＿領
                If fldNameExist("HSXBM" & j & "HT") Then .HSXBM_HT(j) = rs("HSXBM" & j & "HT")  ' 品ＳＸＢＭＤ(n)保証方法＿対
                If fldNameExist("HSXBM" & j & "HS") Then .HSXBM_HS(j) = rs("HSXBM" & j & "HS")  ' 品ＳＸＢＭＤ(n)保証方法＿処
                If fldNameExist("HSXBM" & j & "SZ") Then .HSXBM_SZ(j) = rs("HSXBM" & j & "SZ")  ' 品ＳＸＢＭＤ(n)測定条件
                If fldNameExist("HSXBM" & j & "KM") Then .HSXBM_KM(j) = rs("HSXBM" & j & "KM")  ' 品ＳＸＢＭＤ(n)検査頻度＿枚
                If fldNameExist("HSXBM" & j & "KI") Then .HSXBM_KI(j) = rs("HSXBM" & j & "KI")  ' 品ＳＸＢＭＤ(n)検査頻度＿位
                If fldNameExist("HSXBM" & j & "KH") Then .HSXBM_KH(j) = rs("HSXBM" & j & "KH")  ' 品ＳＸＢＭＤ(n)検査頻度＿保
                If fldNameExist("HSXBM" & j & "KS") Then .HSXBM_KS(j) = rs("HSXBM" & j & "KS")  ' 品ＳＸＢＭＤ(n)検査頻度＿試
                If fldNameExist("HSXBM" & j & "NS") Then .HSXBM_NS(j) = rs("HSXBM" & j & "NS")  ' 品ＳＸＢＭＤ(n)熱処理法
                If fldNameExist("HSXBM" & j & "ET") Then .HSXBM_ET(j) = fncNullCheck(rs("HSXBM" & j & "ET"))  ' 品ＳＸＢＭＤ(n)選択ＥＴ代
                'NULL対応
                If fldNameExist("HSXBMD" & j & "MBP") Then                      ' 品ＳＸＢＭＤ(n)面内分布
                   If IsNull(rs("HSXBMD" & j & "MBP")) = False Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP"))
                End If
            Next
#Else
                If fldNameExist("HSXOF1AX") Then .HSXOF1AX = fncNullCheck(rs("HSXOF1AX"))  ' 品ＳＸＯＳＦ1平均上限
                If fldNameExist("HSXOF1MX") Then .HSXOF1MX = fncNullCheck(rs("HSXOF1MX"))  ' 品ＳＸＯＳＦ1上限
                If fldNameExist("HSXOF1SH") Then .HSXOF1SH = rs("HSXOF1SH")  ' 品ＳＸＯＳＦ1測定位置＿方
                If fldNameExist("HSXOF1ST") Then .HSXOF1ST = rs("HSXOF1ST")  ' 品ＳＸＯＳＦ1測定位置＿点
                If fldNameExist("HSXOF1SR") Then .HSXOF1SR = rs("HSXOF1SR")  ' 品ＳＸＯＳＦ1測定位置＿領
                If fldNameExist("HSXOF1HT") Then .HSXOF1HT = rs("HSXOF1HT")  ' 品ＳＸＯＳＦ1保証方法＿対
                If fldNameExist("HSXOF1HS") Then .HSXOF1HS = rs("HSXOF1HS")  ' 品ＳＸＯＳＦ1保証方法＿処
                If fldNameExist("HSXOF1SZ") Then .HSXOF1SZ = rs("HSXOF1SZ")  ' 品ＳＸＯＳＦ1測定条件
                If fldNameExist("HSXOF1KM") Then .HSXOF1KM = rs("HSXOF1KM")  ' 品ＳＸＯＳＦ1検査頻度＿枚
                If fldNameExist("HSXOF1KI") Then .HSXOF1KI = rs("HSXOF1KI")  ' 品ＳＸＯＳＦ1検査頻度＿位
                If fldNameExist("HSXOF1KH") Then .HSXOF1KH = rs("HSXOF1KH")  ' 品ＳＸＯＳＦ1検査頻度＿保
                If fldNameExist("HSXOF1KS") Then .HSXOF1KS = rs("HSXOF1KS")  ' 品ＳＸＯＳＦ1検査頻度＿試
                If fldNameExist("HSXOF1NS") Then .HSXOF1NS = rs("HSXOF1NS")  ' 品ＳＸＯＳＦ1熱処理法
                If fldNameExist("HSXOF1ET") Then .HSXOF1ET = fncNullCheck(rs("HSXOF1ET"))  ' 品ＳＸＯＳＦ1選択ＥＴ代
                If fldNameExist("HSXOSF1PTK") Then                           ' 品ＳＸＯＳＦ1パタン区分
                   If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")
                End If
                If fldNameExist("HSXOF2AX") Then .HSXOF2AX = fncNullCheck(rs("HSXOF2AX"))  ' 品ＳＸＯＳＦ2平均上限
                If fldNameExist("HSXOF2MX") Then .HSXOF2MX = fncNullCheck(rs("HSXOF2MX"))  ' 品ＳＸＯＳＦ2上限
                If fldNameExist("HSXOF2SH") Then .HSXOF2SH = rs("HSXOF2SH")  ' 品ＳＸＯＳＦ2測定位置＿方
                If fldNameExist("HSXOF2ST") Then .HSXOF2ST = rs("HSXOF2ST")  ' 品ＳＸＯＳＦ2測定位置＿点
                If fldNameExist("HSXOF2SR") Then .HSXOF2SR = rs("HSXOF2SR")  ' 品ＳＸＯＳＦ2測定位置＿領
                If fldNameExist("HSXOF2HT") Then .HSXOF2HT = rs("HSXOF2HT")  ' 品ＳＸＯＳＦ2保証方法＿対
                If fldNameExist("HSXOF2HS") Then .HSXOF2HS = rs("HSXOF2HS")  ' 品ＳＸＯＳＦ2保証方法＿処
                If fldNameExist("HSXOF2SZ") Then .HSXOF2SZ = rs("HSXOF2SZ")  ' 品ＳＸＯＳＦ2測定条件
                If fldNameExist("HSXOF2KM") Then .HSXOF2KM = rs("HSXOF2KM")  ' 品ＳＸＯＳＦ2検査頻度＿枚
                If fldNameExist("HSXOF2KI") Then .HSXOF2KI = rs("HSXOF2KI")  ' 品ＳＸＯＳＦ2検査頻度＿位
                If fldNameExist("HSXOF2KH") Then .HSXOF2KH = rs("HSXOF2KH")  ' 品ＳＸＯＳＦ2検査頻度＿保
                If fldNameExist("HSXOF2KS") Then .HSXOF2KS = rs("HSXOF2KS")  ' 品ＳＸＯＳＦ2検査頻度＿試
                If fldNameExist("HSXOF2NS") Then .HSXOF2NS = rs("HSXOF2NS")  ' 品ＳＸＯＳＦ2熱処理法
                If fldNameExist("HSXOF2ET") Then .HSXOF2ET = fncNullCheck(rs("HSXOF2ET"))  ' 品ＳＸＯＳＦ2選択ＥＴ代
                If fldNameExist("HSXOSF2PTK") Then                           ' 品ＳＸＯＳＦ2パタン区分
                   If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")
                End If
                If fldNameExist("HSXOF3AX") Then .HSXOF3AX = fncNullCheck(rs("HSXOF3AX"))  ' 品ＳＸＯＳＦ3平均上限
                If fldNameExist("HSXOF3MX") Then .HSXOF3MX = fncNullCheck(rs("HSXOF3MX"))  ' 品ＳＸＯＳＦ3上限
                If fldNameExist("HSXOF3SH") Then .HSXOF3SH = rs("HSXOF3SH")  ' 品ＳＸＯＳＦ3測定位置＿方
                If fldNameExist("HSXOF3ST") Then .HSXOF3ST = rs("HSXOF3ST")  ' 品ＳＸＯＳＦ3測定位置＿点
                If fldNameExist("HSXOF3SR") Then .HSXOF3SR = rs("HSXOF3SR")  ' 品ＳＸＯＳＦ3測定位置＿領
                If fldNameExist("HSXOF3HT") Then .HSXOF3HT = rs("HSXOF3HT")  ' 品ＳＸＯＳＦ3保証方法＿対
                If fldNameExist("HSXOF3HS") Then .HSXOF3HS = rs("HSXOF3HS")  ' 品ＳＸＯＳＦ3保証方法＿処
                If fldNameExist("HSXOF3SZ") Then .HSXOF3SZ = rs("HSXOF3SZ")  ' 品ＳＸＯＳＦ3測定条件
                If fldNameExist("HSXOF3KM") Then .HSXOF3KM = rs("HSXOF3KM")  ' 品ＳＸＯＳＦ3検査頻度＿枚
                If fldNameExist("HSXOF3KI") Then .HSXOF3KI = rs("HSXOF3KI")  ' 品ＳＸＯＳＦ3検査頻度＿位
                If fldNameExist("HSXOF3KH") Then .HSXOF3KH = rs("HSXOF3KH")  ' 品ＳＸＯＳＦ3検査頻度＿保
                If fldNameExist("HSXOF3KS") Then .HSXOF3KS = rs("HSXOF3KS")  ' 品ＳＸＯＳＦ3検査頻度＿試
                If fldNameExist("HSXOF3NS") Then .HSXOF3NS = rs("HSXOF3NS")  ' 品ＳＸＯＳＦ3熱処理法
                If fldNameExist("HSXOF3ET") Then .HSXOF3ET = fncNullCheck(rs("HSXOF3ET"))  ' 品ＳＸＯＳＦ3選択ＥＴ代
                If fldNameExist("HSXOSF3PTK") Then                           ' 品ＳＸＯＳＦ3パタン区分
                   If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")
                End If
                If fldNameExist("HSXOF4AX") Then .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))  ' 品ＳＸＯＳＦ4平均上限
                If fldNameExist("HSXOF4MX") Then .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))  ' 品ＳＸＯＳＦ4上限
                If fldNameExist("HSXOF4SH") Then .HSXOF4SH = rs("HSXOF4SH")  ' 品ＳＸＯＳＦ4測定位置＿方
                If fldNameExist("HSXOF4ST") Then .HSXOF4ST = rs("HSXOF4ST")  ' 品ＳＸＯＳＦ4測定位置＿点
                If fldNameExist("HSXOF4SR") Then .HSXOF4SR = rs("HSXOF4SR")  ' 品ＳＸＯＳＦ4測定位置＿領
                If fldNameExist("HSXOF4HT") Then .HSXOF4HT = rs("HSXOF4HT")  ' 品ＳＸＯＳＦ4保証方法＿対
                If fldNameExist("HSXOF4HS") Then .HSXOF4HS = rs("HSXOF4HS")  ' 品ＳＸＯＳＦ4保証方法＿処
                If fldNameExist("HSXOF4SZ") Then .HSXOF4SZ = rs("HSXOF4SZ")  ' 品ＳＸＯＳＦ4測定条件
                If fldNameExist("HSXOF4KM") Then .HSXOF4KM = rs("HSXOF4KM")  ' 品ＳＸＯＳＦ4検査頻度＿枚
                If fldNameExist("HSXOF4KI") Then .HSXOF4KI = rs("HSXOF4KI")  ' 品ＳＸＯＳＦ4検査頻度＿位
                If fldNameExist("HSXOF4KH") Then .HSXOF4KH = rs("HSXOF4KH")  ' 品ＳＸＯＳＦ4検査頻度＿保
                If fldNameExist("HSXOF4KS") Then .HSXOF4KS = rs("HSXOF4KS")  ' 品ＳＸＯＳＦ4検査頻度＿試
                If fldNameExist("HSXOF4NS") Then .HSXOF4NS = rs("HSXOF4NS")  ' 品ＳＸＯＳＦ4熱処理法
                If fldNameExist("HSXOF4ET") Then .HSXOF4ET = fncNullCheck(rs("HSXOF4ET"))  ' 品ＳＸＯＳＦ4選択ＥＴ代
                If fldNameExist("HSXOSF4PTK") Then                           ' 品ＳＸＯＳＦ4パタン区分
                   If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")
                End If
                If fldNameExist("HSXBM1AN") Then .HSXBM1AN = fncNullCheck(rs("HSXBM1AN"))  ' 品ＳＸＢＭＤ1平均下限
                If fldNameExist("HSXBM1AX") Then .HSXBM1AX = fncNullCheck(rs("HSXBM1AX"))  ' 品ＳＸＢＭＤ1平均上限
                If fldNameExist("HSXBM1SH") Then .HSXBM1SH = rs("HSXBM1SH")  ' 品ＳＸＢＭＤ1測定位置＿方
                If fldNameExist("HSXBM1ST") Then .HSXBM1ST = rs("HSXBM1ST")  ' 品ＳＸＢＭＤ1測定位置＿点
                If fldNameExist("HSXBM1SR") Then .HSXBM1SR = rs("HSXBM1SR")  ' 品ＳＸＢＭＤ1測定位置＿領
                If fldNameExist("HSXBM1HT") Then .HSXBM1HT = rs("HSXBM1HT")  ' 品ＳＸＢＭＤ1保証方法＿対
                If fldNameExist("HSXBM1HS") Then .HSXBM1HS = rs("HSXBM1HS")  ' 品ＳＸＢＭＤ1保証方法＿処
                If fldNameExist("HSXBM1SZ") Then .HSXBM1SZ = rs("HSXBM1SZ")  ' 品ＳＸＢＭＤ1測定条件
                If fldNameExist("HSXBM1KM") Then .HSXBM1KM = rs("HSXBM1KM")  ' 品ＳＸＢＭＤ1検査頻度＿枚
                If fldNameExist("HSXBM1KI") Then .HSXBM1KI = rs("HSXBM1KI")  ' 品ＳＸＢＭＤ1検査頻度＿位
                If fldNameExist("HSXBM1KH") Then .HSXBM1KH = rs("HSXBM1KH")  ' 品ＳＸＢＭＤ1検査頻度＿保
                If fldNameExist("HSXBM1KS") Then .HSXBM1KS = rs("HSXBM1KS")  ' 品ＳＸＢＭＤ1検査頻度＿試
                If fldNameExist("HSXBM1NS") Then .HSXBM1NS = rs("HSXBM1NS")  ' 品ＳＸＢＭＤ1熱処理法
                If fldNameExist("HSXBM1ET") Then .HSXBM1ET = fncNullCheck(rs("HSXBM1ET"))  ' 品ＳＸＢＭＤ1選択ＥＴ代
                'NULL対応
                If fldNameExist("HSXBMD1MBP") Then                           ' 品ＳＸＢＭＤ1面内分布
                   If IsNull(rs("HSXBMD1MBP")) = False Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))
                End If
                If fldNameExist("HSXBM2AN") Then .HSXBM2AN = fncNullCheck(rs("HSXBM2AN"))  ' 品ＳＸＢＭＤ2平均下限
                If fldNameExist("HSXBM2AX") Then .HSXBM2AX = fncNullCheck(rs("HSXBM2AX"))  ' 品ＳＸＢＭＤ2平均上限
                If fldNameExist("HSXBM2SH") Then .HSXBM2SH = rs("HSXBM2SH")  ' 品ＳＸＢＭＤ2測定位置＿方
                If fldNameExist("HSXBM2ST") Then .HSXBM2ST = rs("HSXBM2ST")  ' 品ＳＸＢＭＤ2測定位置＿点
                If fldNameExist("HSXBM2SR") Then .HSXBM2SR = rs("HSXBM2SR")  ' 品ＳＸＢＭＤ2測定位置＿領
                If fldNameExist("HSXBM2HT") Then .HSXBM2HT = rs("HSXBM2HT")  ' 品ＳＸＢＭＤ2保証方法＿対
                If fldNameExist("HSXBM2HS") Then .HSXBM2HS = rs("HSXBM2HS")  ' 品ＳＸＢＭＤ2保証方法＿処
                If fldNameExist("HSXBM2SZ") Then .HSXBM2SZ = rs("HSXBM2SZ")  ' 品ＳＸＢＭＤ2測定条件
                If fldNameExist("HSXBM2KM") Then .HSXBM2KM = rs("HSXBM2KM")  ' 品ＳＸＢＭＤ2検査頻度＿枚
                If fldNameExist("HSXBM2KI") Then .HSXBM2KI = rs("HSXBM2KI")  ' 品ＳＸＢＭＤ2検査頻度＿位
                If fldNameExist("HSXBM2KH") Then .HSXBM2KH = rs("HSXBM2KH")  ' 品ＳＸＢＭＤ2検査頻度＿保
                If fldNameExist("HSXBM2KS") Then .HSXBM2KS = rs("HSXBM2KS")  ' 品ＳＸＢＭＤ2検査頻度＿試
                If fldNameExist("HSXBM2NS") Then .HSXBM2NS = rs("HSXBM2NS")  ' 品ＳＸＢＭＤ2熱処理法
                If fldNameExist("HSXBM2ET") Then .HSXBM2ET = fncNullCheck(rs("HSXBM2ET"))  ' 品ＳＸＢＭＤ2選択ＥＴ代
                'NULL対応
                If fldNameExist("HSXBMD2MBP") Then                           ' 品ＳＸＢＭＤ2面内分布
                   If IsNull(rs("HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("HSXBMD2MBP")
                End If
                If fldNameExist("HSXBM3AN") Then .HSXBM3AN = fncNullCheck(rs("HSXBM3AN"))  ' 品ＳＸＢＭＤ3平均下限
                If fldNameExist("HSXBM3AX") Then .HSXBM3AX = fncNullCheck(rs("HSXBM3AX"))  ' 品ＳＸＢＭＤ3平均上限
                If fldNameExist("HSXBM3SH") Then .HSXBM3SH = rs("HSXBM3SH")  ' 品ＳＸＢＭＤ3測定位置＿方
                If fldNameExist("HSXBM3ST") Then .HSXBM3ST = rs("HSXBM3ST")  ' 品ＳＸＢＭＤ3測定位置＿点
                If fldNameExist("HSXBM3SR") Then .HSXBM3SR = rs("HSXBM3SR")  ' 品ＳＸＢＭＤ3測定位置＿領
                If fldNameExist("HSXBM3HT") Then .HSXBM3HT = rs("HSXBM3HT")  ' 品ＳＸＢＭＤ3保証方法＿対
                If fldNameExist("HSXBM3HS") Then .HSXBM3HS = rs("HSXBM3HS")  ' 品ＳＸＢＭＤ3保証方法＿処
                If fldNameExist("HSXBM3SZ") Then .HSXBM3SZ = rs("HSXBM3SZ")  ' 品ＳＸＢＭＤ3測定条件
                If fldNameExist("HSXBM3KM") Then .HSXBM3KM = rs("HSXBM3KM")  ' 品ＳＸＢＭＤ3検査頻度＿枚
                If fldNameExist("HSXBM3KI") Then .HSXBM3KI = rs("HSXBM3KI")  ' 品ＳＸＢＭＤ3検査頻度＿位
                If fldNameExist("HSXBM3KH") Then .HSXBM3KH = rs("HSXBM3KH")  ' 品ＳＸＢＭＤ3検査頻度＿保
                If fldNameExist("HSXBM3KS") Then .HSXBM3KS = rs("HSXBM3KS")  ' 品ＳＸＢＭＤ3検査頻度＿試
                If fldNameExist("HSXBM3NS") Then .HSXBM3NS = rs("HSXBM3NS")  ' 品ＳＸＢＭＤ3熱処理法
                If fldNameExist("HSXBM3ET") Then .HSXBM3ET = fncNullCheck(rs("HSXBM3ET"))  ' 品ＳＸＢＭＤ3選択ＥＴ代
                'NULL対応
                If fldNameExist("HSXBMD3MBP") Then                           ' 品ＳＸＢＭＤ3面内分布
                    If IsNull(rs("HSXBMD3MBP")) = False Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))
                End If
#End If
            If fldNameExist("HSXNOTE") Then .HSXNOTE = rs("HSXNOTE")         ' 品ＳＸ特記
#If False Then  'テーブルの型定義がs_cmzcTableDefs.basで違うため無効とする
            For j = 1 To 10
                If fldNameExist("HSXRS" & j & "N") Then .HSXRS_N(j) = rs("HSXRS" & j & "N")     ' 品ＳＸ予備(n)＿内
                If fldNameExist("HSXRS" & j & "Y") Then .HSXRS_Y(j) = rs("HSXRS" & j & "Y")     ' 品ＳＸ予備(n)＿用
            Next
#Else
                If fldNameExist("HSXRS1N") Then .HSXRS1N = rs("HSXRS1N")     ' 品ＳＸ予備1＿内
                If fldNameExist("HSXRS2N") Then .HSXRS2N = rs("HSXRS2N")     ' 品ＳＸ予備2＿内
                If fldNameExist("HSXRS3N") Then .HSXRS3N = rs("HSXRS3N")     ' 品ＳＸ予備3＿内
                If fldNameExist("HSXRS4N") Then .HSXRS4N = rs("HSXRS4N")     ' 品ＳＸ予備4＿内
                If fldNameExist("HSXRS5N") Then .HSXRS5N = rs("HSXRS5N")     ' 品ＳＸ予備5＿内
                If fldNameExist("HSXRS6N") Then .HSXRS6N = rs("HSXRS6N")     ' 品ＳＸ予備6＿内
                If fldNameExist("HSXRS7N") Then .HSXRS7N = rs("HSXRS7N")     ' 品ＳＸ予備7＿内
                If fldNameExist("HSXRS8N") Then .HSXRS8N = rs("HSXRS8N")     ' 品ＳＸ予備8＿内
                If fldNameExist("HSXRS9N") Then .HSXRS9N = rs("HSXRS9N")     ' 品ＳＸ予備9＿内
                If fldNameExist("HSXRS10N") Then .HSXRS10N = rs("HSXRS10N")  ' 品ＳＸ予備10＿内
                If fldNameExist("HSXRS1Y") Then .HSXRS1Y = rs("HSXRS1Y")     ' 品ＳＸ予備1＿用
                If fldNameExist("HSXRS2Y") Then .HSXRS2Y = rs("HSXRS2Y")     ' 品ＳＸ予備2＿用
                If fldNameExist("HSXRS3Y") Then .HSXRS3Y = rs("HSXRS3Y")     ' 品ＳＸ予備3＿用
                If fldNameExist("HSXRS4Y") Then .HSXRS4Y = rs("HSXRS4Y")     ' 品ＳＸ予備4＿用
                If fldNameExist("HSXRS5Y") Then .HSXRS5Y = rs("HSXRS5Y")     ' 品ＳＸ予備5＿用
                If fldNameExist("HSXRS6Y") Then .HSXRS6Y = rs("HSXRS6Y")     ' 品ＳＸ予備6＿用
                If fldNameExist("HSXRS7Y") Then .HSXRS7Y = rs("HSXRS7Y")     ' 品ＳＸ予備7＿用
                If fldNameExist("HSXRS8Y") Then .HSXRS8Y = rs("HSXRS8Y")     ' 品ＳＸ予備8＿用
                If fldNameExist("HSXRS9Y") Then .HSXRS9Y = rs("HSXRS9Y")     ' 品ＳＸ予備9＿用
                If fldNameExist("HSXRS1YN") Then .HSXRS10Y = rs("HSXRS10Y")     ' 品ＳＸ予備10＿用
#End If
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")     ' 仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")        ' ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")           ' ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")        ' 社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")        ' 登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")        ' 更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")     ' 送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")     ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME020 = FUNCTION_RETURN_SUCCESS

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

Private Function fldNameExist(fldName As String) As Boolean
    Dim sql         As String           'SQL全体
    Dim i As Integer                    'ﾙｰﾌﾟｶｳﾝﾄ


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME***_SQL.bas -- Function fldNameExist"

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

'概要      :テーブル「TBCME018」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME018 ,抽出レコード
'          :formID        ,I  ,String       ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban  ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :2001/06/27作成　長野

Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, formID$, hin() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL全体
Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
Dim sqlWhere    As String           'SQLWhere部
Dim rs          As OraDynaset       'RecordSet
Dim recCnt      As Long             'レコード数
Dim key         As String           '検索KEY
Dim i           As Long             'ﾙｰﾌﾟｶｳﾝﾄ
    

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_TBCME018_SQL.bas -- Function DBDRV_GetTBCME018"

    Select Case formID
        Case "f_cmbc021_1"           '「FTIR(Oi,Cs)実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc022_1"           '「GFA(Oi)実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc023_1"           '「抵抗実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc024_1"           '「BMD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec030_1"           '「BMD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc025_1"           '「OSF実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmec031_1"           '「OSF実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc027_1"           '「ライフタイム実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc028_1"           '「FPD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        Case "f_cmbc029_1"           '「GFA校正情報設定」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    
    End Select
    
    sqlBase = sqlBase & "From TBCME018"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(hin)
        With hin(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(hin) Then
                key = key & ", "
            End If
        End With
    Next
    sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
    sql = sqlBase & sqlWhere
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME018 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''フィールド名を登録する
    fldCnt = rs.Fields.COUNT
    ReDim fldNames(fldCnt)
    For i = 1 To fldCnt
        fldNames(i) = rs.FieldName(i - 1)
    Next
    
    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")           ' 品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")     ' 製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")        ' 工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")        ' 操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")  ' 品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")     ' 品管理社員Ｎｏ
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")     ' 品管理ＳＸ製品番号
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))     ' 品管理ＳＸ製品番号枝番
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG")        ' 確認フラグ
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG")     ' 再付与フラグ
            If fldNameExist("HSXTRWKB") Then .HSXTRWKB = rs("HSXTRWKB")     ' 品ＳＸ統合可否区分
            If fldNameExist("HSXTYPE") Then .HSXTYPE = rs("HSXTYPE")        ' 品ＳＸタイプ
            If fldNameExist("KSXTYPKW") Then .KSXTYPKW = rs("KSXTYPKW")     ' 品ＳＸタイプ検査方法
            If fldNameExist("HSXDOP") Then .HSXDOP = rs("HSXDOP")           ' 品ＳＸドーパント
            If fldNameExist("HSXRMIN") Then .HSXRMIN = fncNullCheck(rs("HSXRMIN"))        ' 品ＳＸ比抵抗下限
            If fldNameExist("HSXRMAX") Then .HSXRMAX = fncNullCheck(rs("HSXRMAX"))        ' 品ＳＸ比抵抗上限
            If fldNameExist("HSXRSPOH") Then .HSXRSPOH = rs("HSXRSPOH")     ' 品ＳＸ比抵抗測定位置＿方
            If fldNameExist("HSXRSPOT") Then .HSXRSPOT = rs("HSXRSPOT")     ' 品ＳＸ比抵抗測定位置＿点
            If fldNameExist("HSXRSPOI") Then .HSXRSPOI = rs("HSXRSPOI")     ' 品ＳＸ比抵抗測定位置＿位
            If fldNameExist("HSXRHWYT") Then .HSXRHWYT = rs("HSXRHWYT")     ' 品ＳＸ比抵抗保証方法＿対
            If fldNameExist("HSXRHWYS") Then .HSXRHWYS = rs("HSXRHWYS")     ' 品ＳＸ比抵抗保証方法＿処
            If fldNameExist("HSXRKWAY") Then .HSXRKWAY = rs("HSXRKWAY")     ' 品ＳＸ比抵抗検査方法
            If fldNameExist("HSXRKHNM") Then .HSXRKHNM = rs("HSXRKHNM")     ' 品ＳＸ比抵抗検査頻度＿枚
            If fldNameExist("HSXRKHNI") Then .HSXRKHNI = rs("HSXRKHNI")     ' 品ＳＸ比抵抗検査頻度＿位
            If fldNameExist("HSXRKHNH") Then .HSXRKHNH = rs("HSXRKHNH")     ' 品ＳＸ比抵抗検査頻度＿保
            If fldNameExist("HSXRKHNS") Then .HSXRKHNS = rs("HSXRKHNS")     ' 品ＳＸ比抵抗検査頻度＿試
            If fldNameExist("HSXRMCAL") Then .HSXRMCAL = rs("HSXRMCAL")     ' 品ＳＸ比抵抗面内計算
            If fldNameExist("HSXRMBNP") Then .HSXRMBNP = fncNullCheck(rs("HSXRMBNP"))     ' 品ＳＸ比抵抗面内分布
            If fldNameExist("HSXRMCL2") Then .HSXRMCL2 = rs("HSXRMCL2")     ' 品ＳＸ比抵抗面内計算２
            If fldNameExist("HSXRMBP2") Then .HSXRMBP2 = fncNullCheck(rs("HSXRMBP2"))     ' 品ＳＸ比抵抗面内分布２
            If fldNameExist("HSXRSDEV") Then .HSXRSDEV = fncNullCheck(rs("HSXRSDEV"))     ' 品ＳＸ比抵抗標準偏差
            If fldNameExist("HSXRAMIN") Then .HSXRAMIN = fncNullCheck(rs("HSXRAMIN"))     ' 品ＳＸ比抵抗平均下限
            If fldNameExist("HSXRAMAX") Then .HSXRAMAX = fncNullCheck(rs("HSXRAMAX"))     ' 品ＳＸ比抵抗平均上限
            If fldNameExist("HSXFORM") Then .HSXFORM = rs("HSXFORM")        ' 品ＳＸ形状
            If fldNameExist("HSXD1CEN") Then .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))     ' 品ＳＸ直径１中心
            If fldNameExist("HSXD1MIN") Then .HSXD1MIN = fncNullCheck(rs("HSXD1MIN"))     ' 品ＳＸ直径１下限
            If fldNameExist("HSXD1MAX") Then .HSXD1MAX = fncNullCheck(rs("HSXD1MAX"))     ' 品ＳＸ直径１上限
            If fldNameExist("HSXD2CEN") Then .HSXD2CEN = fncNullCheck(rs("HSXD2CEN"))     ' 品ＳＸ直径２中心
            If fldNameExist("HSXD2MIN") Then .HSXD2MIN = fncNullCheck(rs("HSXD2MIN"))     ' 品ＳＸ直径２下限
            If fldNameExist("HSXD2MAX") Then .HSXD2MAX = fncNullCheck(rs("HSXD2MAX"))     ' 品ＳＸ直径２上限
            If fldNameExist("HSXCDIR") Then .HSXCDIR = rs("HSXCDIR")        ' 品ＳＸ結晶面方位
            If fldNameExist("HSXCSCEN") Then .HSXCSCEN = fncNullCheck(rs("HSXCSCEN"))     ' 品ＳＸ結晶面傾中心
            If fldNameExist("HSXCSMIN") Then .HSXCSMIN = fncNullCheck(rs("HSXCSMIN"))     ' 品ＳＸ結晶面傾下限
            If fldNameExist("HSXCSMAX") Then .HSXCSMAX = fncNullCheck(rs("HSXCSMAX"))     ' 品ＳＸ結晶面傾上限
            If fldNameExist("HSXCKWAY") Then .HSXCKWAY = rs("HSXCKWAY")     ' 品ＳＸ結晶面検査方法
            If fldNameExist("HSXCKHNM") Then .HSXCKHNM = rs("HSXCKHNM")     ' 品ＳＸ結晶面検査頻度＿枚
            If fldNameExist("HSXCKHNI") Then .HSXCKHNI = rs("HSXCKHNI")     ' 品ＳＸ結晶面検査頻度＿位
            If fldNameExist("HSXCKHNH") Then .HSXCKHNH = rs("HSXCKHNH")     ' 品ＳＸ結晶面検査頻度＿保
            If fldNameExist("HSXCKHNS") Then .HSXCKHNS = rs("HSXCKHNS")     ' 品ＳＸ結晶面検査頻度＿試
            If fldNameExist("HSXCSDIR") Then .HSXCSDIR = rs("HSXCSDIR")     ' 品ＳＸ結晶面傾方位
            If fldNameExist("HSXCSDIS") Then .HSXCSDIS = rs("HSXCSDIS")     ' 品ＳＸ結晶面傾方位指定
            If fldNameExist("HSXCTDIR") Then .HSXCTDIR = rs("HSXCTDIR")     ' 品ＳＸ結晶面傾縦方位
            If fldNameExist("HSXCTCEN") Then .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))   ' 品ＳＸ結晶面傾縦中心
            If fldNameExist("HSXCTMIN") Then .HSXCTMIN = fncNullCheck(rs("HSXCTMIN"))     ' 品ＳＸ結晶面傾縦下限
            If fldNameExist("HSXCTMAX") Then .HSXCTMAX = fncNullCheck(rs("HSXCTMAX"))     ' 品ＳＸ結晶面傾縦上限
            If fldNameExist("HSXCYDIR") Then .HSXCYDIR = rs("HSXCYDIR")     ' 品ＳＸ結晶面傾横方位
            If fldNameExist("HSXCYCEN") Then .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))     ' 品ＳＸ結晶面傾横中心
            If fldNameExist("HSXCYMIN") Then .HSXCYMIN = fncNullCheck(rs("HSXCYMIN"))     ' 品ＳＸ結晶面傾横下限
            If fldNameExist("HSXCYMAX") Then .HSXCYMAX = fncNullCheck(rs("HSXCYMAX"))     ' 品ＳＸ結晶面傾横上限
            If fldNameExist("HSXOF1PD") Then .HSXOF1PD = rs("HSXOF1PD")     ' 品ＳＸＯＦ１位置方位
            If fldNameExist("HSXOF1PN") Then .HSXOF1PN = fncNullCheck(rs("HSXOF1PN"))     ' 品ＳＸＯＦ１位置下限
            If fldNameExist("HSXOF1PX") Then .HSXOF1PX = fncNullCheck(rs("HSXOF1PX"))     ' 品ＳＸＯＦ１位置上限
            If fldNameExist("HSXOF1PW") Then .HSXOF1PW = rs("HSXOF1PW")     ' 品ＳＸＯＦ１位置検査方法
            If fldNameExist("HSXOF1LC") Then .HSXOF1LC = fncNullCheck(rs("HSXOF1LC"))     ' 品ＳＸＯＦ１長中心
            If fldNameExist("HSXOF1LN") Then .HSXOF1LN = fncNullCheck(rs("HSXOF1LN"))     ' 品ＳＸＯＦ１長下限
            If fldNameExist("HSXOF1LX") Then .HSXOF1LX = fncNullCheck(rs("HSXOF1LX"))     ' 品ＳＸＯＦ１長上限
            If fldNameExist("HSXOF1DC") Then .HSXOF1DC = fncNullCheck(rs("HSXOF1DC"))     ' 品ＳＸＯＦ１直径中心
            If fldNameExist("HSXOF1DN") Then .HSXOF1DN = fncNullCheck(rs("HSXOF1DN"))     ' 品ＳＸＯＦ１直径下限
            If fldNameExist("HSXOF1DX") Then .HSXOF1DX = fncNullCheck(rs("HSXOF1DX"))     ' 品ＳＸＯＦ１直径上限
            If fldNameExist("HSXDFORM") Then .HSXDFORM = rs("HSXDFORM")     ' 品ＳＸ溝形状
            If fldNameExist("HSXDPDRC") Then .HSXDPDRC = rs("HSXDPDRC")     ' 品ＳＸ溝位置方向
            If fldNameExist("HSXDPACN") Then .HSXDPACN = fncNullCheck(rs("HSXDPACN"))     ' 品ＳＸ溝位置角度中心
            If fldNameExist("HSXDPAMN") Then .HSXDPAMN = fncNullCheck(rs("HSXDPAMN"))     ' 品ＳＸ溝位置角度下限
            If fldNameExist("HSXDPAMX") Then .HSXDPAMX = fncNullCheck(rs("HSXDPAMX"))     ' 品ＳＸ溝位置角度上限
            If fldNameExist("HSXDPKWY") Then .HSXDPKWY = rs("HSXDPKWY")     ' 品ＳＸ溝位置検査方法
            If fldNameExist("HSXDPDIR") Then .HSXDPDIR = rs("HSXDPDIR")     ' 品ＳＸ溝位置方位
            If fldNameExist("HSXDPMIN") Then .HSXDPMIN = fncNullCheck(rs("HSXDPMIN"))     ' 品ＳＸ溝位置下限
            If fldNameExist("HSXDPMAX") Then .HSXDPMAX = fncNullCheck(rs("HSXDPMAX"))     ' 品ＳＸ溝位置上限
            If fldNameExist("HSXDWCEN") Then .HSXDWCEN = fncNullCheck(rs("HSXDWCEN"))     ' 品ＳＸ溝巾中心
            If fldNameExist("HSXDWMIN") Then .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))     ' 品ＳＸ溝巾下限
            If fldNameExist("HSXDWMAX") Then .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))     ' 品ＳＸ溝巾上限
            If fldNameExist("HSXDDCEN") Then .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))     ' 品ＳＸ溝深中心
            If fldNameExist("HSXDDMIN") Then .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))     ' 品ＳＸ溝深下限
            If fldNameExist("HSXDDMAX") Then .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))     ' 品ＳＸ溝深上限
            If fldNameExist("HSXDACEN") Then .HSXDACEN = fncNullCheck(rs("HSXDACEN"))     ' 品ＳＸ溝角度中心
            If fldNameExist("HSXDAMIN") Then .HSXDAMIN = fncNullCheck(rs("HSXDAMIN"))     ' 品ＳＸ溝角度下限
            If fldNameExist("HSXDAMAX") Then .HSXDAMAX = fncNullCheck(rs("HSXDAMAX"))     ' 品ＳＸ溝角度上限
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")              ' Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")     ' 処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")     ' 仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")        ' ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")           ' ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")        ' 社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")        ' 登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")        ' 更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")     ' 送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")     ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME018 = FUNCTION_RETURN_SUCCESS

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

'概要      :テーブル「TBCMJ002」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMJ002 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
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
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .POSITION = rs("POSITION")       ' 位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .TRANCOND = rs("TRANCOND")       ' 処理条件
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
            .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .hinban = rs("HINBAN")           ' 品番
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
'履歴      :2001/08/24作成　野村
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
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
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
'概要      :テーブル「XSDCS」の条件にあったレコードを更新する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records       ,I   ,typ_XSDCS   ,更新レコード
'          :[sqlWhere]    ,I   ,String         ,更新条件(SQLのWhere節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN  ,更新の成否
'説明      :
'履歴      :2001/07/13作成　伊藤
Public Function DBDRV_UpdateTBCME043(records As typ_XSDCS, Optional sqlWhere$ = vbNullString) As FUNCTION_RETURN
    Dim sql As String
    
    DBDRV_UpdateTBCME043 = FUNCTION_RETURN_FAILURE

    With records
'        sql = "update TBCME043 set "
''        sql = sql & "HINBAN='" & .HINBAN & "', "              ' 品番
''        sql = sql & "REVNUM=" & .REVNUM & ", "                ' 製品番号改訂番号
''        sql = sql & "FACTORY='" & .FACTORY & "', "            ' 工場
''        sql = sql & "OPECOND='" & .OPECOND & "', "            ' 操業条件
''        sql = sql & "KTKBN='" & .KTKBN & "', "                ' 確定区分
''        sql = sql & "CRYINDRS='" & .CRYINDRS & "', "          ' 結晶検査指示（Rs)
''        sql = sql & "CRYINDOI='" & .CRYINDOI & "', "          ' 結晶検査指示（Oi)
''        sql = sql & "CRYINDB1='" & .CRYINDB1 & "', "          ' 結晶検査指示（B1)
''        sql = sql & "CRYINDB2='" & .CRYINDB2 & "', "          ' 結晶検査指示（B2）
''        sql = sql & "CRYINDB3='" & .CRYINDB3 & "', "          ' 結晶検査指示（B3)
''        sql = sql & "CRYINDL1='" & .CRYINDL1 & "', "          ' 結晶検査指示（L1)
''        sql = sql & "CRYINDL2='" & .CRYINDL2 & "', "          ' 結晶検査指示（L2)
''        sql = sql & "CRYINDL3='" & .CRYINDL3 & "', "          ' 結晶検査指示（L3)
''        sql = sql & "CRYINDL4='" & .CRYINDL4 & "', "          ' 結晶検査指示（L4)
''        sql = sql & "CRYINDCS='" & .CRYINDCS & "', "          ' 結晶検査指示（Cs)
''        sql = sql & "CRYINDGD='" & .CRYINDGD & "', "          ' 結晶検査指示（GD)
''        sql = sql & "CRYINDT='" & .CRYINDT & "', "            ' 結晶検査指示（T)
''        sql = sql & "CRYINDEP='" & .CRYINDEP & "', "          ' 結晶検査指示（EPD)
'        sql = sql & "CRYRESRS='" & .CRYRESRS & "', "          ' 結晶検査実績（Rs)
'        sql = sql & "CRYRESOI='" & .CRYRESOI & "', "          ' 結晶検査実績（Oi)
'        sql = sql & "CRYRESB1='" & .CRYRESB1 & "', "          ' 結晶検査実績（B1)
'        sql = sql & "CRYRESB2='" & .CRYRESB2 & "', "          ' 結晶検査実績（B2）
'        sql = sql & "CRYRESB3='" & .CRYRESB3 & "', "          ' 結晶検査実績（B3)
'        sql = sql & "CRYRESL1='" & .CRYRESL1 & "', "          ' 結晶検査実績（L1)
'        sql = sql & "CRYRESL2='" & .CRYRESL2 & "', "          ' 結晶検査実績（L2)
'        sql = sql & "CRYRESL3='" & .CRYRESL3 & "', "          ' 結晶検査実績（L3)
'        sql = sql & "CRYRESL4='" & .CRYRESL4 & "', "          ' 結晶検査実績（L4)
'        sql = sql & "CRYRESCS='" & .CRYRESCS & "', "          ' 結晶検査実績（Cs)
'        sql = sql & "CRYRESGD='" & .CRYRESGD & "', "          ' 結晶検査実績（GD)
'        sql = sql & "CRYREST='" & .CRYREST & "', "            ' 結晶検査実績（T)
'        sql = sql & "CRYRESEP='" & .CRYRESEP & "', "          ' 結晶検査実績（EPD)
''        sql = sql & "SMPLNUM=" & .SMPLNUM & ", "              ' サンプル枚数
''        sql = sql & "SMPLPAT='" & .SMPLPAT & "', "            ' サンプルパターン
'        sql = sql & "UPDDATE=sysdate, "                       ' 更新日付
'        sql = sql & "SENDFLAG='0'"                            ' 送信フラグ


        sql = "update XSDCS set "
        sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "          ' 結晶検査実績（Rs)
        sql = sql & "CRYRESRS2CS='" & .CRYRESRS2CS & "', "          ' 結晶検査実績（Rs)
        sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "          ' 結晶検査実績（Oi)
        sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "          ' 結晶検査実績（B1)
        sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "          ' 結晶検査実績（B2）
        sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "          ' 結晶検査実績（B3)
        sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "          ' 結晶検査実績（L1)
        sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "          ' 結晶検査実績（L2)
        sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "          ' 結晶検査実績（L3)
        sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "          ' 結晶検査実績（L4)
        sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "          ' 結晶検査実績（Cs)
        sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "          ' 結晶検査実績（GD)
        sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "            ' 結晶検査実績（T)
        sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "          ' 結晶検査実績（EPD)
        sql = sql & "KDAYCS=sysdate, "                       ' 更新日付
        sql = sql & "SNDKCS='0'"                            ' 送信フラグ

    End With

    If sqlWhere <> vbNullString Then
        sql = sql & " " & sqlWhere
    End If

    If OraDB.ExecuteSQL(sql) <= 0 Then
        Exit Function
    End If

    DBDRV_UpdateTBCME043 = FUNCTION_RETURN_SUCCESS

End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「XSDCS」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCS ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
'    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
'              " CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
'              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
'              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
'              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
'              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, SMPLNUMCS, SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
              " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME043 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
'            .CRYNUM = rs("CRYNUM")           ' 結晶番号
'            .IngotPos = rs("INGOTPOS")       ' 結晶内位置
'            .SMPKBN = rs("SMPKBN")           ' サンプル区分
'            .SMPLNO = rs("SMPLNO")           ' サンプルNo
'            .hinban = rs("HINBAN")           ' 品番
'            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
'            .factory = rs("FACTORY")         ' 工場
'            .opecond = rs("OPECOND")         ' 操業条件
'            .KTKBN = rs("KTKBN")             ' 確定区分
'            .CRYINDRS = rs("CRYINDRS")       ' 結晶検査指示（Rs)
'            .CRYINDOI = rs("CRYINDOI")       ' 結晶検査指示（Oi)
'            .CRYINDB1 = rs("CRYINDB1")       ' 結晶検査指示（B1)
'            .CRYINDB2 = rs("CRYINDB2")       ' 結晶検査指示（B2）
'            .CRYINDB3 = rs("CRYINDB3")       ' 結晶検査指示（B3)
'            .CRYINDL1 = rs("CRYINDL1")       ' 結晶検査指示（L1)
'            .CRYINDL2 = rs("CRYINDL2")       ' 結晶検査指示（L2)
'            .CRYINDL3 = rs("CRYINDL3")       ' 結晶検査指示（L3)
'            .CRYINDL4 = rs("CRYINDL4")       ' 結晶検査指示（L4)
'            .CRYINDCS = rs("CRYINDCS")       ' 結晶検査指示（Cs)
'            .CRYINDGD = rs("CRYINDGD")       ' 結晶検査指示（GD)
'            .CRYINDT = rs("CRYINDT")         ' 結晶検査指示（T)
'            .CRYINDEP = rs("CRYINDEP")       ' 結晶検査指示（EPD)
'            .CRYRESRS = rs("CRYRESRS")       ' 結晶検査実績（Rs)
'            .CRYRESOI = rs("CRYRESOI")       ' 結晶検査実績（Oi)
'            .CRYRESB1 = rs("CRYRESB1")       ' 結晶検査実績（B1)
'            .CRYRESB2 = rs("CRYRESB2")       ' 結晶検査実績（B2）
'            .CRYRESB3 = rs("CRYRESB3")       ' 結晶検査実績（B3)
'            .CRYRESL1 = rs("CRYRESL1")       ' 結晶検査実績（L1)
'            .CRYRESL2 = rs("CRYRESL2")       ' 結晶検査実績（L2)
'            .CRYRESL3 = rs("CRYRESL3")       ' 結晶検査実績（L3)
'            .CRYRESL4 = rs("CRYRESL4")       ' 結晶検査実績（L4)
'            .CRYRESCS = rs("CRYRESCS")       ' 結晶検査実績（Cs)
'            .CRYRESGD = rs("CRYRESGD")       ' 結晶検査実績（GD)
'            .CRYREST = rs("CRYREST")         ' 結晶検査実績（T)
'            .CRYRESEP = rs("CRYRESEP")       ' 結晶検査実績（EPD)
'            .SMPLNUM = rs("SMPLNUM")         ' サンプル枚数
'            .SMPLPAT = rs("SMPLPAT")         ' サンプルパターン
'            .REGDATE = rs("REGDATE")         ' 登録日付
'            .UPDDATE = rs("UPDDATE")         ' 更新日付
'            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'            .SENDDATE = rs("SENDDATE")       ' 送信日付

            If IsNull(rs("CRYNUMCS")) = False Then .CRYNUMCS = rs("CRYNUMCS")                   ' ブロックID
            If IsNull(rs("SMPKBNCS")) = False Then .SMPKBNCS = rs("SMPKBNCS")                   ' サンプル区分
            If IsNull(rs("TBKBNCS")) = False Then .TBKBNCS = rs("TBKBNCS")                      ' T/B区分
            If IsNull(rs("REPSMPLIDCS")) = False Then .REPSMPLIDCS = rs("REPSMPLIDCS")          ' 代表サンプルID
            If IsNull(rs("XTALCS")) = False Then .XTALCS = rs("XTALCS")                         ' 結晶番号
            If IsNull(rs("INPOSCS")) = False Then .INPOSCS = rs("INPOSCS")                      ' 結晶内位置
            If IsNull(rs("HINBCS")) = False Then .HINBCS = rs("HINBCS")                         ' 品番
            If IsNull(rs("REVNUMCS")) = False Then .REVNUMCS = rs("REVNUMCS")                   ' 製品番号改訂番号
            If IsNull(rs("FACTORYCS")) = False Then .FACTORYCS = rs("FACTORYCS")                ' 工場
            If IsNull(rs("OPECS")) = False Then .OPECS = rs("OPECS")                            ' 操業条件
            If IsNull(rs("KTKBNCS")) = False Then .KTKBNCS = rs("KTKBNCS")                      ' 確定区分
            If IsNull(rs("BLKKTFLAGCS")) = False Then .BLKKTFLAGCS = rs("BLKKTFLAGCS")          ' ブロック確定フラグ
            If IsNull(rs("CRYSMPLIDRSCS")) = False Then .CRYSMPLIDRSCS = rs("CRYSMPLIDRSCS")    ' サンプルID(Rs)
            If IsNull(rs("CRYSMPLIDRS1CS")) = False Then .CRYSMPLIDRS1CS = rs("CRYSMPLIDRS1CS") ' 推定サンプルID1(Rs)
            If IsNull(rs("CRYSMPLIDRS2CS")) = False Then .CRYSMPLIDRS2CS = rs("CRYSMPLIDRS2CS") ' 推定サンプルID2(Rs)
            If IsNull(rs("CRYINDRSCS")) = False Then .CRYINDRSCS = rs("CRYINDRSCS")             ' 状態FLG(Rs)
            If IsNull(rs("CRYRESRS1CS")) = False Then .CRYRESRS1CS = rs("CRYRESRS1CS")          ' 実績FLG1(Rs)
            If IsNull(rs("CRYRESRS2CS")) = False Then .CRYRESRS2CS = rs("CRYRESRS2CS")          ' 実績FLG2(Rs)
            If IsNull(rs("CRYSMPLIDOICS")) = False Then .CRYSMPLIDOICS = rs("CRYSMPLIDOICS")    ' サンプルID(Oi)
            If IsNull(rs("CRYINDOICS")) = False Then .CRYINDOICS = rs("CRYINDOICS")             ' 状態FLG(Oi)
            If IsNull(rs("CRYRESOICS")) = False Then .CRYRESOICS = rs("CRYRESOICS")             ' 実績FLG(Oi)
            If IsNull(rs("CRYSMPLIDB1CS")) = False Then .CRYSMPLIDB1CS = rs("CRYSMPLIDB1CS")    ' サンプルID(B1)
            If IsNull(rs("CRYINDB1CS")) = False Then .CRYINDB1CS = rs("CRYINDB1CS")             ' 状態FLG(B1)
            If IsNull(rs("CRYRESB1CS")) = False Then .CRYRESB1CS = rs("CRYRESB1CS")             ' 実績FLG(B1)
            If IsNull(rs("CRYSMPLIDB2CS")) = False Then .CRYSMPLIDB2CS = rs("CRYSMPLIDB2CS")    ' サンプルID(B2)
            If IsNull(rs("CRYINDB2CS")) = False Then .CRYINDB2CS = rs("CRYINDB2CS")             ' 状態FLG(B2)
            If IsNull(rs("CRYRESB2CS")) = False Then .CRYRESB2CS = rs("CRYRESB2CS")             ' 実績FLG(B2)
            If IsNull(rs("CRYSMPLIDB3CS")) = False Then .CRYSMPLIDB3CS = rs("CRYSMPLIDB3CS")    ' サンプルID(B3)
            If IsNull(rs("CRYINDB3CS")) = False Then .CRYINDB3CS = rs("CRYINDB3CS")             ' 状態FLG(B3)
            If IsNull(rs("CRYRESB3CS")) = False Then .CRYRESB3CS = rs("CRYRESB3CS")             ' 実績FLG(B3)
            If IsNull(rs("CRYSMPLIDL1CS")) = False Then .CRYSMPLIDL1CS = rs("CRYSMPLIDL1CS")    ' サンプルID(L1)
            If IsNull(rs("CRYINDL1CS")) = False Then .CRYINDL1CS = rs("CRYINDL1CS")             ' 状態FLG(L1)
            If IsNull(rs("CRYRESL1CS")) = False Then .CRYRESL1CS = rs("CRYRESL1CS")             ' 実績FLG(L1)
            If IsNull(rs("CRYSMPLIDL2CS")) = False Then .CRYSMPLIDL2CS = rs("CRYSMPLIDL2CS")    ' サンプルID(L2)
            If IsNull(rs("CRYINDL2CS")) = False Then .CRYINDL2CS = rs("CRYINDL2CS")             ' 状態FLG(L2)
            If IsNull(rs("CRYRESL2CS")) = False Then .CRYRESL2CS = rs("CRYRESL2CS")             ' 実績FLG(L2)
            If IsNull(rs("CRYSMPLIDL3CS")) = False Then .CRYSMPLIDL3CS = rs("CRYSMPLIDL3CS")    ' サンプルID(L3)
            If IsNull(rs("CRYINDL3CS")) = False Then .CRYINDL3CS = rs("CRYINDL3CS")             ' 状態FLG(L3)
            If IsNull(rs("CRYRESL3CS")) = False Then .CRYRESL3CS = rs("CRYRESL3CS")             ' 実績FLG(L3)
            If IsNull(rs("CRYSMPLIDL4CS")) = False Then .CRYSMPLIDL4CS = rs("CRYSMPLIDL4CS")    ' サンプルID(L4)
            If IsNull(rs("CRYINDL4CS")) = False Then .CRYINDL4CS = rs("CRYINDL4CS")             ' 状態FLG(L4)
            If IsNull(rs("CRYRESL4CS")) = False Then .CRYRESL4CS = rs("CRYRESL4CS")             ' 実績FLG(L4)
            If IsNull(rs("CRYSMPLIDCSCS")) = False Then .CRYSMPLIDCSCS = rs("CRYSMPLIDCSCS")    ' サンプルID(Cs)
            If IsNull(rs("CRYINDCSCS")) = False Then .CRYINDCSCS = rs("CRYINDCSCS")             ' 状態FLG(Cs)
            If IsNull(rs("CRYRESCSCS")) = False Then .CRYRESCSCS = rs("CRYRESCSCS")             ' 実績FLG(Cs)
            If IsNull(rs("CRYSMPLIDGDCS")) = False Then .CRYSMPLIDGDCS = rs("CRYSMPLIDGDCS")    ' サンプルID(GD)
            If IsNull(rs("CRYINDGDCS")) = False Then .CRYINDGDCS = rs("CRYINDGDCS")             ' 状態FLG(GD)
            If IsNull(rs("CRYRESGDCS")) = False Then .CRYRESGDCS = rs("CRYRESGDCS")             ' 実績FLG(GD)
            If IsNull(rs("CRYSMPLIDTCS")) = False Then .CRYSMPLIDTCS = rs("CRYSMPLIDTCS")       ' サンプルID(T)
            If IsNull(rs("CRYINDTCS")) = False Then .CRYINDTCS = rs("CRYINDTCS")                ' 状態FLG(T)
            If IsNull(rs("CRYRESTCS")) = False Then .CRYRESTCS = rs("CRYRESTCS")                ' 実績FLG(T)
            If IsNull(rs("CRYSMPLIDEPCS")) = False Then .CRYSMPLIDEPCS = rs("CRYSMPLIDEPCS")    ' サンプルID(EPD)
            If IsNull(rs("CRYINDEPCS")) = False Then .CRYINDEPCS = rs("CRYINDEPCS")             ' 状態FLG(EPD)
            If IsNull(rs("CRYRESEPCS")) = False Then .CRYRESEPCS = rs("CRYRESEPCS")             ' 実績FLG(EPD)
            If IsNull(rs("SMPLNUMCS")) = False Then .SMPLNUMCS = rs("SMPLNUMCS")                ' サンプル枚数
            If IsNull(rs("SMPLPATCS")) = False Then .SMPLPATCS = rs("SMPLPATCS")                ' サンプルパターン
            If IsNull(rs("TSTAFFCS")) = False Then .TSTAFFCS = rs("TSTAFFCS")                   ' 登録社員ID
            If IsNull(rs("TDAYCS")) = False Then .TDAYCS = rs("TDAYCS")                         ' 登録日付
            If IsNull(rs("KSTAFFCS")) = False Then .KSTAFFCS = rs("KSTAFFCS")                   ' 更新社員ID
            If IsNull(rs("KDAYCS")) = False Then .KDAYCS = rs("KDAYCS")                         ' 更新日付
            If IsNull(rs("SNDKCS")) = False Then .SNDKCS = rs("SNDKCS")                         ' 送信フラグ
            If IsNull(rs("SNDDAYCS")) = False Then .SNDDAYCS = rs("SNDDAYCS")                   ' 送信日付

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
End Function

'概要      :テーブル「XSDCS」の条件にあったレコードを更新する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records       ,I   ,typ_XSDCS   ,更新レコード
'          :[sqlWhere]    ,I   ,String         ,更新条件(SQLのWhere節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN  ,更新の成否
'説明      :
'履歴      :2001/07/13作成　伊藤
Public Function DBDRV_UpdateXSDCS(sqlUpdate As String) As FUNCTION_RETURN
    
    DBDRV_UpdateXSDCS = FUNCTION_RETURN_FAILURE

    If OraDB.ExecuteSQL(sqlUpdate) <= 0 Then
        Exit Function
    End If

    DBDRV_UpdateXSDCS = FUNCTION_RETURN_SUCCESS

End Function
