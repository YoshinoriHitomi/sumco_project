Attribute VB_Name = "s_kensa_SQL"
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

Public Function DBDRV_GetTBCME019(records() As typ_TBCME019, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
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
        
        
        Case "f_cmbc053_1i"           '「Ｘ線測定 実績入力」
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
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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

Public Function DBDRV_GetTBCME020(records() As typ_TBCME020, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
Dim sql         As String           'SQL全体
Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
Dim sqlWhere    As String           'SQLWhere部
'C－OSF3判定機能追加 2007/04/23 M.Kaga START ---
Dim sqlAnd      As String           'SQLAnd部
'C－OSF3判定機能追加 2007/04/23 M.Kaga End   ---
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
'C－OSF3判定機能追加 2007/04/23 M.Kaga START ---
             sqlBase = "Select T.HINBAN, T.MNOREVNO, T.FACTORY, T.OPECOND, T.HMGSTRRNO, T.HMGSTFNO, T.HMGSXSNO, T.HMGSXSNE, T.HSXDENKU, T.HSXDENMX, T.HSXDENMN," & _
              " T.HSXDENHT, T.HSXDENHS, T.HSXDVDKU, T.HSXDVDMXN, T.HSXDVDMNN, T.HSXDVDHT, T.HSXDVDHS, T.HSXLDLKU, T.HSXLDLMX, T.HSXLDLMN, T.HSXLDLHT," & _
              " T.HSXLDLHS, T.HSXGDSZY, T.HSXGDSPH, T.HSXGDSPT, T.HSXGDSPR, T.HSXGDZAR, T.HSXGDKHM, T.HSXGDKHI, T.HSXGDKHH, T.HSXGDKHS, T.HSXDSOKE," & _
              " T.HSXDSOMX, T.HSXDSOMN, T.HSXDSOAX, T.HSXDSOAN, T.HSXDSOHT, T.HSXDSOHS, T.HSXDSOKM, T.HSXDSOKI, T.HSXDSOKH, T.HSXDSOKS, T.HSXLIFTW," & _
              " T.HSXSDSLP, T.HSXGKKNO, T.HSXCDOP, T.HSXCDOPN, T.HSXCDPNI, T.HSXGSFIN, T.HSXCLMIN, T.HSXCLMAX, T.HSXCLPMN, T.HSXCLPR, T.HSXWFWAR," & _
              " T.HSXOF1AX, T.HSXOF1MX, T.HSXOF1SH, T.HSXOF1ST, T.HSXOF1SR, T.HSXOF1HT, T.HSXOF1HS, T.HSXOF1SZ, T.HSXOF1KM, T.HSXOF1KI, T.HSXOF1KH," & _
              " T.HSXOF1KS, T.HSXOF1NS, T.HSXOF1ET, T.HSXOF2AX, T.HSXOF2MX, T.HSXOF2SH, T.HSXOF2ST, T.HSXOF2SR, T.HSXOF2HT, T.HSXOF2HS, T.HSXOF2SZ," & _
              " T.HSXOF2KM, T.HSXOF2KI, T.HSXOF2KH, T.HSXOF2KS, T.HSXOF2NS, T.HSXOF2ET, T.HSXOF3AX, T.HSXOF3MX, T.HSXOF3SH, T.HSXOF3ST, T.HSXOF3SR,"
            sqlBase = sqlBase & " T.HSXOF3HT, T.HSXOF3HS, T.HSXOF3SZ, T.HSXOF3KM, T.HSXOF3KI, T.HSXOF3KH, T.HSXOF3KS, T.HSXOF3NS, T.HSXOF3ET, T.HSXOF4AX, T.HSXOF4MX," & _
              " T.HSXOF4SH, T.HSXOF4ST, T.HSXOF4SR, T.HSXOF4HT, T.HSXOF4HS, T.HSXOF4SZ, T.HSXOF4KM, T.HSXOF4KI, T.HSXOF4KH, T.HSXOF4KS, T.HSXOF4NS," & _
              " T.HSXOF4ET, T.HSXBM1AN, T.HSXBM1AX, T.HSXBM1SH, T.HSXBM1ST, T.HSXBM1SR, T.HSXBM1HT, T.HSXBM1HS, T.HSXBM1SZ, T.HSXBM1KM, T.HSXBM1KI," & _
              " T.HSXBM1KH, T.HSXBM1KS, T.HSXBM1NS, T.HSXBM1ET, T.HSXBM2AN, T.HSXBM2AX, T.HSXBM2SH, T.HSXBM2ST, T.HSXBM2SR, T.HSXBM2HT, T.HSXBM2HS," & _
              " T.HSXBM2SZ, T.HSXBM2KM, T.HSXBM2KI, T.HSXBM2KH, T.HSXBM2KS, T.HSXBM2NS, T.HSXBM2ET, T.HSXBM3AN, T.HSXBM3AX, T.HSXBM3SH, T.HSXBM3ST," & _
              " T.HSXBM3SR, T.HSXBM3HT, T.HSXBM3HS, T.HSXBM3SZ, T.HSXBM3KM, T.HSXBM3KI, T.HSXBM3KH, T.HSXBM3KS, T.HSXBM3NS, T.HSXBM3ET, T.HSXNOTE," & _
              " T.HSXOSF1PTK, T.HSXOSF2PTK, T.HSXOSF3PTK, T.HSXOSF4PTK,"
' OSF，BMD項目追加対応  ↑　1行分　2002.04.02 yakimura
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
            For i = 1 To 10
                sqlBase = sqlBase & "T.HSXRS" & i & "N, "
                sqlBase = sqlBase & "T.HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "T.SPECRRNO, T.SXLMCNO, T.WFMCNO, T.STAFFID, T.REGDATE, T.UPDDATE, T.SENDFLAG, T.SENDDATE, U.COSF3FLAG, T.HSXCOSF3NS "
        
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
            sqlBase = sqlBase & ", HSXGDPTK "   '' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech
            
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
    
    
        Case "f_cmbc053_1"           '「X線測定 実績入力」
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

        'Add Start 2010/12/17 SMPK Miyata
        Case "f_cmbc054_1"           '「Cu-deco実績入力」
             sqlBase = "Select T.HINBAN, T.MNOREVNO, T.FACTORY, T.OPECOND, T.HMGSTRRNO, T.HMGSTFNO, T.HMGSXSNO, T.HMGSXSNE, T.HSXDENKU, T.HSXDENMX, T.HSXDENMN," & _
              " T.HSXDENHT, T.HSXDENHS, T.HSXDVDKU, T.HSXDVDMXN, T.HSXDVDMNN, T.HSXDVDHT, T.HSXDVDHS, T.HSXLDLKU, T.HSXLDLMX, T.HSXLDLMN, T.HSXLDLHT," & _
              " T.HSXLDLHS, T.HSXGDSZY, T.HSXGDSPH, T.HSXGDSPT, T.HSXGDSPR, T.HSXGDZAR, T.HSXGDKHM, T.HSXGDKHI, T.HSXGDKHH, T.HSXGDKHS, T.HSXDSOKE," & _
              " T.HSXDSOMX, T.HSXDSOMN, T.HSXDSOAX, T.HSXDSOAN, T.HSXDSOHT, T.HSXDSOHS, T.HSXDSOKM, T.HSXDSOKI, T.HSXDSOKH, T.HSXDSOKS, T.HSXLIFTW," & _
              " T.HSXSDSLP, T.HSXGKKNO, T.HSXCDOP, T.HSXCDOPN, T.HSXCDPNI, T.HSXGSFIN, T.HSXCLMIN, T.HSXCLMAX, T.HSXCLPMN, T.HSXCLPR, T.HSXWFWAR," & _
              " T.HSXOF1AX, T.HSXOF1MX, T.HSXOF1SH, T.HSXOF1ST, T.HSXOF1SR, T.HSXOF1HT, T.HSXOF1HS, T.HSXOF1SZ, T.HSXOF1KM, T.HSXOF1KI, T.HSXOF1KH," & _
              " T.HSXOF1KS, T.HSXOF1NS, T.HSXOF1ET, T.HSXOF2AX, T.HSXOF2MX, T.HSXOF2SH, T.HSXOF2ST, T.HSXOF2SR, T.HSXOF2HT, T.HSXOF2HS, T.HSXOF2SZ," & _
              " T.HSXOF2KM, T.HSXOF2KI, T.HSXOF2KH, T.HSXOF2KS, T.HSXOF2NS, T.HSXOF2ET, T.HSXOF3AX, T.HSXOF3MX, T.HSXOF3SH, T.HSXOF3ST, T.HSXOF3SR,"
            sqlBase = sqlBase & " T.HSXOF3HT, T.HSXOF3HS, T.HSXOF3SZ, T.HSXOF3KM, T.HSXOF3KI, T.HSXOF3KH, T.HSXOF3KS, T.HSXOF3NS, T.HSXOF3ET, T.HSXOF4AX, T.HSXOF4MX," & _
              " T.HSXOF4SH, T.HSXOF4ST, T.HSXOF4SR, T.HSXOF4HT, T.HSXOF4HS, T.HSXOF4SZ, T.HSXOF4KM, T.HSXOF4KI, T.HSXOF4KH, T.HSXOF4KS, T.HSXOF4NS," & _
              " T.HSXOF4ET, T.HSXBM1AN, T.HSXBM1AX, T.HSXBM1SH, T.HSXBM1ST, T.HSXBM1SR, T.HSXBM1HT, T.HSXBM1HS, T.HSXBM1SZ, T.HSXBM1KM, T.HSXBM1KI," & _
              " T.HSXBM1KH, T.HSXBM1KS, T.HSXBM1NS, T.HSXBM1ET, T.HSXBM2AN, T.HSXBM2AX, T.HSXBM2SH, T.HSXBM2ST, T.HSXBM2SR, T.HSXBM2HT, T.HSXBM2HS," & _
              " T.HSXBM2SZ, T.HSXBM2KM, T.HSXBM2KI, T.HSXBM2KH, T.HSXBM2KS, T.HSXBM2NS, T.HSXBM2ET, T.HSXBM3AN, T.HSXBM3AX, T.HSXBM3SH, T.HSXBM3ST," & _
              " T.HSXBM3SR, T.HSXBM3HT, T.HSXBM3HS, T.HSXBM3SZ, T.HSXBM3KM, T.HSXBM3KI, T.HSXBM3KH, T.HSXBM3KS, T.HSXBM3NS, T.HSXBM3ET, T.HSXNOTE," & _
              " T.HSXOSF1PTK, T.HSXOSF2PTK, T.HSXOSF3PTK, T.HSXOSF4PTK,"
            sqlBase = sqlBase & "T.HSXCPK, T.HSXCSZ, T.HSXCHT, T.HSXCHS, T.HSXCJPK, T.HSXCJNS, T.HSXCJHT, T.HSXCJHS, " & _
              " T.HSXCJLTPK, T.HSXCJLTNS, T.HSXCJLTHT, T.HSXCJLTHS, T.HSXCJ2PK, T.HSXCJ2NS, T.HSXCJ2HT, T.HSXCJ2HS, "
            For i = 1 To 10
                sqlBase = sqlBase & "T.HSXRS" & i & "N, "
                sqlBase = sqlBase & "T.HSXRS" & i & "Y, "
            Next
            sqlBase = sqlBase & "T.SPECRRNO, T.SXLMCNO, T.WFMCNO, T.STAFFID, T.REGDATE, T.UPDDATE, T.SENDFLAG, T.SENDDATE, U.COSF3FLAG "
        'Add End   2010/12/17 SMPK Miyata

    End Select
       
       
'Chg Start 2010/12/17 SMPK Miyata
'    If Trim(formID) = "f_cmbc025_1" Then
    If Trim(formID) = "f_cmbc025_1" Or Trim(formID) = "f_cmbc054_1" Then
'Chg End   2010/12/17 SMPK Miyata
        sqlBase = sqlBase & "From TBCME020 T , TBCME036 U"
    Else
        sqlBase = sqlBase & "From TBCME020"
    End If
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
                key = key & ", "
            End If
        End With
    Next
    
'Chg Start 2010/12/17 SMPK Miyata
''C－OSF3判定機能追加 2007/04/23 M.Kaga START ---
'    If Trim(formID) = "f_cmbc025_1" Then
    If Trim(formID) = "f_cmbc025_1" Or Trim(formID) = "f_cmbc054_1" Then
'Chg End   2010/12/17 SMPK Miyata
        sqlWhere = " Where(T.HINBAN||TO_CHAR(T.MNOREVNO, 'FM00000')||T.FACTORY||T.OPECOND in(" & key & "))"
        sqlAnd = " And(U.HINBAN||TO_CHAR(U.MNOREVNO, 'FM00000')||U.FACTORY||U.OPECOND in(" & key & "))"
        sql = sqlBase & sqlWhere & sqlAnd
    Else
        sqlWhere = " Where(HINBAN||TO_CHAR(MNOREVNO, 'FM00000')||FACTORY||OPECOND in(" & key & "))"
        sql = sqlBase & sqlWhere
    End If
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
    
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")           ' 品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")       ' 製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")         ' 工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")         ' 操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")     ' 品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")       ' 品管理社員Ｎｏ
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")       ' 品管理ＳＸ製品番号
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))       ' 品管理ＳＸ製品番号枝番
            If fldNameExist("HSXDENKU") Then .HSXDENKU = rs("HSXDENKU")       ' 品ＳＸＤｅｎ検査有無
            If fldNameExist("HSXDENMX") Then .HSXDENMX = fncNullCheck(rs("HSXDENMX"))       ' 品ＳＸＤｅｎ上限
            If fldNameExist("HSXDENMN") Then .HSXDENMN = fncNullCheck(rs("HSXDENMN"))       ' 品ＳＸＤｅｎ下限
            If fldNameExist("HSXDENHT") Then .HSXDENHT = rs("HSXDENHT")       ' 品ＳＸＤｅｎ保証方法＿対
            If fldNameExist("HSXDENHS") Then .HSXDENHS = rs("HSXDENHS")       ' 品ＳＸＤｅｎ保証方法＿処
            If fldNameExist("HSXDVDKU") Then .HSXDVDKU = rs("HSXDVDKU")       ' 品ＳＸＤＶＤ２検査有無
            If fldNameExist("HSXDVDMXN") Then .HSXDVDMX = fncNullCheck(rs("HSXDVDMXN"))       ' 品ＳＸＤＶＤ２上限    ＷＦサンプル処理変更 2003.05.20 yakimura
            If fldNameExist("HSXDVDMNN") Then .HSXDVDMN = fncNullCheck(rs("HSXDVDMNN"))       ' 品ＳＸＤＶＤ２下限    ＷＦサンプル処理変更 2003.05.20 yakimura
            If fldNameExist("HSXDVDHT") Then .HSXDVDHT = rs("HSXDVDHT")       ' 品ＳＸＤＶＤ２保証方法＿対
            If fldNameExist("HSXDVDHS") Then .HSXDVDHS = rs("HSXDVDHS")       ' 品ＳＸＤＶＤ２保証方法＿処
            If fldNameExist("HSXLDLKU") Then .HSXLDLKU = rs("HSXLDLKU")       ' 品ＳＸＬ／ＤＬ検査有無
            If fldNameExist("HSXLDLMX") Then .HSXLDLMX = fncNullCheck(rs("HSXLDLMX"))       ' 品ＳＸＬ／ＤＬ上限
            If fldNameExist("HSXLDLMN") Then .HSXLDLMN = fncNullCheck(rs("HSXLDLMN"))       ' 品ＳＸＬ／ＤＬ下限
            If fldNameExist("HSXLDLHT") Then .HSXLDLHT = rs("HSXLDLHT")       ' 品ＳＸＬ／ＤＬ保証方法＿対
            If fldNameExist("HSXLDLHS") Then .HSXLDLHS = rs("HSXLDLHS")       ' 品ＳＸＬ／ＤＬ保証方法＿処
            If fldNameExist("HSXGDSZY") Then .HSXGDSZY = rs("HSXGDSZY")       ' 品ＳＸＧＤ測定条件
            If fldNameExist("HSXGDSPH") Then .HSXGDSPH = rs("HSXGDSPH")       ' 品ＳＸＧＤ測定位置＿方
            If fldNameExist("HSXGDSPT") Then .HSXGDSPT = rs("HSXGDSPT")       ' 品ＳＸＧＤ測定位置＿点
            If fldNameExist("HSXGDSPR") Then .HSXGDSPR = rs("HSXGDSPR")       ' 品ＳＸＧＤ測定位置＿領
            If fldNameExist("HSXGDZAR") Then .HSXGDZAR = fncNullCheck(rs("HSXGDZAR"))       ' 品ＳＸＧＤ除外領域
            If fldNameExist("HSXGDKHM") Then .HSXGDKHM = rs("HSXGDKHM")       ' 品ＳＸＧＤ検査頻度＿枚
            If fldNameExist("HSXGDKHI") Then .HSXGDKHI = rs("HSXGDKHI")       ' 品ＳＸＧＤ検査頻度＿位
            If fldNameExist("HSXGDKHH") Then .HSXGDKHH = rs("HSXGDKHH")       ' 品ＳＸＧＤ検査頻度＿保
            If fldNameExist("HSXGDKHS") Then .HSXGDKHS = rs("HSXGDKHS")       ' 品ＳＸＧＤ検査頻度＿試
            If fldNameExist("HSXDSOKE") Then .HSXDSOKE = rs("HSXDSOKE")       ' 品ＳＸＤＳＯＤ検査
            If fldNameExist("HSXDSOMX") Then .HSXDSOMX = fncNullCheck(rs("HSXDSOMX"))       ' 品ＳＸＤＳＯＤ上限
            If fldNameExist("HSXDSOMN") Then .HSXDSOMN = fncNullCheck(rs("HSXDSOMN"))       ' 品ＳＸＤＳＯＤ下限
            If fldNameExist("HSXDSOAX") Then .HSXDSOAX = fncNullCheck(rs("HSXDSOAX"))       ' 品ＳＸＤＳＯＤ領域上限
            If fldNameExist("HSXDSOAN") Then .HSXDSOAN = fncNullCheck(rs("HSXDSOAN"))       ' 品ＳＸＤＳＯＤ領域下限
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
            If fldNameExist("HSXCLMIN") Then .HSXCLMIN = fncNullCheck(rs("HSXCLMIN"))       ' 品ＳＸ結晶長下限
            If fldNameExist("HSXCLMAX") Then .HSXCLMAX = fncNullCheck(rs("HSXCLMAX"))       ' 品ＳＸ結晶長上限
            If fldNameExist("HSXCLPMN") Then .HSXCLPMN = fncNullCheck(rs("HSXCLPMN"))       ' 品ＳＸ結晶長許容下限
            If fldNameExist("HSXCLPR") Then .HSXCLPR = fncNullCheck(rs("HSXCLPR"))         ' 品ＳＸ結晶長許容比率
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
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOSF" & j & "PTK") Then                       ' 品ＳＸＯＳＦ(n)パタン区分
                   If IsNull(rs("HSXOSF" & j & "PTK")) = False Then .HSXOSF_PTK(j) = rs("HSXOSF" & j & "PTK")
                   End If
' OSF，BMD項目追加対応  2002.04.02 yakimura
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
'NULL対応 2003/12/21
' OSF，BMD項目追加対応  2002.04.02 yakimura
'                If fldNameExist("HSXBMD" & j & "MBP") Then                      ' 品ＳＸＢＭＤ(n)面内分布
'                   If IsNull(rs("HSXBMD" & j & "MBP")) = False Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP"))
'                   End If
                If fldNameExist("HSXBMD" & j & "MBP") Then .HSXBMD_MBP(j) = fncNullCheck(rs("HSXBMD" & j & "MBP")) ' 品ＳＸＢＭＤ(n)面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
'NULL対応 2003/12/21
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
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOSF1PTK") Then                           ' 品ＳＸＯＳＦ1パタン区分
                   If IsNull(rs("HSXOSF1PTK")) = False Then .HSXOSF1PTK = rs("HSXOSF1PTK")
                   End If
' OSF，BMD項目追加対応  2002.04.02 yakimura
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
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOSF2PTK") Then                           ' 品ＳＸＯＳＦ2パタン区分
                   If IsNull(rs("HSXOSF2PTK")) = False Then .HSXOSF2PTK = rs("HSXOSF2PTK")
                   End If
' OSF，BMD項目追加対応  2002.04.02 yakimura
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
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOSF3PTK") Then                           ' 品ＳＸＯＳＦ3パタン区分
                   If IsNull(rs("HSXOSF3PTK")) = False Then .HSXOSF3PTK = rs("HSXOSF3PTK")
                   End If
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOF4AX") Then .HSXOF4AX = fncNullCheck(rs("HSXOF4AX"))  ' 品ＳＸＯＳＦ4平均上限
                If fldNameExist("HSXOF4MX") Then .HSXOF4MX = fncNullCheck(rs("HSXOF4MX"))  ' 品ＳＸＯＳＦ4上限
                If fldNameExist("HSXOF4SH") Then .HSXOF4SH = rs("HSXOF4SH")  ' 品ＳＸＯＳＦ4測定位置＿方
                If fldNameExist("HSXOF4ST") Then .HSXOF4ST = rs("HSXOF4ST")  ' 品ＳＸＯＳＦ4測定位置＿点
                If fldNameExist("HSXOF4SR") Then .HSXOF4SR = rs("HSXOF4SR")  ' 品ＳＸＯＳＦ4測定位置＿領
                If fldNameExist("HSXOF4HT") Then .HSXOF4HT = rs("HSXOF4HT")  ' 品ＳＸＯＳＦ4保証方法＿対
'C－OSF3判定機能追加 2007/04/23 M.Kaga START ---
                If fldNameExist("COSF3FLAG") Then
                    If IsNull(rs("COSF3FLAG")) = False Then .HSXOF4HS = rs("COSF3FLAG") Else .HSXOF4HS = " "
                End If
'C－OSF3判定機能追加 2007/04/23 M.Kaga END   ---
                If fldNameExist("HSXOF4SZ") Then .HSXOF4SZ = rs("HSXOF4SZ")  ' 品ＳＸＯＳＦ4測定条件
                If fldNameExist("HSXOF4KM") Then .HSXOF4KM = rs("HSXOF4KM")  ' 品ＳＸＯＳＦ4検査頻度＿枚
                If fldNameExist("HSXOF4KI") Then .HSXOF4KI = rs("HSXOF4KI")  ' 品ＳＸＯＳＦ4検査頻度＿位
                If fldNameExist("HSXOF4KH") Then .HSXOF4KH = rs("HSXOF4KH")  ' 品ＳＸＯＳＦ4検査頻度＿保
                If fldNameExist("HSXOF4KS") Then .HSXOF4KS = rs("HSXOF4KS")  ' 品ＳＸＯＳＦ4検査頻度＿試
                If fldNameExist("HSXOF4NS") Then .HSXOF4NS = rs("HSXOF4NS")  ' 品ＳＸＯＳＦ4熱処理法
                If fldNameExist("HSXOF4ET") Then .HSXOF4ET = fncNullCheck(rs("HSXOF4ET"))  ' 品ＳＸＯＳＦ4選択ＥＴ代
' OSF，BMD項目追加対応  2002.04.02 yakimura
                If fldNameExist("HSXOSF4PTK") Then                           ' 品ＳＸＯＳＦ4パタン区分
                   If IsNull(rs("HSXOSF4PTK")) = False Then .HSXOSF4PTK = rs("HSXOSF4PTK")
                   End If
' OSF，BMD項目追加対応  2002.04.02 yakimura

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
'NULL対応 2003/12/21
' OSF，BMD項目追加対応  2002.04.02 yakimura
'                If fldNameExist("HSXBMD1MBP") Then                           ' 品ＳＸＢＭＤ1面内分布
'                   If IsNull(rs("HSXBMD1MBP")) = False Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP"))
'                   End If
                If fldNameExist("HSXBMD1MBP") Then .HSXBMD1MBP = fncNullCheck(rs("HSXBMD1MBP")) ' 品ＳＸＢＭＤ1面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
'NULL対応 2003/12/21
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
'NULL対応 2003/12/21
' OSF，BMD項目追加対応  2002.04.02 yakimura
'                If fldNameExist("HSXBMD2MBP") Then                           ' 品ＳＸＢＭＤ2面内分布
'                   If IsNull(rs("HSXBMD2MBP")) = False Then .HSXBMD2MBP = rs("HSXBMD2MBP")
'                   End If
                If fldNameExist("HSXBMD2MBP") Then .HSXBMD2MBP = rs("HSXBMD2MBP") ' 品ＳＸＢＭＤ2面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
'NULL対応 2003/12/21
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
'NULL対応 2003/12/21
' OSF，BMD項目追加対応  2002.04.02 yakimura
'                If fldNameExist("HSXBMD3MBP") Then                           ' 品ＳＸＢＭＤ3面内分布
'                   If IsNull(rs("HSXBMD3MBP")) = False Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP"))
'                   End If
                If fldNameExist("HSXBMD3MBP") Then .HSXBMD3MBP = fncNullCheck(rs("HSXBMD3MBP")) ' 品ＳＸＢＭＤ3面内分布
' OSF，BMD項目追加対応  2002.04.02 yakimura
'NULL対応 2003/12/21
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
                If fldNameExist("HSXRS10N") Then .HSXRS10N = rs("HSXRS10N")     ' 品ＳＸ予備10＿内

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
            
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
                If fldNameExist("HSXGDPTK") Then         ' 品ＳＸＧＤパタン区分
                If IsNull(rs("HSXGDPTK")) = False Then .HSXGDPTK = rs("HSXGDPTK") Else .HSXGDPTK = " "
            End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

            'Add Start 2011/01/26 SMPK Miyata
            If fldNameExist("HSXCPK") Then
                If IsNull(rs("HSXCPK")) = False Then .HSXCPK = rs("HSXCPK")         '品ＳＸＣパターン区分
            End If
            If fldNameExist("HSXCSZ") Then
                If IsNull(rs("HSXCSZ")) = False Then .HSXCSZ = rs("HSXCSZ")         '品ＳＸＣ測定条件
            End If
            If fldNameExist("HSXCHT") Then
                If IsNull(rs("HSXCHT")) = False Then .HSXCHT = rs("HSXCHT")         '品ＳＸＣ保証方法＿対
            End If
            If fldNameExist("HSXCHS") Then
                If IsNull(rs("HSXCHS")) = False Then .HSXCHS = rs("HSXCHS")         '品ＳＸＣ保証方法＿処
            End If
            If fldNameExist("HSXCJPK") Then
                If IsNull(rs("HSXCJPK")) = False Then .HSXCJPK = rs("HSXCJPK")       '品ＳＸＣＪパターン区分
            End If
            If fldNameExist("HSXCJNS") Then
                If IsNull(rs("HSXCJNS")) = False Then .HSXCJNS = rs("HSXCJNS")       '品ＳＸＣＪ熱処理法
            End If
            If fldNameExist("HSXCJHT") Then
                If IsNull(rs("HSXCJHT")) = False Then .HSXCJHT = rs("HSXCJHT")       '品ＳＸＣＪ保証方法＿対
            End If
            If fldNameExist("HSXCJHS") Then
                If IsNull(rs("HSXCJHS")) = False Then .HSXCJHS = rs("HSXCJHS")       '品ＳＸＣＪ保証方法＿処
            End If
            If fldNameExist("HSXCJLTPK") Then
                If IsNull(rs("HSXCJLTPK")) = False Then .HSXCJLTPK = rs("HSXCJLTPK")   '品ＳＸＣＪＬＴパターン区分
            End If
            If fldNameExist("HSXCJLTNS") Then
                If IsNull(rs("HSXCJLTNS")) = False Then .HSXCJLTNS = rs("HSXCJLTNS")   '品ＳＸＣＪＬＴ熱処理法
            End If
            If fldNameExist("HSXCJLTHT") Then
                If IsNull(rs("HSXCJLTHT")) = False Then .HSXCJLTHT = rs("HSXCJLTHT")   '品ＳＸＣＪＬＴ保証方法＿対
            End If
            If fldNameExist("HSXCJLTHS") Then
                If IsNull(rs("HSXCJLTHS")) = False Then .HSXCJLTHS = rs("HSXCJLTHS")   '品ＳＸＣＪＬＴ保証方法＿処
            End If
            If fldNameExist("HSXCJ2PK") Then
                If IsNull(rs("HSXCJ2PK")) = False Then .HSXCJ2PK = rs("HSXCJ2PK")     '品ＳＸＣＪ２パターン区分
            End If
            If fldNameExist("HSXCJ2NS") Then
                If IsNull(rs("HSXCJ2NS")) = False Then .HSXCJ2NS = rs("HSXCJ2NS")     '品ＳＸＣＪ２熱処理法
            End If
            If fldNameExist("HSXCJ2HT") Then
                If IsNull(rs("HSXCJ2HT")) = False Then .HSXCJ2HT = rs("HSXCJ2HT")     '品ＳＸＣＪ２保証方法＿対
            End If
            If fldNameExist("HSXCJ2HS") Then
                If IsNull(rs("HSXCJ2HS")) = False Then .HSXCJ2HS = rs("HSXCJ2HS")     '品ＳＸＣＪ２保証方法＿処
            End If
            'Add End   2011/01/26 SMPK Miyata
            
            'Add Start 2011/02/17 Y.Hitomi
            If fldNameExist("HSXCOSF3NS") Then
                If IsNull(rs("HSXCOSF3NS")) = False Then .HSXCOSF3NS = rs("HSXCOSF3NS")     '品ＳＸＣＪ２保証方法＿処
            End If
            'Add End   2011/02/17 Y.Hitomi
        
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

'概要      :テーブル「TBCME021」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records()     ,O  ,typ_TBCME021    ,抽出レコード
'          :formID        ,I  ,String          ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban     ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :05/03/01 ooba
Public Function DBDRV_GetTBCME021(records() As typ_TBCME021, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME021"

    Select Case formID
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, CONFLAG, " & _
              "REINFLAG, HWFTRWKB, HWFFACES, HWFBACKS, HWFBDSWY, HWFTYPE, HWFTYPKW, HWFDOP, HWFFKBWK, " & _
              "HWFFKBWS, HWFRMIN, HWFRMAX, HWFRSPOH, HWFRSPOT, HWFRSPOI, HWFRHWYT, HWFRHWYS, HWFRKWAY, " & _
              "HWFRKHNM, HWFRKHNN, HWFRKHNH, HWFRKHNU, HWFRSDEV, HWFRAMIN, HWFRAMAX, HWFRMBNP, HWFRMCAL, " & _
              "HWFRMBP2, HWFRMCL2, HWFRKBSH, HWFRKBST, HWFRKBSI, HWFRKBHT, HWFRKBHS, HWFSTMAX, HWFSTSPH, " & _
              "HWFSTSPT, HWFSTSPI, HWFSTHWT, HWFSTHWS, HWFSTKWY, HWFSTKHM, HWFSTKHN, HWFSTKHH, HWFSTKHU, "
            sqlBase = sqlBase & "HWFACEN, HWFAMIN, HWFAMAX, HWFASPOH, HWFASPOT, HWFASPOI, HWFAHWYT, HWFAHWYS, HWFAKWAY, " & _
              "HWFAKHNM, HWFAKHNN, HWFAKHNH, HWFAKHNU, HWFASDEV, HWFAAMIN, HWFAAMAX, HWFAMBNP, HWFAMCAL, " & _
              "HWFALTBP, HWFALTCL, HWFALTRA, HWFAMRAN, HWFDIVS, HWFAKBSH, HWFAKBST, HWFAKBSI, HWFAKBHT, " & _
              "HWFAKBHS, HWFWFORM, HWFD1CEN, HWFD1MIN, HWFD1MAX, HWFD2CEN, HWFD2MIN, HWFD2MAX, HWFDKHNM, " & _
              "HWFDKHNN, HWFDKHNH, HWFDKHNU, HWFLPMNP, HWFSGMNP, HWFETMNP, HWFMPMNP, HWFLPKS1, HWFLPKS2, " & _
              "HWFLPKZ1, HWFLPKZ2, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        '追加 2005/06/15 ffc)tanabe start
        Case "f_cmec067_1"           '「SPV実績参照」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFTYPE, HWFD1CEN "
        '追加 2005/06/15 ffc)tanabe end
        
    End Select
       
    sqlBase = sqlBase & "From TBCME021"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME021 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '品管理社員Ｎｏ
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '品管理ＷＦ製品番号
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '品管理ＷＦ製品番号枝番
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG") '確認フラグ
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG") '再付与フラグ
            If fldNameExist("HWFTRWKB") Then .HWFTRWKB = rs("HWFTRWKB") '品ＷＦ統合可否区分
            If fldNameExist("HWFFACES") Then .HWFFACES = rs("HWFFACES") '品ＷＦ表面仕上げ
            If fldNameExist("HWFBACKS") Then .HWFBACKS = rs("HWFBACKS") '品ＷＦ裏仕上げ
            If fldNameExist("HWFBDSWY") Then .HWFBDSWY = rs("HWFBDSWY") '品ＷＦＢＤ処理方法
            If fldNameExist("HWFTYPE") Then .HWFTYPE = rs("HWFTYPE") '品ＷＦタイプ
            If fldNameExist("HWFTYPKW") Then .HWFTYPKW = rs("HWFTYPKW") '品ＷＦタイプ検査方法
            If fldNameExist("HWFDOP") Then .HWFDOP = rs("HWFDOP") '品ＷＦドーパント
            If fldNameExist("HWFFKBWK") Then .HWFFKBWK = rs("HWFFKBWK") '品ＷＦ表面区分方法＿区
            If fldNameExist("HWFFKBWS") Then .HWFFKBWS = rs("HWFFKBWS") '品ＷＦ表面区分方法＿指
            If fldNameExist("HWFRMIN") Then .HWFRMIN = fncNullCheck(rs("HWFRMIN")) '品ＷＦ比抵抗下限
            If fldNameExist("HWFRMAX") Then .HWFRMAX = fncNullCheck(rs("HWFRMAX")) '品ＷＦ比抵抗上限
            If fldNameExist("HWFRSPOH") Then .HWFRSPOH = rs("HWFRSPOH") '品ＷＦ比抵抗測定位置＿方
            If fldNameExist("HWFRSPOT") Then .HWFRSPOT = rs("HWFRSPOT") '品ＷＦ比抵抗測定位置＿点
            If fldNameExist("HWFRSPOI") Then .HWFRSPOI = rs("HWFRSPOI") '品ＷＦ比抵抗測定位置＿位
            If fldNameExist("HWFRHWYT") Then .HWFRHWYT = rs("HWFRHWYT") '品ＷＦ比抵抗保証方法＿対
            If fldNameExist("HWFRHWYS") Then .HWFRHWYS = rs("HWFRHWYS") '品ＷＦ比抵抗保証方法＿処
            If fldNameExist("HWFRKWAY") Then .HWFRKWAY = rs("HWFRKWAY") '品ＷＦ比抵抗検査方法
            If fldNameExist("HWFRKHNM") Then .HWFRKHNM = rs("HWFRKHNM") '品ＷＦ比抵抗検査頻度＿枚
            If fldNameExist("HWFRKHNN") Then .HWFRKHNN = rs("HWFRKHNN") '品ＷＦ比抵抗検査頻度＿抜
            If fldNameExist("HWFRKHNH") Then .HWFRKHNH = rs("HWFRKHNH") '品ＷＦ比抵抗検査頻度＿保
            If fldNameExist("HWFRKHNU") Then .HWFRKHNU = rs("HWFRKHNU") '品ＷＦ比抵抗検査頻度＿ウ
            If fldNameExist("HWFRSDEV") Then .HWFRSDEV = fncNullCheck(rs("HWFRSDEV")) '品ＷＦ比抵抗標準偏差
            If fldNameExist("HWFRAMIN") Then .HWFRAMIN = fncNullCheck(rs("HWFRAMIN")) '品ＷＦ比抵抗平均下限
            If fldNameExist("HWFRAMAX") Then .HWFRAMAX = fncNullCheck(rs("HWFRAMAX")) '品ＷＦ比抵抗平均上限
            If fldNameExist("HWFRMBNP") Then .HWFRMBNP = fncNullCheck(rs("HWFRMBNP")) '品ＷＦ比抵抗面内分布
            If fldNameExist("HWFRMCAL") Then .HWFRMCAL = rs("HWFRMCAL") '品ＷＦ比抵抗面内計算
            If fldNameExist("HWFRMBP2") Then .HWFRMBP2 = fncNullCheck(rs("HWFRMBP2")) '品ＷＦ比抵抗面内分布２
            If fldNameExist("HWFRMCL2") Then .HWFRMCL2 = rs("HWFRMCL2") '品ＷＦ比抵抗面内計算２
            If fldNameExist("HWFRKBSH") Then .HWFRKBSH = rs("HWFRKBSH") '品ＷＦ比抵抗振区分測定位置＿方
            If fldNameExist("HWFRKBST") Then .HWFRKBST = rs("HWFRKBST") '品ＷＦ比抵抗振区分測定位置＿点
            If fldNameExist("HWFRKBSI") Then .HWFRKBSI = rs("HWFRKBSI") '品ＷＦ比抵抗振区分測定位置＿位
            If fldNameExist("HWFRKBHT") Then .HWFRKBHT = rs("HWFRKBHT") '品ＷＦ比抵抗振区分保証方法＿対
            If fldNameExist("HWFRKBHS") Then .HWFRKBHS = rs("HWFRKBHS") '品ＷＦ比抵抗振区分保証方法＿処
            If fldNameExist("HWFSTMAX") Then .HWFSTMAX = fncNullCheck(rs("HWFSTMAX")) '品ＷＦストリエ上限
            If fldNameExist("HWFSTSPH") Then .HWFSTSPH = rs("HWFSTSPH") '品ＷＦストリエ測定位置＿方
            If fldNameExist("HWFSTSPT") Then .HWFSTSPT = rs("HWFSTSPT") '品ＷＦストリエ測定位置＿点
            If fldNameExist("HWFSTSPI") Then .HWFSTSPI = rs("HWFSTSPI") '品ＷＦストリエ測定位置＿位
            If fldNameExist("HWFSTHWT") Then .HWFSTHWT = rs("HWFSTHWT") '品ＷＦストリエ保証方法＿対
            If fldNameExist("HWFSTHWS") Then .HWFSTHWS = rs("HWFSTHWS") '品ＷＦストリエ保証方法＿処
            If fldNameExist("HWFSTKWY") Then .HWFSTKWY = rs("HWFSTKWY") '品ＷＦストリエ検査方法
            If fldNameExist("HWFSTKHM") Then .HWFSTKHM = rs("HWFSTKHM") '品ＷＦストリエ検査頻度＿枚
            If fldNameExist("HWFSTKHN") Then .HWFSTKHN = rs("HWFSTKHN") '品ＷＦストリエ検査頻度＿抜
            If fldNameExist("HWFSTKHH") Then .HWFSTKHH = rs("HWFSTKHH") '品ＷＦストリエ検査頻度＿保
            If fldNameExist("HWFSTKHU") Then .HWFSTKHU = rs("HWFSTKHU") '品ＷＦストリエ検査頻度＿ウ
            If fldNameExist("HWFACEN") Then .HWFACEN = fncNullCheck(rs("HWFACEN")) '品ＷＦ厚中心
            If fldNameExist("HWFAMIN") Then .HWFAMIN = fncNullCheck(rs("HWFAMIN")) '品ＷＦ厚下限
            If fldNameExist("HWFAMAX") Then .HWFAMAX = fncNullCheck(rs("HWFAMAX")) '品ＷＦ厚上限
            If fldNameExist("HWFASPOH") Then .HWFASPOH = rs("HWFASPOH") '品ＷＦ厚測定位置＿方
            If fldNameExist("HWFASPOT") Then .HWFASPOT = rs("HWFASPOT") '品ＷＦ厚測定位置＿点
            If fldNameExist("HWFASPOI") Then .HWFASPOI = rs("HWFASPOI") '品ＷＦ厚測定位置＿位
            If fldNameExist("HWFAHWYT") Then .HWFAHWYT = rs("HWFAHWYT") '品ＷＦ厚保証方法＿対
            If fldNameExist("HWFAHWYS") Then .HWFAHWYS = rs("HWFAHWYS") '品ＷＦ厚保証方法＿処
            If fldNameExist("HWFAKWAY") Then .HWFAKWAY = rs("HWFAKWAY") '品ＷＦ厚検査方法
            If fldNameExist("HWFAKHNM") Then .HWFAKHNM = rs("HWFAKHNM") '品ＷＦ厚検査頻度＿枚
            If fldNameExist("HWFAKHNN") Then .HWFAKHNN = rs("HWFAKHNN") '品ＷＦ厚検査頻度＿抜
            If fldNameExist("HWFAKHNH") Then .HWFAKHNH = rs("HWFAKHNH") '品ＷＦ厚検査頻度＿保
            If fldNameExist("HWFAKHNU") Then .HWFAKHNU = rs("HWFAKHNU") '品ＷＦ厚検査頻度＿ウ
            If fldNameExist("HWFASDEV") Then .HWFASDEV = fncNullCheck(rs("HWFASDEV")) '品ＷＦ厚標準偏差
            If fldNameExist("HWFAAMIN") Then .HWFAAMIN = fncNullCheck(rs("HWFAAMIN")) '品ＷＦ厚平均下限
            If fldNameExist("HWFAAMAX") Then .HWFAAMAX = fncNullCheck(rs("HWFAAMAX")) '品ＷＦ厚平均上限
            If fldNameExist("HWFAMBNP") Then .HWFAMBNP = fncNullCheck(rs("HWFAMBNP")) '品ＷＦ厚面内分布
            If fldNameExist("HWFAMCAL") Then .HWFAMCAL = rs("HWFAMCAL") '品ＷＦ厚面内計算
            If fldNameExist("HWFALTBP") Then .HWFALTBP = fncNullCheck(rs("HWFALTBP")) '品ＷＦ厚ＬＴ分布
            If fldNameExist("HWFALTCL") Then .HWFALTCL = rs("HWFALTCL") '品ＷＦ厚ＬＴ計算
            If fldNameExist("HWFALTRA") Then .HWFALTRA = fncNullCheck(rs("HWFALTRA")) '品ＷＦ厚ＬＴ範囲
            If fldNameExist("HWFAMRAN") Then .HWFAMRAN = fncNullCheck(rs("HWFAMRAN")) '品ＷＦ厚面内範囲
            If fldNameExist("HWFDIVS") Then .HWFDIVS = fncNullCheck(rs("HWFDIVS")) '品ＷＦ分割数
            If fldNameExist("HWFAKBSH") Then .HWFAKBSH = rs("HWFAKBSH") '品ＷＦ厚振区分測定位置＿方
            If fldNameExist("HWFAKBST") Then .HWFAKBST = rs("HWFAKBST") '品ＷＦ厚振区分測定位置＿点
            If fldNameExist("HWFAKBSI") Then .HWFAKBSI = rs("HWFAKBSI") '品ＷＦ厚振区分測定位置＿位
            If fldNameExist("HWFAKBHT") Then .HWFAKBHT = rs("HWFAKBHT") '品ＷＦ厚振区分保証方法＿対
            If fldNameExist("HWFAKBHS") Then .HWFAKBHS = rs("HWFAKBHS") '品ＷＦ厚振区分保証方法＿処
            If fldNameExist("HWFWFORM") Then .HWFWFORM = rs("HWFWFORM") '品ＷＦウェーハ形状
            If fldNameExist("HWFD1CEN") Then .HWFD1CEN = fncNullCheck(rs("HWFD1CEN")) '品ＷＦ直径１中心
            If fldNameExist("HWFD1MIN") Then .HWFD1MIN = fncNullCheck(rs("HWFD1MIN")) '品ＷＦ直径１下限
            If fldNameExist("HWFD1MAX") Then .HWFD1MAX = fncNullCheck(rs("HWFD1MAX")) '品ＷＦ直径１上限
            If fldNameExist("HWFD2CEN") Then .HWFD2CEN = fncNullCheck(rs("HWFD2CEN")) '品ＷＦ直径２中心
            If fldNameExist("HWFD2MIN") Then .HWFD2MIN = fncNullCheck(rs("HWFD2MIN")) '品ＷＦ直径２下限
            If fldNameExist("HWFD2MAX") Then .HWFD2MAX = fncNullCheck(rs("HWFD2MAX")) '品ＷＦ直径２上限
            If fldNameExist("HWFDKHNM") Then .HWFDKHNM = rs("HWFDKHNM") '品ＷＦ直径検査頻度＿枚
            If fldNameExist("HWFDKHNN") Then .HWFDKHNN = rs("HWFDKHNN") '品ＷＦ直径検査頻度＿抜
            If fldNameExist("HWFDKHNH") Then .HWFDKHNH = rs("HWFDKHNH") '品ＷＦ直径検査頻度＿保
            If fldNameExist("HWFDKHNU") Then .HWFDKHNU = rs("HWFDKHNU") '品ＷＦ直径検査頻度＿ウ
            If fldNameExist("HWFLPMNP") Then .HWFLPMNP = fncNullCheck(rs("HWFLPMNP")) '品ＷＦＬＰ厚最小加工代
            If fldNameExist("HWFSGMNP") Then .HWFSGMNP = fncNullCheck(rs("HWFSGMNP")) '品ＷＦＳＧ厚最小加工代
            If fldNameExist("HWFETMNP") Then .HWFETMNP = fncNullCheck(rs("HWFETMNP")) '品ＷＦＥＴ厚最小加工代
            If fldNameExist("HWFMPMNP") Then .HWFMPMNP = fncNullCheck(rs("HWFMPMNP")) '品ＷＦＭＰ厚最小加工代
            If fldNameExist("HWFLPKS1") Then .HWFLPKS1 = rs("HWFLPKS1") '品ＷＦＬＰ研磨材種１
            If fldNameExist("HWFLPKS2") Then .HWFLPKS2 = rs("HWFLPKS2") '品ＷＦＬＰ研磨材種２
            If fldNameExist("HWFLPKZ1") Then .HWFLPKZ1 = rs("HWFLPKZ1") '品ＷＦＬＰ研磨材粒度種１
            If fldNameExist("HWFLPKZ2") Then .HWFLPKZ2 = rs("HWFLPKZ2") '品ＷＦＬＰ研磨材粒度種２
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") 'Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") 'ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") 'ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME021 = FUNCTION_RETURN_SUCCESS

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

'概要      :テーブル「TBCME022」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records()     ,O  ,typ_TBCME022    ,抽出レコード
'          :formID        ,I  ,String          ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban     ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :05/03/01 ooba
Public Function DBDRV_GetTBCME022(records() As typ_TBCME022, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME022"

    Select Case formID
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, HWFCDIR, " & _
              "HWFCSCEN, HWFCSMIN, HWFCSMAX, HWFCSDIS, HWFCSDIR, HWFCKWAY, HWFCKHNM, HWFCKHNN, HWFCKHNH, " & _
              "HWFCKHNU, HWFCTDIR, HWFCTCEN, HWFCTMIN, HWFCTMAX, HWFCYDIR, HWFCYCEN, HWFCYMIN, HWFCYMAX, " & _
              "HWFKPTNN, HWFOFPKM, HWFOFPKN, HWFOFPKH, HWFOFPKU, HWFOFLKM, HWFOFLKN, HWFOFLKH, HWFOFLKU, " & _
              "HWFOF1PD, HWFOF1PN, HWFOF1PX, HWFOF1PW, HWFOF1LC, HWFOF1LN, HWFOF1LX, HWFOF1RF, HWFOFRRC, " & _
              "HWFOFRRN, HWFOFRRX, HWFOFRLC, HWFOFRLN, HWFOFRLX, HWFOF1DC, HWFOF1DN, HWFOF1DX, HWFZFORM, "
            sqlBase = sqlBase & "HWFD3CEN, HWFD3MIN, HWFD3MAX, HWFDFKJ, HWFDFKHM, HWFDFKHN, HWFDFKHH, " & _
              "HWFDFKHU, HWFDPDRC, HWFDPACN, HWFDPAMN, HWFDPAMX, HWFDPDIR, HWFDPMIN, HWFDPMAX, HWFDPKWY, " & _
              "HWFDPKHM, HWFDPKHB, HWFDPKHH, HWFDPKHU, HWFDACEN, HWFDAMIN, HWFDAMAX, HWFDWCEN, HWFDWMIN, " & _
              "HWFDWMAX, HWFDDCEN, HWFDDMIN, HWFDDMAX, HWFDBRCN, HWFDBRMN, HWFDBRMX, HWFDRRCN, HWFDRRMN, " & _
              "HWFDRRMX, HWFDLRCN, HWFDLRMN, HWFDLRMX, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, " & _
              "STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        
        '追加 2005/06/15 ffc)tanabe start
        Case "f_cmec067_1"           '「SPV実績参照」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFCDIR "
        '追加 2005/06/15 ffc)tanabe end

    End Select
       
    sqlBase = sqlBase & "From TBCME022"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME022 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '品管理社員Ｎｏ
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '品管理ＷＦ製品番号
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '品管理ＷＦ製品番号枝番
            If fldNameExist("HWFCDIR") Then .HWFCDIR = rs("HWFCDIR") '品ＷＦ結晶面方位
            If fldNameExist("HWFCSCEN") Then .HWFCSCEN = fncNullCheck(rs("HWFCSCEN")) '品ＷＦ結晶面傾中心
            If fldNameExist("HWFCSMIN") Then .HWFCSMIN = fncNullCheck(rs("HWFCSMIN")) '品ＷＦ結晶面傾下限
            If fldNameExist("HWFCSMAX") Then .HWFCSMAX = fncNullCheck(rs("HWFCSMAX")) '品ＷＦ結晶面傾上限
            If fldNameExist("HWFCSDIS") Then .HWFCSDIS = rs("HWFCSDIS") '品ＷＦ結晶面傾方位指定
            If fldNameExist("HWFCSDIR") Then .HWFCSDIR = rs("HWFCSDIR") '品ＷＦ結晶面傾方位
            If fldNameExist("HWFCKWAY") Then .HWFCKWAY = rs("HWFCKWAY") '品ＷＦ結晶面検査方法
            If fldNameExist("HWFCKHNM") Then .HWFCKHNM = rs("HWFCKHNM") '品ＷＦ結晶面検査頻度＿枚
            If fldNameExist("HWFCKHNN") Then .HWFCKHNN = rs("HWFCKHNN") '品ＷＦ結晶面検査頻度＿抜
            If fldNameExist("HWFCKHNH") Then .HWFCKHNH = rs("HWFCKHNH") '品ＷＦ結晶面検査頻度＿保
            If fldNameExist("HWFCKHNU") Then .HWFCKHNU = rs("HWFCKHNU") '品ＷＦ結晶面検査頻度＿ウ
            If fldNameExist("HWFCTDIR") Then .HWFCTDIR = rs("HWFCTDIR") '品ＷＦ結晶面傾縦方位
            If fldNameExist("HWFCTCEN") Then .HWFCTCEN = fncNullCheck(rs("HWFCTCEN")) '品ＷＦ結晶面傾縦中心
            If fldNameExist("HWFCTMIN") Then .HWFCTMIN = fncNullCheck(rs("HWFCTMIN")) '品ＷＦ結晶面傾縦下限
            If fldNameExist("HWFCTMAX") Then .HWFCTMAX = fncNullCheck(rs("HWFCTMAX")) '品ＷＦ結晶面傾縦上限
            If fldNameExist("HWFCYDIR") Then .HWFCYDIR = rs("HWFCYDIR") '品ＷＦ結晶面傾横方位
            If fldNameExist("HWFCYCEN") Then .HWFCYCEN = fncNullCheck(rs("HWFCYCEN")) '品ＷＦ結晶面傾横中心
            If fldNameExist("HWFCYMIN") Then .HWFCYMIN = fncNullCheck(rs("HWFCYMIN")) '品ＷＦ結晶面傾横下限
            If fldNameExist("HWFCYMAX") Then .HWFCYMAX = fncNullCheck(rs("HWFCYMAX")) '品ＷＦ結晶面傾横上限
            If fldNameExist("HWFKPTNN") Then .HWFKPTNN = rs("HWFKPTNN") '品ＷＦ光像パタン名
            If fldNameExist("HWFOFPKM") Then .HWFOFPKM = rs("HWFOFPKM") '品ＷＦＯＦ位置検査頻度＿枚
            If fldNameExist("HWFOFPKN") Then .HWFOFPKN = rs("HWFOFPKN") '品ＷＦＯＦ位置検査頻度＿抜
            If fldNameExist("HWFOFPKH") Then .HWFOFPKH = rs("HWFOFPKH") '品ＷＦＯＦ位置検査頻度＿保
            If fldNameExist("HWFOFPKU") Then .HWFOFPKU = rs("HWFOFPKU") '品ＷＦＯＦ位置検査頻度＿ウ
            If fldNameExist("HWFOFLKM") Then .HWFOFLKM = rs("HWFOFLKM") '品ＷＦＯＦ長検査頻度＿枚
            If fldNameExist("HWFOFLKN") Then .HWFOFLKN = rs("HWFOFLKN") '品ＷＦＯＦ長検査頻度＿抜
            If fldNameExist("HWFOFLKH") Then .HWFOFLKH = rs("HWFOFLKH") '品ＷＦＯＦ長検査頻度＿保
            If fldNameExist("HWFOFLKU") Then .HWFOFLKU = rs("HWFOFLKU") '品ＷＦＯＦ長検査頻度＿ウ
            If fldNameExist("HWFOF1PD") Then .HWFOF1PD = rs("HWFOF1PD") '品ＷＦＯＦ１位置方位
            If fldNameExist("HWFOF1PN") Then .HWFOF1PN = fncNullCheck(rs("HWFOF1PN")) '品ＷＦＯＦ１位置下限
            If fldNameExist("HWFOF1PX") Then .HWFOF1PX = fncNullCheck(rs("HWFOF1PX")) '品ＷＦＯＦ１位置上限
            If fldNameExist("HWFOF1PW") Then .HWFOF1PW = rs("HWFOF1PW") '品ＷＦＯＦ１位置検査方法
            If fldNameExist("HWFOF1LC") Then .HWFOF1LC = fncNullCheck(rs("HWFOF1LC")) '品ＷＦＯＦ１長中心
            If fldNameExist("HWFOF1LN") Then .HWFOF1LN = fncNullCheck(rs("HWFOF1LN")) '品ＷＦＯＦ１長下限
            If fldNameExist("HWFOF1LX") Then .HWFOF1LX = fncNullCheck(rs("HWFOF1LX")) '品ＷＦＯＦ１長上限
            If fldNameExist("HWFOF1RF") Then .HWFOF1RF = rs("HWFOF1RF") '品ＷＦＯＦ１両端Ｒ形状
            If fldNameExist("HWFOFRRC") Then .HWFOFRRC = fncNullCheck(rs("HWFOFRRC")) '品ＷＦＯＦ両端Ｒ右中心
            If fldNameExist("HWFOFRRN") Then .HWFOFRRN = fncNullCheck(rs("HWFOFRRN")) '品ＷＦＯＦ両端Ｒ右下限
            If fldNameExist("HWFOFRRX") Then .HWFOFRRX = fncNullCheck(rs("HWFOFRRX")) '品ＷＦＯＦ両端Ｒ右上限
            If fldNameExist("HWFOFRLC") Then .HWFOFRLC = fncNullCheck(rs("HWFOFRLC")) '品ＷＦＯＦ両端Ｒ左中心
            If fldNameExist("HWFOFRLN") Then .HWFOFRLN = fncNullCheck(rs("HWFOFRLN")) '品ＷＦＯＦ両端Ｒ左下限
            If fldNameExist("HWFOFRLX") Then .HWFOFRLX = fncNullCheck(rs("HWFOFRLX")) '品ＷＦＯＦ両端Ｒ左上限
            If fldNameExist("HWFOF1DC") Then .HWFOF1DC = fncNullCheck(rs("HWFOF1DC")) '品ＷＦＯＦ１直径中心
            If fldNameExist("HWFOF1DN") Then .HWFOF1DN = fncNullCheck(rs("HWFOF1DN")) '品ＷＦＯＦ１直径下限
            If fldNameExist("HWFOF1DX") Then .HWFOF1DX = fncNullCheck(rs("HWFOF1DX")) '品ＷＦＯＦ１直径上限
            If fldNameExist("HWFZFORM") Then .HWFZFORM = rs("HWFZFORM") '品ＷＦ材料形状
            If fldNameExist("HWFD3CEN") Then .HWFD3CEN = fncNullCheck(rs("HWFD3CEN")) '品ＷＦ直径３中心
            If fldNameExist("HWFD3MIN") Then .HWFD3MIN = fncNullCheck(rs("HWFD3MIN")) '品ＷＦ直径３下限
            If fldNameExist("HWFD3MAX") Then .HWFD3MAX = fncNullCheck(rs("HWFD3MAX")) '品ＷＦ直径３上限
            If fldNameExist("HWFDFKJ") Then .HWFDFKJ = rs("HWFDFKJ") '品ＷＦ溝形状
            If fldNameExist("HWFDFKHM") Then .HWFDFKHM = rs("HWFDFKHM") '品ＷＦ溝形状検査頻度＿枚
            If fldNameExist("HWFDFKHN") Then .HWFDFKHN = rs("HWFDFKHN") '品ＷＦ溝形状検査頻度＿抜
            If fldNameExist("HWFDFKHH") Then .HWFDFKHH = rs("HWFDFKHH") '品ＷＦ溝形状検査頻度＿保
            If fldNameExist("HWFDFKHU") Then .HWFDFKHU = rs("HWFDFKHU") '品ＷＦ溝形状検査頻度＿ウ
            If fldNameExist("HWFDPDRC") Then .HWFDPDRC = rs("HWFDPDRC") '品ＷＦ溝位置方向
            If fldNameExist("HWFDPACN") Then .HWFDPACN = fncNullCheck(rs("HWFDPACN")) '品ＷＦ溝位置角度中心
            If fldNameExist("HWFDPAMN") Then .HWFDPAMN = fncNullCheck(rs("HWFDPAMN")) '品ＷＦ溝位置角度下限
            If fldNameExist("HWFDPAMX") Then .HWFDPAMX = fncNullCheck(rs("HWFDPAMX")) '品ＷＦ溝位置角度上限
            If fldNameExist("HWFDPDIR") Then .HWFDPDIR = rs("HWFDPDIR") '品ＷＦ溝位置方位
            If fldNameExist("HWFDPMIN") Then .HWFDPMIN = fncNullCheck(rs("HWFDPMIN")) '品ＷＦ溝位置下限
            If fldNameExist("HWFDPMAX") Then .HWFDPMAX = fncNullCheck(rs("HWFDPMAX")) '品ＷＦ溝位置上限
            If fldNameExist("HWFDPKWY") Then .HWFDPKWY = rs("HWFDPKWY") '品ＷＦ溝位置検査方法
            If fldNameExist("HWFDPKHM") Then .HWFDPKHM = rs("HWFDPKHM") '品ＷＦ溝位置検査頻度＿枚
            If fldNameExist("HWFDPKHB") Then .HWFDPKHB = rs("HWFDPKHB") '品ＷＦ溝位置検査頻度＿抜
            If fldNameExist("HWFDPKHH") Then .HWFDPKHH = rs("HWFDPKHH") '品ＷＦ溝位置検査頻度＿保
            If fldNameExist("HWFDPKHU") Then .HWFDPKHU = rs("HWFDPKHU") '品ＷＦ溝位置検査頻度＿ウ
            If fldNameExist("HWFDACEN") Then .HWFDACEN = fncNullCheck(rs("HWFDACEN")) '品ＷＦ溝角度中心
            If fldNameExist("HWFDAMIN") Then .HWFDAMIN = fncNullCheck(rs("HWFDAMIN")) '品ＷＦ溝角度下限
            If fldNameExist("HWFDAMAX") Then .HWFDAMAX = fncNullCheck(rs("HWFDAMAX")) '品ＷＦ溝角度上限
            If fldNameExist("HWFDWCEN") Then .HWFDWCEN = fncNullCheck(rs("HWFDWCEN")) '品ＷＦ溝巾中心
            If fldNameExist("HWFDWMIN") Then .HWFDWMIN = fncNullCheck(rs("HWFDWMIN")) '品ＷＦ溝巾下限
            If fldNameExist("HWFDWMAX") Then .HWFDWMAX = fncNullCheck(rs("HWFDWMAX")) '品ＷＦ溝巾上限
            If fldNameExist("HWFDDCEN") Then .HWFDDCEN = fncNullCheck(rs("HWFDDCEN")) '品ＷＦ溝深中心
            If fldNameExist("HWFDDMIN") Then .HWFDDMIN = fncNullCheck(rs("HWFDDMIN")) '品ＷＦ溝深下限
            If fldNameExist("HWFDDMAX") Then .HWFDDMAX = fncNullCheck(rs("HWFDDMAX")) '品ＷＦ溝深上限
            If fldNameExist("HWFDBRCN") Then .HWFDBRCN = fncNullCheck(rs("HWFDBRCN")) '品ＷＦ溝底Ｒ中心
            If fldNameExist("HWFDBRMN") Then .HWFDBRMN = fncNullCheck(rs("HWFDBRMN")) '品ＷＦ溝底Ｒ下限
            If fldNameExist("HWFDBRMX") Then .HWFDBRMX = fncNullCheck(rs("HWFDBRMX")) '品ＷＦ溝底Ｒ上限
            If fldNameExist("HWFDRRCN") Then .HWFDRRCN = fncNullCheck(rs("HWFDRRCN")) '品ＷＦ溝右Ｒ中心
            If fldNameExist("HWFDRRMN") Then .HWFDRRMN = fncNullCheck(rs("HWFDRRMN")) '品ＷＦ溝右Ｒ下限
            If fldNameExist("HWFDRRMX") Then .HWFDRRMX = fncNullCheck(rs("HWFDRRMX")) '品ＷＦ溝右Ｒ上限
            If fldNameExist("HWFDLRCN") Then .HWFDLRCN = fncNullCheck(rs("HWFDLRCN")) '品ＷＦ溝左Ｒ中心
            If fldNameExist("HWFDLRMN") Then .HWFDLRMN = fncNullCheck(rs("HWFDLRMN")) '品ＷＦ溝左Ｒ下限
            If fldNameExist("HWFDLRMX") Then .HWFDLRMX = fncNullCheck(rs("HWFDLRMX")) '品ＷＦ溝左Ｒ上限
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") 'Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") 'ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") 'ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME022 = FUNCTION_RETURN_SUCCESS

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

'概要      :テーブル「TBCME026」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records()     ,O  ,typ_TBCME026    ,抽出レコード
'          :formID        ,I  ,String          ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban     ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :05/03/01 ooba
Public Function DBDRV_GetTBCME026(records() As typ_TBCME026, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

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
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME026"

    Select Case formID
        Case "f_cmbc026_1"           '「GD実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGWFSNO, HMGWFSNE, HWFBDOMN, HWFBDOMX, " & _
              "HWFBDOSH, HWFBDOST, HWFBDOSR, HWFBDOHT, HWFBDOHS, HWFBDOSZ, HWFBDONS, HWFBDOKM, HWFBDOKN, HWFBDOKH, " & _
              "HWFBDOKU, HWFBDOET, HWFBDSMN, HWFBDSMX, HWFBDSSH, HWFBDSST, HWFBDSSR, HWFBDSHT, HWFBDSHS, HWFBDSSZ, " & _
              "HWFBDSNS, HWFBDSKM, HWFBDSKN, HWFBDSKH, HWFBDSKU, HWFBDSET, HWFRNFMX, HWFRNFSH, HWFRNFST, HWFRNFSI, " & _
              "HWFRNFKW, HWFRNFZA, HWFRNBMX, HWFRNBSH, HWFRNBST, HWFRNBSI, HWFRNBKW, HWFRNBZA, HWFDENKU, HWFDENMX, " & _
              "HWFDENMN, HWFDENHT, HWFDENHS, HWFDVDKU, HWFDVDMX, HWFDVDMN, HWFDVDHT, HWFDVDHS, HWFLDLKU, HWFLDLMX, " & _
              "HWFLDLMN, HWFLDLHT, HWFLDLHS, HWFGDSPH, HWFGDSPT, HWFGDSPR, HWFGDSZY, HWFGDZAR, HWFGDKHM, HWFGDKHN, "
            sqlBase = sqlBase & "HWFGDKHH, HWFGDKHU, HWFDSOKE, HWFDSOMX, HWFDSOMN, HWFDSOAX, HWFDSOAN, HWFDSOHT, HWFDSOHS, HWFDSOKM, " & _
              "HWFDSOKN, HWFDSOKH, HWFDSOKU, HWFNTPUM, HWFNTPK1, HWFNTPP1, HWFNTPS1, HWFNTPK2, HWFNTPP2, HWFNTPS2, " & _
              "HWFNTPK3, HWFNTPP3, HWFNTPS3, HWFNTPZA, HWFNTPHT, HWFNTPHS, HWFNTPKM, HWFNTPKN, HWFNTPKH, HWFNTPKU, " & _
              "HWFCRSSK, HWFMDCEN, HWFMDMAX, HWFMDMIN, HWFMDSPH, HWFMDSPT, HWFMDSPI, HWFMDHWT, HWFMDHWS, HWFMDKHM, " & _
              "HWFMDKHN, HWFMDKHH, HWFMDKHU, IFKBN, SYORIKBN, SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, " & _
              "SENDFLAG, SENDDATE, HWFDVDMXN, HWFDVDMNN, HWFDSONWY, HWFMSUMX, HWFMSUZY, HWFMSUKW, HWFMSUSZ, " & _
              "HWFNP1AR, HWFNP1MAX, HWFNP2AR, HWFNP2MAX "
    End Select
       
    sqlBase = sqlBase & "From TBCME026"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME026 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN") '品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO") '製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY") '工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND") '操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO") '品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO") '品管理社員Ｎｏ
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO") '品管理ＷＦ製品番号
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE")) '品管理ＷＦ製品番号枝番
            If fldNameExist("HWFBDOMN") Then .HWFBDOMN = fncNullCheck(rs("HWFBDOMN")) '品ＷＦＢＤＯＳＦ下限
            If fldNameExist("HWFBDOMX") Then .HWFBDOMX = fncNullCheck(rs("HWFBDOMX")) '品ＷＦＢＤＯＳＦ上限
            If fldNameExist("HWFBDOSH") Then .HWFBDOSH = rs("HWFBDOSH") '品ＷＦＢＤＯＳＦ測定位置＿方
            If fldNameExist("HWFBDOST") Then .HWFBDOST = rs("HWFBDOST") '品ＷＦＢＤＯＳＦ測定位置＿点
            If fldNameExist("HWFBDOSR") Then .HWFBDOSR = rs("HWFBDOSR") '品ＷＦＢＤＯＳＦ測定位置＿領
            If fldNameExist("HWFBDOHT") Then .HWFBDOHT = rs("HWFBDOHT") '品ＷＦＢＤＯＳＦ保証方法＿対
            If fldNameExist("HWFBDOHS") Then .HWFBDOHS = rs("HWFBDOHS") '品ＷＦＢＤＯＳＦ保証方法＿処
            If fldNameExist("HWFBDOSZ") Then .HWFBDOSZ = rs("HWFBDOSZ") '品ＷＦＢＤＯＳＦ測定条件
            If fldNameExist("HWFBDONS") Then .HWFBDONS = rs("HWFBDONS") '品ＷＦＢＤＯＳＦ熱処理法
            If fldNameExist("HWFBDOKM") Then .HWFBDOKM = rs("HWFBDOKM") '品ＷＦＢＤＯＳＦ検査頻度＿枚
            If fldNameExist("HWFBDOKN") Then .HWFBDOKN = rs("HWFBDOKN") '品ＷＦＢＤＯＳＦ検査頻度＿抜
            If fldNameExist("HWFBDOKH") Then .HWFBDOKH = rs("HWFBDOKH") '品ＷＦＢＤＯＳＦ検査頻度＿保
            If fldNameExist("HWFBDOKU") Then .HWFBDOKU = rs("HWFBDOKU") '品ＷＦＢＤＯＳＦ検査頻度＿ウ
            If fldNameExist("HWFBDOET") Then .HWFBDOET = fncNullCheck(rs("HWFBDOET")) '品ＷＦＢＤＯＳＦ選択ＥＴ代
            If fldNameExist("HWFBDSMN") Then .HWFBDSMN = fncNullCheck(rs("HWFBDSMN")) '品ＷＦＢＤＳＴ跡下限
            If fldNameExist("HWFBDSMX") Then .HWFBDSMX = fncNullCheck(rs("HWFBDSMX")) '品ＷＦＢＤＳＴ跡上限
            If fldNameExist("HWFBDSSH") Then .HWFBDSSH = rs("HWFBDSSH") '品ＷＦＢＤＳＴ跡測定位置＿方
            If fldNameExist("HWFBDSST") Then .HWFBDSST = rs("HWFBDSST") '品ＷＦＢＤＳＴ跡測定位置＿点
            If fldNameExist("HWFBDSSR") Then .HWFBDSSR = rs("HWFBDSSR") '品ＷＦＢＤＳＴ跡測定位置＿領
            If fldNameExist("HWFBDSHT") Then .HWFBDSHT = rs("HWFBDSHT") '品ＷＦＢＤＳＴ跡保証方法＿対
            If fldNameExist("HWFBDSHS") Then .HWFBDSHS = rs("HWFBDSHS") '品ＷＦＢＤＳＴ跡保証方法＿処
            If fldNameExist("HWFBDSSZ") Then .HWFBDSSZ = rs("HWFBDSSZ") '品ＷＦＢＤＳＴ跡測定条件
            If fldNameExist("HWFBDSNS") Then .HWFBDSNS = rs("HWFBDSNS") '品ＷＦＢＤＳＴ跡熱処理法
            If fldNameExist("HWFBDSKM") Then .HWFBDSKM = rs("HWFBDSKM") '品ＷＦＢＤＳＴ跡検査頻度＿枚
            If fldNameExist("HWFBDSKN") Then .HWFBDSKN = rs("HWFBDSKN") '品ＷＦＢＤＳＴ跡検査頻度＿抜
            If fldNameExist("HWFBDSKH") Then .HWFBDSKH = rs("HWFBDSKH") '品ＷＦＢＤＳＴ跡検査頻度＿保
            If fldNameExist("HWFBDSKU") Then .HWFBDSKU = rs("HWFBDSKU") '品ＷＦＢＤＳＴ跡検査頻度＿ウ
            If fldNameExist("HWFBDSET") Then .HWFBDSET = fncNullCheck(rs("HWFBDSET")) '品ＷＦＢＤＳＴ跡選択ＥＴ代
            If fldNameExist("HWFRNFMX") Then .HWFRNFMX = fncNullCheck(rs("HWFRNFMX")) '品ＷＦラフネス表上限
            If fldNameExist("HWFRNFSH") Then .HWFRNFSH = rs("HWFRNFSH") '品ＷＦラフネス表測定位置＿方
            If fldNameExist("HWFRNFST") Then .HWFRNFST = rs("HWFRNFST") '品ＷＦラフネス表測定位置＿点
            If fldNameExist("HWFRNFSI") Then .HWFRNFSI = rs("HWFRNFSI") '品ＷＦラフネス表測定位置＿位
            If fldNameExist("HWFRNFKW") Then .HWFRNFKW = rs("HWFRNFKW") '品ＷＦラフネス表検査方法
            If fldNameExist("HWFRNFZA") Then .HWFRNFZA = fncNullCheck(rs("HWFRNFZA")) '品ＷＦラフネス表除外領域
            If fldNameExist("HWFRNBMX") Then .HWFRNBMX = fncNullCheck(rs("HWFRNBMX")) '品ＷＦラフネス裏上限
            If fldNameExist("HWFRNBSH") Then .HWFRNBSH = rs("HWFRNBSH") '品ＷＦラフネス裏測定位置＿方
            If fldNameExist("HWFRNBST") Then .HWFRNBST = rs("HWFRNBST") '品ＷＦラフネス裏測定位置＿点
            If fldNameExist("HWFRNBSI") Then .HWFRNBSI = rs("HWFRNBSI") '品ＷＦラフネス裏測定位置＿位
            If fldNameExist("HWFRNBKW") Then .HWFRNBKW = rs("HWFRNBKW") '品ＷＦラフネス裏検査方法
            If fldNameExist("HWFRNBZA") Then .HWFRNBZA = fncNullCheck(rs("HWFRNBZA")) '品ＷＦラフネス裏除外領域
            If fldNameExist("HWFDENKU") Then .HWFDENKU = rs("HWFDENKU") '品ＷＦＤｅｎ検査有無
            If fldNameExist("HWFDENMX") Then .HWFDENMX = fncNullCheck(rs("HWFDENMX")) '品ＷＦＤｅｎ上限
            If fldNameExist("HWFDENMN") Then .HWFDENMN = fncNullCheck(rs("HWFDENMN")) '品ＷＦＤｅｎ下限
            If fldNameExist("HWFDENHT") Then .HWFDENHT = rs("HWFDENHT") '品ＷＦＤｅｎ保証方法＿対
            If fldNameExist("HWFDENHS") Then .HWFDENHS = rs("HWFDENHS") '品ＷＦＤｅｎ保証方法＿処
            If fldNameExist("HWFDVDKU") Then .HWFDVDKU = rs("HWFDVDKU") '品ＷＦＤＶＤ２検査有無
            If fldNameExist("HWFDVDMX") Then .HWFDVDMX = fncNullCheck(rs("HWFDVDMX")) '品ＷＦＤＶＤ２上限
            If fldNameExist("HWFDVDMN") Then .HWFDVDMN = fncNullCheck(rs("HWFDVDMN")) '品ＷＦＤＶＤ２下限
            If fldNameExist("HWFDVDHT") Then .HWFDVDHT = rs("HWFDVDHT") '品ＷＦＤＶＤ２保証方法＿対
            If fldNameExist("HWFDVDHS") Then .HWFDVDHS = rs("HWFDVDHS") '品ＷＦＤＶＤ２保証方法＿処
            If fldNameExist("HWFLDLKU") Then .HWFLDLKU = rs("HWFLDLKU") '品ＷＦＬ／ＤＬ検査有無
            If fldNameExist("HWFLDLMX") Then .HWFLDLMX = fncNullCheck(rs("HWFLDLMX")) '品ＷＦＬ／ＤＬ上限
            If fldNameExist("HWFLDLMN") Then .HWFLDLMN = fncNullCheck(rs("HWFLDLMN")) '品ＷＦＬ／ＤＬ下限
            If fldNameExist("HWFLDLHT") Then .HWFLDLHT = rs("HWFLDLHT") '品ＷＦＬ／ＤＬ保証方法＿対
            If fldNameExist("HWFLDLHS") Then .HWFLDLHS = rs("HWFLDLHS") '品ＷＦＬ／ＤＬ保証方法＿処
            If fldNameExist("HWFGDSPH") Then .HWFGDSPH = rs("HWFGDSPH") '品ＷＦＧＤ測定位置＿方
            If fldNameExist("HWFGDSPT") Then .HWFGDSPT = rs("HWFGDSPT") '品ＷＦＧＤ測定位置＿点
            If fldNameExist("HWFGDSPR") Then .HWFGDSPR = rs("HWFGDSPR") '品ＷＦＧＤ測定位置＿領
            If fldNameExist("HWFGDSZY") Then .HWFGDSZY = rs("HWFGDSZY") '品ＷＦＧＤ測定条件
            If fldNameExist("HWFGDZAR") Then .HWFGDZAR = fncNullCheck(rs("HWFGDZAR")) '品ＷＦＧＤ除外領域
            If fldNameExist("HWFGDKHM") Then .HWFGDKHM = rs("HWFGDKHM") '品ＷＦＧＤ検査頻度＿枚
            If fldNameExist("HWFGDKHN") Then .HWFGDKHN = rs("HWFGDKHN") '品ＷＦＧＤ検査頻度＿抜
            If fldNameExist("HWFGDKHH") Then .HWFGDKHH = rs("HWFGDKHH") '品ＷＦＧＤ検査頻度＿保
            If fldNameExist("HWFGDKHU") Then .HWFGDKHU = rs("HWFGDKHU") '品ＷＦＧＤ検査頻度＿ウ
            If fldNameExist("HWFDSOKE") Then .HWFDSOKE = rs("HWFDSOKE") '品ＷＦＤＳＯＤ検査
            If fldNameExist("HWFDSOMX") Then .HWFDSOMX = fncNullCheck(rs("HWFDSOMX")) '品ＷＦＤＳＯＤ上限
            If fldNameExist("HWFDSOMN") Then .HWFDSOMN = fncNullCheck(rs("HWFDSOMN")) '品ＷＦＤＳＯＤ下限
            If fldNameExist("HWFDSOAX") Then .HWFDSOAX = fncNullCheck(rs("HWFDSOAX")) '品ＷＦＤＳＯＤ領域上限
            If fldNameExist("HWFDSOAN") Then .HWFDSOAN = fncNullCheck(rs("HWFDSOAN")) '品ＷＦＤＳＯＤ領域下限
            If fldNameExist("HWFDSOHT") Then .HWFDSOHT = rs("HWFDSOHT") '品ＷＦＤＳＯＤ保証方法＿対
            If fldNameExist("HWFDSOHS") Then .HWFDSOHS = rs("HWFDSOHS") '品ＷＦＤＳＯＤ保証方法＿処
            If fldNameExist("HWFDSOKM") Then .HWFDSOKM = rs("HWFDSOKM") '品ＷＦＤＳＯＤ検査頻度＿枚
            If fldNameExist("HWFDSOKN") Then .HWFDSOKN = rs("HWFDSOKN") '品ＷＦＤＳＯＤ検査頻度＿抜
            If fldNameExist("HWFDSOKH") Then .HWFDSOKH = rs("HWFDSOKH") '品ＷＦＤＳＯＤ検査頻度＿保
            If fldNameExist("HWFDSOKU") Then .HWFDSOKU = rs("HWFDSOKU") '品ＷＦＤＳＯＤ検査頻度＿ウ
            If fldNameExist("HWFNTPUM") Then .HWFNTPUM = rs("HWFNTPUM") '品ＷＦ平坦ナノトポ有無
            If fldNameExist("HWFNTPK1") Then .HWFNTPK1 = fncNullCheck(rs("HWFNTPK1")) '品ＷＦ平坦ナノトポ規格１
            If fldNameExist("HWFNTPP1") Then .HWFNTPP1 = fncNullCheck(rs("HWFNTPP1")) '品ＷＦ平坦ナノトポＰＵＡ１
            If fldNameExist("HWFNTPS1") Then .HWFNTPS1 = fncNullCheck(rs("HWFNTPS1")) '品ＷＦ平坦ナノトポサイト１
            If fldNameExist("HWFNTPK2") Then .HWFNTPK2 = fncNullCheck(rs("HWFNTPK2")) '品ＷＦ平坦ナノトポ規格２
            If fldNameExist("HWFNTPP2") Then .HWFNTPP2 = fncNullCheck(rs("HWFNTPP2")) '品ＷＦ平坦ナノトポＰＵＡ２
            If fldNameExist("HWFNTPS2") Then .HWFNTPS2 = fncNullCheck(rs("HWFNTPS2")) '品ＷＦ平坦ナノトポサイト２
            If fldNameExist("HWFNTPK3") Then .HWFNTPK3 = fncNullCheck(rs("HWFNTPK3")) '品ＷＦ平坦ナノトポ規格３
            If fldNameExist("HWFNTPP3") Then .HWFNTPP3 = fncNullCheck(rs("HWFNTPP3")) '品ＷＦ平坦ナノトポＰＵＡ３
            If fldNameExist("HWFNTPS3") Then .HWFNTPS3 = fncNullCheck(rs("HWFNTPS3")) '品ＷＦ平坦ナノトポサイト３
            If fldNameExist("HWFNTPZA") Then .HWFNTPZA = fncNullCheck(rs("HWFNTPZA")) '品ＷＦ平坦ナノトポ除外領域
            If fldNameExist("HWFNTPHT") Then .HWFNTPHT = rs("HWFNTPHT") '品ＷＦ平坦ナノトポ保証方法＿対
            If fldNameExist("HWFNTPHS") Then .HWFNTPHS = rs("HWFNTPHS") '品ＷＦ平坦ナノトポ保証方法＿処
            If fldNameExist("HWFNTPKM") Then .HWFNTPKM = rs("HWFNTPKM") '品ＷＦ平坦ナノトポ検査頻度＿枚
            If fldNameExist("HWFNTPKN") Then .HWFNTPKN = rs("HWFNTPKN") '品ＷＦ平坦ナノトポ検査頻度＿抜
            If fldNameExist("HWFNTPKH") Then .HWFNTPKH = rs("HWFNTPKH") '品ＷＦ平坦ナノトポ検査頻度＿保
            If fldNameExist("HWFNTPKU") Then .HWFNTPKU = rs("HWFNTPKU") '品ＷＦ平坦ナノトポ検査頻度＿ウ
            If fldNameExist("HWFCRSSK") Then .HWFCRSSK = rs("HWFCRSSK") '品ＷＦ平坦クロスＳＳ検査
            If fldNameExist("HWFMDCEN") Then .HWFMDCEN = fncNullCheck(rs("HWFMDCEN")) '品ＷＦ平坦面ダレ高低差中心
            If fldNameExist("HWFMDMAX") Then .HWFMDMAX = fncNullCheck(rs("HWFMDMAX")) '品ＷＦ平坦面ダレ高低差上限
            If fldNameExist("HWFMDMIN") Then .HWFMDMIN = fncNullCheck(rs("HWFMDMIN")) '品ＷＦ平坦面ダレ高低差下限
            If fldNameExist("HWFMDSPH") Then .HWFMDSPH = rs("HWFMDSPH") '品ＷＦ平坦面ダレ測定位置＿方
            If fldNameExist("HWFMDSPT") Then .HWFMDSPT = rs("HWFMDSPT") '品ＷＦ平坦面ダレ測定位置＿点
            If fldNameExist("HWFMDSPI") Then .HWFMDSPI = rs("HWFMDSPI") '品ＷＦ平坦面ダレ測定位置＿位
            If fldNameExist("HWFMDHWT") Then .HWFMDHWT = rs("HWFMDHWT") '品ＷＦ平坦面ダレ保証方法＿対
            If fldNameExist("HWFMDHWS") Then .HWFMDHWS = rs("HWFMDHWS") '品ＷＦ平坦面ダレ保証方法＿処
            If fldNameExist("HWFMDKHM") Then .HWFMDKHM = rs("HWFMDKHM") '品ＷＦ平坦面ダレ検査頻度＿枚
            If fldNameExist("HWFMDKHN") Then .HWFMDKHN = rs("HWFMDKHN") '品ＷＦ平坦面ダレ検査頻度＿抜
            If fldNameExist("HWFMDKHH") Then .HWFMDKHH = rs("HWFMDKHH") '品ＷＦ平坦面ダレ検査頻度＿保
            If fldNameExist("HWFMDKHU") Then .HWFMDKHU = rs("HWFMDKHU") '品ＷＦ平坦面ダレ検査頻度＿ウ
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN") 'Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN") '処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO") '仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO") 'ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO") 'ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID") '社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE") '登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE") '更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG") '送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE") '送信日付
            If fldNameExist("HWFDVDMXN") Then .HWFDVDMXN = fncNullCheck(rs("HWFDVDMXN")) '品ＷＦＤＶＤ２上限
            If fldNameExist("HWFDVDMNN") Then .HWFDVDMNN = fncNullCheck(rs("HWFDVDMNN")) '品ＷＦＤＶＤ２下限
'            If fldNameExist("HWFDSONWY") Then .HWFDSONWY = rs("HWFDSONWY") '品ＷＦＤＳＯＤ熱処理法
'            If fldNameExist("HWFMSUMX") Then .HWFMSUMX = fncNullCheck(rs("HWFMSUMX")) '品ＷＦＭスクラッチ上限
'            If fldNameExist("HWFMSUZY") Then .HWFMSUZY = rs("HWFMSUZY") '品ＷＦＭスクラッチ測定条件
'            If fldNameExist("HWFMSUKW") Then .HWFMSUKW = rs("HWFMSUKW") '品ＷＦＭスクラッチ検査方法
'            If fldNameExist("HWFMSUSZ") Then .HWFMSUSZ = fncNullCheck(rs("HWFMSUSZ")) '品ＷＦＭスクラッチサイズ
'            If fldNameExist("HWFNP1AR") Then .HWFNP1AR = fncNullCheck(rs("HWFNP1AR")) '品WFナノトポ１エリア
'            If fldNameExist("HWFNP1MAX") Then .HWFNP1MAX = fncNullCheck(rs("HWFNP1MAX")) '品WFナノトポ１上限
'            If fldNameExist("HWFNP2AR") Then .HWFNP2AR = fncNullCheck(rs("HWFNP2AR")) '品WFナノトポ２エリア
'            If fldNameExist("HWFNP2MAX") Then .HWFNP2MAX = fncNullCheck(rs("HWFNP2MAX")) '品WFナノトポ２上限
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME026 = FUNCTION_RETURN_SUCCESS

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

'概要      :テーブル「TBCME028」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records()     ,O  ,typ_TBCME028    ,抽出レコード
'          :formID        ,I  ,String          ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban     ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :新規作成 2005/06/15 ffc)tanabe
Public Function DBDRV_GetTBCME028(records() As typ_TBCME028, formID$, HIN() As tFullHinban) As FUNCTION_RETURN

    Dim sql         As String           'SQL全体
    Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
    Dim sqlWhere    As String           'SQLWhere部
    Dim rs          As OraDynaset       'RecordSet
    Dim recCnt      As Long             'レコード数
    Dim key         As String           '検索KEY
    Dim i           As Long             'ﾙｰﾌﾟｶｳﾝﾄ
    Dim j           As Long             'ﾙｰﾌﾟｶｳﾝﾄ2


    DBDRV_GetTBCME028 = FUNCTION_RETURN_FAILURE
            
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME028"

    Select Case formID
        Case "f_cmec067_1"           '「SPV実測参照」
            sqlBase = "SELECT HINBAN, MNOREVNO, FACTORY, OPECOND, HWFSPVMX, HWFSPVKM, HWFSPVKN, HWFSPVKH, HWFSPVKU, HWFSPVSH, HWFSPVST, HWFSPVSI, HWFSPVHT, " & _
                "HWFSPVHS, HWFDLMIN, HWFDLMAX, HWFDLKHM, HWFDLKHN, HWFDLKHH, HWFDLKHU, HWFDLSPH, HWFDLSPT, HWFDLSPI, HWFDLHWT, HWFDLHWS, HWFSPVMXN "
    End Select
       
    sqlBase = sqlBase & "From TBCME028"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                           '品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                     '製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                        '工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                        '操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")                  '品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")                     '品管理社員Ｎｏ
            If fldNameExist("HMGWFSNO") Then .HMGWFSNO = rs("HMGWFSNO")                     '品管理ＷＦ製品番号
            If fldNameExist("HMGWFSNE") Then .HMGWFSNE = fncNullCheck(rs("HMGWFSNE"))       '品管理ＷＦ製品番号枝番
            If fldNameExist("HWFMK1SI") Then .HWFMK1SI = fncNullCheck(rs("HWFMK1SI"))       '品ＷＦ面検欠陥１サイズ
            If fldNameExist("HWFMK1MX") Then .HWFMK1MX = fncNullCheck(rs("HWFMK1MX"))       '品ＷＦ面検欠陥１上限
            If fldNameExist("HWFMK1SZ") Then .HWFMK1SZ = rs("HWFMK1SZ")                     '品ＷＦ面検欠陥１測定条件
            If fldNameExist("HWFMK1ZA") Then .HWFMK1ZA = fncNullCheck(rs("HWFMK1ZA"))       '品ＷＦ面検欠陥１除外領域
            If fldNameExist("HWFMK1HT") Then .HWFMK1HT = rs("HWFMK1HT")                     '品ＷＦ面検欠陥１保証方法＿対
            If fldNameExist("HWFMK1HS") Then .HWFMK1HS = rs("HWFMK1HS")                     '品ＷＦ面検欠陥１保証方法＿処
            If fldNameExist("HWFMK1KM") Then .HWFMK1KM = rs("HWFMK1KM")                     '品ＷＦ面検欠陥１検査頻度＿枚
            If fldNameExist("HWFMK1KN") Then .HWFMK1KN = rs("HWFMK1KN")                     '品ＷＦ面検欠陥１検査頻度＿抜
            If fldNameExist("HWFMK1KH") Then .HWFMK1KH = rs("HWFMK1KH")                     '品ＷＦ面検欠陥１検査頻度＿保
            If fldNameExist("HWFMK1KU") Then .HWFMK1KU = rs("HWFMK1KU")                     '品ＷＦ面検欠陥１検査頻度＿ウ
            If fldNameExist("HWFM1B1") Then .HWFM1B1 = fncNullCheck(rs("HWFM1B1"))          '品ＷＦ面検欠陥１境界１
            If fldNameExist("HWFM1B1B") Then .HWFM1B1B = fncNullCheck(rs("HWFM1B1B"))       '品ＷＦ面検欠陥１境界１下
            If fldNameExist("HWFM1B2") Then .HWFM1B2 = fncNullCheck(rs("HWFM1B2"))          '品ＷＦ面検欠陥１境界２
            If fldNameExist("HWFM1B2B") Then .HWFM1B2B = fncNullCheck(rs("HWFM1B2B"))       '品ＷＦ面検欠陥１境界２下
            If fldNameExist("HWFM1B3") Then .HWFM1B3 = fncNullCheck(rs("HWFM1B3"))          '品ＷＦ面検欠陥１境界３
            If fldNameExist("HWFM1B3B") Then .HWFM1B3B = fncNullCheck(rs("HWFM1B3B"))       '品ＷＦ面検欠陥１境界３下
            If fldNameExist("HWFMK2SI") Then .HWFMK2SI = fncNullCheck(rs("HWFMK2SI"))       '品ＷＦ面検欠陥２サイズ
            If fldNameExist("HWFMK2MX") Then .HWFMK2MX = fncNullCheck(rs("HWFMK2MX"))       '品ＷＦ面検欠陥２上限
            If fldNameExist("HWFMK2HT") Then .HWFMK2HT = rs("HWFMK2HT")                     '品ＷＦ面検欠陥２保証方法＿対
            If fldNameExist("HWFMK2HS") Then .HWFMK2HS = rs("HWFMK2HS")                     '品ＷＦ面検欠陥２保証方法＿処
            If fldNameExist("HWFMK2KM") Then .HWFMK2KM = rs("HWFMK2KM")                     '品ＷＦ面検欠陥２検査頻度＿枚
            If fldNameExist("HWFMK2KN") Then .HWFMK2KN = rs("HWFMK2KN")                     '品ＷＦ面検欠陥２検査頻度＿抜
            If fldNameExist("HWFMK2KH") Then .HWFMK2KH = rs("HWFMK2KH")                     '品ＷＦ面検欠陥２検査頻度＿保
            If fldNameExist("HWFMK2KU") Then .HWFMK2KU = rs("HWFMK2KU")                     '品ＷＦ面検欠陥２検査頻度＿ウ
            If fldNameExist("HWFM2B1") Then .HWFM2B1 = fncNullCheck(rs("HWFM2B1"))          '品ＷＦ面検欠陥２境界１
            If fldNameExist("HWFM2B1B") Then .HWFM2B1B = fncNullCheck(rs("HWFM2B1B"))       '品ＷＦ面検欠陥２境界１下
            If fldNameExist("HWFM2B2") Then .HWFM2B2 = fncNullCheck(rs("HWFM2B2"))          '品ＷＦ面検欠陥２境界２
            If fldNameExist("HWFM2B2B") Then .HWFM2B2B = fncNullCheck(rs("HWFM2B2B"))       '品ＷＦ面検欠陥２境界２下
            If fldNameExist("HWFM2B3") Then .HWFM2B3 = fncNullCheck(rs("HWFM2B3"))          '品ＷＦ面検欠陥２境界３
            If fldNameExist("HWFM2B3B") Then .HWFM2B3B = fncNullCheck(rs("HWFM2B3B"))       '品ＷＦ面検欠陥２境界３下
            If fldNameExist("HWFMK3SI") Then .HWFMK3SI = fncNullCheck(rs("HWFMK3SI"))       '品ＷＦ面検欠陥３サイズ
            If fldNameExist("HWFMK3MX") Then .HWFMK3MX = fncNullCheck(rs("HWFMK3MX"))       '品ＷＦ面検欠陥３上限
            If fldNameExist("HWFMK3HT") Then .HWFMK3HT = rs("HWFMK3HT")                     '品ＷＦ面検欠陥３保証方法＿対
            If fldNameExist("HWFMK3HS") Then .HWFMK3HS = rs("HWFMK3HS")                     '品ＷＦ面検欠陥３保証方法＿処
            If fldNameExist("HWFMK3KM") Then .HWFMK3KM = rs("HWFMK3KM")                     '品ＷＦ面検欠陥３検査頻度＿枚
            If fldNameExist("HWFMK3KN") Then .HWFMK3KN = rs("HWFMK3KN")                     '品ＷＦ面検欠陥３検査頻度＿抜
            If fldNameExist("HWFMK3KH") Then .HWFMK3KH = rs("HWFMK3KH")                     '品ＷＦ面検欠陥３検査頻度＿保
            If fldNameExist("HWFMK3KU") Then .HWFMK3KU = rs("HWFMK3KU")                     '品ＷＦ面検欠陥３検査頻度＿ウ
            If fldNameExist("HWFM3B1") Then .HWFM3B1 = fncNullCheck(rs("HWFM3B1"))          '品ＷＦ面検欠陥３境界１
            If fldNameExist("HWFM3B1B") Then .HWFM3B1B = fncNullCheck(rs("HWFM3B1B"))       '品ＷＦ面検欠陥３境界１下
            If fldNameExist("HWFM3B2") Then .HWFM3B2 = fncNullCheck(rs("HWFM3B2"))          '品ＷＦ面検欠陥３境界２
            If fldNameExist("HWFM3B2B") Then .HWFM3B2B = fncNullCheck(rs("HWFM3B2B"))       '品ＷＦ面検欠陥３境界２下
            If fldNameExist("HWFM3B3") Then .HWFM3B3 = fncNullCheck(rs("HWFM3B3"))          '品ＷＦ面検欠陥３境界３
            If fldNameExist("HWFM3B3B") Then .HWFM3B3B = fncNullCheck(rs("HWFM3B3B"))       '品ＷＦ面検欠陥３境界３下
            If fldNameExist("HWFMK4SI") Then .HWFMK4SI = fncNullCheck(rs("HWFMK4SI"))       '品ＷＦ面検欠陥４サイズ
            If fldNameExist("HWFMK4MX") Then .HWFMK4MX = fncNullCheck(rs("HWFMK4MX"))       '品ＷＦ面検欠陥４上限
            If fldNameExist("HWFMK4HT") Then .HWFMK4HT = rs("HWFMK4HT")                     '品ＷＦ面検欠陥４保証方法＿対
            If fldNameExist("HWFMK4HS") Then .HWFMK4HS = rs("HWFMK4HS")                     '品ＷＦ面検欠陥４保証方法＿処
            If fldNameExist("HWFMK4KM") Then .HWFMK4KM = rs("HWFMK4KM")                     '品ＷＦ面検欠陥４検査頻度＿枚
            If fldNameExist("HWFMK4KN") Then .HWFMK4KN = rs("HWFMK4KN")                     '品ＷＦ面検欠陥４検査頻度＿抜
            If fldNameExist("HWFMK4KH") Then .HWFMK4KH = rs("HWFMK4KH")                     '品ＷＦ面検欠陥４検査頻度＿保
            If fldNameExist("HWFMK4KU") Then .HWFMK4KU = rs("HWFMK4KU")                     '品ＷＦ面検欠陥４検査頻度＿ウ
            If fldNameExist("HWFM4B1") Then .HWFM4B1 = fncNullCheck(rs("HWFM4B1"))          '品ＷＦ面検欠陥４境界１
            If fldNameExist("HWFM4B1B") Then .HWFM4B1B = fncNullCheck(rs("HWFM4B1B"))       '品ＷＦ面検欠陥４境界１下
            If fldNameExist("HWFM4B2") Then .HWFM4B2 = fncNullCheck(rs("HWFM4B2"))          '品ＷＦ面検欠陥４境界２
            If fldNameExist("HWFM4B2B") Then .HWFM4B2B = fncNullCheck(rs("HWFM4B2B"))       '品ＷＦ面検欠陥４境界２下
            If fldNameExist("HWFM4B3") Then .HWFM4B3 = fncNullCheck(rs("HWFM4B3"))          '品ＷＦ面検欠陥４境界３
            If fldNameExist("HWFM4B3B") Then .HWFM4B3B = fncNullCheck(rs("HWFM4B3B"))       '品ＷＦ面検欠陥４境界３下
            If fldNameExist("HWFMB1SI") Then .HWFMB1SI = fncNullCheck(rs("HWFMB1SI"))       '品ＷＦ面検欠陥裏１サイズ
            If fldNameExist("HWFMB1MX") Then .HWFMB1MX = fncNullCheck(rs("HWFMB1MX"))       '品ＷＦ面検欠陥裏１上限
            If fldNameExist("HWFMB1SZ") Then .HWFMB1SZ = rs("HWFMB1SZ")                     '品ＷＦ面検欠陥裏１測定条件
            If fldNameExist("HWFMB1ZA") Then .HWFMB1ZA = fncNullCheck(rs("HWFMB1ZA"))       '品ＷＦ面検欠陥裏１除外領域
            If fldNameExist("HWFMB1HT") Then .HWFMB1HT = rs("HWFMB1HT")                     '品ＷＦ面検欠陥裏１保証方法＿対
            If fldNameExist("HWFMB1HS") Then .HWFMB1HS = rs("HWFMB1HS")                     '品ＷＦ面検欠陥裏１保証方法＿処
            If fldNameExist("HWFMB1KM") Then .HWFMB1KM = rs("HWFMB1KM")                     '品ＷＦ面検欠陥裏１検査頻度＿枚
            If fldNameExist("HWFMB1KN") Then .HWFMB1KN = rs("HWFMB1KN")                     '品ＷＦ面検欠陥裏１検査頻度＿抜
            If fldNameExist("HWFMB1KH") Then .HWFMB1KH = rs("HWFMB1KH")                     '品ＷＦ面検欠陥裏１検査頻度＿保
            If fldNameExist("HWFMB1KU") Then .HWFMB1KU = rs("HWFMB1KU")                     '品ＷＦ面検欠陥裏１検査頻度＿ウ
            If fldNameExist("HWFMB2SI") Then .HWFMB2SI = fncNullCheck(rs("HWFMB2SI"))       '品ＷＦ面検欠陥裏２サイズ
            If fldNameExist("HWFMB2MX") Then .HWFMB2MX = fncNullCheck(rs("HWFMB2MX"))       '品ＷＦ面検欠陥裏２上限
            If fldNameExist("HWFMB2SZ") Then .HWFMB2SZ = rs("HWFMB2SZ")                     '品ＷＦ面検欠陥裏２測定条件
            If fldNameExist("HWFMB2ZA") Then .HWFMB2ZA = fncNullCheck(rs("HWFMB2ZA"))       '品ＷＦ面検欠陥裏２除外領域
            If fldNameExist("HWFMB2HT") Then .HWFMB2HT = rs("HWFMB2HT")                     '品ＷＦ面検欠陥裏２保証方法＿対
            If fldNameExist("HWFMB2HS") Then .HWFMB2HS = rs("HWFMB2HS")                     '品ＷＦ面検欠陥裏２保証方法＿処
            If fldNameExist("HWFMB2KM") Then .HWFMB2KM = rs("HWFMB2KM")                     '品ＷＦ面検欠陥裏２検査頻度＿枚
            If fldNameExist("HWFMB2KN") Then .HWFMB2KN = rs("HWFMB2KN")                     '品ＷＦ面検欠陥裏２検査頻度＿抜
            If fldNameExist("HWFMB2KH") Then .HWFMB2KH = rs("HWFMB2KH")                     '品ＷＦ面検欠陥裏２検査頻度＿保
            If fldNameExist("HWFMB2KU") Then .HWFMB2KU = rs("HWFMB2KU")                     '品ＷＦ面検欠陥裏２検査頻度＿ウ
            If fldNameExist("HWFMKSRE") Then .HWFMKSRE = rs("HWFMKSRE")                     '品ＷＦ面検欠陥測定器
            If fldNameExist("HWFMKKW") Then .HWFMKKW = rs("HWFMKKW")                        '品ＷＦ面検欠陥検査方法
            If fldNameExist("HWFMPIPT") Then .HWFMPIPT = rs("HWFMPIPT")                     '品ＷＦ面検欠陥ＰＩＰ検査
            If fldNameExist("HWFMPIPK") Then .HWFMPIPK = fncNullCheck(rs("HWFMPIPK"))       '品ＷＦ面検欠陥ＰＩＰ個数
            If fldNameExist("HWFMPISH") Then .HWFMPISH = rs("HWFMPISH")                     '品ＷＦ面検ＰＩＰ測定位置＿方
            If fldNameExist("HWFMPIST") Then .HWFMPIST = rs("HWFMPIST")                     '品ＷＦ面検ＰＩＰ測定位置＿点
            If fldNameExist("HWFMPISI") Then .HWFMPISI = rs("HWFMPISI")                     '品ＷＦ面検ＰＩＰ測定位置＿位
            If fldNameExist("HWFMPIKM") Then .HWFMPIKM = rs("HWFMPIKM")                     '品ＷＦ面検ＰＩＰ検査頻度＿枚
            If fldNameExist("HWFMPIKN") Then .HWFMPIKN = rs("HWFMPIKN")                     '品ＷＦ面検ＰＩＰ検査頻度＿抜
            If fldNameExist("HWFMPIKH") Then .HWFMPIKH = rs("HWFMPIKH")                     '品ＷＦ面検ＰＩＰ検査頻度＿保
            If fldNameExist("HWFMPIKU") Then .HWFMPIKU = rs("HWFMPIKU")                     '品ＷＦ面検ＰＩＰ検査頻度＿ウ
            If fldNameExist("HWFMNMAX") Then .HWFMNMAX = fncNullCheck(rs("HWFMNMAX"))       '品ＷＦ金属濃度上限
            If fldNameExist("HWFMNALX") Then .HWFMNALX = fncNullCheck(rs("HWFMNALX"))       '品ＷＦ金属濃度ＡＬ上限
            If fldNameExist("HWFMNCAX") Then .HWFMNCAX = fncNullCheck(rs("HWFMNCAX"))       '品ＷＦ金属濃度ＣＡ上限
            If fldNameExist("HWFMNCRX") Then .HWFMNCRX = fncNullCheck(rs("HWFMNCRX"))       '品ＷＦ金属濃度ＣＲ上限
            If fldNameExist("HWFMNCUX") Then .HWFMNCUX = fncNullCheck(rs("HWFMNCUX"))       '品ＷＦ金属濃度ＣＵ上限
            If fldNameExist("HWFMNFEX") Then .HWFMNFEX = fncNullCheck(rs("HWFMNFEX"))       '品ＷＦ金属濃度ＦＥ上限
            If fldNameExist("HWFMNKMX") Then .HWFMNKMX = fncNullCheck(rs("HWFMNKMX"))       '品ＷＦ金属濃度Ｋ上限
            If fldNameExist("HWFMNMGX") Then .HWFMNMGX = fncNullCheck(rs("HWFMNMGX"))       '品ＷＦ金属濃度ＭＧ上限
            If fldNameExist("HWFMNNAX") Then .HWFMNNAX = fncNullCheck(rs("HWFMNNAX"))       '品ＷＦ金属濃度ＮＡ上限
            If fldNameExist("HWFMNNIX") Then .HWFMNNIX = fncNullCheck(rs("HWFMNNIX"))       '品ＷＦ金属濃度ＮＩ上限
            If fldNameExist("HWFMNZNX") Then .HWFMNZNX = fncNullCheck(rs("HWFMNZNX"))       '品ＷＦ金属濃度ＺＮ上限
            If fldNameExist("HWFMNKWY") Then .HWFMNKWY = rs("HWFMNKWY")                     '品ＷＦ金属濃度検査方法
            If fldNameExist("HWFMNSPH") Then .HWFMNSPH = rs("HWFMNSPH")                     '品ＷＦ金属濃度測定位置＿方
            If fldNameExist("HWFMNSPT") Then .HWFMNSPT = rs("HWFMNSPT")                     '品ＷＦ金属濃度測定位置＿点
            If fldNameExist("HWFMNSPI") Then .HWFMNSPI = rs("HWFMNSPI")                     '品ＷＦ金属濃度測定位置＿位
            If fldNameExist("HWFMNHWT") Then .HWFMNHWT = rs("HWFMNHWT")                     '品ＷＦ金属濃度保証方法＿対
            If fldNameExist("HWFMNHWS") Then .HWFMNHWS = rs("HWFMNHWS")                     '品ＷＦ金属濃度保証方法＿処
            If fldNameExist("HWFMNKHM") Then .HWFMNKHM = rs("HWFMNKHM")                     '品ＷＦ金属濃度検査頻度＿枚
            If fldNameExist("HWFMNKHN") Then .HWFMNKHN = rs("HWFMNKHN")                     '品ＷＦ金属濃度検査頻度＿抜
            If fldNameExist("HWFMNKHH") Then .HWFMNKHH = rs("HWFMNKHH")                     '品ＷＦ金属濃度検査頻度＿保
            If fldNameExist("HWFMNKHU") Then .HWFMNKHU = rs("HWFMNKHU")                     '品ＷＦ金属濃度検査頻度＿ウ
            If fldNameExist("HWFSPVMX") Then .HWFSPVMX = fncNullCheck(rs("HWFSPVMX"))       '品ＷＦＳＰＶＦＥ上限
            If fldNameExist("HWFSPVKM") Then .HWFSPVKM = rs("HWFSPVKM")                     '品ＷＦＳＰＶＦＥ検査頻度＿枚
            If fldNameExist("HWFSPVKN") Then .HWFSPVKN = rs("HWFSPVKN")                     '品ＷＦＳＰＶＦＥ検査頻度＿抜
            If fldNameExist("HWFSPVKH") Then .HWFSPVKH = rs("HWFSPVKH")                     '品ＷＦＳＰＶＦＥ検査頻度＿保
            If fldNameExist("HWFSPVKU") Then .HWFSPVKU = rs("HWFSPVKU")                     '品ＷＦＳＰＶＦＥ検査頻度＿ウ
            If fldNameExist("HWFSPVSH") Then .HWFSPVSH = rs("HWFSPVSH")                     '品ＷＦＳＰＶＦＥ測定位置＿方
            If fldNameExist("HWFSPVST") Then .HWFSPVST = rs("HWFSPVST")                     '品ＷＦＳＰＶＦＥ測定位置＿点
            If fldNameExist("HWFSPVSI") Then .HWFSPVSI = rs("HWFSPVSI")                     '品ＷＦＳＰＶＦＥ測定位置＿位
            If fldNameExist("HWFSPVHT") Then .HWFSPVHT = rs("HWFSPVHT")                     '品ＷＦＳＰＶＦＥ保証方法＿対
            If fldNameExist("HWFSPVHS") Then .HWFSPVHS = rs("HWFSPVHS")                     '品ＷＦＳＰＶＦＥ保証方法＿処
            If fldNameExist("HWFDLMIN") Then .HWFDLMIN = fncNullCheck(rs("HWFDLMIN"))       '品ＷＦ拡散長下限
            If fldNameExist("HWFDLMAX") Then .HWFDLMAX = fncNullCheck(rs("HWFDLMAX"))       '品ＷＦ拡散長上限
            If fldNameExist("HWFDLKHM") Then .HWFDLKHM = rs("HWFDLKHM")                     '品ＷＦ拡散長検査頻度＿枚
            If fldNameExist("HWFDLKHN") Then .HWFDLKHN = rs("HWFDLKHN")                     '品ＷＦ拡散長検査頻度＿抜
            If fldNameExist("HWFDLKHH") Then .HWFDLKHH = rs("HWFDLKHH")                     '品ＷＦ拡散長検査頻度＿保
            If fldNameExist("HWFDLKHU") Then .HWFDLKHU = rs("HWFDLKHU")                     '品ＷＦ拡散長検査頻度＿ウ
            If fldNameExist("HWFDLSPH") Then .HWFDLSPH = rs("HWFDLSPH")                     '品ＷＦ拡散長測定位置＿方
            If fldNameExist("HWFDLSPT") Then .HWFDLSPT = rs("HWFDLSPT")                     '品ＷＦ拡散長測定位置＿点
            If fldNameExist("HWFDLSPI") Then .HWFDLSPI = rs("HWFDLSPI")                     '品ＷＦ拡散長測定位置＿位
            If fldNameExist("HWFDLHWT") Then .HWFDLHWT = rs("HWFDLHWT")                     '品ＷＦ拡散長保証方法＿対
            If fldNameExist("HWFDLHWS") Then .HWFDLHWS = rs("HWFDLHWS")                     '品ＷＦ拡散長保証方法＿処
            If fldNameExist("HWFGKNO1") Then .HWFGKNO1 = rs("HWFGKNO1")                     '品ＷＦ外観規格Ｎｏ１
            If fldNameExist("HWFGKNO2") Then .HWFGKNO2 = rs("HWFGKNO2")                     '品ＷＦ外観規格Ｎｏ２
            If fldNameExist("HWFOTMIN") Then .HWFOTMIN = fncNullCheck(rs("HWFOTMIN"))       '品ＷＦ酸化膜耐圧下限
            If fldNameExist("HWFOTMX1") Then .HWFOTMX1 = fncNullCheck(rs("HWFOTMX1"))       '品ＷＦ酸化膜耐圧上限１
            If fldNameExist("HWFOTMX2") Then .HWFOTMX2 = fncNullCheck(rs("HWFOTMX2"))       '品ＷＦ酸化膜耐圧上限２
            If fldNameExist("HWFOTSPH") Then .HWFOTSPH = rs("HWFOTSPH")                     '品ＷＦ酸化膜耐圧測定位置＿方
            If fldNameExist("HWFOTSPT") Then .HWFOTSPT = rs("HWFOTSPT")                     '品ＷＦ酸化膜耐圧測定位置＿点
            If fldNameExist("HWFOTSPI") Then .HWFOTSPI = rs("HWFOTSPI")                     '品ＷＦ酸化膜耐圧測定位置＿位
            If fldNameExist("HWFOTHWT") Then .HWFOTHWT = rs("HWFOTHWT")                     '品ＷＦ酸化膜耐圧保証方法＿対
            If fldNameExist("HWFOTHWS") Then .HWFOTHWS = rs("HWFOTHWS")                     '品ＷＦ酸化膜耐圧保証方法＿処
            If fldNameExist("HWFOTKWY") Then .HWFOTKWY = rs("HWFOTKWY")                     '品ＷＦ酸化膜耐圧検査方法
            If fldNameExist("HWFOTKW1") Then .HWFOTKW1 = rs("HWFOTKW1")                     '品ＷＦ酸化膜耐圧検査方法１
            If fldNameExist("HWFOTKW2") Then .HWFOTKW2 = rs("HWFOTKW2")                     '品ＷＦ酸化膜耐圧検査方法２
            If fldNameExist("HWFOTKHM") Then .HWFOTKHM = rs("HWFOTKHM")                     '品ＷＦ酸化膜耐圧検査頻度＿枚
            If fldNameExist("HWFOTKHN") Then .HWFOTKHN = rs("HWFOTKHN")                     '品ＷＦ酸化膜耐圧検査頻度＿抜
            If fldNameExist("HWFOTKHH") Then .HWFOTKHH = rs("HWFOTKHH")                     '品ＷＦ酸化膜耐圧検査頻度＿保
            If fldNameExist("HWFOTKHU") Then .HWFOTKHU = rs("HWFOTKHU")                     '品ＷＦ酸化膜耐圧検査頻度＿ウ
            If fldNameExist("HWFTSPHM") Then .HWFTSPHM = rs("HWFTSPHM")                     '品ＷＦトレスサンプル頻度＿枚
            If fldNameExist("HWFTSPHN") Then .HWFTSPHN = rs("HWFTSPHN")                     '品ＷＦトレスサンプル頻度＿抜
            If fldNameExist("HWFTSPHH") Then .HWFTSPHH = rs("HWFTSPHH")                     '品ＷＦトレスサンプル頻度＿保
            If fldNameExist("HWFTSPHU") Then .HWFTSPHU = rs("HWFTSPHU")                     '品ＷＦトレスサンプル頻度＿ウ
            If fldNameExist("HWFLTDCX") Then .HWFLTDCX = fncNullCheck(rs("HWFLTDCX"))       '品ＷＦＬＴＤ濃度ＣＵ上限
            If fldNameExist("HWFLTDIN") Then .HWFLTDIN = rs("HWFLTDIN")                     '品ＷＦＬＴＤ濃度指数
            If fldNameExist("HWFLTDKW") Then .HWFLTDKW = rs("HWFLTDKW")                     '品ＷＦＬＴＤ濃度検査方法
            If fldNameExist("HWFLTDSH") Then .HWFLTDSH = rs("HWFLTDSH")                     '品ＷＦＬＴＤ濃度測定位置＿方
            If fldNameExist("HWFLTDST") Then .HWFLTDST = rs("HWFLTDST")                     '品ＷＦＬＴＤ濃度測定位置＿点
            If fldNameExist("HWFLTDSI") Then .HWFLTDSI = rs("HWFLTDSI")                     '品ＷＦＬＴＤ濃度測定位置＿位
            If fldNameExist("HWFLTDHT") Then .HWFLTDHT = rs("HWFLTDHT")                     '品ＷＦＬＴＤ濃度保証方法＿対
            If fldNameExist("HWFLTDHS") Then .HWFLTDHS = rs("HWFLTDHS")                     '品ＷＦＬＴＤ濃度保証方法＿処
            If fldNameExist("HWFLTDKM") Then .HWFLTDKM = rs("HWFLTDKM")                     '品ＷＦＬＴＤ濃度検査頻度＿枚
            If fldNameExist("HWFLTDKN") Then .HWFLTDKN = rs("HWFLTDKN")                     '品ＷＦＬＴＤ濃度検査頻度＿抜
            If fldNameExist("HWFLTDKH") Then .HWFLTDKH = rs("HWFLTDKH")                     '品ＷＦＬＴＤ濃度検査頻度＿保
            If fldNameExist("HWFLTDKU") Then .HWFLTDKU = rs("HWFLTDKU")                     '品ＷＦＬＴＤ濃度検査頻度＿ウ
            If fldNameExist("IFKBN") Then .IFKBN = rs("IFKBN")                              'Ｉ／Ｆ区分
            If fldNameExist("SYORIKBN") Then .SYORIKBN = rs("SYORIKBN")                     '処理区分
            If fldNameExist("SPECRRNO") Then .SPECRRNO = rs("SPECRRNO")                     '仕様登録依頼番号
            If fldNameExist("SXLMCNO") Then .SXLMCNO = rs("SXLMCNO")                        'ＳＸＬ製作条件番号
            If fldNameExist("WFMCNO") Then .WFMCNO = rs("WFMCNO")                           'ＷＦ製作条件番号
            If fldNameExist("STAFFID") Then .StaffID = rs("STAFFID")                        '社員ID
            If fldNameExist("REGDATE") Then .REGDATE = rs("REGDATE")                        '登録日付
            If fldNameExist("UPDDATE") Then .UPDDATE = rs("UPDDATE")                        '更新日付
            If fldNameExist("SENDFLAG") Then .SENDFLAG = rs("SENDFLAG")                     '送信フラグ
            If fldNameExist("SENDDATE") Then .SENDDATE = rs("SENDDATE")                     '送信日付
            If fldNameExist("HWFSPVAM") Then .HWFSPVAM = fncNullCheck(rs("HWFSPVAM"))       '品ＷＦＳＰＶＦＥ平均
            If fldNameExist("HWFMK1MC") Then .HWFMK1MC = rs("HWFMK1MC")                     '品ＷＦ面検欠陥１面指定
            If fldNameExist("HWFMK2MC") Then .HWFMK2MC = rs("HWFMK2MC")                     '品ＷＦ面検欠陥２面指定
            If fldNameExist("HWFMK3MC") Then .HWFMK3MC = rs("HWFMK3MC")                     '品ＷＦ面検欠陥３面指定
            If fldNameExist("HWFMK4MC") Then .HWFMK4MC = rs("HWFMK4MC")                     '品ＷＦ面検欠陥４面指定
            If fldNameExist("HWFMK5MC") Then .HWFMK5MC = rs("HWFMK5MC")                     '品ＷＦ面検欠陥５面指定
            If fldNameExist("HWFMK6MC") Then .HWFMK6MC = rs("HWFMK6MC")                     '品ＷＦ面検欠陥６面指定
            If fldNameExist("HWFMK2SZ") Then .HWFMK2SZ = rs("HWFMK2SZ")                     '品ＷＦ面検欠陥２測定条件
            If fldNameExist("HWFMK3SZ") Then .HWFMK3SZ = rs("HWFMK3SZ")                     '品ＷＦ面検欠陥３測定条件
            If fldNameExist("HWFMK4SZ") Then .HWFMK4SZ = rs("HWFMK4SZ")                     '品ＷＦ面検欠陥４測定条件
            If fldNameExist("HWFMK2ZAR") Then .HWFMK2ZAR = fncNullCheck(rs("HWFMK2ZAR"))    '品ＷＦ面検欠陥２除外領域
            If fldNameExist("HWFMK3ZAR") Then .HWFMK3ZAR = fncNullCheck(rs("HWFMK3ZAR"))    '品ＷＦ面検欠陥３除外領域
            If fldNameExist("HWFMK4ZAR") Then .HWFMK4ZAR = fncNullCheck(rs("HWFMK4ZAR"))    '品ＷＦ面検欠陥４除外領域
            If fldNameExist("HWFMK5B1") Then .HWFMK5B1 = fncNullCheck(rs("HWFMK5B1"))       '品ＷＦ面検欠陥５境界１
            If fldNameExist("HWFMK5B1B") Then .HWFMK5B1B = fncNullCheck(rs("HWFMK5B1B"))    '品ＷＦ面検欠陥５境界１下
            If fldNameExist("HWFMK5B2") Then .HWFMK5B2 = fncNullCheck(rs("HWFMK5B2"))       '品ＷＦ面検欠陥５境界２
            If fldNameExist("HWFMK5B2B") Then .HWFMK5B2B = fncNullCheck(rs("HWFMK5B2B"))    '品ＷＦ面検欠陥５境界２下
            If fldNameExist("HWFMK5B3") Then .HWFMK5B3 = fncNullCheck(rs("HWFMK5B3"))       '品ＷＦ面検欠陥５境界３
            If fldNameExist("HWFMK5B3B") Then .HWFMK5B3B = fncNullCheck(rs("HWFMK5B3B"))    '品ＷＦ面検欠陥５境界３下
            If fldNameExist("HWFMK6B1") Then .HWFMK6B1 = fncNullCheck(rs("HWFMK6B1"))       '品ＷＦ面検欠陥６境界１
            If fldNameExist("HWFMK6B1B") Then .HWFMK6B1B = fncNullCheck(rs("HWFMK6B1B"))    '品ＷＦ面検欠陥６境界１下
            If fldNameExist("HWFMK6B2") Then .HWFMK6B2 = fncNullCheck(rs("HWFMK6B2"))       '品ＷＦ面検欠陥６境界２
            If fldNameExist("HWFMK6B2B") Then .HWFMK6B2B = fncNullCheck(rs("HWFMK6B2B"))    '品ＷＦ面検欠陥６境界２下
            If fldNameExist("HWFMK6B3") Then .HWFMK6B3 = fncNullCheck(rs("HWFMK6B3"))       '品ＷＦ面検欠陥６境界３
            If fldNameExist("HWFMK6B3B") Then .HWFMK6B3B = fncNullCheck(rs("HWFMK6B3B"))    '品ＷＦ面検欠陥６境界３下
            If fldNameExist("HWFMK7MC") Then .HWFMK7MC = HWFMK7MC                           '品ＷＦ面検欠陥７面指定
            If fldNameExist("HWFMK7SI") Then .HWFMK7SI = fncNullCheck(rs("HWFMK7SI"))       '品ＷＦ面検欠陥７サイズ
            If fldNameExist("HWFMK7MX") Then .HWFMK7MX = fncNullCheck(rs("HWFMK7MX"))       '品ＷＦ面検欠陥７上限
            If fldNameExist("HWFMK7SZ") Then .HWFMK7SZ = HWFMK7SZ                           '品ＷＦ面検欠陥７測定条件
            If fldNameExist("HWFMK7ZA") Then .HWFMK7ZA = fncNullCheck(rs("HWFMK7ZA"))       '品ＷＦ面検欠陥７除外領域
            If fldNameExist("HWFMK7HT") Then .HWFMK7HT = HWFMK7HT                           '品ＷＦ面検欠陥７保証方法＿対
            If fldNameExist("HWFMK7HS") Then .HWFMK7HS = HWFMK7HS                           '品ＷＦ面検欠陥７保証方法＿処
            If fldNameExist("HWFMK8MC") Then .HWFMK8MC = HWFMK8MC                           '品ＷＦ面検欠陥８面指定
            If fldNameExist("HWFMK8SI") Then .HWFMK8SI = fncNullCheck(rs("HWFMK8SI"))       '品ＷＦ面検欠陥８サイズ
            If fldNameExist("HWFMK8MX") Then .HWFMK8MX = fncNullCheck(rs("HWFMK8MX"))       '品ＷＦ面検欠陥８上限
            If fldNameExist("HWFMK8SZ") Then .HWFMK8SZ = HWFMK8SZ                           '品ＷＦ面検欠陥８測定条件
            If fldNameExist("HWFMK8ZA") Then .HWFMK8ZA = fncNullCheck(rs("HWFMK8ZA"))       '品ＷＦ面検欠陥８除外領域
            If fldNameExist("HWFMK8HT") Then .HWFMK8HT = HWFMK8HT                           '品ＷＦ面検欠陥８保証方法＿対
            If fldNameExist("HWFMK8HS") Then .HWFMK8HS = HWFMK8HS                           '品ＷＦ面検欠陥８保証方法＿処
            If fldNameExist("HWFMK9MC") Then .HWFMK9MC = HWFMK9MC                           '品ＷＦ面検欠陥９面指定
            If fldNameExist("HWFMK9SI") Then .HWFMK9SI = fncNullCheck(rs("HWFMK9SI"))       '品ＷＦ面検欠陥９サイズ
            If fldNameExist("HWFMK9MX") Then .HWFMK9MX = fncNullCheck(rs("HWFMK9MX"))       '品ＷＦ面検欠陥９上限
            If fldNameExist("HWFMK9SZ") Then .HWFMK9SZ = HWFMK9SZ                           '品ＷＦ面検欠陥９測定条件
            If fldNameExist("HWFMK9ZA") Then .HWFMK9ZA = fncNullCheck(rs("HWFMK9ZA"))       '品ＷＦ面検欠陥９除外領域
            If fldNameExist("HWFMK9HT") Then .HWFMK9HT = HWFMK9HT                           '品ＷＦ面検欠陥９保証方法＿対
            If fldNameExist("HWFMK9HS") Then .HWFMK9HS = HWFMK9HS                           '品ＷＦ面検欠陥９保証方法＿処
            If fldNameExist("HWFMK10MC") Then .HWFMK10MC = HWFMK10MC                        '品ＷＦ面検欠陥１０面指定
            If fldNameExist("HWFMK10SI") Then .HWFMK10SI = fncNullCheck(rs("HWFMK10SI"))    '品ＷＦ面検欠陥１０サイズ
            If fldNameExist("HWFMK10MX") Then .HWFMK10MX = fncNullCheck(rs("HWFMK10MX"))    '品ＷＦ面検欠陥１０上限
            If fldNameExist("HWFMK10SZ") Then .HWFMK10SZ = HWFMK10SZ                        '品ＷＦ面検欠陥１０測定条件
            If fldNameExist("HWFMK10ZA") Then .HWFMK10ZA = fncNullCheck(rs("HWFMK10ZA"))    '品ＷＦ面検欠陥１０除外領域
            If fldNameExist("HWFMK10HT") Then .HWFMK10HT = HWFMK10HT                        '品ＷＦ面検欠陥１０保証方法＿対
            If fldNameExist("HWFMK10HS") Then .HWFMK10HS = HWFMK10HS                        '品ＷＦ面検欠陥１０保証方法＿処
            If fldNameExist("HWFMK11MC") Then .HWFMK11MC = HWFMK11MC                        '品ＷＦ面検欠陥１１面指定
            If fldNameExist("HWFMK11SI") Then .HWFMK11SI = fncNullCheck(rs("HWFMK11SI"))    '品ＷＦ面検欠陥１１サイズ
            If fldNameExist("HWFMK11MX") Then .HWFMK11MX = fncNullCheck(rs("HWFMK11MX"))    '品ＷＦ面検欠陥１１上限
            If fldNameExist("HWFMK11SZ") Then .HWFMK11SZ = HWFMK11SZ                        '品ＷＦ面検欠陥１１測定条件
            If fldNameExist("HWFMK11ZA") Then .HWFMK11ZA = fncNullCheck(rs("HWFMK11ZA"))    '品ＷＦ面検欠陥１１除外領域
            If fldNameExist("HWFMK11HT") Then .HWFMK11HT = HWFMK11HT                        '品ＷＦ面検欠陥１１保証方法＿対
            If fldNameExist("HWFMK11HS") Then .HWFMK11HS = HWFMK11HS                        '品ＷＦ面検欠陥１１保証方法＿処
            If fldNameExist("HWFMK12MC") Then .HWFMK12MC = HWFMK12MC                        '品ＷＦ面検欠陥１２面指定
            If fldNameExist("HWFMK12SI") Then .HWFMK12SI = fncNullCheck(rs("HWFMK12SI"))    '品ＷＦ面検欠陥１２サイズ
            If fldNameExist("HWFMK12MX") Then .HWFMK12MX = fncNullCheck(rs("HWFMK12MX"))    '品ＷＦ面検欠陥１２上限
            If fldNameExist("HWFMK12SZ") Then .HWFMK12SZ = HWFMK12SZ                        '品ＷＦ面検欠陥１２測定条件
            If fldNameExist("HWFMK12ZA") Then .HWFMK12ZA = fncNullCheck(rs("HWFMK12ZA"))    '品ＷＦ面検欠陥１２除外領域
            If fldNameExist("HWFMK12HT") Then .HWFMK12HT = HWFMK12HT                        '品ＷＦ面検欠陥１２保証方法＿対
            If fldNameExist("HWFMK12HS") Then .HWFMK12HS = HWFMK12HS                        '品ＷＦ面検欠陥１２保証方法＿処
            If fldNameExist("HWFMK13MC") Then .HWFMK13MC = HWFMK13MC                        '品ＷＦ面検欠陥１３面指定
            If fldNameExist("HWFMK13SI") Then .HWFMK13SI = fncNullCheck(rs("HWFMK13SI"))    '品ＷＦ面検欠陥１３サイズ
            If fldNameExist("HWFMK13MX") Then .HWFMK13MX = fncNullCheck(rs("HWFMK13MX"))    '品ＷＦ面検欠陥１３上限
            If fldNameExist("HWFMK13SZ") Then .HWFMK13SZ = HWFMK13SZ                        '品ＷＦ面検欠陥１３測定条件
            If fldNameExist("HWFMK13ZA") Then .HWFMK13ZA = fncNullCheck(rs("HWFMK13ZA"))    '品ＷＦ面検欠陥１３除外領域
            If fldNameExist("HWFMK13HT") Then .HWFMK13HT = HWFMK13HT                        '品ＷＦ面検欠陥１３保証方法＿対
            If fldNameExist("HWFMK13HS") Then .HWFMK13HS = HWFMK13HS                        '品ＷＦ面検欠陥１３保証方法＿処
            If fldNameExist("HWFMK14MC") Then .HWFMK14MC = HWFMK14MC                        '品ＷＦ面検欠陥１４面指定
            If fldNameExist("HWFMK14SI") Then .HWFMK14SI = fncNullCheck(rs("HWFMK14SI"))    '品ＷＦ面検欠陥１４サイズ
            If fldNameExist("HWFMK14MX") Then .HWFMK14MX = fncNullCheck(rs("HWFMK14MX"))    '品ＷＦ面検欠陥１４上限
            If fldNameExist("HWFMK14SZ") Then .HWFMK14SZ = HWFMK14SZ                        '品ＷＦ面検欠陥１４測定条件
            If fldNameExist("HWFMK14ZA") Then .HWFMK14ZA = fncNullCheck(rs("HWFMK14ZA"))    '品ＷＦ面検欠陥１４除外領域
            If fldNameExist("HWFMK14HT") Then .HWFMK14HT = HWFMK14HT                        '品ＷＦ面検欠陥１４保証方法＿対
            If fldNameExist("HWFMK14HS") Then .HWFMK14HS = HWFMK14HS                        '品ＷＦ面検欠陥１４保証方法＿処
            If fldNameExist("HWFMK15MC") Then .HWFMK15MC = HWFMK15MC                        '品ＷＦ面検欠陥１５面指定
            If fldNameExist("HWFMK15SI") Then .HWFMK15SI = fncNullCheck(rs("HWFMK15SI"))    '品ＷＦ面検欠陥１５サイズ
            If fldNameExist("HWFMK15MX") Then .HWFMK15MX = fncNullCheck(rs("HWFMK15MX"))    '品ＷＦ面検欠陥１５上限
            If fldNameExist("HWFMK15SZ") Then .HWFMK15SZ = HWFMK15SZ                        '品ＷＦ面検欠陥１５測定条件
            If fldNameExist("HWFMK15ZA") Then .HWFMK15ZA = fncNullCheck(rs("HWFMK15ZA"))    '品ＷＦ面検欠陥１５除外領域
            If fldNameExist("HWFMK15HT") Then .HWFMK15HT = HWFMK15HT                        '品ＷＦ面検欠陥１５保証方法＿対
            If fldNameExist("HWFMK15HS") Then .HWFMK15HS = HWFMK15HS                        '品ＷＦ面検欠陥１５保証方法＿処
            If fldNameExist("HWFSPVMXN") Then .HWFSPVMXN = fncNullCheck(rs("HWFSPVMXN"))    '品ＷＦＳＰＶＦＥ上限
            If fldNameExist("HWFSPVAMN") Then .HWFSPVAMN = fncNullCheck(rs("HWFSPVAMN"))    '品ＷＦＳＰＶＦＥ平均
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME028 = FUNCTION_RETURN_SUCCESS

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

Public Function DBDRV_GetTBCME018(records() As typ_TBCME018, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
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
        
        '2009/08 SUMCO Akizuki 追加
        Case "f_cmbc053_1"           '「Ｘ線測定実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "

        'Add Start 2010/12/17 SMPK Miyata
        Case "f_cmbc054_1"           '「Cu-deco実績入力」
            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HMGSTRRNO, HMGSTFNO, HMGSXSNO, HMGSXSNE, CONFLAG, REINFLAG, HSXTRWKB," & _
              " HSXTYPE, KSXTYPKW, HSXDOP, HSXRMIN, HSXRMAX, HSXRSPOH, HSXRSPOT, HSXRSPOI, HSXRHWYT, HSXRHWYS, HSXRKWAY, HSXRKHNM," & _
              " HSXRKHNI, HSXRKHNH, HSXRKHNS, HSXRMCAL, HSXRMBNP, HSXRMCL2, HSXRMBP2, HSXRSDEV, HSXRAMIN, HSXRAMAX, HSXFORM," & _
              " HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXCDIR, HSXCSCEN, HSXCSMIN, HSXCSMAX, HSXCKWAY," & _
              " HSXCKHNM, HSXCKHNI, HSXCKHNH, HSXCKHNS, HSXCSDIR, HSXCSDIS, HSXCTDIR, HSXCTCEN, HSXCTMIN, HSXCTMAX, HSXCYDIR," & _
              " HSXCYCEN, HSXCYMIN, HSXCYMAX, HSXOF1PD, HSXOF1PN, HSXOF1PX, HSXOF1PW, HSXOF1LC, HSXOF1LN, HSXOF1LX, HSXOF1DC," & _
              " HSXOF1DN, HSXOF1DX, HSXDFORM, HSXDPDRC, HSXDPACN, HSXDPAMN, HSXDPAMX, HSXDPKWY, HSXDPDIR, HSXDPMIN, HSXDPMAX," & _
              " HSXDWCEN, HSXDWMIN, HSXDWMAX, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDACEN, HSXDAMIN, HSXDAMAX, IFKBN, SYORIKBN," & _
              " SPECRRNO, SXLMCNO, WFMCNO, STAFFID, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
        'Add End   2010/12/17 SMPK Miyata

    End Select
    
    sqlBase = sqlBase & "From TBCME018"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                       ' 品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                 ' 製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                    ' 工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                    ' 操業条件
            If fldNameExist("HMGSTRRNO") Then .HMGSTRRNO = rs("HMGSTRRNO")              ' 品管理仕様登録依頼番号
            If fldNameExist("HMGSTFNO") Then .HMGSTFNO = rs("HMGSTFNO")                 ' 品管理社員Ｎｏ
            If fldNameExist("HMGSXSNO") Then .HMGSXSNO = rs("HMGSXSNO")                 ' 品管理ＳＸ製品番号
            If fldNameExist("HMGSXSNE") Then .HMGSXSNE = fncNullCheck(rs("HMGSXSNE"))   ' 品管理ＳＸ製品番号枝番
            If fldNameExist("CONFLAG") Then .CONFLAG = rs("CONFLAG")                    ' 確認フラグ
            If fldNameExist("REINFLAG") Then .REINFLAG = rs("REINFLAG")                 ' 再付与フラグ
            If fldNameExist("HSXTRWKB") Then .HSXTRWKB = rs("HSXTRWKB")                 ' 品ＳＸ統合可否区分
            If fldNameExist("HSXTYPE") Then .HSXTYPE = rs("HSXTYPE")                    ' 品ＳＸタイプ
            If fldNameExist("KSXTYPKW") Then .KSXTYPKW = rs("KSXTYPKW")                 ' 品ＳＸタイプ検査方法
            If fldNameExist("HSXDOP") Then .HSXDOP = rs("HSXDOP")                       ' 品ＳＸドーパント
            If fldNameExist("HSXRMIN") Then .HSXRMIN = fncNullCheck(rs("HSXRMIN"))      ' 品ＳＸ比抵抗下限
            If fldNameExist("HSXRMAX") Then .HSXRMAX = fncNullCheck(rs("HSXRMAX"))      ' 品ＳＸ比抵抗上限
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

'*** UPDATE ↓ Y.SIMIZU 2005/10/1
'概要      :テーブル「TBCME036」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :records()     ,O  ,typ_TBCME036    ,抽出レコード
'          :formID        ,I  ,String          ,使用フォームID
'          :sqlOrder      ,I  ,tFullHinban     ,抽出品番（配列）
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :検索キーは、｢HINBAN｣+「MNOREVNO」+「FACTORY」+「OPECOND」の文字列とする
'履歴      :05/03/01 ooba
Public Function DBDRV_GetTBCME036(records() As typ_TBCME036, formID$, HIN() As tFullHinban) As FUNCTION_RETURN
    Dim sql         As String           'SQL全体
    Dim sqlBase     As String           'SQL基本部(WHERE節の前まで)
    Dim sqlWhere    As String           'SQLWhere部
    Dim rs          As OraDynaset       'RecordSet
    Dim recCnt      As Long             'レコード数
    Dim key         As String           '検索KEY
    Dim i           As Long             'ﾙｰﾌﾟｶｳﾝﾄ

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_kensa_SQL.bas -- Function DBDRV_GetTBCME036"

'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
''    Select Case formID
''        Case "f_cmbc026_1"           '「GD実績入力」
''            sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HSXGDLINE,HWFGDLINE "
''    End Select
    'GD実績入力用のようだが、結晶内側管理に追加された項目はOSF実績入力、総合判定、
    'WFｾﾝﾀｰ総合判定でも使用するので画面指定での検索を無しにする
    sqlBase = "Select HINBAN, MNOREVNO, FACTORY, OPECOND, HSXGDLINE,HWFGDLINE "
    sqlBase = sqlBase & ",HSXLDLRMN,HSXLDLRMX,HWFLDLRMN,HWFLDLRMX "
    sqlBase = sqlBase & ",HSXOF1ARPTK,HSXOFARMIN,HSXOFARMAX,HSXOFARMHMX "
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
    'Add Start 2011/01/27 SMPK Miyata
    sqlBase = sqlBase & ",HSXCJLTBND "
    'Add End   2011/01/27 SMPK Miyata

    sqlBase = sqlBase & "From TBCME036"
    
    '''SQLのWhere文作成
    For i = 0 To UBound(HIN)
        With HIN(i)
            key = key & "'" & .hinban & Format(.mnorevno, "00000") & .factory & .opecond & "'"
            If i <> UBound(HIN) Then
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
        DBDRV_GetTBCME036 = FUNCTION_RETURN_FAILURE
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
            If fldNameExist("HINBAN") Then .hinban = rs("HINBAN")                           '品番
            If fldNameExist("MNOREVNO") Then .mnorevno = rs("MNOREVNO")                     '製品番号改訂番号
            If fldNameExist("FACTORY") Then .factory = rs("FACTORY")                        '工場
            If fldNameExist("OPECOND") Then .opecond = rs("OPECOND")                        '操業条件
            If fldNameExist("HSXGDLINE") Then .HSXGDLINE = fncNullCheck(rs("HSXGDLINE"))    '品管理仕様登録依頼番号
            If fldNameExist("HWFGDLINE") Then .HWFGDLINE = fncNullCheck(rs("HWFGDLINE"))    '品管理社員Ｎｏ
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            If fldNameExist("HSXLDLRMN") Then .HSXLDLRMN = fncNullCheck(rs("HSXLDLRMN"))    '品SXL/DL連続0下限
            If fldNameExist("HSXLDLRMX") Then .HSXLDLRMX = fncNullCheck(rs("HSXLDLRMX"))    '品SXL/DL連続0上限
            If fldNameExist("HWFLDLRMN") Then .HWFLDLRMN = fncNullCheck(rs("HWFLDLRMN"))    '品WFL/DL連続0下限
            If fldNameExist("HWFLDLRMX") Then .HWFLDLRMX = fncNullCheck(rs("HWFLDLRMX"))    '品WFL/DL連続0上限
            If fldNameExist("HSXOF1ARPTK") Then If IsNull(rs("HSXOF1ARPTK")) = False Then .HSXOF1ARPTK = rs("HSXOF1ARPTK")                '品SXOSF1(ArAN)パタン区分
            If fldNameExist("HSXOFARMIN") Then .HSXOFARMIN = fncNullCheck(rs("HSXOFARMIN"))     '品SXOSF(ArAN)下限
            If fldNameExist("HSXOFARMAX") Then .HSXOFARMAX = fncNullCheck(rs("HSXOFARMAX"))     '品SXOSF(ArAN)上限
            If fldNameExist("HSXOFARMHMX") Then .HSXOFARMHMX = fncNullCheck(rs("HSXOFARMHMX"))  '品SXOSF(ArAN)面内比上限
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
            'Add Start 2011/01/27 SMPK Miyata
            If fldNameExist("HSXCJLTBND") Then .HSXCJLTBND = fncNullCheck(rs("HSXCJLTBND"))  '品SXL/CJLTバンド幅
            'Add End   2011/01/27 SMPK Miyata
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME036 = FUNCTION_RETURN_SUCCESS

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
'*** UPDATE ↑ Y.SIMIZU 2005/10/1


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
''Add Start 2011/07/13 LT10Ω換算判定追加 T.Koi(SETsw)
        sql = sql & "CRYREST10CS='" & .CRYREST10CS & "', "        ' 結晶検査実績（LT10)
''Add End   2011/07/13 LT10Ω換算判定追加 T.Koi(SETsw)
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
'          :records()     ,O  ,typ_XSDCS    ,抽出レコード
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
    'Chg Start 2010/12/17 SMPK Miyata Cu-deco検査項目(C,CJ,CJLT,CJ2)追加
    'sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
    '          " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
    '          " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
    '          " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
    '          " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
    '          " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, CRYSMPLIDXCS, CRYINDXCS, CRYRESXCS, SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, BLKKTFLAGCS, " & _
              " CRYSMPLIDRSCS, nvl(CRYSMPLIDRS1CS, 0) as CRYSMPLIDRS1CS, nvl(CRYSMPLIDRS2CS, 0) as CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, nvl(CRYRESRS2CS, ' ') as CRYRESRS2CS, CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, " & _
              " CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, " & _
              " CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, " & _
              " CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, " & _
              " CRYSMPLIDEPCS, CRYINDEPCS, CRYRESEPCS, CRYSMPLIDXCS, CRYINDXCS, CRYRESXCS, " & _
              " CRYSMPLIDCCS, CRYINDCCS, CRYRESCCS, CRYSMPLIDCJCS, CRYINDCJCS, CRYRESCJCS, " & _
              " CRYSMPLIDCJLTCS , CRYINDCJLTCS, CRYRESCJLTCS, CRYSMPLIDCJ2CS, CRYINDCJ2CS, CRYRESCJ2CS, " & _
              " SMPLNUMCS, SMPLPATCS, nvl(TSTAFFCS, ' ') as TSTAFFCS, TDAYCS, nvl(KSTAFFCS, ' ') as KSTAFFCS, KDAYCS, nvl(SNDKCS, ' ') as SNDKCS, nvl(SNDDAYCS, sysdate) as SNDDAYCS "
    'Chg End   2010/12/17 SMPK Miyata
    sqlBase = sqlBase & ",QCKBNCS "
'Add Start 2011/07/13 LT10Ω換算判定 T.Koi(SETsw)
    sqlBase = sqlBase & ",CRYREST10CS "
'Add End   2011/07/13 LT10Ω換算判定 T.Koi(SETsw)
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
            
            ' サンプルID(X線)   2009/08 SUMCO Akizuki ｘ線測定実績入力　項目追加
            If IsNull(rs("CRYSMPLIDXCS")) = True Then
                .CRYSMPLIDXCS = 999999
            Else
                .CRYSMPLIDXCS = rs("CRYSMPLIDXCS")
            End If
            
            ' 状態FLG(X線)      2009/08 SUMCO Akizuki ｘ線測定実績入力　項目追加
            If IsNull(rs("CRYINDXCS")) = True Then
                .CRYINDXCS = "0"
            Else
                .CRYINDXCS = rs("CRYINDXCS")
            End If
            
            ' 実績FLG(X線)      2009/08 SUMCO Akizuki ｘ線測定実績入力　項目追加
            If IsNull(rs("CRYRESXCS")) = True Then
                .CRYRESXCS = "0"
            Else
                .CRYRESXCS = rs("CRYRESXCS")
            End If

            'Add Start 2010/12/17 SMPK Miyata
            If IsNull(rs("CRYSMPLIDCCS")) = False Then .CRYSMPLIDCCS = rs("CRYSMPLIDCCS")           ' サンプルID(C)
            If IsNull(rs("CRYINDCCS")) = False Then .CRYINDCCS = rs("CRYINDCCS")                    ' 状態FLG(C)
            If IsNull(rs("CRYRESCCS")) = False Then .CRYRESCCS = rs("CRYRESCCS")                    ' 実績FLG(C)
            If IsNull(rs("CRYSMPLIDCJCS")) = False Then .CRYSMPLIDCJCS = rs("CRYSMPLIDCJCS")        ' サンプルID(CJ)
            If IsNull(rs("CRYINDCJCS")) = False Then .CRYINDCJCS = rs("CRYINDCJCS")                 ' 状態FLG(CJ)
            If IsNull(rs("CRYRESCJCS")) = False Then .CRYRESCJCS = rs("CRYRESCJCS")                 ' 実績FLG(CJ)
            If IsNull(rs("CRYSMPLIDCJLTCS")) = False Then .CRYSMPLIDCJLTCS = rs("CRYSMPLIDCJLTCS")  ' サンプルID(CJLT)
            If IsNull(rs("CRYINDCJLTCS")) = False Then .CRYINDCJLTCS = rs("CRYINDCJLTCS")           ' 状態FLG(CJLT)
            If IsNull(rs("CRYRESCJLTCS")) = False Then .CRYRESCJLTCS = rs("CRYRESCJLTCS")           ' 実績FLG(CJLT)
            If IsNull(rs("CRYSMPLIDCJ2CS")) = False Then .CRYSMPLIDCJ2CS = rs("CRYSMPLIDCJ2CS")     ' サンプルID(CJ2)
            If IsNull(rs("CRYINDCJ2CS")) = False Then .CRYINDCJ2CS = rs("CRYINDCJ2CS")              ' 状態FLG(CJ2)
            If IsNull(rs("CRYRESCJ2CS")) = False Then .CRYRESCJ2CS = rs("CRYRESCJ2CS")              ' 実績FLG(CJ2)
            'Add End   2010/12/17 SMPK Miyata

            If IsNull(rs("SMPLNUMCS")) = False Then .SMPLNUMCS = rs("SMPLNUMCS")                ' サンプル枚数
            If IsNull(rs("SMPLPATCS")) = False Then .SMPLPATCS = rs("SMPLPATCS")                ' サンプルパターン
            If IsNull(rs("TSTAFFCS")) = False Then .TSTAFFCS = rs("TSTAFFCS")                   ' 登録社員ID
            If IsNull(rs("TDAYCS")) = False Then .TDAYCS = rs("TDAYCS")                         ' 登録日付
            If IsNull(rs("KSTAFFCS")) = False Then .KSTAFFCS = rs("KSTAFFCS")                   ' 更新社員ID
            If IsNull(rs("KDAYCS")) = False Then .KDAYCS = rs("KDAYCS")                         ' 更新日付
            If IsNull(rs("SNDKCS")) = False Then .SNDKCS = rs("SNDKCS")                         ' 送信フラグ
            If IsNull(rs("SNDDAYCS")) = False Then .SNDDAYCS = rs("SNDDAYCS")                   ' 送信日付

            ' 管理区分     2009/11/06追加 SETsw kubota
            If IsNull(rs("QCKBNCS")) = False Then .QCKBNCS = rs("QCKBNCS")

'Add Start 2011/07/13 LT10Ω換算判定 T.Koi(SETsw)
            If IsNull(rs("CRYREST10CS")) = False Then .CRYREST10CS = rs("CRYREST10CS")
'Add End   2011/07/13 LT10Ω換算判定 T.Koi(SETsw)
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
