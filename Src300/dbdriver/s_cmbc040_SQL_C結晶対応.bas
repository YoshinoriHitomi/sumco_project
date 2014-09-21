Attribute VB_Name = "s_cmbc040_SQL"
Option Explicit

Public Type typ_cmlc001e_Disp
    ' SXL管理
    CRYNUM As String * 12           ' 結晶番号
    INGOTPOS As Integer             ' 結晶内開始位置
    SXLID As String * 13            ' 仮SXLID
    hinban As String * 8            ' 品番
    LENGTH As Integer               ' 長さ
    COUNT As Integer                ' 予定枚数
    ENDDATE As Date                 ' 完了日付（SXL管理.登録日付）
    ' WFホールド（解除）実績
    HLDCLASSOLD As String * 1       ' 旧ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
    HLDCLASS As String * 1          ' ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
    HLDDATE As Date                 ' ホールド日付(.登録日付)
    HLDSTAFFNAME As String          ' ホールド担当者(.登録社員ID)
    HLDCAUSE As String * 2          ' ホールド理由 ('SC','17')
    HLDCMNT As String               ' ホールドコメント
    MUKESAKI As String              ' 向先 2007/09/04 SPK Tsutsumi Add
    AGRSTATUS As String             ' 承認確認区分      add SETkimizuka
    STOP    As String               ' 停止 add SETkimizuka
    CAUSE   As String               ' 停止理由 add SETkimizuka
    PRINTNO As String               ' 先行評価 add SETkimizuka
    '■EDI情報ﾘﾝｸ対応 2009/12/4 Add Strat SPK habuki↓↓↓
    EDIFLG As String                ' EDIﾌﾗｸﾞ(△:全件null、OK:送信対象有、NG:送信対象無)
    '■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki↑↑↑
End Type

'2002/09/05 ADD hitec)N.MATSUMOTO Start

'ブロック管理
Public Type typ_cmkc001f_Block
    'E040 ブロック管理
    INGOTPOS As Integer         ' 結晶内開始位置
    LENGTH As Integer           ' 長さ
    REALLEN As Integer          ' 実長さ
    KRPROCCD As String * 5      ' 現在管理工程
    NOWPROC As String * 5       ' 現在工程
    LPKRPROCCD As String * 5    ' 最終通過管理工程
    LASTPASS As String * 5      ' 最終通過工程
    DELCLS As String * 1        ' 削除区分
    RSTATCLS As String * 1      ' 流動状態区分
    LSTATCLS As String * 1      ' 最終状態区分 */
    'E037 結晶情報管理
    SEED As String              'SEED
End Type


'仕様取得用
Public Type typ_cmkc001f_Disp
    '品番管理
    hinban As String * 8              ' 品番
    INGOTPOS As Integer               ' 結晶内開始位置
    REVNUM As Integer                 ' 製品番号改訂番号
    factory As String * 1             ' 工場
    opecond As String * 1             ' 操業条件
    LENGTH As Integer                 ' 長さ
    '製品仕様SXLデータ
    HSXD1CEN As Double                ' 品ＳＸ直径１中心
    HSXRMIN As Double                 ' 品ＳＸ比抵抗下限
    HSXRMAX As Double                 ' 品ＳＸ比抵抗上限
    HSXRMBNP As Double                ' 品ＳＸ比抵抗面内分布
    HSXRHWYS As String * 1            ' 品ＳＸ比抵抗保証方法＿処
    HSXONMIN As Double                ' 品ＳＸ酸素濃度下限
    HSXONMAX As Double                ' 品ＳＸ酸素濃度上限
    HSXONMBP As Double                ' 品ＳＸ酸素濃度面内分布
    HSXONHWS As String * 1            ' 品ＳＸ酸素濃度保証方法＿処
    HSXCNMIN As Double                ' 品ＳＸ炭素濃度下限
    HSXCNMAX As Double                ' 品ＳＸ炭素濃度上限
    HSXCNHWS As String * 1            ' 品ＳＸ炭素濃度保証方法＿処
    HSXTMMAX As Double                ' 品ＳＸ転位密度上限          項目追加，修正対応 2003.05.20 yakimura
    HSXBMnAN(1 To 3) As Double        ' 品ＳＸＢＭＤn 平均下限
    HSXBMnAX(1 To 3) As Double        ' 品ＳＸＢＭＤn 平均上限
    HSXBMnHS(1 To 3) As String * 1    ' 品ＳＸＢＭＤn 保証方法＿処
    HSXOFnAX(1 To 4) As Double        ' 品ＳＸＯＳＦn平均上限
    HSXOFnMX(1 To 4) As Double        ' 品ＳＸＯＳＦn上限
    HSXOFnHS(1 To 4) As String * 1    ' 品ＳＸＯＳＦn 保証方法＿処
    HSXDENMX As Integer               ' 品ＳＸＤｅｎ上限
    HSXDENMN As Integer               ' 品ＳＸＤｅｎ下限
    HSXDENHS As String * 1            ' 品ＳＸＤｅｎ保証方法＿処
    HSXDVDMX As Integer               ' 品ＳＸＤＶＤ２上限
    HSXDVDMN As Integer               ' 品ＳＸＤＶＤ２下限
    HSXDVDHS As String * 1            ' 品ＳＸＤＶＤ２保証方法＿処
    HSXLDLMX As Integer               ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN As Integer               ' 品ＳＸＬ／ＤＬ下限
    HSXLDLHS As String * 1            ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXLTMIN As Integer               ' 品ＳＸＬタイム下限
    HSXLTMAX As Integer               ' 品ＳＸＬタイム上限
    HSXLTHWS As String * 1            ' 品ＳＸＬタイム保証方法＿処
    HSXDPDIR As String * 2            ' 品ＳＸ溝位置方位
    HSXDPDRC As String * 1            ' 品ＳＸ溝位置方向
    HSXDWMIN As Double                ' 品ＳＸ溝巾下限
    HSXDWMAX As Double                ' 品ＳＸ溝巾上限
    HSXDDMIN As Double                ' 品ＳＸ溝深下限
    HSXDDMAX As Double                ' 品ＳＸ溝深上限
    HSXD1MIN As Double                ' 品ＳＸ直径１下限
    HSXD1MAX As Double                ' 品ＳＸ直径１上限
    HSXCTCEN As Double                ' 品ＳＸ結晶面傾縦中心
    HSXCYCEN As Double                ' 品ＳＸ結晶面傾横中心
    EPDUP As Integer                  ' 結晶内側管理 EPD　上限
End Type

Private Type tCsData
    'Cs実績
    SXL_CS_SMPPOS As Integer        ' 測定位置
    SXLCS_CSMEAS As Double          ' Cs値
    SXLCS_70PPRE As Double          ' 70%推定値
End Type

Public strBlockID()    As String
Public sProSXLID       As String    '処理SXLID　06/03/24 ooba

Public Const PROCD_WFC_SAINUKISI = "CW760"              ' WFセンター再抜試
Public Const PROCD_SXL_MAP = "TX860"                    ' シングルマップ
'2002/09/05 ADD hitec)N.MATSUMOTO end

''***********************************************************
'Micron対応 2011/01/14 Add start tkimura
Public Type typ_Y011
    LOTID As String          '' ブロックID
    BLOCKSEQ As Integer      '' ブロック内連番
    RITOP_POS As String       '' 推定位置
End Type
'Micron対応 2011/01/14 Add end tkimura
''***********************************************************

Public Function DBDRV_fcmlc001e_Disp(records() As typ_cmlc001e_Disp) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset
    Dim i As Integer
    Dim j As Integer    ' 2007/09/13 SPK Tsutsumi Add
    
    '2004/07/15 koyama
    Dim sWFHOLDDATE   As String
    Dim sUSER_ID As String
    
    Dim sOldID      As String   '09/03/17 add SETkimizuka
    Dim iCnt        As Integer  '09/03/17 add SETkimizuka

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function DBDRV_fcmlc001e_Disp"

''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   WFS.XTALCB as CRYNUM"                               ''結晶番号
    sql = sql & "  ,WFS.INPOSCB as INGOTPOS"                            ''結晶内開始位置
    sql = sql & "  ,WFS.SXLIDCB as SXLID"                               ''SXLID
    sql = sql & "  ,WFS.HINBCB as HINBAN"                               ''品番
    sql = sql & "  ,WFS.RLENCB as LENGTH"                               ''理論長さ
    sql = sql & "  ,WFS.maicb as COUNT"                                 ''実枚数
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka Start  09/03/17
    'sql = sql & "  ,NVL(HLD.HLDCLASS,'0') as HLDCLASS"
    'sql = sql & "  ,NVL(HLD.REGDATE,SYSDATE) as HLDDATE"
    'sql = sql & "  ,NVL(HLD.STAFFNAME,' ') as HLDSTAFFNAME"
    'sql = sql & "  ,NVL(HLD.HLDCAUSE,' ') as HLDCAUSE"
    'sql = sql & "  ,NVL(HLD.HLDCMNT,' ') as HLDCMNT"
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka End  09/03/17
    sql = sql & "  ,NVL(PASS.REGDATE,SYSDATE) as ENDDATE"
    sql = sql & "  ,WFS.PLANTCATCB as PLANTCAT"                         ''向先 2007/09/04 SPK Tsutsumi Add
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka Start  09/03/17
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
    sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUSY4"
    sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSEY4"
    sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNOY4"
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/30
    sql = sql & "  ,NVL(Y4_DISPHLD.STOPY4,'0') as HLDCLASS"
    sql = sql & "  ,NVL(Y4_DISPHLD.SETTDAYY4,SYSDATE) as HLDDATE"
    sql = sql & "  ,NVL(Y4_DISPHLD.STAFFNAME,' ') as HLDSTAFFNAME"
    sql = sql & "  ,NVL(Y4_DISPHLD.CAUSEY4,' ') as HLDCAUSE"
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Strat SPK habuki↓↓↓
    sql = sql & "  ,case NVL(EDI.EDIFLG,'@')"
    sql = sql & "     when '2' then 'OK'"
    sql = sql & "     when '1' then 'NG'"
    sql = sql & "     when '0' then 'NG'"
    '2010/1/22 Null時ロックするよう改修 SPK Hitomi
    sql = sql & "     when '@' then 'NG'"
    sql = sql & "     else '  '"
    sql = sql & "   end  as EDIFLG "
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki↑↑↑
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka End  09/03/17
    sql = sql & " FROM"
    sql = sql & "   XSDCB WFS"
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka Start  09/03/17
    'sql = sql & "  ,("
    'sql = sql & "    SELECT"
    'sql = sql & "      DAT.SNGLID"
    'sql = sql & "     ,DAT.HLDCLASS"
    'sql = sql & "     ,DAT.REGDATE"
    'sql = sql & "     ,DAT.TSTAFFID"
    'sql = sql & "     ,DAT.HLDCAUSE"
    'sql = sql & "     ,DAT.HLDCMNT"
    'sql = sql & "     ,rtrim(STAFF.JFMLNAME)||rtrim(STAFF.JFSTNAME) as STAFFNAME"
    'sql = sql & "    FROM"
    'sql = sql & "      TBCMW008 DAT"
    'sql = sql & "     ,("
    'sql = sql & "       SELECT"
    'sql = sql & "         SNGLID"
    'sql = sql & "        ,MAX(TRANCNT) as MAX_TRANCNT"
    'sql = sql & "       FROM"
    'sql = sql & "         TBCMW008"
    'sql = sql & "       GROUP BY"
    'sql = sql & "         SNGLID"
    'sql = sql & "      ) W008"
    'sql = sql & "     ,TBCMB001 STAFF"
    'sql = sql & "    WHERE DAT.SNGLID   = W008.SNGLID"
    'sql = sql & "      AND DAT.TRANCNT  = W008.MAX_TRANCNT"
    'sql = sql & "      AND DAT.TSTAFFID = STAFF.STAFFID"
    'sql = sql & "   ) HLD"
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka End  09/03/17
    ' 処理速度改善対応  Add Y.Hitomi 09/12/4
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      *"
    sql = sql & "    FROM"
    sql = sql & "    ("
    sql = sql & "        SELECT"
    sql = sql & "             SXLID,"
    sql = sql & "             TRANCNT,"
    sql = sql & "             REGDATE,"
    sql = sql & "             rank() over(partition by SXLID order by TRANCNT desc ) as RANK"
    sql = sql & "         FROM"
    sql = sql & "             TBCMW005"
    sql = sql & "    )"
    sql = sql & "    WHERE"
    sql = sql & "         RANK = 1"
    sql = sql & "   ) PASS"
'処理速度改善対応  del Y.Hitomi 09/12/4
'    sql = sql & "    SELECT"
'    sql = sql & "      DAT.SXLID"
'    sql = sql & "     ,DAT.REGDATE"
'    sql = sql & "    FROM"
'    sql = sql & "      TBCMW005 DAT"
'    sql = sql & "     ,("
'    sql = sql & "       SELECT"
'    sql = sql & "         SXLID"
'    sql = sql & "        ,max(TRANCNT) as MAX_TRANCNT"
'    sql = sql & "       FROM"
'    sql = sql & "         TBCMW005"
'    sql = sql & "       GROUP BY"
'    sql = sql & "         SXLID"
'    sql = sql & "      ) W005"
'    sql = sql & "    WHERE DAT.SXLID    = W005.SXLID"
'    sql = sql & "      AND DAT.TRANCNT  = W005.MAX_TRANCNT"
'    sql = sql & "   ) PASS"

    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
    ' 流動停止項目追加 add SETkimizuka Start  09/03/17
'    sql = sql & "    ,( "
'    sql = sql & "      SELECT SXLIDY3 as SXLIDY4 ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
'    sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
'    sql = sql & "      FROM XSDCB,XODY3  "
'    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 ='CW000')"
''    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = SXLIDY3 AND GNWKNTCB    = 'CW800' AND  "
'    sql = sql & "       LIVKY3    = '0'  "
'    sql = sql & "       GROUP BY SXLIDY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9 "
'
'    sql = sql & "      UNION SELECT SXLIDY3 as SXLIDY4 ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
'    sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
'    sql = sql & "      FROM XSDCB,XODY3  "
'    sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND SXLIDY3 = SXLIDY4 AND STOPY4 <> '2' AND LIVKY4 = '0' AND WKKTY4 = 'CW800')"
'    sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
'    sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
'    sql = sql & "                FROM XODY3,XODY4,XODY5 "
'    sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
'    sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
'    sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
'    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = SXLIDY3 AND GNWKNTCB    = 'CW800' AND  "
'    sql = sql & "       LIVKY3    = '0'  "
'    sql = sql & "       GROUP BY SXLIDY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9 "
'    sql = sql & "       ) Y4 "
    sql = sql & "    ,( "
    sql = sql & "      SELECT SXLIDY3 as SXLIDY4 ,DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4) as AGRSTATUS  "
    sql = sql & "      ,STOPY4 as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y4.PRINTNOY4 as PRINTNO,Y4.PRINTKINDY4 as PRINTKIND"
    sql = sql & "      ,WKKTY4 "
    sql = sql & "      FROM XSDCB,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = Y3.SXLIDY3 AND GNWKNTCB    = 'CW800' "
    sql = sql & "       AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & "       AND Y3.SXLIDY3 = Y4.SXLIDY4(+) "
    sql = sql & "       AND Y3.LIVKY3(+) = '0' "
    sql = sql & "       AND Y4.LIVKY4(+) = '0' "
    sql = sql & "       AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    sql = sql & " UNION SELECT SXLIDY3 as SXLIDY4 ,DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4) as AGRSTATUS  "
    sql = sql & "      ,STOPY4 as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y4.PRINTNOY4 as PRINTNO,Y4.PRINTKINDY4 as PRINTKIND"
    sql = sql & "      ,WKKTY4 "
    sql = sql & "      FROM XSDCB,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "      WHERE LIVKCB     <>'1' AND SXLIDCB = Y3.SXLIDY3 AND GNWKNTCB    = 'CW800' "
    sql = sql & "       AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & "       AND Y4.WKKTY4(+) = 'CW000' "
    sql = sql & "       AND Y3.LIVKY3(+) = '0' "
    sql = sql & "       AND Y4.LIVKY4(+) = '0' "
    sql = sql & "       AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    sql = sql & "       ) Y4 "
    ' 流動停止項目追加 add SETkimizuka End  09/03/17
    sql = sql & "    ,(SELECT SXLIDY4,SETTDAYY4,NAMEJA9 as STAFFNAME,CAUSEY4,STOPY4 FROM XODY4,KODA9 "
    sql = sql & "       WHERE STOPY4 <> '2' AND WKKTY4 = '" & DISP_HOLD & "'"
    sql = sql & "         AND SYSCA9(+) = 'K' AND SHUCA9(+) = '55' AND CODEA9(+) = SETSTAFFIDY4 ) Y4_DISPHLD "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/30
    
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Strat SPK habuki↓↓↓
    '/* EDIﾌﾗｸﾞの設定状況確認 */
    sql = sql & "    ,( "
    sql = sql & "      select "
    sql = sql & "          substr(c61.XTALC6,1,7) XTAL"
    sql = sql & "         ,max(c61.EDIFLGC6)      EDIFLG"
    sql = sql & "      from "
    sql = sql & "          XODC6_1  c61"
    sql = sql & "         ,TBCMH001 t01"
    sql = sql & "      where"
    sql = sql & "           substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)"
    sql = sql & "       and t01.CODE > '4'"
    sql = sql & "      group by"
    sql = sql & "         substr(c61.XTALC6,1,7)"
    sql = sql & "     ) EDI "
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki↑↑↑

    sql = sql & " WHERE WFS.LIVKCB     <>'1'"
    sql = sql & "   AND WFS.GNWKNTCB    = 'CW800'"
    'sql = sql & "   AND WFS.SXLIDCB     = HLD.SNGLID(+)"           'del 09/03/17 SETkimizuka
    sql = sql & "   AND WFS.SXLIDCB     = PASS.SXLID(+)"
    sql = sql & "   AND WFS.SXLIDCB     = Y4.SXLIDY4(+)"            'add 09/03/17 SETkimizuka
    sql = sql & "   AND WFS.SXLIDCB     = Y4_DISPHLD.SXLIDY4(+)"    'add 09/03/17 SETkimizuka
    
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Strat SPK habuki↓↓↓
    sql = sql & "   AND substr(WFS.SXLIDCB,1,7) = EDI.XTAL(+)"
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki↑↑↑
    
    ' 向先 2007/09/04 SPK Tsutsumi Add Start
    If sCmbMukesaki <> "ALL" Then
        sql = sql & "   AND WFS.PLANTCATCB      = '" & sCmbMukesaki & "'"
    End If
    ' 2007/09/04 SPK Tsutsumi Add End
    
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka Start  09/03/17
    sql = sql & "   ORDER BY NVL(Y4_DISPHLD.STOPY4,0),WFS.SXLIDCB"
    'sql = sql & "   ORDER BY HLD.HLDCLASS,WFS.SXLIDCB"
    ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka End  09/03/17
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'    sql = sql & "select  SXL.CRYNUM, SXL.INGOTPOS, SXL.SXLID, SXL.HINBAN, SXL.LENGTH, WFS.MAICB as COUNT , "
'    sql = sql & "  nvl(HLD.HLDCLASS,'0') as HLDCLASS, nvl(HLD.REGDATE,SYSDATE) as HLDDATE, "
'    sql = sql & "  nvl(HLD.STAFFNAME,' ') as HLDSTAFFNAME, nvl(HLD.HLDCAUSE,' ') as HLDCAUSE, "
'    sql = sql & "  nvl(HLD.HLDCMNT,' ') as HLDCMNT,  nvl(PASS.REGDATE,SYSDATE) as ENDDATE  "
'    sql = sql & "from TBCME042 SXL,XSDCB WFS, "
'    sql = sql & "  (select DAT.SNGLID, DAT.HLDCLASS, DAT.REGDATE, DAT.TSTAFFID, DAT.HLDCAUSE, DAT.HLDCMNT, "
'    sql = sql & "     rtrim(STAFF.JFMLNAME)||rtrim(STAFF.JFSTNAME) as STAFFNAME "
'    sql = sql & "   from TBCMW008 DAT,  "
'    sql = sql & "     (select SNGLID, MAX(TRANCNT) as MAX_TRANCNT from TBCMW008 group by SNGLID) W008, "
'    sql = sql & "     TBCMB001 STAFF "
'    sql = sql & "   Where (DAT.SNGLID = W008.SNGLID) and (DAT.TRANCNT = W008.MAX_TRANCNT) and (DAT.TSTAFFID = STAFF.STAFFID) "
'    sql = sql & "  ) HLD, "
'    sql = sql & "  (select  DAT.SXLID, DAT.REGDATE "
'    sql = sql & "   from  TBCMW005 DAT, "
'    sql = sql & "     (select SXLID, max(TRANCNT) as MAX_TRANCNT from TBCMW005 group by SXLID) W005 "
'    sql = sql & "   where (DAT.SXLID = W005.SXLID) and (DAT.TRANCNT = W005.MAX_TRANCNT) "
'    sql = sql & "  ) PASS "
'    sql = sql & " where (SXL.DELCLS<>'1') "
'    sql = sql & " and (SXL.NOWPROC='CW800') "
'    sql = sql & " and (SXL.SXLID = HLD.SNGLID(+)) "
'    sql = sql & " and (SXL.SXLID = PASS.SXLID(+)) "
'    sql = sql & " and (SXL.SXLID = WFS.SXLIDCB(+)) "
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    
    ReDim records(0)
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'If rs.RecordCount = 0 Then
    '    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_FAILURE
    '    rs.Close
    '    GoTo PROC_EXIT
    'End If
    
    'ReDim records(rs.RecordCount)  'del 09/03/17 SETkimizuka
    iCnt = 0
    For i = 1 To rs.RecordCount
        If sOldID <> rs("SXLID") Then  'add 09/03/17 SETkimizuka
            iCnt = iCnt + 1        'add 09/03/17 SETkimizuka
            ReDim Preserve records(iCnt)
            With records(iCnt)
                ' SXL管理
                .CRYNUM = rs("CRYNUM")          ' 結晶番号
                .INGOTPOS = rs("INGOTPOS")      ' 結晶内位置
                .SXLID = rs("SXLID")            ' 仮SXLID
                .hinban = rs("HINBAN")          ' 品番
                .LENGTH = rs("LENGTH")          ' 長さ
                .COUNT = rs("COUNT")            ' 予定枚数
                .ENDDATE = rs("ENDDATE")        ' 完了日付（SXL管理.登録日付）
                
                ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
                '.AGRSTATUS = rs("AGRSTATUSY4")               ' 流動停止区分  add 09/03/17 SETkimizuka
                '.STOP = rs("STOP")               ' 流動停止区分  add 09/03/17 SETkimizuka
                'If Trim(rs("CAUSEY4")) <> "" Then
                '    .CAUSE = rs("CAUSEY4") & vbTab      ' 流動停止理由  add 09/03/17 SETkimizuka
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW800" Or rs("WKKTY4") = "CW000") Then
                    .AGRSTATUS = rs("AGRSTATUSY4")               ' 流動停止区分
                    .STOP = rs("STOP")               ' 流動停止区分
                    If Trim(rs("CAUSEY4")) <> "" Then
                        .CAUSE = rs("CAUSEY4") & vbTab      ' 流動停止理由  add 09/03/17 SETkimizuka
                    End If
                Else
                    .STOP = "0"               ' 流動停止区分
                End If
                ' 流動監視SQL修正 upd SETkimizuka End  09/06/30
                
                If Trim(rs("PRINTNOY4")) <> "" Then
                    .PRINTNO = rs("PRINTNOY4") & vbTab  ' 先行評価No    add 09/03/17 SETkimizuka
                End If
                
                ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka Start  09/03/17
                .HLDCLASSOLD = rs("HLDCLASS")
                .HLDCLASS = rs("HLDCLASS")
                .HLDDATE = rs("HLDDATE")
                .HLDCAUSE = rs("HLDCAUSE")
                .HLDSTAFFNAME = rs("HLDSTAFFNAME")
                ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka End  09/03/17
                
                '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Strat SPK habuki↓↓↓
                .EDIFLG = rs("EDIFLG")
                '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki↑↑↑
                
                ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka Start  09/03/17
                '' WFホールド（解除）実績
                'If (CStr(rs("HLDCLASS")) = "1") Then
                '    .HLDCLASSOLD = rs("HLDCLASS")       ' 旧ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '    .HLDCLASS = rs("HLDCLASS")          ' ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '    .HLDDATE = rs("HLDDATE")            ' ホールド日付(.登録日付)
                '    .HLDSTAFFNAME = rs("HLDSTAFFNAME")  ' ホールド担当者名(.登録社員ID)
                '    .HLDCAUSE = rs("HLDCAUSE")          ' ホールド理由 ('SC','17')
                '    .HLDCMNT = rs("HLDCMNT")            ' ホールドコメント
                'Else
                '
                '    'WFﾎｰﾙﾄﾞ以外の表示処理　2004/07/15 koyama
                '    If DBDRV_s_cmbc040_SQL_Y019XSDCB(rs("SXLID"), rs("CRYNUM"), rs("HINBAN"), _
                                                     rs("INGOTPOS"), sWFHOLDDATE, sUSER_ID _
                                                    ) = FUNCTION_RETURN_FAILURE Then
    
    '                   .HLDCLASSOLD = vbNullString         ' 旧ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '        .HLDCLASSOLD = "0"                  ' 旧ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '        .HLDCLASS = vbNullString            ' ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
    '                    .HLDDATE = " "                      ' ホールド処理日
                '        .HLDSTAFFNAME = vbNullString        ' ホールド担当者名(.登録社員ID)
                '        .HLDCAUSE = vbNullString            ' ホールド理由 ('SC','17')
                '        .HLDCMNT = vbNullString             ' ホールドコメント
    
                '    Else
                '
                '        If sWFHOLDDATE <> "" And IsNull(sWFHOLDDATE) = False Then
                '            .HLDDATE = sWFHOLDDATE      ' ホールド日付(.登録日付)
                '       End If
                '        .HLDSTAFFNAME = sUSER_ID    ' ホールド担当者名(.登録社員ID)
                '
                '        '2004/07/21 koyama
                '        .HLDCLASSOLD = "0"                  ' 旧ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '        .HLDCLASS = vbNullString            ' ホールド処理区分 (0:ホールド解除処理、1:ホールド処理)
                '        .HLDCAUSE = vbNullString            ' ホールド理由 ('SC','17')
                '       .HLDCMNT = vbNullString             ' ホールドコメント
                '    End If
                '
                'End If
                ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka End  09/03/17
                
                ' 2007/09/13 SPK Tsutsumi Add Start
                If IsNull(rs("PLANTCAT")) = False Then
                    For j = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs("PLANTCAT") Then
                            .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                ' 2007/09/13 SPK Tsutsumi Add End
                
                sOldID = rs("SXLID")   'add 09/03/17 SETkimizuka
                rs.MoveNext
            End With
        Else
            ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
            '' 流動停止理由  add 09/03/17 SETkimizuka
            'If InStr(records(iCnt).CAUSE, rs("CAUSEY4")) = 0 And Trim(rs("CAUSEY4")) <> "" Then
            '    records(iCnt).CAUSE = records(iCnt).CAUSE & rs("CAUSEY4") & vbTab
            'End If
            If rs("STOP") <> "2" And (rs("WKKTY4") = "CW800" Or rs("WKKTY4") = "CW000") Then
                If Trim(records(iCnt).AGRSTATUS) = "" Or (rs("AGRSTATUSY4") < records(iCnt).AGRSTATUS) Then
                    records(iCnt).AGRSTATUS = rs("AGRSTATUSY4")               ' 承認確認区分
                    records(iCnt).STOP = rs("STOP")               ' 流動停止区分
                End If
                If InStr(records(iCnt).CAUSE, rs("CAUSEY4")) = 0 And Trim(rs("CAUSEY4")) <> "" Then
                    records(iCnt).CAUSE = records(iCnt).CAUSE & rs("CAUSEY4") & vbTab
                End If
            End If
            ' 流動監視SQL修正 upd SETkimizuka Start  09/06/30
            ' 先行評価No    add 09/03/17 SETkimizuka
            If InStr(records(iCnt).PRINTNO, rs("PRINTNOY4")) = 0 And Trim(rs("PRINTNOY4")) <> "" Then
                records(iCnt).PRINTNO = records(iCnt).PRINTNO & rs("PRINTNOY4") & vbTab
            End If
            rs.MoveNext
        End If
    Next
    rs.Close
    
    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmlc001e_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'Public Function DBDRV_fcmlc001e_Exec(records() As typ_cmlc001e_Disp, StaffID$, errTbl$) As FUNCTION_RETURN
'records()→record　06/10/20 ooba
Public Function DBDRV_fcmlc001e_Exec(record As typ_cmlc001e_Disp, StaffID$, errTbl$) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset
Dim i As Integer
Dim smpId(2) As String
Dim blkID As String
Dim errmsg As String
Dim typXODY3()  As typ_XODY3    'add 09/03/17 SETkimizuka
Dim typXODY4()  As typ_XODY4    'add 09/03/17 SETkimizuka
Dim tXODY4      As typ_XODY4    'add 09/03/17 SETkimizuka
Dim IRow        As Integer      'add 09/03/17 SETkimizuka

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function DBDRV_fcmlc001e_Exec"

    ''画面の行毎に処理
'    For i = 1 To UBound(records)
'        With records(i)
        With record     '06/10/20 ooba
        
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    ''XSDCB.ホールド区分の更新を行なう。
            sql = " UPDATE XSDCB SET" & _
                  "  KDAYCB = SYSDATE"
            If .HLDCLASS = "0" Then
                sql = sql & " ,SHOLDCLSCB = '0' "
            Else
                sql = sql & " ,SHOLDCLSCB = '1' "
            End If
            sql = sql & " WHERE SXLIDCB = '" & .SXLID & "'"
            If 0 >= OraDB.ExecuteSQL(sql) Then
                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                errTbl = "XSDCB"
                GoTo proc_exit
            End If
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'            ''SXL管理情報の更新
'            sql = "update TBCME042 set " & _
'                  "KRPROCCD = '" & MGPRCD_SXL_KAKUTEI & "', " & _
'                  "NOWPROC = '" & PROCD_SXL_KAKUTEI & "', " & _
'                  "LPKRPROCCD = '" & MGPRCD_SXL_KAKUTEI & "', " & _
'                  "LASTPASS = '" & PROCD_SXL_KAKUTEI & "', " & _
'                  "UPDDATE = SYSDATE, "
'            If .HLDCLASS = "0" Then     'SXL確定
'                sql = sql & "HOLDCLS = '0', " & _
'                    "LSTATCLS='S', " & _
'                    "SENDFLAG='3', " & _
'                    "DELCLS='1' "
'            Else                        'WFホールド
'                sql = sql & "HOLDCLS = '1', " & _
'                    "SENDFLAG = '0' "
'            End If
'            sql = sql & "where (SXLID='" & .SXLID & "')"
'            If 0 >= OraDB.ExecuteSQL(sql) Then
'                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
'                errTbl = "E042"
'                GoTo proc_exit
'            End If
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
            
            
            ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka Start  09/03/17
            '''WFホールド実績の追加
            'If .HLDCLASS <> .HLDCLASSOLD Then
            '    sql = "insert into TBCMW008 (" & _
            '          " CRYNUM, INGOTPOS, TRANCNT," & _
            '          " CRYLEN, KRPROCCD, PROCCODE," & _
            '          " SNGLID, HLDCLASS, HLDCAUSE, HLDCMNT," & _
            '          " TSTAFFID , REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
            '          ") select " & _
            '          "'" & .CRYNUM & "', " & .INGOTPOS & ", nvl(max(TRANCNT),0)+1, " & _
            '          .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
            '          "'" & .SXLID & "', '" & .HLDCLASS & "', " & NoNullStr(.HLDCAUSE) & ", " & NoNullStr(.HLDCMNT) & ", " & _
            '          "'" & STAFFID & "', SYSDATE, ' ', SYSDATE, '0', SYSDATE " & _
            '          "From TBCMW008 " & _
            '          "where (CRYNUM='" & .CRYNUM & "') and (INGOTPOS=" & .INGOTPOS & ") "
            '    If 0 >= OraDB.ExecuteSQL(sql) Then
            '        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
            '        errTbl = "W008"
            '        GoTo proc_exit
            '    End If
            'End If
            ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え del SETkimizuka End  09/03/17
            
            ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka Start  09/03/17
            If .HLDCLASS <> .HLDCLASSOLD Then
                Call GetSysdate
                ReDim typXODY4(1)
                If .HLDCLASS = "1" Then
                    If GetXODY3(typXODY3, "WHERE SXLIDY3 ='" & .SXLID & "' AND LIVKY3 = '0' ") = True Then
                        For IRow = 1 To UBound(typXODY3)
                            typXODY4(1).AGRSTATUSY4 = VB_KBN
                            typXODY4(1).CAUSEY4 = .HLDCAUSE
                            typXODY4(1).LIVKY4 = "0"
                            typXODY4(1).SETSTAFFIDY4 = StaffID
                            typXODY4(1).WKKTY4 = DISP_HOLD
                            typXODY4(1).STOPY4 = STOP_KBN
                            typXODY4(1).XTALNOY4 = typXODY3(IRow).XTALNOY3
                            typXODY4(1).RCNTY4 = typXODY3(IRow).RCNTY3
                            typXODY4(1).SXLIDY4 = typXODY3(IRow).SXLIDY3
                            If InsertXODY4(typXODY4) = False Then
                                DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                                errTbl = "XODY4"
                                GoTo proc_exit
                            End If
                        Next
                    End If
                Else
                    tXODY4.STOPY4 = KAIJO_KBN
                    tXODY4.KSTAFFIDY4 = StaffID
                    tXODY4.KDAYY4 = gsSysdate
                    If UpdateXODY4(tXODY4, _
                        "WHERE SXLIDY4 ='" & .SXLID & "' AND STOPY4 = '" & STOP_KBN & "' " _
                        & "AND WKKTY4 = '" & DISP_HOLD & "'") = False Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "XODY4"
                        GoTo proc_exit
                    End If
                End If
                
            End If
            ' 既存ﾎｰﾙﾄﾞを流動監視データへ置き換え add SETkimizuka End  09/03/17
            
            ''SXL確定実績の追加
            If .HLDCLASS = "0" Then
                smpId(1) = vbNullString
                smpId(2) = vbNullString
                blkID = vbNullString
                
                'サンプルIDを得る
 '               sql = "select E044SMPLID from VECME011 where (E042SXLID='" & .SXLID & "') order by E044INGOTPOS"
 '　　　　　　　 サンプル管理としてサンプル指示が全てないもの確定９のものは除く
                ''　ブロックIDセット 2003/09/17 Motegi ===========================================> START
                '' サンプルID(From)、サンプルID(To)が見つからなければ、ブロックIDをセットし、
                '' サンプルID(From)、サンプルID(To)が見つかれば、ブロックIDをセットしない。
 
'                sql = "select E044SMPLID from VECME011 where (E042SXLID='" & .SXLID & "' and E044KTKBN != '9') order by E044INGOTPOS"
'                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'                If rs.RecordCount = 2 Then
'                    smpId(1) = rs("E044SMPLID")
'                    rs.MoveNext
'                    smpId(2) = rs("E044SMPLID")
'                    rs.Close
'                    Set rs = Nothing
'                Else
'                    rs.Close
'
'                    'ブロックIDを得る
'                    sql = "select BLOCKID from TBCME040 " & _
'                          "where (crynum='" & .CRYNUM & "') and (INGOTPOS<=" & .INGOTPOS & ") and (" & .INGOTPOS & "<INGOTPOS+LENGTH)"
'                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'                    If rs.RecordCount = 1 Then
'                        blkID = rs("BLOCKID")
'                    End If
'                    rs.Close
'                    Set rs = Nothing
'                End If
                ''-----------------------------------------------------------------------------
                sql = "select REPSMPLIDCW from XSDCW where SXLIDCW='" & .SXLID & "' and KTKBNCW != '9' order by INPOSCW"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 2 Then
                    smpId(1) = rs("REPSMPLIDCW")
                    rs.MoveNext
                    smpId(2) = rs("REPSMPLIDCW")
                    rs.Close
                    Set rs = Nothing
                
                    '共有サンプルチェック処理
                    If chkComSAMPL(.SXLID, smpId(1), smpId(1)) Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "DBDRV_fcmlc001e_Exec:共有ｻﾝﾌﾟﾙID取得(From)"
                        GoTo proc_exit
                    End If
                    If chkComSAMPL(.SXLID, smpId(2), smpId(2)) Then
                        DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                        errTbl = "DBDRV_fcmlc001e_Exec:共有ｻﾝﾌﾟﾙID取得(To)"
                        GoTo proc_exit
                    End If
                Else
                    rs.Close

                    'ブロックIDを得る
                    sql = "select BLOCKID from TBCME040 " & _
                          "where (crynum='" & .CRYNUM & "') and (INGOTPOS<=" & .INGOTPOS & ") and (" & .INGOTPOS & "<INGOTPOS+LENGTH)"
                    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                    If rs.RecordCount = 1 Then
                        blkID = rs("BLOCKID")
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
                ''　ブロックIDセット 2003/09/17 Motegi ===========================================> END
                
'                '確定実績を書き込む
' 2007/09/04 SPK Tsutsumi Add Start
                sql = "insert into TBCMW007 (" & _
                      "CRYNUM, INGOTPOS, " & _
                      "CRYLEN, KRPROCCD, PROCCODE, " & _
                      "SXLID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, " & _
                      "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
                      ",PLANTCAT  " & _
                      ") values (" & _
                      NoNullStr(.CRYNUM) & ", " & .INGOTPOS & ", " & _
                      .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
                      NoNullStr(.SXLID) & ", " & NoNullStr(smpId(1)) & ", " & NoNullStr(smpId(2)) & ", " & NoNullStr(blkID) & ", " & _
                      NoNullStr(StaffID) & ", SYSDATE, ' ', SYSDATE, '0', SYSDATE, " & sCmbMukesaki & " " & _
                      ")"
'                sql = "insert into TBCMW007 (" & _
'                      "CRYNUM, INGOTPOS, " & _
'                      "CRYLEN, KRPROCCD, PROCCODE, " & _
'                      "SXLID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, " & _
'                      "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE" & _
'                      ",PLANTCAT  " & _
'                      ") values (" & _
'                      NoNullStr(.CRYNUM) & ", " & .INGOTPOS & ", " & _
'                      .LENGTH & ", '" & MGPRCD_SXL_KAKUTEI & "', '" & PROCD_SXL_KAKUTEI & "', " & _
'                      NoNullStr(.SXLID) & ", " & NoNullStr(smpId(1)) & ", " & NoNullStr(smpId(2)) & ", " & NoNullStr(BlkId) & ", " & _
'                      NoNullStr(StaffID) & ", SYSDATE, ' ', SYSDATE, '0', SYSDATE " & _
'                      ")"
' 2007/09/04 SPK Tsutsumi Add End
                If 0 >= OraDB.ExecuteSQL(sql) Then
                    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                    errTbl = "W007"
                    GoTo proc_exit
                End If
                
                'SXL検査書とその関連データを書き込む
                If WriteX00n(.SXLID, .COUNT, errmsg) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
                    errTbl = errmsg
                    GoTo proc_exit
                End If
            End If
        End With
'    Next
        
    
    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_fcmlc001e_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'全面変更 2003/10/17 SystemBrain
'履歴    :2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
'                         ②サンプルIDを反映元のサンプルID(結晶サンプルID含む)ではなく、代表サンプルIDに変更する。
'                         ③GB7/GB8/GB9のSXL確定日付を一致させる。
Private Function WriteX00n(ByVal SXLID$, ByVal WfCnt%, errmsg$) As FUNCTION_RETURN
    Dim recX001(1 To 2)     As c_cmzcrec
    Dim recX002(1 To 2)     As c_cmzcrec
    Dim recX003(1 To 2)     As c_cmzcrec        'GD検査測定点データ 2005/02/15 ffc)tanabe
    Dim recX004(1 To 2)     As c_cmzcrec        'EP検査書　06/08/10 ooba
    Dim recX005(1 To 2)     As c_cmzcrec        'EP測定点ﾃﾞｰﾀ　06/08/10 ooba
    Dim recX006(1 To 2)     As c_cmzcrec        'CuDeco対応 2011/02/14 tkimura
    Dim recX007()           As c_cmzcrec        'GBG対応 2011/06/23 Marushita
    Dim i                   As Integer
    Dim j                   As Integer
    Dim rs                  As OraDynaset
    Dim sql                 As String
    Dim XlSmpPos(1 To 2)    As Integer
    Dim CRYNUM              As String
    Dim blkID               As String
    Dim sBlkId(1 To 2)      As String       'XSDCW BLOCKID格納
    Dim smpId(2)            As String
    Dim HIN                 As tFullHinban
'2003/10/19 ﾛｰｶﾙ変数追加 SystemBrain ==================================▽
    Dim recXSDCS(1 To 2)    As c_cmzcrec        '新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ)
    Dim recXSDCW(1 To 2)    As c_cmzcrec        '新ｻﾝﾌﾟﾙ管理(SXL)
    Dim recXSDCW_1()        As c_cmzcrec        '新ｻﾝﾌﾟﾙ管理(SXL中間抜試) GBG対応 2011/06/23 Marushita
    Dim recE037             As c_cmzcrec        '結晶情報
    Dim recH001             As c_cmzcrec        '引上指示実績 08/12/01
    Dim recXSDC1            As c_cmzcrec        '結晶引上
'2003/10/19 ﾛｰｶﾙ変数追加 SystemBrain ==================================△

'2003/10/19 ﾛｰｶﾙ変数削除 SystemBrain ==================================▽
'    Dim recW009(1 To 2)     As c_cmzcrec
'    Dim recJ014(1 To 2)     As c_cmzcrec
'    Dim recY013(1 To 2)     As c_cmzcrec                '測定評価結果
'    Dim fld As c_cmzcfld
'    Dim SXLPOS(1 To 2) As Integer
'    Dim fldNo As Integer
'    Dim fldCnt As Integer
'    Dim FldName As String
'    Dim specName As Variant
'    Dim CSDATA(1 To 2) As tCsData
'    Dim dMin(3)         As Double
'    Dim dMeas(9)        As Double
'    Dim strMeasPos      As String
'    Dim iRet            As Integer
'2003/10/19 ﾛｰｶﾙ変数削除 SystemBrain ==================================△

    Dim RsHIN       As tFullHinban  '比抵抗(Rs)仕様取得品番　04/02/12 ooba
    Dim sRsData(10) As String       '比抵抗(Rs)ﾃﾞｰﾀ　04/02/12 ooba
'    Dim sRsPtn      As String       '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ　04/02/12 ooba
    Dim sRsPtn(2)   As String       '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ　04/04/15 ooba
    Dim sPos        As String       'SXL位置(TOP/BOT)　04/04/15 ooba
    Dim gSmpID(2)   As String       'TBCMX003用サンプルID   2005/02/15 ffc)tanabe
    Dim sErrMsg     As String       'ｴﾗｰﾒｯｾｰｼﾞ　06/04/20 ooba
    Dim nowtime     As Date  ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
    
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Start SPK habuki　↓↓↓
    Dim flgEDI      As Boolean      'EDI情報有無判定用(True:有、False：無)
    Dim dbPN        As Double       '不純物濃度(P:ﾘﾝ)
    Dim dbBN        As Double       '不純物濃度(B:ﾎﾞﾛﾝ)
    Dim dbASN       As Double       '不純物濃度(AS:砒素)
    Dim dbCN        As Double       '不純物濃度(C:炭素)
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Start SPK habuki　↑↑↑
    
    ''***********************************************************
    'Micron対応 2011/01/14 Add start tkimura
    Dim dd          As type_Coefficient_new2    ''推定抵抗,推定引上率計算構造体
    Dim sRsPos(2)   As String                   '比抵抗(Rs)位置[TOP/BOT]
    Dim data        As String                   'シングル側で抵抗値を取得しているときはSXLID,それ以外のときはBLOCKIDを取得する。
    ''2011/01/14 Add end tkimura
    ''***********************************************************
    'Add Start 2011/05/31 Y.Hitomi
    Dim sUP_RATIO(2) As String                  'SXL検査書/引上げ率[TOP/BOT]
    'Add End   2011/05/31 Y.Hitomi
    '>>>>> 2011/06/24 SETsw)Marushita
    Dim iMCUTUNIT As Integer                    '中間抜試単位
    Dim sMSMPFLG  As Integer                    '中間抜試フラグ
    '<<<<< 2011/06/24 SETsw)Marushita
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function WriteX00n"
    
    WriteX00n = FUNCTION_RETURN_FAILURE
    
    ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
    nowtime = getSvrTime()    'サーバーの時間を取得


    ''SXLの品番を取得する
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   HINBCB as HINBAN"           ''品番
    sql = sql & "  ,REVNUMCB as REVNUM"         ''製品番号改訂番号
    sql = sql & "  ,FACTORYCB as FACTORY"       ''工場
    sql = sql & "  ,OPECB as OPECOND"           ''操業条件
    sql = sql & "  ,PLANTCATCB as PLANTCAT"     ''向先  2007/09/04 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    sql = sql & " WHERE SXLIDCB = '" & SXLID & "'"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'    sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME042 where SXLID = '" & SXLID & "'"
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
        errmsg = "XSDCB:" & rs.RecordCount
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'        errmsg = "E042:" & rs.RecordCount
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
        rs.Close
        GoTo proc_exit
    End If
    HIN.hinban = rs!hinban
    HIN.mnorevno = rs!REVNUM
    HIN.factory = rs!factory
    HIN.opecond = rs!opecond
    HIN.sMukesaki = rs!PLANTCAT
    Set rs = Nothing
    
    ''残存酸素仕様チェック追加　03/12/19 ooba START =====================>
    iChkAoi = ChkAoiSiyou(HIN)
    If iChkAoi < 0 Then
        errmsg = "残存酸素(AOi)仕様エラー" & "  (" & HIN.hinban & Format(HIN.mnorevno, "00") & _
                                                    HIN.factory & HIN.opecond & ")"
        WriteX00n = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ''残存酸素仕様チェック追加　03/12/19 ooba END =======================>
        
    '-------------------- XSDCWの読み込み ----------------------------------------
    sql = "select * from XSDCW where SXLIDCW = '" & SXLID & "' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <> 2 Then
        errmsg = "XSDCW:" & rs.RecordCount
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDCW(1) = New c_cmzcrec
    recXSDCW(1).CopyFromRs "XSDCW", rs
    rs.MoveNext
    Set recXSDCW(2) = New c_cmzcrec
    recXSDCW(2).CopyFromRs "XSDCW", rs
    Set rs = Nothing
    
    CRYNUM = left$(SXLID, 9) & "000"        '結晶番号
    
    '-------------------- ｻﾝﾌﾟﾙIDとﾌﾞﾛｯｸIDの取得 ----------------------------------------
    ' ｻﾝﾌﾟﾙID(From)、ｻﾝﾌﾟﾙID(To)が見つからなければ、ﾌﾞﾛｯｸIDをｾｯﾄし、
    ' ｻﾝﾌﾟﾙID(From)、ｻﾝﾌﾟﾙID(To)が見つかれば、ﾌﾞﾛｯｸIDをｾｯﾄしない。
    smpId(1) = ""       'ｻﾝﾌﾟﾙID(From)初期化
    smpId(2) = ""       'ｻﾝﾌﾟﾙID(To)初期化
    blkID = ""          'ﾌﾞﾛｯｸID初期化
    gSmpID(1) = ""      'TBCMX003用ｻﾝﾌﾟﾙID(From)初期化  2005/02/18 ffc)tanabe
    gSmpID(2) = ""      'TBCMX003用ｻﾝﾌﾟﾙID(To)初期化    2005/02/18 ffc)tanabe
    
    sql = "select REPSMPLIDCW,WFSMPLIDGDCW from XSDCW where SXLIDCW = '" & SXLID & "' and KTKBNCW != '9' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    'ﾌﾞﾛｯｸIDを必ずセットするように変更（XSDCW） 2005/3/22 TUKU START ------------------------------------------
    sBlkId(1) = Trim$(recXSDCW(1)("SMCRYNUMCW").Value)      'TOP BLOCKID
    sBlkId(2) = Trim$(recXSDCW(2)("SMCRYNUMCW").Value)      'BOT BLOCKID

    If rs.RecordCount = 2 Then
        smpId(1) = rs("REPSMPLIDCW")
        ':2008/03/31 青柳 ②サンプルIDを反映元のサンプルID(結晶サンプルID含む)ではなく、代表サンプルIDに変更する。
        ''gSmpID(1) = rs("WFSMPLIDGDCW")      '追加 2005/02/18 ffc)tanabe
        gSmpID(1) = rs("REPSMPLIDCW")
        rs.MoveNext
        
        smpId(2) = rs("REPSMPLIDCW")
        ':2008/03/31 青柳 ②サンプルIDを反映元のサンプルID(結晶サンプルID含む)ではなく、代表サンプルIDに変更する。
        ''gSmpID(2) = rs("WFSMPLIDGDCW")      '追加 2005/02/18 ffc)tanabe
        gSmpID(2) = rs("REPSMPLIDCW")
        Set rs = Nothing
    
        '共有サンプルチェック処理
        If chkComSAMPL(SXLID, smpId(1), smpId(1)) Then
            errmsg = "WriteX00n:共有ｻﾝﾌﾟﾙID取得(From)"
            GoTo proc_exit
        End If
        If chkComSAMPL(SXLID, smpId(2), smpId(2)) Then
            errmsg = "WriteX00n:共有ｻﾝﾌﾟﾙID取得(To)"
            GoTo proc_exit
        End If
    Else
        Set rs = Nothing
        
        
        ':2008/03/31 青柳 ②サンプルIDを反映元のサンプルID(結晶サンプルID含む)ではなく、代表サンプルIDに変更する。
''''        '確定区分=｢9｣で結晶GDを引継いでいる場合の対応　05/10/24 ooba START ================>
''''        sql = "select WFSMPLIDGDCW from XSDCW "
''''        sql = sql & "where SXLIDCW = '" & SXLID & "' "
''''        sql = sql & "and (KTKBNCW != '9' "
''''        sql = sql & "or (KTKBNCW = '9' and WFINDGDCW <> '0' and WFRESGDCW <> '0')) "
''''        sql = sql & "and LIVKCW = '0' "
''''        sql = sql & "order by INPOSCW"
''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''        If rs.RecordCount = 2 Then
''''            gSmpID(1) = rs("WFSMPLIDGDCW")
''''            rs.MoveNext
''''            gSmpID(2) = rs("WFSMPLIDGDCW")
''''            Set rs = Nothing
''''        Else
''''            Set rs = Nothing
''''        End If
''''        '確定区分=｢9｣で結晶GDを引継いでいる場合の対応　05/10/24 ooba END ==================>
        
        
        
        ''ﾌﾞﾛｯｸIDは必ず取得するのでコメント化 2005/3/22 TUKU
        'ブロックIDを得る
''''        sql = "select BLOCKID from TBCME040 "
''''        sql = sql & "where crynum = '" & CRYNUM & "' and "
''''        sql = sql & "INGOTPOS <= " & recXSDCW(1)("INPOSCW").Value & " and "
''''        sql = sql & "INGOTPOS + LENGTH > " & recXSDCW(1)("INPOSCW").Value
''''        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''        If rs.RecordCount = 1 Then
''''            BlkId = rs("BLOCKID")
''''        End If
''''        Set rs = Nothing
    End If
    'ﾌﾞﾛｯｸIDを必ずセットするように変更（XSDCW） 2005/3/22 TUKU END ------------------------------------------
    
    '-------------------- XSDCSの読み込み ----------------------------------------
    For j = 1 To 2
        If j = 1 Then
            '近いXL測定位置(FROM)を求める
            sql = "select * from XSDCS where CRYNUMCS = '" & Trim$(recXSDCW(j)("SMCRYNUMCW").Value) & "' and "
            sql = sql & "TBKBNCS = 'T' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:From"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(1) = New c_cmzcrec
            recXSDCS(1).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(1) = recXSDCS(1)("INPOSCS").Value
        ElseIf j = 2 Then
            '近いXL測定位置(TO)を求める
            sql = "select * from XSDCS where CRYNUMCS = '" & Trim$(recXSDCW(j)("SMCRYNUMCW").Value) & "' and "
            sql = sql & "TBKBNCS = 'B' and LIVKCS = '0'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount <> 1 Then
                errmsg = "XSDCS:To"
                Set rs = Nothing
                GoTo proc_exit
            End If
            Set recXSDCS(2) = New c_cmzcrec
            recXSDCS(2).CopyFromRs "XSDCS", rs
            Set rs = Nothing
            XlSmpPos(2) = recXSDCS(2)("INPOSCS").Value
        End If
    Next j

    '-------------------- TBCME037の読み込み ----------------------------------------
    sql = "select * from TBCME037 where (CRYNUM='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCME037"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recE037 = New c_cmzcrec
    recE037.CopyFromRs "TBCME037", rs
    Set rs = Nothing

    '-------------------- TBCMH001の読み込み ---------------------------------------- (08/12/01)
    sql = "select * from TBCMH001 where UPINDNO = '"
      '8桁目を0とする。 2009/12/23 Change Y.Hitomi
    sql = sql & Mid(CRYNUM, 1, 7) & "0" & Mid(CRYNUM, 9, 1) & "'"
'    Del 2009/12/23 Y.Hitomi
'    If Mid(CRYNUM, 9, 1) = "A" Or Mid(CRYNUM, 9, 1) = "B" Or Mid(CRYNUM, 9, 1) = "C" Then
'        '残量引上(8桁目9桁目を0)
'        sql = sql & Mid(CRYNUM, 1, 7) & "00'"
'    Else
'        '通常品,ﾘﾁｬｰｼﾞ(8桁目を0)
'        sql = sql & Mid(CRYNUM, 1, 7) & "0" & Mid(CRYNUM, 9, 1) & "'"
'    End If
    

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "TBCMH001"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recH001 = New c_cmzcrec
    recH001.CopyFromRs "TBCMH001", rs
    Set rs = Nothing
    
    '-------------------- XSDC1の読み込み ----------------------------------------
    sql = "select * from XSDC1 where (XTALC1='" & CRYNUM & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        errmsg = "XSDC1"
        Set rs = Nothing
        GoTo proc_exit
    End If
    Set recXSDC1 = New c_cmzcrec
    recXSDC1.CopyFromRs "XSDC1", rs
    Set rs = Nothing

    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Start SPK habuki　↓↓↓
    'EDI情報取得
    If Not fncGetEdiInfo(SXLID, dbPN, dbBN, dbASN, dbCN, flgEDI) Then
        errmsg = "XODC6_1"
        GoTo proc_exit
    End If
    '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki　↑↑↑
    
    '==============================================
    '　各種実績データの取得・設定
    '==============================================
    For i = 1 To 2
        '-------------------- TBCMX001固定情報データ設定 ----------------------------------------
        Set recX001(i) = New c_cmzcrec
        recX001(i).TABLENAME = "TBCMX001"
        recX001(i).SetRecDefault
        
        With recX001(i)
            .Fields("SXLID").Value = SXLID                                                  'SXLID
            .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
            .Fields("SAMPLE_FROM").Value = smpId(1)                                         'サンプルID(From)
            .Fields("SAMPLE_TO").Value = smpId(2)                                           'サンプルID(To)
            'XODCWより必ずBOLCKID設定するように変更 2005/03/22 TUKU
            '.Fields("BLOCKID").Value = BlkId                                                'ブロックID
            .Fields("BLOCKID").Value = sBlkId(i)                                                'ブロックID
            .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
            
            ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
            '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID確定日付
            .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
            .nowtime = nowtime
            
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '引上日付
            .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '結晶内開始位置
            .Fields("HINBAN").Value = HIN.hinban                                            '品番
            .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
            .Fields("FACTORY").Value = HIN.factory                                          '工場
            .Fields("OPECOND").Value = HIN.opecond                                          '操業条件
            .Fields("PRODCOND").Value = recE037("PRODCOND").Value                           '製作条件
            .Fields("PGID").Value = Mid(recE037("PGID"), 1, 8)                              'PG-ID
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
            .Fields("SXLPOS").Value = 0                                                     'SXL位置
            .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
            .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
            .Fields("FREELENG").Value = recE037("FREELENG").Value                           'フリー長
            .Fields("DIAMETER").Value = recE037("DIAMETER").Value                           '直径
'            .Fields("CHARGE").Value = recE037("CHARGE").Value                               'チャージ量
            .Fields("SEED").Value = recE037("SEED").Value                                   'シード
            If i = 1 Then                                                                   'ｻﾝﾌﾟﾙID
                .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP側の値
            Else
                .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL側の値
            End If
            .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
            .Fields("CHARGE").Value = recXSDC1("PUCHAGC1").Value                            'ﾁｬｰｼﾞ量 08/12/01
            .Fields("ROCHARGE").Value = recH001("CHARGE").Value                             '認定炉仕込量 08/12/01
            
            '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Start SPK habuki　↓↓↓
            If flgEDI Then
                .Fields("PXL_BORON").Value = dbBN                                           '不純物濃度（B:ﾎﾞﾛﾝ）
                .Fields("PXL_PHOSPHOR").Value = dbPN                                        '不純物濃度（P:ﾘﾝ）
                .Fields("PXL_CARBON").Value = dbCN                                          '不純物濃度（C:炭素）
                .Fields("PXL_ARSENIC").Value = dbASN                                        '不純物濃度（AS:砒素）
            End If
            '■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki　↑↑↑
            
            'Add Start 2011/05/30 Y.Hitomi マルチ対応
            'マルチフラグ
            If Int(recH001("SIJICNT").Value) = 1 Then
                .Fields("MULTI_FLG") = "A"
            ElseIf Int(recH001("SIJICNT").Value) >= 2 Then
                .Fields("MULTI_FLG") = "M"
            End If
            '残量引きフラグ
            If IsNumeric(Mid(CRYNUM, 9, 1)) = True Then
                .Fields("ZANRYO_FLG") = Mid(CRYNUM, 9, 1)
            ElseIf Mid(CRYNUM, 9, 1) = "A" Then
                .Fields("ZANRYO_FLG") = "7"
            ElseIf Mid(CRYNUM, 9, 1) = "B" Then
                .Fields("ZANRYO_FLG") = "8"
            ElseIf Mid(CRYNUM, 9, 1) = "C" Then
                .Fields("ZANRYO_FLG") = "9"
            End If
            'Add End   2011/05/30 Y.Hitomi
        
        End With
        
        '-------------------- TBCMX002固定情報データ設定 ----------------------------------------
        Set recX002(i) = New c_cmzcrec
        recX002(i).TABLENAME = "TBCMX002"
        recX002(i).SetRecDefault
        
        With recX002(i)
            .Fields("SXLID").Value = SXLID                                                  'SXLID
            .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
            .Fields("SAMPLE_FROM").Value = smpId(1)                                         'サンプルID(From)
            .Fields("SAMPLE_TO").Value = smpId(2)                                           'サンプルID(To)
            'XODCWより必ずBOLCKID設定するように変更 2005/03/22 TUKU
            '.Fields("BLOCKID").Value = BlkId                                                'ブロックID
            .Fields("BLOCKID").Value = sBlkId(i)                                                'ブロックID
            .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
            
            ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
            '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID確定日付
            .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
            .nowtime = nowtime

            
            .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '引上日付
            .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '結晶内開始位置
            .Fields("HINBAN").Value = HIN.hinban                                            '品番
            .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
            .Fields("FACTORY").Value = HIN.factory                                          '工場
            .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
            .Fields("SXLPOS").Value = 0                                                     'SXL位置
            .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
            .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
            .Fields("FREELENG").Value = recE037("FREELENG").Value                           'フリー長
            If i = 1 Then                                                                   'ｻﾝﾌﾟﾙID
                .Fields("SAMPID_1").Value = .Fields("SAMPLE_FROM").Value                    'TOP側の値
            Else
                .Fields("SAMPID_1").Value = .Fields("SAMPLE_TO").Value                      'TAIL側の値
            End If
            .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
        End With
                
        If i = 1 Then sPos = "TOP" Else sPos = "BOT"    '04/04/15 ooba
        
        '-------------------- (結晶Rs)結晶抵抗実績(TBCMJ002)データ取得設定 ----------------------------------------
        If getTBCMJ002(CRYNUM, recXSDCS(), i, HIN, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J002:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (結晶Oi)結晶Oi実績(TBCMJ003)データ取得設定 ----------------------------------------
        If getTBCMJ003(CRYNUM, recXSDCS(i), HIN, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J003:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (Cs)Cs実績(TBCMJ004)データ取得設定 ----------------------------------------
'        If getTBCMJ004(CRYNUM, recXSDCS(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
        '品番＆ｴﾗｰﾒｯｾｰｼﾞ追加　06/04/20 ooba
        If getTBCMJ004(CRYNUM, recXSDCS(i), HIN, recX001(i), sErrMsg) = FUNCTION_RETURN_FAILURE Then
            If sErrMsg = "" Then
                errmsg = "J004:" & XlSmpPos(i)
            Else
                errmsg = sErrMsg
            End If
            GoTo proc_exit
        End If

        '-------------------- (結晶OSF1～4)結晶OSF実績(TBCMJ005)データ取得設定 ----------------------------------------
        For j = 1 To 4
            If getTBCMJ005(CRYNUM, recXSDCS(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (結晶BMD1～3)結晶BMD実績(TBCMJ008)データ取得設定 ----------------------------------------
        For j = 1 To 3
            If getTBCMJ008(CRYNUM, recXSDCS(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J008-" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (GD)GD実績(TBCMJ006)データ取得設定 ----------------------------------------
        If getTBCMJ006(CRYNUM, recXSDCS(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J006:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (LT)LT実績(TBCMJ007)データ取得設定 ----------------------------------------
'        If getTBCMJ007(CRYNUM, recXSDCS(i), i, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
        If getTBCMJ007(CRYNUM, recXSDCS(i), HIN, i, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then  '05/12/05 ooba
            errmsg = "J007:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFOi)WFOi実績(TBCMY013)データ取得設定 ----------------------------------------
        If getTBCMY013WFOi(recXSDCW(i), HIN, sPos, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-Oi:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFRs)WFRs実績(TBCMY013)データ取得設定 ----------------------------------------
        If getTBCMY013WFRs(recXSDCW(i), HIN, sPos, recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-Rs:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        '-------------------- (WFDOi1～3)WFDOi実績(TBCMY013)データ取得設定 ----------------------------------------
        For j = 1 To 3
            If getTBCMY013WFDOi(recXSDCW(i), j, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-DOi" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFOSF1～4)WFOSF実績(TBCMY013)データ取得設定 ----------------------------------------
'        For j = 1 To 4  Change  2010/04/19 SIRD対応
        For j = 1 To 3
            If getTBCMY013WFOSF(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i), recX001(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
            'If getTBCMY013WFOSF(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-OSF" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFBMD1～3)WFBMD実績(TBCMY013)データ取得設定 ----------------------------------------
        For j = 1 To 3
            If getTBCMY013WFBMD(recXSDCW(i), j, HIN, sPos, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y013-BMD" & j & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
        Next

        '-------------------- (WFDSOD)WFDSOD実績(TBCMY013)データ取得設定 ----------------------------------------
        If getTBCMY013WFDSOD(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-DSOD:" & XlSmpPos(i)
            GoTo proc_exit
        End If

    ''Upd start 2005/06/23 (TCS)T.Terauchi      SPV9点対応
'        '-------------------- (WFSPV)WFSPV実績(TBCMY013)データ取得設定 ----------------------------------------
'        If getTBCMY013WFSPV(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
'            errmsg = "Y013-SPV:" & XlSmpPos(i)
'            GoTo proc_exit
'        End If
        '-------------------- (WFSPV)WFSPV実績(TBCMJ016)データ取得設定 ----------------------------------------
        If getTBCMJ016WFSPV(CRYNUM, recXSDCW(i), HIN, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J016-SPV:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        
        '-------------------- 標準測定(TBCMY018)よりWarp実績設定 ----------------------------------------------
        If getTBCMY018WARP(sBlkId(i), recXSDCW(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y018-WARP:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        
    ''Upd end   2005/06/23 (TCS)T.Terauchi      SPV9点対応
    
        '-------------------- (WFDZ)WFDZ実績(TBCMY013)データ取得設定 ----------------------------------------
        '品番追加　06/09/06 ooba
        If getTBCMY013WFDZ(recXSDCW(i), HIN, recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
'        If getTBCMY013WFDZ(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-DZ:" & XlSmpPos(i)
            GoTo proc_exit
        End If

        ''残存酸素実績取得追加　03/12/19 ooba START ================================================>
        '-------------------- (WFAOi)WFAOi実績(TBCMY013)データ取得設定 ----------------------------------------
        If getTBCMY013WFAOi(recXSDCW(i), recX001(i), recX002(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "Y013-AOi:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        ''残存酸素実績取得追加　03/12/19 ooba END ==================================================>
        
        ''↓SIRD評価実績取得追加　10/04/19 Y.Hitomi
        If getTBCMJ022SIRD(CRYNUM, recXSDCW(i), recX001(i)) = FUNCTION_RETURN_FAILURE Then
            errmsg = "J022-SIRD:" & XlSmpPos(i)
            GoTo proc_exit
        End If
        ''↑SIRD評価実績取得追加　10/04/19 Y.Hitomi
        

        '==============================================
        '　TBCMX001 に書き込む
        '==============================================
        With recX001(i)
            .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
''            .Fields("SENDFLAG").Value = "3"                                       '送信フラグ
''            .Fields("SENDFLAG").Value = "0"                 '送信ﾀｲﾐﾝｸﾞ変更　05/11/25 ooba
            .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
            .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
            sql = .SqlInsert
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "X001-" & i
                GoTo proc_exit
            End If
        End With

        '==============================================
        '　TBCMX002 に書き込む
        '==============================================
        With recX002(i)
            .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
''            .Fields("SENDFLAG").Value = "3"                                       '送信フラグ
''            .Fields("SENDFLAG").Value = "0"                 '送信ﾀｲﾐﾝｸﾞ変更　05/11/25 ooba
            .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
            .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付

            sql = .SqlInsert
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "X002-" & i
                GoTo proc_exit
            End If
        End With

        ''GD検査測定点データ(TBCMX003)の追加 2005/02/15 ffc)tanabe START ===========================>
        ''XSDCWのGD先行評価の状態フラグ=1且つ実績フラグ=1の場合TBCMX003に登録する。
        If (recXSDCW(i)("WFINDGDCW").Value <> "0") And (recXSDCW(i)("WFRESGDCW").Value <> "0") Then
            
            '-------------------- TBCMX003固定情報データ設定 ----------------------------------------
            Set recX003(i) = New c_cmzcrec
            recX003(i).TABLENAME = "TBCMX003"
            recX003(i).SetRecDefault
            
            With recX003(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
                .Fields("SAMPLE_FROM").Value = gSmpID(1)                                        'サンプルID(From)
                .Fields("SAMPLE_TO").Value = gSmpID(2)                                          'サンプルID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            'ブロックID
                .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
            
                ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID確定日付
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '引上日付
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '結晶内開始位置
                .Fields("HINBAN").Value = HIN.hinban                                            '品番
                .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
                .Fields("FACTORY").Value = HIN.factory                                          '工場
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
                .Fields("SXLPOS").Value = 0                                                     'SXL位置
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           'フリー長
                If i = 1 Then                                                                   'ｻﾝﾌﾟﾙID
                    .Fields("SAMPID_1").Value = .Fields("SAMPLE_FROM").Value                    'TOP側の値
                Else
                    .Fields("SAMPID_1").Value = .Fields("SAMPLE_TO").Value                      'TAIL側の値
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
            End With
            
            '保証フラグ=1(結晶GD実績を保証)の場合
            If recXSDCW(i).Fields("WFHSGDCW") = "1" Then
                If getTBCMJ006GD(CRYNUM, recXSDCW(i), recX003(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "J006-GD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            '保証フラグ=0(WFGD実績を保証)の場合
            Else
                If getTBCMJ015WFGD(CRYNUM, recXSDCW(i), recX003(i), recX003(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                'If getTBCMJ015WFGD(CRYNUM, recXSDCW(i), recX003(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "J015-GD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            End If
        
            '==============================================
            '　TBCMX003 に書き込む
            '==============================================
            With recX003(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
''                .Fields("SENDFLAG").Value = "3"                                       '送信フラグ
''                .Fields("SENDFLAG").Value = "0"             '送信ﾀｲﾐﾝｸﾞ変更　05/11/25 ooba
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
                .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
                
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X003-" & i
                    GoTo proc_exit
                End If
            End With
        End If
        
        ''GD検査測定点データ(TBCMX003)の追加 2005/02/15 ffc)tanabe END ============================>
        
        
        'EP検査書(X004)/EP測定点ﾃﾞｰﾀ(X005)作成　06/08/10 ooba START ===========================>
        If ((recXSDCW(i)("EPINDB1CW").Value <> "0") And (recXSDCW(i)("EPRESB1CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDB2CW").Value <> "0") And (recXSDCW(i)("EPRESB2CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDB3CW").Value <> "0") And (recXSDCW(i)("EPRESB3CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL1CW").Value <> "0") And (recXSDCW(i)("EPRESL1CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL2CW").Value <> "0") And (recXSDCW(i)("EPRESL2CW").Value <> "0")) Or _
            ((recXSDCW(i)("EPINDL3CW").Value <> "0") And (recXSDCW(i)("EPRESL3CW").Value <> "0")) Then
        
            '-------------------- TBCMX004固定情報ﾃﾞｰﾀ設定 ---------------------------------------
            Set recX004(i) = New c_cmzcrec
            recX004(i).TABLENAME = "TBCMX004"
            recX004(i).SetRecDefault
            
            With recX004(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
                .Fields("SAMPLE_FROM").Value = smpId(1)                                         'ｻﾝﾌﾟﾙID(From)
                .Fields("SAMPLE_TO").Value = smpId(2)                                           'ｻﾝﾌﾟﾙID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            'ﾌﾞﾛｯｸID
                .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
                
                ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID確定日付
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '引上日付
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '結晶内開始位置
                .Fields("HINBAN").Value = HIN.hinban                                            '品番
                .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
                .Fields("FACTORY").Value = HIN.factory                                          '工場
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
                .Fields("SXLPOS").Value = 0                                                     'SXL位置
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           'ﾌﾘｰ長
                If i = 1 Then                                                                   'ｻﾝﾌﾟﾙID
                    .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP側の値
                Else
                    .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL側の値
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
            End With
            
            '-------------------- TBCMX005固定情報ﾃﾞｰﾀ設定 ---------------------------------------
            Set recX005(i) = New c_cmzcrec
            recX005(i).TABLENAME = "TBCMX005"
            recX005(i).SetRecDefault
            
            With recX005(i)
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
                .Fields("SAMPLE_FROM").Value = smpId(1)                                         'ｻﾝﾌﾟﾙID(From)
                .Fields("SAMPLE_TO").Value = smpId(2)                                           'ｻﾝﾌﾟﾙID(To)
                .Fields("BLOCKID").Value = sBlkId(i)                                            'ﾌﾞﾛｯｸID
                .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
                
                ':2008/03/31 青柳 ③GB7/GB8/GB9のSXL確定日付を一致させる。
                '' .Fields("SXLDECDATE").Value = "SYSDATE"                                         'SXL-ID確定日付
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
                .nowtime = nowtime
                
                .Fields("PLUPDATE").Value = recXSDC1("TDAYC1").Value                            '引上日付
                .Fields("INGOTPOS").Value = recXSDCW(i)("INPOSCW").Value                        '結晶内開始位置
                .Fields("HINBAN").Value = HIN.hinban                                            '品番
                .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
                .Fields("FACTORY").Value = HIN.factory                                          '工場
                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
                .Fields("SXLPOS").Value = 0                                                     'SXL位置
                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
                .Fields("FREELENG").Value = recE037("FREELENG").Value                           'ﾌﾘｰ長
                If i = 1 Then                                                                   'ｻﾝﾌﾟﾙID
                    .Fields("SAMPID").Value = .Fields("SAMPLE_FROM").Value                      'TOP側の値
                Else
                    .Fields("SAMPID").Value = .Fields("SAMPLE_TO").Value                        'TAIL側の値
                End If
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
            End With
            
            '-------------------- ｴﾋﾟOSF1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定 ---------------------------
            For j = 1 To 3
                If getTBCMY022EPOSF(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i), recX004(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                'If getTBCMY022EPOSF(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y022-EPOSF" & j & ":" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            Next j
            
            '-------------------- ｴﾋﾟBMD1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定 ---------------------------
            For j = 1 To 3
                If getTBCMY022EPBMD(recXSDCW(i), j, HIN, sPos, recX004(i), recX005(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y022-EPBMD" & j & ":" & XlSmpPos(i)
                    GoTo proc_exit
                End If
            Next j
            
            '==============================================
            '　TBCMX004 に書き込む
            '==============================================
            With recX004(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
''                .Fields("SENDFLAG").Value = "0"                                       '送信ﾌﾗｸﾞ
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
                .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
    
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X004-" & i
                    GoTo proc_exit
                End If
            End With
            
            '==============================================
            '　TBCMX005 に書き込む
            '==============================================
            With recX005(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
''                .Fields("SENDFLAG").Value = "0"                                       '送信ﾌﾗｸﾞ
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
                .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
    
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X005-" & i
                    GoTo proc_exit
                End If
            End With
        End If
        'EP検査書(X004)/EP測定点ﾃﾞｰﾀ(X005)作成　06/08/10 ooba END =============================>
        
'=================================================================================
' 2011/02/14 tkimura ADD START
        '-------------------- TBCMX006固定情報データ設定 ----------------------------------------
        'Add 2011/03/01 Y.Hitomi C-OSF3指示有りであれば、X006を作成する。
        If recXSDCS(i)("CRYINDL4CS").Value = "1" Or recXSDCS(i)("CRYINDL4CS").Value = "2" Then
        
            Set recX006(i) = New c_cmzcrec
            recX006(i).TABLENAME = "TBCMX006"
            recX006(i).SetRecDefault
            
            With recX006(i)
                .Fields("PLANTCAT").Value = HIN.sMukesaki
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = IIf(recXSDCW(i)("TBKBNCW").Value = "T", "1", "2")  'FROMTO区分
                'XODCWより必ずBOLCKID設定するように変更 2005/03/22 TUKU
                '.Fields("BLOCKID").Value = BlkId                                               'ブロックID
                .Fields("BLOCKID").Value = sBlkId(i)                                            'ブロックID
                .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
                
                ':2011/02/14 tkimura ③GB7/GB8/GB9/GBFのSXL確定日付を一致させる。
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
                .nowtime = nowtime
                
                .Fields("HINBAN").Value = HIN.hinban                                            '品番
                .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
                .Fields("FACTORY").Value = HIN.factory                                          '工場
                .Fields("OPECOND").Value = HIN.opecond                                          '操業条件
                
                .Fields("SXL_SMPLICHI").Value = recXSDCS(i)("INPOSCS").Value                    '結晶内位置
                .Fields("SXL_SMPLNO").Value = recXSDCS(i)("CRYSMPLIDL4CS").Value                'C-OSF3サンプルNo
    
            End With
    
            '-------------------- (結晶OSF3)結晶OSF実績(TBCMJ005)データ取得設定 ----------------------------------------
            If getTBCMJ005CuDeco(CRYNUM, recXSDCS(i), recX006(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J005-" & "CuDeco" & ":" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
            '-------------------- (CLESTA)CLESTA評価実績(TBCMJ023)データ取得設定 ----------------------------------------
            '入力の際に必要なものは結晶番号+サンプルID
            If getTBCMJ023(CRYNUM, recXSDCS(i), recX006(i)) = FUNCTION_RETURN_FAILURE Then
                errmsg = "J023-:" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
            '==============================================
            '　TBCMX006 に書き込む
            '==============================================
            With recX006(i)
                .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
                .Fields("SENDFLAG").Value = "3"                                         ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する
                .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
                sql = .SqlInsert
                Debug.Print (sql)
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X006-" & i
                    GoTo proc_exit
                End If
            End With
        End If
' 2011/02/14 tkimura ADD END
'=================================================================================
        
        ''TOP/BOT別に比抵抗ﾃﾞｰﾀ取得　04/04/15 ooba START ======================================>
        '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝを求める。
        If SxlRsPattern(HIN, sPos, sRsPtn(i)) = FUNCTION_RETURN_FAILURE Then
            '取得ﾃﾞｰﾀなし
            sRsPtn(i) = "C"
        End If
        '比抵抗ﾃﾞｰﾀを取得する。
        If cmbc040_GetSxlRsData(SXLID, sPos, sRsPtn(i), sRsData()) = FUNCTION_RETURN_FAILURE Then
            If sRsPtn(i) = "A" Then errmsg = "WF"
            If sRsPtn(i) = "B" Then errmsg = "結晶"
            errmsg = errmsg & "比抵抗実績ﾃﾞｰﾀを取得できません(Y007)"
            GoTo proc_exit
        End If
        ''TOP/BOT別に比抵抗ﾃﾞｰﾀ取得　04/04/15 ooba END ========================================>
        
'=================================================================================
' 2011/01/14 tkimura ADD START
''①抵抗のTop位置取得方法,抵抗のBot位置取得方法
        'パターンAのときはシングルID,BのときはブロックID
        If sRsPtn(i) = "A" Then
            data = SXLID
        ElseIf sRsPtn(i) = "B" Then
            data = sBlkId(i)
        Else
            data = ""
        End If

        '関数名:cmbc040_GetSxlRsPos(シングルIDorブロックID,sPos,sRsPtn(i),
        'sRsPos(i))+代表サンプルID[smpId(1),smpId(2)]
        'sRsPos(i)にρトップ位置とρボトム位置を格納する。
        If cmbc040_GetSxlRsPos(data, _
                               sPos, _
                               sRsPtn(i), _
                               sRsPos()) = FUNCTION_RETURN_FAILURE Then
            If sRsPtn(i) = "A" Then errmsg = "WF"
            If sRsPtn(i) = "B" Then errmsg = "結晶"
            errmsg = errmsg & "比抵抗位置ﾃﾞｰﾀを取得できません"
            GoTo proc_exit
        End If
' 2011/01/14 tkimura ADD END
'=================================================================================
'        If i = 1 Then
        If i = 2 Then   '04/04/15 ooba
            If smpId(1) = vbNullString Then
                '' ブロックID取得　2003/09/16 Motegi ==================================> START
'                    blkID = Trim$(recW009(1)("BLOCKID").Value)
                blkID = Trim$(recXSDCW(1)("SMCRYNUMCW").Value)
                '' ブロックID取得　2003/09/16 Motegi ==================================> END
            Else
                blkID = vbNullString
            End If
            
'''''            ''比抵抗ﾃﾞｰﾀ取得　04/02/12 ooba START ===========================================>
'''''            If SXLID <> vbNullString Then
'''''                'SXLの品番を取得
'''''                RsHIN.HINBAN = ""
'''''                sql = "select HINBAN, REVNUM, FACTORY, OPECOND from TBCME042 "
'''''                sql = sql & "where SXLID = '" & SXLID & "' "
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount = 1 Then
'''''                    RsHIN.HINBAN = rs("HINBAN")
'''''                    RsHIN.mnorevno = rs("REVNUM")
'''''                    RsHIN.factory = rs("FACTORY")
'''''                    RsHIN.opecond = rs("OPECOND")
'''''                End If
'''''                Set rs = Nothing
'''''                '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝを求める。
'''''                If SxlRsPattern(RsHIN, sPos, sRsPtn) = FUNCTION_RETURN_FAILURE Then
'''''                    '取得ﾃﾞｰﾀなし
'''''                    sRsPtn = "C"
'''''                End If
'''''                '比抵抗ﾃﾞｰﾀを取得する。
'''''                If cmbc040_GetSxlRsData(SXLID, sRsPtn, sRsData()) = FUNCTION_RETURN_FAILURE Then
'''''                    If sRsPtn = "A" Then errmsg = "WF"
'''''                    If sRsPtn = "B" Then errmsg = "結晶"
'''''                    errmsg = errmsg & "比抵抗実績ﾃﾞｰﾀを取得できません(Y007)"
'''''                    GoTo proc_exit
'''''                End If
'''''            End If
'''''            ''比抵抗ﾃﾞｰﾀ取得　04/02/12 ooba END =============================================>
            
            ''TBCMY007
            ''比抵抗ﾃﾞｰﾀ登録追加　04/02/12 ooba

'' 2007/09/04 SPK Tsutsumi Add Start
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
            sql = "Insert into TBCMY007 " & _
                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT, PLANTCAT) " & _
                  "values (" & _
                  NoNullStr(SXLID) & ", " & _
                  NoNullStr(smpId(1)) & ", " & _
                  NoNullStr(smpId(2)) & ", " & _
                  NoNullStr(blkID) & ", " & _
                  "(select distinct HINBCB||to_char(REVNUMCB,'FM00') from XSDCB where SXLIDCB=" & NoNullStr(SXLID) & "), " & _
                  "'00', " & _
                  "'TX853I', " & _
                  "SYSDATE, '0', '0', SYSDATE, " & _
                  NoNullStr(sRsData(1)) & ", " & _
                  NoNullStr(sRsData(2)) & ", " & _
                  NoNullStr(sRsData(3)) & ", " & _
                  NoNullStr(sRsData(4)) & ", " & _
                  NoNullStr(sRsData(5)) & ", " & _
                  NoNullStr(sRsData(6)) & ", " & _
                  NoNullStr(sRsData(7)) & ", " & _
                  NoNullStr(sRsData(8)) & ", " & _
                  NoNullStr(sRsData(9)) & ", " & _
                  NoNullStr(sRsData(10)) & "," & _
                  "'" & sCmbMukesaki & "'" & _
                  ")"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'' 2007/09/04 SPK Tsutsumi Add End
            
'''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'            sql = "Insert into TBCMY007 " & _
'                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
'                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
'                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT) " & _
'                  "values (" & _
'                  NoNullStr(SXLID) & ", " & _
'                  NoNullStr(smpId(1)) & ", " & _
'                  NoNullStr(smpId(2)) & ", " & _
'                  NoNullStr(BlkId) & ", " & _
'                  "(select distinct HINBCB||to_char(REVNUMCB,'FM00') from XSDCB where SXLIDCB=" & NoNullStr(SXLID) & "), " & _
'                  "'00', " & _
'                  "'TX853I', " & _
'                  "SYSDATE, '0', '0', SYSDATE, " & _
'                  NoNullStr(sRsData(1)) & ", " & _
'                  NoNullStr(sRsData(2)) & ", " & _
'                  NoNullStr(sRsData(3)) & ", " & _
'                  NoNullStr(sRsData(4)) & ", " & _
'                  NoNullStr(sRsData(5)) & ", " & _
'                  NoNullStr(sRsData(6)) & ", " & _
'                  NoNullStr(sRsData(7)) & ", " & _
'                  NoNullStr(sRsData(8)) & ", " & _
'                  NoNullStr(sRsData(9)) & ", " & _
'                  NoNullStr(sRsData(10)) & _
'                  ")"
'''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'            sql = "Insert into TBCMY007 " & _
'                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE, " & _
'                  "MESDATA1TOP, MESDATA2TOP, MESDATA3TOP, MESDATA4TOP, MESDATA5TOP, " & _
'                  "MESDATA1BOT, MESDATA2BOT, MESDATA3BOT, MESDATA4BOT, MESDATA5BOT) " & _
'                  "values (" & _
'                  NoNullStr(SXLID) & ", " & _
'                  NoNullStr(smpId(1)) & ", " & _
'                  NoNullStr(smpId(2)) & ", " & _
'                  NoNullStr(BlkId) & ", " & _
'                  "(select distinct HINBAN||to_char(REVNUM,'FM00') from TBCME042 where SXLID=" & NoNullStr(SXLID) & "), " & _
'                  "'00', " & _
'                  "'TX853I', " & _
'                  "SYSDATE, '0', '0', SYSDATE, " & _
'                  NoNullStr(sRsData(1)) & ", " & _
'                  NoNullStr(sRsData(2)) & ", " & _
'                  NoNullStr(sRsData(3)) & ", " & _
'                  NoNullStr(sRsData(4)) & ", " & _
'                  NoNullStr(sRsData(5)) & ", " & _
'                  NoNullStr(sRsData(6)) & ", " & _
'                  NoNullStr(sRsData(7)) & ", " & _
'                  NoNullStr(sRsData(8)) & ", " & _
'                  NoNullStr(sRsData(9)) & ", " & _
'                  NoNullStr(sRsData(10)) & _
'                  ")"
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''''                  "(SXL_ID,SAMPLE_FROM,SAMPLE_TO,BLOCKID,HINBAN,KUBUN,TXID,REGDATE,SUMMITSENDFLAG,SENDFLAG,SENDDATE) " &
''''                  "SYSDATE, '0', '0', SYSDATE" &
            If OraDB.ExecuteSQL(sql) < 1 Then
                errmsg = "Y007"
                GoTo proc_exit
            End If

            '=================================================================================
            ' 2011/01/14 tkimura ADD START
            ''②TBCMY007取得後、ρTop位置引上率,ρBot位置引上率,実効偏析,基準抵抗値をもとめる。
            If GetStandardPosRes(CRYNUM, _
                                 sRsData(), _
                                 sRsPos(), _
                                 dd) = FUNCTION_RETURN_FAILURE Then
                errmsg = errmsg & "基準抵抗値を取得できません"
                GoTo proc_exit
            End If
  
            '③推定対象引上率,推定位置比抵抗値を求める。
            '推定抵抗データ計算、更新(Y011)    SuiteiResDataCalculation(SXLID[inputのみ],dd[inputのみ])

            If SuiteiResDataCalculation(SXLID, _
                                        dd, sUP_RATIO) = FUNCTION_RETURN_FAILURE Then
                errmsg = errmsg & "推定抵抗値の計算に失敗しました。"
                GoTo proc_exit
            End If
            
            '⑤TBCMY011テーブルの品質システム送信フラグを更新する。
            If UpdateTBCMY011SendFlag(SXLID, HIN) = FUNCTION_RETURN_FAILURE Then
                errmsg = "Y011:" & XlSmpPos(i)
                GoTo proc_exit
            End If
            
        End If
' 2011/01/14 tkimura ADD END
'=================================================================================

    Next
'Add Start 2011/05/31 Y.Hitomi　引上げ率更新
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMX001" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " UP_RATIO='" & sUP_RATIO(1) & "'" & vbCrLf  'SXL最BOT位置の引き上げ率
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " SXLID='" & SXLID & "'" & vbCrLf           'SXLID
        
    If OraDB.ExecuteSQL(sql) < 1 Then
        errmsg = "X001-" & i
        GoTo proc_exit
    End If
'Add End   2011/05/31 Y.Hitomi

'>>>>> 2011/06/00 SETsw)Marushita 中間抜試サンプル送信対応
    Dim recCnt As Integer
    Dim iPosC2 As Integer
    '-------------------- XSDCW_1の読み込み ----------------------------------------
    sql = "select * from XSDCW_1 where SXLIDCW = '" & SXLID & "' and LIVKCW = '0' order by INPOSCW"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim recXSDCW_1(recCnt)
    ReDim recX007(recCnt)
    If recCnt > 0 Then
        For i = 1 To recCnt
            Set recXSDCW_1(i) = New c_cmzcrec
            recXSDCW_1(i).CopyFromRs "XSDCW_1", rs
            rs.MoveNext
        Next
    End If
    Set rs = Nothing
    
    If getTBCME036(HIN, iMCUTUNIT, sMSMPFLG) = FUNCTION_RETURN_FAILURE Then
        errmsg = "X007-Get_TBCME036"
        GoTo proc_exit
    End If
    
    If recCnt > 0 And sMSMPFLG = 1 Then
        For i = 1 To recCnt
            '-------------------- TBCMX007固定情報データ設定 ----------------------------------------
            Set recX007(i) = New c_cmzcrec
            recX007(i).TABLENAME = "TBCMX007"
            recX007(i).SetRecDefault
            'ダミー初期化
            Set recX002(1) = New c_cmzcrec
            recX002(1).TABLENAME = "TBCMX002"
            recX002(1).SetRecDefault
            Set recX005(1) = New c_cmzcrec
            recX005(1).TABLENAME = "TBCMX005"
            recX005(1).SetRecDefault
        
            With recX007(i)
                .Fields("PLANTCAT").Value = HIN.sMukesaki                                       '向先 2007/09/04 SPK Tsutsumi Add
                .Fields("SXLID").Value = SXLID                                                  'SXLID
                .Fields("FROMTOKBN").Value = "C"                                                'FROMTO区分(C固定)
                .Fields("SAMPLE_ID").Value = recXSDCW_1(i)("REPSMPLIDCW").Value                 'サンプルID(代表)
                .Fields("BLOCKID").Value = Trim$(recXSDCW_1(i)("SMCRYNUMCW").Value)             'ブロックID
                .Fields("CRYNUM").Value = CRYNUM                                                '結晶番号
                .Fields("SXLDECDATE").Value = nowtime                                           'SXL-ID確定日付
                .nowtime = nowtime
            
                'Cng Start 2011/07/11 Y.Hitomi ブロック内位置は、XSDC2の位置を取得し、結晶内位置から引く
                iPosC2 = getXSDC2Pos(Trim$(recXSDCW_1(i)("SMCRYNUMCW").Value))
                .Fields("BLOCKPOS").Value = recXSDCW_1(i)("INPOSCW").Value - iPosC2             'ブロック内抜試位置
                '.Fields("BLOCKPOS").Value = recXSDCW_1(i)("INPOSCW").Value                      'ブロック内抜試位置
                'Cng End   2011/07/11 Y.Hitomi
                
                .Fields("INGOTPOS").Value = recXSDCW_1(i)("INPOSCW").Value                      '結晶内開始位置
                .Fields("HINBAN").Value = HIN.hinban                                            '品番
                .Fields("REVNUM").Value = HIN.mnorevno                                          '製品番号改訂番号
                .Fields("FACTORY").Value = HIN.factory                                          '工場
                .Fields("OPECOND").Value = HIN.opecond                                          '操業条件
                .Fields("MCUTUNIT").Value = iMCUTUNIT                                           '中間抜試単位
                
                '-------------------- TBCME036データ取得設定 ----------------------------------------
'                .Fields("PRODCOND").Value = recE037("PRODCOND").Value                           '製作条件
'                .Fields("PGID").Value = Mid(recE037("PGID"), 1, 8)                              'PG-ID
'                .Fields("UPLENGTH").Value = recE037("UPLENGTH").Value                           '引上げ長さ
'                .Fields("SXLPOS").Value = 0                                                     'SXL位置
'                .Fields("SXLLENGTH").Value = XlSmpPos(2) - XlSmpPos(1)                          'SXL-ID確定長さ
'                .Fields("SXLWAFERCNT").Value = WfCnt                                            'SXL-ID確定時のWF枚数
'                .Fields("FREELENG").Value = recE037("FREELENG").Value                           'フリー長
'                .Fields("DIAMETER").Value = recE037("DIAMETER").Value                           '直径
'                .Fields("SEED").Value = recE037("SEED").Value                                   'シード
                '-------------------- (WFOi)WFOi実績(TBCMY013)データ取得設定 ----------------------------------------
                If getTBCMY013WFOi(recXSDCW_1(i), HIN, sPos, recX007(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-Oi:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFRs)WFRs実績(TBCMY013)データ取得設定 ----------------------------------------
                If getTBCMY013WFRs(recXSDCW_1(i), HIN, sPos, recX007(i)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-Rs:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFDOi1～3)WFDOi実績(TBCMY013)データ取得設定 ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFDOi(recXSDCW_1(i), j, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-DOi" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFOSF1～4)WFOSF実績(TBCMY013)データ取得設定 ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFOSF(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX002(1), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-OSF" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFBMD1～3)WFBMD実績(TBCMY013)データ取得設定 ----------------------------------------
                For j = 1 To 3
                    If getTBCMY013WFBMD(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y013-BMD" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next
                '-------------------- (WFDSOD)WFDSOD実績(TBCMY013)データ取得設定 ----------------------------------------
                If getTBCMY013WFDSOD(recXSDCW_1(i), recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-DSOD:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                '-------------------- (WFDZ)WFDZ実績(TBCMY013)データ取得設定 ----------------------------------------
                If getTBCMY013WFDZ(recXSDCW_1(i), HIN, recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-DZ:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                ''残存酸素実績取得追加　03/12/19 ooba START ================================================>
                '-------------------- (WFAOi)WFAOi実績(TBCMY013)データ取得設定 ----------------------------------------
                If getTBCMY013WFAOi(recXSDCW_1(i), recX007(i), recX002(1)) = FUNCTION_RETURN_FAILURE Then
                    errmsg = "Y013-AOi:" & XlSmpPos(i)
                    GoTo proc_exit
                End If
                ''XSDCWのGD先行評価の状態フラグ=1且つ実績フラグ=1の場合TBCMX007に登録する。
                If (recXSDCW_1(i)("WFINDGDCW").Value <> "0") And (recXSDCW_1(i)("WFRESGDCW").Value <> "0") Then
                    '-------------------- (GD)GD実績(TBCMJ015)データ取得設定 ----------------------------------------
                    If getTBCMJ015WFGD(CRYNUM, recXSDCW_1(i), recX007(i), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "J015-GD:" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                End If
                '-------------------- ｴﾋﾟOSF1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定 ---------------------------
                For j = 1 To 3
                    If getTBCMY022EPOSF(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX005(1), recX007(i).TABLENAME) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y022-EPOSF" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next j
                '-------------------- ｴﾋﾟBMD1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定 ---------------------------
                For j = 1 To 3
                    If getTBCMY022EPBMD(recXSDCW_1(i), j, HIN, sPos, recX007(i), recX005(1)) = FUNCTION_RETURN_FAILURE Then
                        errmsg = "Y022-EPBMD" & j & ":" & XlSmpPos(i)
                        GoTo proc_exit
                    End If
                Next j
                '==============================================
                '　TBCMX007 に書き込む
                '==============================================
                .Fields("REGDATE").Value = "SYSDATE"                                    '登録日付
                .Fields("SENDFLAG").Value = "3"  ':2008/03/31 青柳 ①現行SXL確定実行時からSXLマップ受信時に送信ﾀｲﾐﾝｸﾞ変更する。
                .Fields("SENDDATE").Value = "SYSDATE"                                   '送信日付
                sql = .SqlInsert
                If OraDB.ExecuteSQL(sql) < 1 Then
                    errmsg = "X007-" & i
                    GoTo proc_exit
                End If
            End With
        Next
    End If
    
    WriteX00n = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    WriteX00n = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Private Function NoNullStr(s$) As String
    If s = vbNullString Then
        NoNullStr = "' '"
    Else
        NoNullStr = "'" & s & "'"
    End If
End Function

'分割結晶（ブロック）前工程実績取得＆構造体作成 2002/09/03 ADD hitec)N.MATSUMOTO
Public Function cmbc040_CreateXSDC2(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim intLoopCnt  As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim iRtn    As Integer
    Dim dblDiameter As Double
    Dim intNum  As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    'ブロックIDを得る
    sql = " SELECT * FROM XSDC2" '
    sql = sql & " WHERE CRYNUMC2='" & strBlockID(iBlockCnt) & "'"
''''    sql = sql & "   AND NEWKNTC2='" & BeforeProc & "'"
    sql = sql & "   AND LIVKC2= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc040_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    If rs.EOF = False Then
        '前工程取得
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '工程連番
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '現在長さ
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '現在重量
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '現在枚数
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")   ' 2007/09/04 SPK Tsutsumi Add
        End With

        '分割結晶（ﾌﾞﾛｯｸ）の前工程を、分割結晶（ﾌﾞﾛｯｸ）の現在工程へコピー
        BlkNow = BlkOld

        '分割結晶（ﾌﾞﾛｯｸ）現在工程の編集
        With BlkNow
            .KCNTC2 = CInt(.KCNTC2) + 1     '工程連番]
            'Cng Start 2010/09/02 Y.Hitomi
            ''ブロック内SXLが1つでも完了していた場合、工程コードを更新しないようにする
            If .GNWKNTC2 <> "     " Then
                .NEWKNTC2 = Kihon.NOWPROC        '前工程
                .GNWKNTC2 = "CW800"              '現在工程
            End If
            
'            .NEWKNTC2 = Kihon.NOWPROC           '前工程
'            .GNWKNTC2 = Kihon.NEWPROC           '現在工程
            'Cng End 2010/09/02 Y.Hitomi

            '現在重量を求める
            If GetDiameter(strBlockID(iBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            Kihon.DIAMETER = dblDiameter
            '取得した直径を元に重量を求める
''''            .GNWC2 = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLC2)))
''''            '現在枚数を求める
''''            If WfCount(strBlockID(iBlockCnt), CInt(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
''''                .GNMC2 = 0
''''''''                GNMCA proc_wxit
''''            Else
''''                .GNMC2 = intNum
''''            End If
            .SUMITBC2 = "0"
            .SUMITLC2 = "0"
            .SUMITMC2 = "0"
            .SUMITWC2 = "0"
        End With

    End If
    
    rs.Close

    cmbc040_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc040_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'分割結晶（品番）前工程実績取得＆構造体作成 2002/09/03 ADD hitec)N.MATSUMOTO
'Cng Start 2010/10/14 Y.Hitomi
'Public Function cmbc040_CreateXSDCA(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean) As FUNCTION_RETURN
Public Function cmbc040_CreateXSDCA(ByVal iBlockCnt As Integer, ByRef bNoData As Boolean, strSxlId As String) As FUNCTION_RETURN
'Cng End   2010/10/14 Y.Hitomi
    Dim iLoopCnt    As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim dblDiameter As Double
    Dim intNum  As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    'ブロックIDを得る
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID(iBlockCnt) & "'"
''''    sql = sql & "   AND NEWKNTCA='" & BeforeProc & "'"
    sql = sql & "   AND LIVKCA= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc040_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    rs.MoveFirst
    iLoopCnt = 0
    
    Do While Not rs.EOF
        ReDim Preserve HinOld(iLoopCnt)
        ReDim Preserve HinNow(iLoopCnt)
        With HinOld(iLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA")   ' 2007/09/04 SPK Tsutsumi Add
        End With
        
        '前工程の構造体を現在工程の構造体へコピー
        HinNow(iLoopCnt) = HinOld(iLoopCnt)
        
        '現在工程構造体の工程連番の編集
        With HinNow(iLoopCnt)
            .KCKNTCA = BlkNow.KCNTC2
            'Cng Start 2010/10/14 Y.Hitomi
            '実行指示SXLIDのみ工程コードを変更し、それ以外は、前工程を引き継ぐ
            If strSxlId = .SXLIDCA Then
                .NEWKNTCA = Kihon.NOWPROC             '前工程コードを最終通過工程にセット
                .GNWKNTCA = Kihon.NEWPROC             '現在工程コードを現在工程へセット
            Else
                .NEWKNTCA = rs.Fields("NEWKNTCA")     '前工程コードを最終通過工程にセット
                .GNWKNTCA = rs.Fields("GNWKNTCA")     '現在工程コードを現在工程へセット
            End If
'            .NEWKNTCA = Kihon.NOWPROC       '現在工程
'            .GNWKNTCA = Kihon.NEWPROC       '次工程
            'Cng End   2010/10/14 Y.Hitomi
            
            '現在重量を求める
            If GetDiameter(strBlockID(iBlockCnt), dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '取得した直径を元に重量を求める
''''            HinNow(iLoopCnt).GNWCA = CStr(WeightOfCylinder(dblDiameter, CDbl(.GNLCA)))
'''''            '現在枚数を求める
'''''            If WfCount(strBlockID(iBlockCnt), CInt(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
'''''                .GNMCA = 0
'''''''''                GNMCA proc_wxit
'''''            Else
'''''                .GNMCA = intNum
'''''            End If
            .SUMITBCA = "0"
            .SUMITLCA = HinOld(iLoopCnt).SUMITLCA    ''03/05/13 後藤
            .SUMITMCA = HinOld(iLoopCnt).SUMITMCA    ''03/05/14 後藤
            .SUMITWCA = HinOld(iLoopCnt).SUMITWCA    ''03/05/13 後藤
        End With
        
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop
    
    With Kihon  '基本情報　品番構造体のカウントをセット
        .CNTHINOLD = iLoopCnt
        .CNTHINNOW = iLoopCnt
    End With
    
    rs.Close
    cmbc040_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc040_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'構造体作成処理 2002/09/03 ADD hitec)N.MATSUMOTO
Public Function CreateTable(ByVal strSxlId As String, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim sql     As String
    Dim rsMain  As OraDynaset
    Dim iBlockCnt   As Integer
    Dim strDBName   As String
    Dim bNoData     As Boolean
    Dim sTmpSxl() As String     '仕掛工程再ﾁｪｯｸ用SXLID　06/03/14 ooba

    On Error GoTo proc_err

    bNoData = False
    'ブロックID取得
    strDBName = "XSDCA"
    sql = "select DISTINCT(CRYNUMCA) from XSDCA " & _
          "where SXLIDCA='" & strSxlId & "' " & _
          "  and LIVKCA= '0'"
    Set rsMain = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rsMain.RecordCount = 0 Then
        Debug.Print "XSDC2：前工程実績無し"
        Debug.Print sql
        rsMain.Close
        CreateTable = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    sProSXLID = strSxlId    '処理SXLIDｾｯﾄ　06/03/24 ooba
    
    'ブロックID取得
''''    bNoData = False
''''    strDBName = "E040"
''''    sql = "select BLOCKID from TBCME040 " & _
''''          "where (crynum='" & strCrynum & "') and (INGOTPOS<=" & intIngotpos & ") and (" & intIngotpos & "<INGOTPOS+LENGTH)"
''''    Set rsMain = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rsMain.RecordCount = 0 Then
''''        rsMain.Close
''''        strErrMsg = GetMsgStr("EAPLY") & strDBName
''''        GoTo proc_exit
''''    End If
    
    '仕掛工程再チェック機能追加　06/03/14 ooba
    ReDim sTmpSxl(1)
    sTmpSxl(1) = strSxlId
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_SXL_KAKUTEI, strErrMsg) = FUNCTION_RETURN_FAILURE Then
        CreateTable = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    iBlockCnt = 0
    
    Do While Not rsMain.EOF
        iBlockCnt = iBlockCnt + 1
        ReDim strBlockID(iBlockCnt)
        strBlockID(iBlockCnt) = rsMain("CRYNUMCA")
        With Kihon
            .StaffID = Trim(f_cmbc040_1.txtStaffID.text)    '担当者コード
''''            BeforeProc = PROCD_WFC_SOUGOUHANTEI '前工程
            .NEWPROC = PROCD_SXL_MAP
            .NOWPROC = PROCD_SXL_KAKUTEI
            .DIAMETER = 0   '------------------保留
            .ALLSCRAP = "N"     '全数スクラップ無し
            .FURYOUMU = "N"       '不良無し
        End With
        
        '分割結晶（ブロック）から前工程実績取得
        strDBName = "XSDC2"
        If cmbc040_CreateXSDC2(iBlockCnt, bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
'                CreateTable = FUNCTION_RETURN_SUCCESS
                CreateTable = FUNCTION_RETURN_FAILURE       '07/02/06 ooba
                strErrMsg = GetMsgStr("EAPLY") & strDBName  '07/02/06 ooba
                Debug.Print "cmbc040_CreateXSDC2(" & iBlockCnt & "," & bNoData & ")：XSDC2前工程実績無し"
                GoTo proc_exit
            Else
                CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc040_CreateXSDC2(" & iBlockCnt & "," & bNoData & ")：XSDC2前工程実績読み込みエラー"
                GoTo proc_exit
            End If
        End If
        
        '分割結晶（品番）から前工程実績取得
        strDBName = "XSDCA"
        'Cng Start 2010/10/14 Y.Hitomi
'        If cmbc040_CreateXSDCA(iBlockCnt, bNoData) = FUNCTION_RETURN_FAILURE Then
        If cmbc040_CreateXSDCA(iBlockCnt, bNoData, strSxlId) = FUNCTION_RETURN_FAILURE Then
        'Cng End   2010/10/14 Y.Hitomi
            If bNoData = True Then
'                CreateTable = FUNCTION_RETURN_SUCCESS
                CreateTable = FUNCTION_RETURN_FAILURE       '07/02/06 ooba
                strErrMsg = GetMsgStr("EAPLY") & strDBName  '07/02/06 ooba
                Debug.Print "cmbc040_CreateXSDCA(" & iBlockCnt & "," & bNoData & ")：XSDCA前工程実績無し "
                GoTo proc_exit
            Else
                CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc040_CreateXSDCA(" & iBlockCnt & "," & bNoData & ")：XSDCA前工程実績読み込みエラー"
                GoTo proc_exit
            End If
        End If
        
        '基本処理
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            CreateTable = FUNCTION_RETURN_FAILURE           '08/04/04 ooba
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "KihonProc()：基本処理に失敗しました"
            Exit Function
        End If
        
        rsMain.MoveNext
    Loop
    rsMain.Close
                
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    CreateTable = FUNCTION_RETURN_FAILURE
    Resume proc_exit
                
End Function
'2002/09/03 ADD hitec)N.MATSUMOTO

'###  WF通信処理（基本処理パラメータ作成） ### '2002/09/03 ADD hitec)N.MATSUMOTO    Start
'Public Function MakeParameter() As FUNCTION_RETURN
'f_cmbc040_1.frm→s_cmbc040_SQL.bas関数移動,record追加　06/10/20 ooba
Public Function MakeParameter(record As typ_cmlc001e_Disp) As FUNCTION_RETURN
    Dim lng     As Long
    Dim dat     As Variant
    Dim lRowCnt As Long
    Dim rsMain      As OraDynaset
    Dim sql     As String
    Dim iBlockCnt   As Integer
    Dim sErrMsg As String

'    For lRowCnt = 1 To f_cmbc040_1.sprList.MaxRows
'        With rec(lRowCnt)
        With record     '06/10/20 ooba
            If .HLDCLASS = "0" Then     'ホールドチェックがOFFの場合
                strSxlData = .SXLID     '03/05/01 Add.後藤
                If CreateTable(.SXLID, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                    MakeParameter = FUNCTION_RETURN_FAILURE
                    f_cmbc040_1.lblMsg.Caption = sErrMsg
                    GoTo proc_exit
                End If
            End If
        End With
'    Next
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

End Function
'2002/09/03 ADD hitec)N.MATSUMOTO    End

'2003/10/19 使ってないので削除 SystemBrain ==========================================================▽
''''''概要      :Cs実績データの取得ドライバ
''''''ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
''''''          :sxlid           , I  ,String            , SXLID
''''''          :Cs()            , I  ,tCsData           , 結晶Cs測定結果
''''''      　　:戻り値          , O  , FUNCTION_RETURN　, 読み込みの成否
''''''説明      :Csの上下実績を取得する
''''''履歴      :2002/10/03 野村 作成
'''''Private Function getSXLCs(SXLID$, Cs() As tCsData) As FUNCTION_RETURN
'''''    Dim rs As OraDynaset
'''''    Dim sql As String
'''''    Dim CRYNUM As String
'''''    Dim sxlFrom As Integer
'''''    Dim sxlLen As Integer
'''''    Dim SpecCsMin As Double
'''''    Dim specCsH As String
'''''
'''''    '' エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmzcF_cmlc001e_SQL.bas -- Function getSXLCs"
'''''    getSXLCs = FUNCTION_RETURN_FAILURE
'''''
'''''    '実績初期化
'''''    With Cs(1)
'''''        .SXL_CS_SMPPOS = -1
'''''        .SXLCS_CSMEAS = -1
'''''        .SXLCS_70PPRE = -1
'''''    End With
'''''    With Cs(2)
'''''        .SXL_CS_SMPPOS = -1
'''''        .SXLCS_CSMEAS = -1
'''''        .SXLCS_70PPRE = -1
'''''    End With
'''''
'''''    '結晶番号,SXL範囲,SXL品番のCs仕様(下限値)を取得
'''''    sql = "select CRYNUM, INGOTPOS, LENGTH, HSXCNMIN, HSXCNHWS "
'''''    sql = sql & "from TBCME019 SPEC, TBCME042 SXL "
'''''    sql = sql & "where SXL.SXLID='" & SXLID & "'"
'''''    sql = sql & "  and SPEC.HINBAN=SXL.HINBAN and SPEC.MNOREVNO=SXL.REVNUM and SPEC.FACTORY=SXL.FACTORY and SPEC.OPECOND=SXL.OPECOND"
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''    If rs.RecordCount > 0 Then
'''''        CRYNUM = rs("CRYNUM")
'''''        sxlFrom = rs("INGOTPOS")
'''''        sxlLen = rs("LENGTH")
'''''        SpecCsMin = rs("HSXCNMIN")
'''''        specCsH = rs("HSXCNHWS")
'''''    Else
'''''        GoTo proc_exit  'SXLもしくは品番仕様なし
'''''    End If
'''''    rs.Close
'''''    '仕様あり（H　OR　S）の場合のみ実績を書き込む
'''''    If specCsH = "H" Or specCsH = "S" Then
''''''        If Left(CRYNUM, 1) <> "8" Then                 '2003/10/18 削除 SystemBrain
'''''            '引上結晶の実績取得
'''''            If SpecCsMin > 0 Then
'''''                'FromTo仕様の場合は、ブロックのTop/Bot測定値を検索する(引継不可)
'''''                'Top側
'''''                sql = vbNullString
'''''                sql = sql & "select J.POSITION, J.CSMEAS, J.PRE70P "
'''''                sql = sql & "from TBCME040 B, TBCMJ004 J "
'''''                sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "  and B.INGOTPOS<=" & sxlFrom
'''''                sql = sql & "  and " & sxlFrom & "<B.INGOTPOS+B.LENGTH"
'''''                sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS "
'''''                sql = sql & "order by TRANCNT desc"
'''''                sql = "select * from (" & sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(1)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Top側実績なし(SXL確定不可)
'''''                End If
'''''                'Bot側
'''''                sql = vbNullString
'''''                sql = sql & "select J.POSITION, J.CSMEAS, J.PRE70P "
'''''                sql = sql & "from TBCME040 B, TBCMJ004 J "
'''''                sql = sql & "where B.CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "  and B.INGOTPOS<" & sxlFrom + sxlLen
'''''                sql = sql & "  and " & sxlFrom + sxlLen & "<=B.INGOTPOS+B.LENGTH"
'''''                sql = sql & "  and J.CRYNUM=B.CRYNUM and J.POSITION=B.INGOTPOS+B.LENGTH "
'''''                sql = sql & "order by TRANCNT desc"
'''''                sql = "select * from (" & sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(2)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Tail側実績なし(SXL確定不可)
'''''                End If
'''''            Else
'''''                'FromTo仕様でなければ、なるべく近い下側から検索する
'''''                sql = vbNullString
'''''                sql = sql & "select * from ("
'''''                sql = sql & "  select POSITION, CSMEAS, PRE70P"
'''''                sql = sql & "  from TBCMJ004 J"
'''''                sql = sql & "  where CRYNUM='" & CRYNUM & "'"
'''''                sql = sql & "    and POSITION>=" & sxlFrom + sxlLen
'''''                sql = sql & "  order by POSITION, TRANCOND, SMPKBN, TRANCNT desc"
'''''                sql = sql & ") where rownum=1"
'''''                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''''                If rs.RecordCount > 0 Then
'''''                    With Cs(2)
'''''                        .SXL_CS_SMPPOS = rs("POSITION")
'''''                        .SXLCS_CSMEAS = rs("CSMEAS")
'''''                        .SXLCS_70PPRE = rs("PRE70P")
'''''                    End With
'''''                ElseIf specCsH = "H" Then
'''''                    GoTo proc_exit       'Tail側実績なし(SXL確定不可)
'''''                End If
'''''                rs.Close
'''''            End If
''''''2003/10/18 削除 SystemBrain -------------------------------------------▽
''''''        Else
''''''            '購入単結晶の実績取得
''''''            sql = vbNullString
''''''            sql = sql & "select * from ("
''''''            sql = sql & " select B.INGOTPOS, B.LENGTH, XL.CSTOP, XL.CSTAIL"
''''''            sql = sql & " from TBCMG002 XL, TBCME040 B "
''''''            sql = sql & " where B.CRYNUM='" & CRYNUM & "' and B.INGOTPOS<=" & sxlFrom & " and " & sxlFrom + sxlLen & "<=B.INGOTPOS+B.LENGTH"
''''''            sql = sql & "   and XL.CRYNUM=B.BLOCKID"
''''''            sql = sql & " order by TRANCNT desc"
''''''            sql = sql & ") where rownum=1"
''''''            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''''            If rs.RecordCount > 0 Then
''''''                If rs("CSTOP") >= 0 Then    'Top側の値が入っていたら
''''''                    With Cs(1)
''''''                        .SXL_CS_SMPPOS = rs("INGOTPOS")
''''''                        .SXLCS_CSMEAS = rs("CSTOP")
''''''                        .SXLCS_70PPRE = 0
''''''                    End With
''''''                ElseIf (specCsH = "H") And (SpecCsMin > 0) Then 'FromTo保証
''''''                    GoTo proc_exit      'Top側実績なし(SXL確定不可)
''''''                End If
''''''                If rs("CSTAIL") >= 0 Then    'Tail側の値が入っていたら
''''''                    With Cs(2)
''''''                        .SXL_CS_SMPPOS = Val(rs("INGOTPOS")) + Val(rs("LENGTH"))
''''''                        .SXLCS_CSMEAS = rs("CSTAIL")
''''''                        .SXLCS_70PPRE = 0
''''''                    End With
''''''                ElseIf specCsH = "H" Then
''''''                    GoTo proc_exit      'Tail側実績なし(SXL確定不可)
''''''                End If
''''''            ElseIf specCsH = "H" Then
''''''                GoTo proc_exit          'Tail側実績なし(SXL確定不可)
''''''            End If
''''''            rs.Close
''''''        End If
''''''2003/10/18 削除 SystemBrain -------------------------------------------△
'''''    End If
'''''    getSXLCs = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' 終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    getSXLCs = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function
'2003/10/19 使ってないので削除 SystemBrain ==========================================================△

'概要      :BMD実績のMin値を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :dMin          ,O   ,Double    ,Min値
'          :strMeasPos    ,I   ,String    ,結晶欠陥測定位置コード（3byte）
'          :dMeas()       ,I   ,Double    ,測定位置配列
'          :戻り値        ,O   ,Integer     ,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function getSXLBMDMIN(dMin As Double, strMeasPos As String, dMeas() As Double) As Integer
    Dim dConv       As Double
    Dim iMeasNum    As Integer
    Dim Index       As Integer
    Dim dForMin()   As Double
    Dim strParam    As String

    On Error GoTo Err
    getSXLBMDMIN = FUNCTION_RETURN_FAILURE

    If strMeasPos = "" Then
        dMin = -1
        getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    '' 結晶欠陥測定位置（測定方法）より換算係数を取得
    strParam = GetCodeField("GP", "01", Mid(strMeasPos, 1, 1), "INFO8")
    If strParam = vbNullString Then strParam = "1"
    dConv = val(strParam)

    '' 結晶欠陥測定位置（測定点）の取得
    iMeasNum = GetMeasureNum(Mid(strMeasPos, 2, 1), 1)
    If iMeasNum < 1 Then Exit Function

    '' Min値計算
    ReDim dForMin(iMeasNum - 1)
    For Index = 0 To UBound(dForMin)
        dForMin(Index) = dMeas(Index)
    Next Index
    dMin = GetMin(dForMin) * dConv / 10000

    getSXLBMDMIN = FUNCTION_RETURN_SUCCESS
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
End Function

Private Function NtoS(strWk As String) As String
    If Mid(strWk, 1, 1) = Chr(0) Then
        NtoS = " "
        Exit Function
    End If
    NtoS = strWk
End Function

Private Function NtoZ2(strWk As String) As Double
    If Trim(strWk) = "" Then
        NtoZ2 = -1
        Exit Function
    End If
    NtoZ2 = CDbl(strWk)
End Function

Private Function CryRES_Judg(CRs() As Double, GarRes As Guarantee) As Double
    Dim pt As Integer

    ''RRG判定
    Select Case GarRes.cPos
      Case "B", "C", "D", "E", "F", "K", "S", "Y"
          Select Case GarRes.cBunp
          Case "A", "B", "C", "M"
             ''RRG計算
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

          Case "", " "                                          'ｽﾍﾟｰｽ追加　05/07/05 ooba
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarRes.cBunp = "" Or GarRes.cBunp = " " Then   '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
                CryRES_Judg = -1
                Exit Function
'             End If                                            '→ｺﾒﾝﾄ化　05/07/05 ooba

          Case Else
             ''RRG計算　　　コード "A" にて計算
             If Trim(GarRes.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarRes.cCount)
             End If
             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)

         End Select

      Case Else
         Select Case GarRes.cBunp
         Case "A", "B", "C", "D", "E", "M", "N"
             ''RRG計算
             CryRES_Judg = MENNAI_Cal(RES_JUDG, CRs(), GarRes, GarRes.cBunp)

         Case "", " "                                           'ｽﾍﾟｰｽ追加　05/07/05 ooba
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarRes.cBunp = "" Or GarRes.cBunp = " " Then   '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
                CryRES_Judg = -1
                Exit Function
'             End If                                            '→ｺﾒﾝﾄ化　05/07/05 ooba

         Case Else
             ''RRG計算　　　コード "A" にて計算
             If Trim(GarRes.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarRes.cCount)
             End If
             CryRES_Judg = RoundUp((RGCal(CRs(), pt)), 4)

         End Select
    End Select
Cal_Escp:
        
End Function

Private Function CryOi_Judg(COi() As Double, GarOi As Guarantee) As Double
    Dim pt As Integer
    ReDim JData(UBound(COi())) As Double
    
    ''ORG判定
    
    Select Case GarOi.cPos
      Case "B", "C", "D", "E", "F", "K", "Y"
          Select Case GarOi.cBunp
          Case "A", "B", "C"
             ''ORG計算
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

          Case "", " "                                              'ｽﾍﾟｰｽ追加　05/07/05 ooba
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
                CryOi_Judg = -1
                Exit Function
'             End If                                                '→ｺﾒﾝﾄ化　05/07/05 ooba

          Case Else
             ''ORG計算　　　コード "A" にて計算
             If Trim(GarOi.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarOi.cCount)
             End If
             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)

         End Select

      Case Else

         Select Case GarOi.cBunp
         Case "A", "B", "C", "D", "E", "N"
             ''ORG計算
             CryOi_Judg = MENNAI_Cal(OI_JUDG, COi(), GarOi, GarOi.cBunp)

         Case "", " "                                               'ｽﾍﾟｰｽ追加　05/07/05 ooba
             ''計算区分がスペースの場合は、計算，判定を行わない
'             If GarOi.cBunp = "" Or GarOi.cBunp = " " Then         '→ｺﾒﾝﾄ化　05/07/05 ooba
'                    GoTo Cal_Escp
                CryOi_Judg = -1
                Exit Function
'             End If                                                '→ｺﾒﾝﾄ化　05/07/05 ooba

         Case Else
             ''ORG計算　　　コード "A" にて計算
             If Trim(GarOi.cCount) = "" Then
                pt = 3
             Else
                pt = val(GarOi.cCount)
             End If
             CryOi_Judg = RoundUp((RGCal(COi(), pt)), 4)

         End Select
    End Select
Cal_Escp:

End Function

'概要      :結晶抵抗実績(TBCMJ002)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS()      , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :i               , I  ,Integer           , Top/Bot種別(1:Top, 2:Bot)
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶抵抗実績(TBCMJ002)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ002(CRYNUM As String, recXSDCS() As c_cmzcrec, i As Integer, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim k           As Integer
    Dim wMeas1(2)   As Double
    Dim wgtCharge   As Long                 '偏析計算用パラメータ
    Dim wgtCharge1  As Long                 '偏析計算用パラメータ   '' 2008/11/26 SXL検査書チャージ量追加 ADD By Systech
    Dim wgtCharge2  As Long                 '偏析計算用パラメータ   '' 2008/11/26 SXL検査書チャージ量追加 ADD By Systech
    Dim wgtTop      As Double               '偏析計算用パラメータ
    Dim wgtTopCut   As Double               '偏析計算用パラメータ
    Dim DM          As Double               '偏析計算用パラメータ
    Dim cc          As type_Coefficient
    Dim CRes        As C_RES                '結晶RS判定構造体
    Dim wComp       As Double
    Dim wHSXRHWYS   As String               '保証方法＿処
    Dim RET As FUNCTION_RETURN
    Dim wStaff      As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ002"
    
    getTBCMJ002 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("SXL_RS_SMPPOS").Value = -1                 'SXLRSサンプル測定位置(SXL測定情報)
        .Fields("SXLRS_MEAS1").Value = -1                   'SXLRS_測定値1
        .Fields("SXLRS_MEAS2").Value = -1                   'SXLRS_測定値2
        .Fields("SXLRS_MEAS3").Value = -1                   'SXLRS_測定値3
        .Fields("SXLRS_MEAS4").Value = -1                   'SXLRS_測定値4
        .Fields("SXLRS_MEAS5").Value = -1                   'SXLRS_測定値5
        .Fields("SXLRS_EFEHS").Value = -1                   'SXLRS_実効偏析
        .Fields("SXLRS_RRG").Value = -1                     'SXLRS_RRG
    
        '-------------------- TBCMJ002の読み込み(Rs) ----------------------------------------
        If (recXSDCS(i)("CRYINDRSCS").Value <> "0") And (recXSDCS(i)("CRYRESRS1CS").Value <> "0") Then
            '実効偏析算出の為、Top/Botの両方を取得
            For k = 1 To 2
                sql = "select * from TBCMJ002 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & recXSDCS(k)("CRYSMPLIDRSCS").Value & " "
                sql = sql & "order by TRANCNT desc"
                sql = "select * from (" & sql & ") where rownum = 1"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                If k = i Then
                    .Fields("SXL_RS_SMPPOS").Value = rs("POSITION")             'SXLRSサンプル測定位置(SXL測定情報)
                    .Fields("SXLRS_MEAS1").Value = rs("MEAS1")                  'SXLRS_測定値1
                    .Fields("SXLRS_MEAS2").Value = rs("MEAS2")                  'SXLRS_測定値2
                    .Fields("SXLRS_MEAS3").Value = rs("MEAS3")                  'SXLRS_測定値3
                    .Fields("SXLRS_MEAS4").Value = rs("MEAS4")                  'SXLRS_測定値4
                    .Fields("SXLRS_MEAS5").Value = rs("MEAS5")                  'SXLRS_測定値5
                    wStaff = rs("KSTAFFID")                                     '---TEST2004/10
                End If
                wMeas1(k) = rs("MEAS1")                             '実効偏析算出用
                Set rs = Nothing
            Next k
            
            'SXLRS_EFEHS
            'マルチ引上対応 関数参照先変更 2008/05/26 SETsw Nakada
            If GetCoeffParams_new(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo proc_exit
'            If GetCoeffParams(CRYNUM, wgtCharge, wgtTop, wgtTopCut, DM) = FUNCTION_RETURN_FAILURE Then GoTo PROC_EXIT

'' 2008/11/26 SXL検査書チャージ量追加 DEL By Systech Start
''            .Fields("CHARGE").Value = wgtCharge 'チャージ量 取得先変更 2008/05/26 SETsw Nakada
'' 2008/11/26 SXL検査書チャージ量追加 DEL By Systech End
            
            cc.DUNMENSEKI = AreaOfCircle(DM)
            cc.TOPSMPLPOS = recXSDCS(1)("INPOSCS").Value
            cc.BOTSMPLPOS = recXSDCS(2)("INPOSCS").Value
            cc.CHARGEWEIGHT = wgtCharge
            cc.TOPWEIGHT = wgtTop + wgtTopCut
            cc.TOPRES = wMeas1(1)
            cc.BOTRES = wMeas1(2)
            wComp = CoefficientCalculation(cc)
        
            If wComp = -9999 Then
                wComp = 0                                       'SXLRS_実効偏析
            End If
            .Fields("SXLRS_EFEHS").Value = wComp                'SXLRS_実効偏析
            
'''' 2008/11/26 SXL検査書チャージ量追加 UPD By Systech Start
''            'チャージ量
''            sql = " SELECT C1.SUICHARGE, C1.PUCHAGC1 "
''            sql = sql & " FROM XSDC1 C1 "
''            sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"
''            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''            If rs.RecordCount = 0 Then
''                wgtCharge1 = 0                      ''推定チャージ
''                wgtCharge2 = 0                      ''チャージ量
''            Else
''                wgtCharge1 = rs("SUICHARGE")        ''推定チャージ
''                wgtCharge2 = rs("PUCHAGC1")         ''チャージ量
''            End If
''            Set rs = Nothing
''
''            .Fields("CHARGE").Value = wgtCharge2    'チャージ量
''            .Fields("ROCHARGE").Value = wgtCharge1  '推定チャージ
'''' 2008/11/26 SXL検査書チャージ量追加 UPD By Systech End

            'SXLRS_RRG
            sql = "select HSXRHWYS, HSXRSPOH, HSXRSPOT, HSXRSPOI from TBCME018 where "
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            CRes.GuaranteeRes.cBunp = rs("HSXRSPOH")                    ' 品ＳＸ比抵抗測定位置＿方
            CRes.GuaranteeRes.cCount = rs("HSXRSPOT")                   ' 品ＳＸ比抵抗測定位置＿点
            CRes.GuaranteeRes.cPos = rs("HSXRSPOI")                     ' 品ＳＸ比抵抗測定位置＿位
            wHSXRHWYS = rs("HSXRHWYS")                                  ' 品ＳＸ比抵抗保証方法＿処
            Set rs = Nothing
            
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs測定値1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs測定値2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs測定値3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs測定値4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs測定値5
            
            ''------TEST2004/10 -> 2004/12 測定データ順に更新のため削除
            ''-----> 2006/06 測定位置による計算は必要なためコメントを外し測定順にデータを戻す処理を追加する
            If Trim(wStaff) <> KSTAFF_J002 Then   '新測定データの場合だけ処理する
                RET = Set_Rs_Ichi(CRes.GuaranteeRes.cCount, CRes.GuaranteeRes.cPos, CRes.Res(0), CRes.Res(1), CRes.Res(2), _
                               CRes.Res(3), CRes.Res(4))
            End If
            
            .Fields("SXLRS_RRG").Value = CryRES_Judg(CRes.Res(), CRes.GuaranteeRes)     'SXLRS_RRG
        
            '2006/06 追加----
            CRes.Res(0) = NtoZ2(.Fields("SXLRS_MEAS1").Value)           'Rs測定値1
            CRes.Res(1) = NtoZ2(.Fields("SXLRS_MEAS2").Value)           'Rs測定値2
            CRes.Res(2) = NtoZ2(.Fields("SXLRS_MEAS3").Value)           'Rs測定値3
            CRes.Res(3) = NtoZ2(.Fields("SXLRS_MEAS4").Value)           'Rs測定値4
            CRes.Res(4) = NtoZ2(.Fields("SXLRS_MEAS5").Value)           'Rs測定値5
            '--------
            
            '保証方法="H"、かつ、SXLRS_RRG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
            If (wHSXRHWYS = "H") And (.Fields("SXLRS_RRG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMJ002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ002 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶Oi実績(TBCMJ003)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶Oi実績(TBCMJ003)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ003(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim COi         As C_Oi                 '結晶Oi判定構造体
    Dim wHSXONHWS   As String               '保証方法＿処
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ003"
    
    getTBCMJ003 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("SXL_OI_SMPPOS").Value = -1                 'SXLOIサンプル測定位置(SXL測定情報)
        .Fields("SXLOI_OIMEAS1").Value = -1                 'SXLOI_Oi測定値1
        .Fields("SXLOI_OIMEAS2").Value = -1                 'SXLOI_Oi測定値2
        .Fields("SXLOI_OIMEAS3").Value = -1                 'SXLOI_Oi測定値3
        .Fields("SXLOI_OIMEAS4").Value = -1                 'SXLOI_Oi測定値4
        .Fields("SXLOI_OIMEAS5").Value = -1                 'SXLOI_Oi測定値5
        .Fields("SXLOI_ORGRES").Value = -1                  'SXLOI_ORG結果
        .Fields("SXLOI_INSPECTWAY").Value = -1              'SXLOI検査方法
    
        '-------------------- TBCMJ003の読み込み(Oi) ----------------------------------------
        If (recXSDCS("CRYINDOICS").Value <> "0") And (recXSDCS("CRYRESOICS").Value <> "0") Then
            sql = "select * from TBCMJ003 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDOICS").Value & " "
            sql = sql & "  and TRANCOND = 0 "   'GFAのFTIR換算値取得異常対応 2011/02/28 SETsw kubota
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_OI_SMPPOS").Value = rs("POSITION")             'SXLOIサンプル測定位置(SXL測定情報)
''''            .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1")              'SXLOI_Oi測定値1
''''            .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2")              'SXLOI_Oi測定値2
''''            .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3")              'SXLOI_Oi測定値3
''''            .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4")              'SXLOI_Oi測定値4
''''            .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5")              'SXLOI_Oi測定値5
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("OIMEAS1")) = False Then .Fields("SXLOI_OIMEAS1").Value = rs("OIMEAS1") Else .Fields("SXLOI_OIMEAS1").Value = -1  'SXLOI_Oi測定値1
            If IsNull(rs("OIMEAS2")) = False Then .Fields("SXLOI_OIMEAS2").Value = rs("OIMEAS2") Else .Fields("SXLOI_OIMEAS2").Value = -1  'SXLOI_Oi測定値2
            If IsNull(rs("OIMEAS3")) = False Then .Fields("SXLOI_OIMEAS3").Value = rs("OIMEAS3") Else .Fields("SXLOI_OIMEAS3").Value = -1  'SXLOI_Oi測定値3
            If IsNull(rs("OIMEAS4")) = False Then .Fields("SXLOI_OIMEAS4").Value = rs("OIMEAS4") Else .Fields("SXLOI_OIMEAS4").Value = -1  'SXLOI_Oi測定値4
            If IsNull(rs("OIMEAS5")) = False Then .Fields("SXLOI_OIMEAS5").Value = rs("OIMEAS5") Else .Fields("SXLOI_OIMEAS5").Value = -1  'SXLOI_Oi測定値5
            'OI_NULL対応　2005/03/08 TUKU END   --------------------------------------------------
            .Fields("SXLOI_INSPECTWAY").Value = rs("INSPECTWAY")        'SXLOI検査方法
            Set rs = Nothing
        
            'SXLOI_ORG
            sql = "select HSXONHWS, HSXONSPH, HSXONSPT, HSXONSPI from TBCME019 where "
            sql = sql & " HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " FACTORY = '" & HIN.factory & "' and "
            sql = sql & " OPECOND = '" & HIN.opecond & "' "
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ReDim COi.Oi(4) As Double
            COi.GuaranteeOi.cBunp = rs("HSXONSPH")                      ' 品ＳＸ酸素濃度測定位置＿方
            COi.GuaranteeOi.cCount = rs("HSXONSPT")                     ' 品ＳＸ酸素濃度測定位置＿点
            COi.GuaranteeOi.cPos = rs("HSXONSPI")                       ' 品ＳＸ酸素濃度測定位置＿位
            wHSXONHWS = rs("HSXONHWS")                                  ' 品ＳＸ酸素濃度保証方法＿処
            Set rs = Nothing

            COi.Oi(0) = NtoZ2(.Fields("SXLOI_OIMEAS1").Value)           'Oi測定値1
            COi.Oi(1) = NtoZ2(.Fields("SXLOI_OIMEAS2").Value)           'Oi測定値2
            COi.Oi(2) = NtoZ2(.Fields("SXLOI_OIMEAS3").Value)           'Oi測定値3
            COi.Oi(3) = NtoZ2(.Fields("SXLOI_OIMEAS4").Value)           'Oi測定値4
            COi.Oi(4) = NtoZ2(.Fields("SXLOI_OIMEAS5").Value)           'Oi測定値5
            
            .Fields("SXLOI_ORGRES").Value = CryOi_Judg(COi.Oi(), COi.GuaranteeOi)       'SXLOI_ORG結果
            
            '保証方法="H"、かつ、SXLOI_ORG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
            If (wHSXONHWS = "H") And (.Fields("SXLOI_ORGRES").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMJ003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ003 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :Cs実績(TBCMJ004)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :HIN             , I  ,tFullHinban       , 品番　06/04/20 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :sErrMsg         , O  ,String            , ｴﾗｰﾒｯｾｰｼﾞ　06/04/20 ooba
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :Cs実績(TBCMJ004)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ004(CRYNUM As String, recXSDCS As c_cmzcrec, HIN As tFullHinban, _
                             recX001 As c_cmzcrec, sErrMsg As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    '06/04/20 ooba START =========================================>
    Dim rs2         As OraDynaset
    Dim dCmax       As Double           '仕様(上限値)
    Dim dCmin       As Double           '仕様(下限値)
    Dim iSmpNo      As Long             '推定元ｻﾝﾌﾟﾙNo      'Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    Dim tCsSuitei   As CS_SUITEI_TYPE   'CS推定計算用構造体
    Dim dCsSuitei   As Double           'Cs推定値
    '06/04/20 ooba END ===========================================>
    
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ004"
    
    getTBCMJ004 = FUNCTION_RETURN_FAILURE

    sErrMsg = ""        '06/04/20 ooba
    
    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("SXL_CS_SMPPOS").Value = -1                 'SXLCSサンプル測定位置(SXL測定情報)
        .Fields("SXLCS_CSMEAS").Value = -1                  'SXLCS_Cs実測値
        .Fields("SXLCS_70PPRE").Value = -1                  'SXLCS_70%推定値
        .Fields("SXLCS_BSUIMEAS").Value = -1                'SXLCS_Csﾌﾞﾛｯｸ推定値　06/04/20 ooba
    
        '-------------------- TBCMJ004の読み込み(Cs) ----------------------------------------
        If (recXSDCS("CRYINDCSCS").Value <> "0") And (recXSDCS("CRYRESCSCS").Value <> "0") Then
            sql = "select * from TBCMJ004 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDCSCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("SXL_CS_SMPPOS").Value = rs("POSITION")             'SXLCSサンプル測定位置(SXL測定情報)
''''            .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS")                'SXLCS_Cs実測値
''''            .Fields("SXLCS_70PPRE").Value = rs("PRE70P")                'SXLCS_70%推定値
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            If IsNull(rs("CSMEAS")) = False Then .Fields("SXLCS_CSMEAS").Value = rs("CSMEAS") Else .Fields("SXLCS_CSMEAS").Value = -1  'SXLCS_Cs実測値
            If IsNull(rs("PRE70P")) = False Then .Fields("SXLCS_70PPRE").Value = rs("PRE70P") Else .Fields("SXLCS_70PPRE").Value = -1  'SXLCS_70%推定値
            'OI_NULL対応　2005/03/08 TUKU START --------------------------------------------------
            
            Set rs = Nothing
            
            ''Csﾌﾞﾛｯｸ推定値計算対応　06/04/20 ooba START ======================================>
        
            '実測の場合は｢ﾌﾞﾛｯｸ推定値＝実測値｣
            If recXSDCS("CRYINDCSCS").Value = "1" Then
                .Fields("SXLCS_BSUIMEAS").Value = .Fields("SXLCS_CSMEAS").Value
            Else
                '①推定位置
                tCsSuitei.sInfPos = CStr(recXSDCS("INPOSCS").Value)
                
                '②ｻﾝﾌﾟﾙ位置
                '③ｻﾝﾌﾟﾙ測定値
                '推定元ｻﾝﾌﾟﾙNo取得
                iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
                
'''                If recXSDCS("CRYINDCSCS").Value <> "0" Then
'''                    iSmpNo = recXSDCS("CRYSMPLIDCSCS").Value
'''                Else
'''                    '仕様値取得
'''                    sql = "select HSXCNMAX, HSXCNMIN from TBCME019 "
'''                    sql = sql & "where HINBAN = '" & HIN.HINBAN & "' "
'''                    sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
'''                    sql = sql & "and FACTORY = '" & HIN.factory & "' "
'''                    sql = sql & "and OPECOND = '" & HIN.opecond & "' "
'''                    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                    If rs.RecordCount = 0 Then
'''                        sErrMsg = GetMsgStr("ECLC3") & " <ｻﾝﾌﾟﾙNo> "
'''                        Set rs = Nothing
'''                        GoTo PROC_EXIT
'''                    End If
'''                    dCmax = fncNullCheck(rs("HSXCNMAX"))    '品SX炭素濃度上限
'''                    dCmin = fncNullCheck(rs("HSXCNMIN"))    '品SX炭素濃度下限
'''                    Set rs = Nothing
'''
'''                    'Cs反映ﾙｰﾙにより推定元ｻﾝﾌﾟﾙNo取得
'''                    '反映ﾊﾟﾀｰﾝ1
'''                    If dCmax > 0 And dCmin > 0 Then
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        'TOP側検索
'''                        If recXSDCS("TBKBNCS").Value = "T" Then
'''                            sql = sql & "where tbkbncs = 'T' and "
'''                            sql = sql & "      xtalcs = '" & CRYNUM & "' and "
'''                            sql = sql & "      inposcs <= " & recXSDCS("INPOSCS").Value & " and "
'''                            sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                            sql = sql & "      CRYRESCSCS <> '0' "
'''                            sql = sql & "order by inposcs desc"
'''                        'BOT側検索
'''                        Else
'''                            sql = sql & "where tbkbncs = 'B' and "
'''                            sql = sql & "      xtalcs = '" & CRYNUM & "' and "
'''                            sql = sql & "      inposcs >= " & recXSDCS("INPOSCS").Value & " and "
'''                            sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                            sql = sql & "      CRYRESCSCS <> '0' "
'''                            sql = sql & "order by inposcs asc"
'''                        End If
'''                    '反映ﾊﾟﾀｰﾝ2
'''                    Else
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        sql = sql & "where xtalcs = '" & CRYNUM & "' and "
'''                        sql = sql & "      inposcs >= " & recXSDCS("INPOSCS").Value & " and "
'''                        sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                        sql = sql & "      CRYRESCSCS <> '0' "
'''                        sql = sql & "order by inposcs asc"
'''                    End If
'''
'''                    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                    If rs.RecordCount > 0 Then
'''                        iSmpNo = rs("CRYSMPLIDCSCS")
'''                    '取得出来なかった場合はTOP側検索
'''                    Else
'''                        sql = "select CRYSMPLIDCSCS from XSDCS "
'''                        sql = sql & "where xtalcs = '" & CRYNUM & "' and "
'''                        sql = sql & "      inposcs <= " & recXSDCS("INPOSCS").Value & " and "
'''                        sql = sql & "      (CRYINDCSCS = '1' or CRYINDCSCS = '2') and "
'''                        sql = sql & "      CRYRESCSCS <> '0' "
'''                        sql = sql & "order by inposcs desc"
'''
'''                        Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''                        If rs2.RecordCount = 0 Then
'''                            sErrMsg = GetMsgStr("ECLC3") & " <ｻﾝﾌﾟﾙNo> "
'''                            Set rs = Nothing
'''                            Set rs2 = Nothing
'''                            GoTo PROC_EXIT
'''                        End If
'''                        iSmpNo = rs2("CRYSMPLIDCSCS")
'''                        Set rs2 = Nothing
'''                    End If
'''                    Set rs = Nothing
'''                End If
                
                'ｻﾝﾌﾟﾙ位置＆ｻﾝﾌﾟﾙ測定値取得
                sql = "select POSITION, CSMEAS from TBCMJ004 "
                sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
                sql = sql & "      SMPLNO = " & iSmpNo & " "
                sql = sql & "order by TRANCNT desc"
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ｻﾝﾌﾟﾙ測定値> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sSamplePos = rs("POSITION")       'ｻﾝﾌﾟﾙ位置
                tCsSuitei.sResCs = rs("CSMEAS")             'ｻﾝﾌﾟﾙ測定値
                Set rs = Nothing
                
                '④ﾁｬｰｼﾞ量
                '⑤TOP重量
                sql = "select SUICHARGE, WGHTTOC1, PUTCUTWC1 from XSDC1 "
                sql = sql & "where XTALC1 = '" & CRYNUM & "' "
                
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                'ﾃﾞｰﾀ不正
                If (IsNull(rs("SUICHARGE")) Or IsNull(rs("WGHTTOC1")) Or IsNull(rs("PUTCUTWC1"))) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                
                tCsSuitei.sSiWeight = rs("SUICHARGE")       '推定ﾁｬｰｼﾞ量
                tCsSuitei.sTopWT = CLng(rs("WGHTTOC1")) + CLng(rs("PUTCUTWC1"))     'TOP重量
                Set rs = Nothing
                '｢推定ﾁｬｰｼﾞ量=0｣or｢推定ﾁｬｰｼﾞ量≦TOP重量｣の場合はｴﾗｰとする
                If CLng(tCsSuitei.sSiWeight) = 0 Or _
                   (CLng(tCsSuitei.sSiWeight) <= CLng(tCsSuitei.sTopWT)) Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ﾁｬｰｼﾞ量> "
                    GoTo proc_exit
                End If
                
                '⑥直径
                sql = "select HSXD1CEN from TBCME018 "
                sql = sql & "where HINBAN = '" & HIN.hinban & "' "
                sql = sql & "and MNOREVNO = " & HIN.mnorevno & " "
                sql = sql & "and FACTORY = '" & HIN.factory & "' "
                sql = sql & "and OPECOND = '" & HIN.opecond & "' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <直径> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sUpDm = rs("HSXD1CEN")            '品SX直径1中心
                
                '⑦ｶｰﾎﾞﾝ偏析係数
                sql = "select CTR01A9 from KODA9 "
                sql = sql & "where SYSCA9 = 'K' "
                sql = sql & "and SHUCA9 = 'AP' "
                sql = sql & "and CODEA9 = '1' "
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    sErrMsg = GetMsgStr("ECLC3") & " <ｶｰﾎﾞﾝ偏析係数> "
                    Set rs = Nothing
                    GoTo proc_exit
                End If
                tCsSuitei.sCsHenseki = rs("CTR01A9")        'ｶｰﾎﾞﾝ偏析係数
                
                '⑧Csﾌﾞﾛｯｸ推定値計算
                If Not GetCsSuiteiMain(tCsSuitei, dCsSuitei) Then
                    sErrMsg = GetMsgStr("ECLC3")
                    GoTo proc_exit
                End If
                .Fields("SXLCS_BSUIMEAS").Value = dCsSuitei
            End If
            ''Csﾌﾞﾛｯｸ推定値計算対応　06/04/20 ooba END ========================================>
        End If
    End With

    getTBCMJ004 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ004 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶OSF実績(TBCMJ005)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :j               , I  ,Integer           , OSF No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶OSF実績(TBCMJ005)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ005(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005"
    
    getTBCMJ005 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLOSF_SMPPOS").Value = -1             'OSFサンプル測定位置(SXL測定情報)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'OSFx結晶欠陥測定位置
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'OSFx熱処理法
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'OSFx結晶欠陥測定条件+選択ET代
        .Fields("SXLOSF" & j & "_CALCMAX").Value = -1       'OSFxSXL計算結果 Max_x
        .Fields("SXLOSF" & j & "_CALCAVE").Value = -1       'OSFxSXL計算結果 Ave_x
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        If j = 1 Then
            .Fields("SXLOSF1_PTNJUDGRES").Value = ""            'OSF1パターン判定結果
        End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLOSF1_SMPPOS").Value = -1            'SXLOSFサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLOSF" & j & "_KKSP").Value = ""          'SXLOSFx結晶欠陥確定位置
        .Fields("SXLOSF" & j & "_NETU").Value = ""          'SXLOSFx熱処理法
        .Fields("SXLOSF" & j & "_KKSET").Value = ""         'SXLOSFx結晶欠陥測定条件+選択ET代
        .Fields("SXLOSF" & j & "_MEAS1").Value = -1         'SXLOSFx測定点1
        .Fields("SXLOSF" & j & "_MEAS2").Value = -1         'SXLOSFx測定点2
        .Fields("SXLOSF" & j & "_MEAS3").Value = -1         'SXLOSFx測定点3
        .Fields("SXLOSF" & j & "_MEAS4").Value = -1         'SXLOSFx測定点4
        .Fields("SXLOSF" & j & "_MEAS5").Value = -1         'SXLOSFx測定点5
        .Fields("SXLOSF" & j & "_MEAS6").Value = -1         'SXLOSFx測定点6
        .Fields("SXLOSF" & j & "_MEAS7").Value = -1         'SXLOSFx測定点7
        .Fields("SXLOSF" & j & "_MEAS8").Value = -1         'SXLOSFx測定点8
        .Fields("SXLOSF" & j & "_MEAS9").Value = -1         'SXLOSFx測定点9
        .Fields("SXLOSF" & j & "_MEAS10").Value = -1        'SXLOSFx測定点10
        .Fields("SXLOSF" & j & "_MEAS11").Value = -1        'SXLOSFx測定点11
        .Fields("SXLOSF" & j & "_MEAS12").Value = -1        'SXLOSFx測定点12
        .Fields("SXLOSF" & j & "_MEAS13").Value = -1        'SXLOSFx測定点13
        .Fields("SXLOSF" & j & "_MEAS14").Value = -1        'SXLOSFx測定点14
        .Fields("SXLOSF" & j & "_MEAS15").Value = -1        'SXLOSFx測定点15
        .Fields("SXLOSF" & j & "_MEAS16").Value = -1        'SXLOSFx測定点16
        .Fields("SXLOSF" & j & "_MEAS17").Value = -1        'SXLOSFx測定点17
        .Fields("SXLOSF" & j & "_MEAS18").Value = -1        'SXLOSFx測定点18
        .Fields("SXLOSF" & j & "_MEAS19").Value = -1        'SXLOSFx測定点19
        .Fields("SXLOSF" & j & "_MEAS20").Value = -1        'SXLOSFx測定点20
    End With
    
    '-------------------- TBCMJ005の読み込み(OSF1～4) ----------------------------------------
    If (recXSDCS("CRYINDL" & j & "CS").Value <> "0") And (recXSDCS("CRYRESL" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ005 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
    
        'TBCMX001
        With recX001
            If .Fields("SXLOSF_SMPPOS").Value = -1 Then
                .Fields("SXLOSF_SMPPOS").Value = rs("POSITION")         'OSFサンプル測定位置(SXL測定情報)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'OSFx結晶欠陥測定位置
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'OSFx熱処理法
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'OSFx結晶欠陥測定条件+選択ET代
            .Fields("SXLOSF" & j & "_CALCMAX").Value = rs("CALCMAX")    'OSFxSXL計算結果 Max_x
            .Fields("SXLOSF" & j & "_CALCAVE").Value = rs("CALCAVE")    'OSFxSXL計算結果 Ave_x
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            'OSF1パターン判定結果
            If j = 1 Then
                If IsNull(rs("PTNJUDGRES")) = True Then
                    .Fields("SXLOSF1_PTNJUDGRES").Value = " "
                Else
                    .Fields("SXLOSF1_PTNJUDGRES").Value = rs("PTNJUDGRES")
                End If
            End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLOSF1_SMPPOS").Value = -1 Then
                .Fields("SXLOSF1_SMPPOS").Value = rs("POSITION")        'SXLOSFサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLOSF" & j & "_KKSP").Value = rs("KKSP")          'SXLOSFx結晶欠陥確定位置
            .Fields("SXLOSF" & j & "_NETU").Value = rs("HTPRC")         'SXLOSFx熱処理法
            .Fields("SXLOSF" & j & "_KKSET").Value = rs("KKSET")        'SXLOSFx結晶欠陥測定条件+選択ET代
            .Fields("SXLOSF" & j & "_MEAS1").Value = rs("MEAS1")        'SXLOSFx測定点1
            .Fields("SXLOSF" & j & "_MEAS2").Value = rs("MEAS2")        'SXLOSFx測定点2
            .Fields("SXLOSF" & j & "_MEAS3").Value = rs("MEAS3")        'SXLOSFx測定点3
            .Fields("SXLOSF" & j & "_MEAS4").Value = rs("MEAS4")        'SXLOSFx測定点4
            .Fields("SXLOSF" & j & "_MEAS5").Value = rs("MEAS5")        'SXLOSFx測定点5
            .Fields("SXLOSF" & j & "_MEAS6").Value = rs("MEAS6")        'SXLOSFx測定点6
            .Fields("SXLOSF" & j & "_MEAS7").Value = rs("MEAS7")        'SXLOSFx測定点7
            .Fields("SXLOSF" & j & "_MEAS8").Value = rs("MEAS8")        'SXLOSFx測定点8
            .Fields("SXLOSF" & j & "_MEAS9").Value = rs("MEAS9")        'SXLOSFx測定点9
            .Fields("SXLOSF" & j & "_MEAS10").Value = rs("MEAS10")      'SXLOSFx測定点10
            .Fields("SXLOSF" & j & "_MEAS11").Value = rs("MEAS11")      'SXLOSFx測定点11
            .Fields("SXLOSF" & j & "_MEAS12").Value = rs("MEAS12")      'SXLOSFx測定点12
            .Fields("SXLOSF" & j & "_MEAS13").Value = rs("MEAS13")      'SXLOSFx測定点13
            .Fields("SXLOSF" & j & "_MEAS14").Value = rs("MEAS14")      'SXLOSFx測定点14
            .Fields("SXLOSF" & j & "_MEAS15").Value = rs("MEAS15")      'SXLOSFx測定点15
            .Fields("SXLOSF" & j & "_MEAS16").Value = rs("MEAS16")      'SXLOSFx測定点16
            .Fields("SXLOSF" & j & "_MEAS17").Value = rs("MEAS17")      'SXLOSFx測定点17
            .Fields("SXLOSF" & j & "_MEAS18").Value = rs("MEAS18")      'SXLOSFx測定点18
            .Fields("SXLOSF" & j & "_MEAS19").Value = rs("MEAS19")      'SXLOSFx測定点19
            .Fields("SXLOSF" & j & "_MEAS20").Value = rs("MEAS20")      'SXLOSFx測定点20
        End With
        Set rs = Nothing
    End If

    getTBCMJ005 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶BMD実績(TBCMJ008)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :j               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶BMD実績(TBCMJ008)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ008(CRYNUM As String, recXSDCS As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim dMeas(9)    As Double
    Dim strMeasPos  As String
    Dim iRet        As Integer
    Dim wComp       As Double
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ008"
    
    getTBCMJ008 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'BMDサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'BMDx結晶欠陥測定位置
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'BMDx熱処理法
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'BMDx結晶欠陥測定条件＋選択ET代
        .Fields("SXLBMD" & j & "_CALCMAX").Value = -1       'BMDxSXL計算結果 Max
        .Fields("SXLBMD" & j & "_CALCAVE").Value = -1       'BMDxSXL計算結果 Ave
        .Fields("SXLBMD" & j & "_CALCMIN").Value = -1       'BMDxSXL計算結果 Min
        .Fields("SXLBMD" & j & "_CALCMB").Value = -1        'BMDxSXL計算結果 面内分布
    End With
        
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("SXLBMD_SMPPOS").Value = -1             'SXLBMDサンプル測定位置(SXL位置情報)
        End If
        .Fields("SXLBMD" & j & "_KKSP").Value = ""          'SXLBMD1結晶欠陥測定位置
        .Fields("SXLBMD" & j & "_NETU").Value = ""          'SXLBMD1熱処理法
        .Fields("SXLBMD" & j & "_KKSET").Value = ""         'SXLBMD1結晶欠陥測定条件+選択ET代
        .Fields("SXLBMD" & j & "_MEAS1").Value = -1         'SXLBMD1測定点1
        .Fields("SXLBMD" & j & "_MEAS2").Value = -1         'SXLBMD1測定点2
        .Fields("SXLBMD" & j & "_MEAS3").Value = -1         'SXLBMD1測定点3
        .Fields("SXLBMD" & j & "_MEAS4").Value = -1         'SXLBMD1測定点4
        .Fields("SXLBMD" & j & "_MEAS5").Value = -1         'SXLBMD1測定点5
    End With
    
    '-------------------- TBCMJ008の読み込み(BMD1～3) ----------------------------------------
    If (recXSDCS("CRYINDB" & j & "CS").Value <> "0") And (recXSDCS("CRYRESB" & j & "CS").Value <> "0") Then
        sql = "select * from TBCMJ008 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDB" & j & "CS").Value & " and "
        sql = sql & "      TRANCOND = '" & j & "' "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'BMDサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'BMDx結晶欠陥測定位置
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'BMDx熱処理法
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'BMDx結晶欠陥測定条件＋選択ET代
            .Fields("SXLBMD" & j & "_CALCMAX").Value = rs("MEASMAX")    'BMDxSXL計算結果 Max
            .Fields("SXLBMD" & j & "_CALCAVE").Value = rs("MEASAVE")    'BMDxSXL計算結果 Ave
'            .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL計算結果 面内分布
            If IsNull(rs("BMDMNBUNP")) = False Then .Fields("SXLBMD" & j & "_CALCMB").Value = rs("BMDMNBUNP")   'BMDxSXL計算結果 面内分布
        End With
            
        'TBCMX002
        With recX002
            If .Fields("SXLBMD_SMPPOS").Value = -1 Then
                .Fields("SXLBMD_SMPPOS").Value = rs("POSITION")         'SXLBMDサンプル測定位置(SXL位置情報)
            End If
            .Fields("SXLBMD" & j & "_KKSP").Value = rs("KKSP")          'SXLBMDx結晶欠陥測定位置
            .Fields("SXLBMD" & j & "_NETU").Value = rs("HTPRC")         'SXLBMDx熱処理法
            .Fields("SXLBMD" & j & "_KKSET").Value = rs("KKSET")        'SXLBMDx結晶欠陥測定条件+選択ET代
            .Fields("SXLBMD" & j & "_MEAS1").Value = rs("MEAS1")        'SXLBMDx測定点1
            .Fields("SXLBMD" & j & "_MEAS2").Value = rs("MEAS2")        'SXLBMDx測定点2
            .Fields("SXLBMD" & j & "_MEAS3").Value = rs("MEAS3")        'SXLBMDx測定点3
            .Fields("SXLBMD" & j & "_MEAS4").Value = rs("MEAS4")        'SXLBMDx測定点4
            .Fields("SXLBMD" & j & "_MEAS5").Value = rs("MEAS5")        'SXLBMDx測定点5
        End With
        Set rs = Nothing
    
        'BMD最小値の取得 2003/05/31 tuku                START
        dMeas(0) = recX002.Fields("SXLBMD" & j & "_MEAS1").Value
        dMeas(1) = recX002.Fields("SXLBMD" & j & "_MEAS2").Value
        dMeas(2) = recX002.Fields("SXLBMD" & j & "_MEAS3").Value
        dMeas(3) = recX002.Fields("SXLBMD" & j & "_MEAS4").Value
        dMeas(4) = recX002.Fields("SXLBMD" & j & "_MEAS5").Value
        ''結晶欠陥測定位置コード
        strMeasPos = Trim(recX002.Fields("SXLBMD" & j & "_KKSP").Value)
        ''最小値を計算する。
        iRet = getSXLBMDMIN(wComp, strMeasPos, dMeas)
        ''計算結果を格納する
        recX001.Fields("SXLBMD" & j & "_CALCMIN").Value = wComp
    End If

    getTBCMJ008 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ008 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :GD実績(TBCMJ006)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :GD実績(TBCMJ006)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ006(CRYNUM As String, recXSDCS As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006"
    
    getTBCMJ006 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLGD_SMPPOS").Value = -1                  'GDサンプル測定位置(SXL位置情報)
        .Fields("SXLGD_MSRSDEN").Value = -1                 'SXLGD_測定結果 Den
        .Fields("SXLGD_MSRSLDL").Value = -1                 'SXLGD_測定結果 L/DL
        .Fields("SXLGD_MSRSDVD2").Value = -1                'SXLGD_測定結果 DVD2
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLGD_SMPPOS").Value = -1                                  'SXLGDサンプル測定位置(SXL位置情報)
        For i = 1 To 15
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = -1       'SXLGD_測定値xx L/DL1
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = -1       'SXLGD_測定値xx L/DL2
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = -1       'SXLGD_測定値xx L/DL3
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = -1       'SXLGD_測定値xx L/DL4
            .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = -1       'SXLGD_測定値xx L/DL5
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = -1       'SXLGD_測定値xx Den1
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = -1       'SXLGD_測定値xx Den2
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = -1       'SXLGD_測定値xx Den3
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = -1       'SXLGD_測定値xx Den4
            .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = -1       'SXLGD_測定値xx Den5
        Next
    End With
        
    '-------------------- TBCMJ006の読み込み(GD) ----------------------------------------
    If (recXSDCS("CRYINDGDCS").Value <> "0") And (recXSDCS("CRYRESGDCS").Value <> "0") Then
        sql = "select * from TBCMJ006 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDGDCS").Value & " "
        sql = sql & "order by TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum = 1"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'GDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")              'SXLGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")              'SXLGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")            'SXLGD_測定結果 DVD2
        End With
            
        'TBCMX002
        With recX002
            .Fields("SXLGD_SMPPOS").Value = rs("POSITION")                                                      'SXLGDサンプル測定位置(SXL位置情報)
            For i = 1 To 15
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_測定値xx L/DL1
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_測定値xx L/DL2
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_測定値xx L/DL3
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_測定値xx L/DL4
                .Fields("SXLGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_測定値xx L/DL5
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_測定値xx Den1
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_測定値xx Den2
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_測定値xx Den3
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_測定値xx Den4
                .Fields("SXLGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_測定値xx Den5
            Next
        End With
        Set rs = Nothing
    End If

    getTBCMJ006 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :LT実績(TBCMJ007)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :ChkHin          , I  ,tFullHinban       , LT仕様取得用品番　05/12/05 ooba
'          :i               , I  ,Integer           , BMD No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :LT実績(TBCMJ007)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ007(CRYNUM As String, recXSDCS As c_cmzcrec, ChkHin As tFullHinban, i As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim j           As Integer      '                       '05/12/05 ooba START =======>
    Dim rs2         As OraDynaset   '
    Dim sql2        As String       '
    Dim iRet        As Integer      '
    Dim iTmpMes(9)  As Integer      'LT実績ﾃﾞｰﾀ(1～10)
    Dim iCalcMeas   As Integer      'LT計算結果
    Dim sIchi       As String       '品SXLﾀｲﾑ測定位置_位
    Dim iOldFlg     As Integer      '旧ﾃﾞｰﾀﾌﾗｸﾞ             '05/12/05 ooba END =========>
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ007"
    
    getTBCMJ007 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("SXLLT_SMPPOS").Value = -1                  'LTサンプル測定位置(SXL位置情報)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_測定値 ピーク値
        .Fields("SXLLT_CALCMEAS").Value = -1                'SXLLT_計算結果
    End With
        
    'TBCMX002
    With recX002
        .Fields("SXLT_SMPPOS").Value = -1                   'SXLLTサンプル測定位置(SXL位置情報)
        .Fields("SXLLT_MEASPEAK").Value = -1                'SXLLT_測定値 ピーク値
        .Fields("SXLLT_MEAS1").Value = -1                   'SXLLT_測定値1
        .Fields("SXLLT_MEAS2").Value = -1                   'SXLLT_測定値2
        .Fields("SXLLT_MEAS3").Value = -1                   'SXLLT_測定値3
        .Fields("SXLLT_MEAS4").Value = -1                   'SXLLT_測定値4
        .Fields("SXLLT_MEAS5").Value = -1                   'SXLLT_測定値5
    End With
        
    'BOT側のみﾃﾞｰﾀ取得
    If i <> 1 Then
        '-------------------- TBCMJ007の読み込み(LT) ----------------------------------------
        If (recXSDCS("CRYINDTCS").Value <> "0") And (recXSDCS("CRYRESTCS").Value <> "0") Then
            sql = "select * from TBCMJ007 "
            sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
            sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDTCS").Value & " "
            sql = sql & "order by TRANCNT desc"
            sql = "select * from (" & sql & ") where rownum = 1"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            ''LT実績送信ﾃﾞｰﾀ変更　05/12/05 ooba START ==================================>
            If IsNull(rs("LTSPIFLG")) Then iOldFlg = 1 Else iOldFlg = 0
            
            '初期化
            iCalcMeas = -1
            For j = 0 To 9
                iTmpMes(j) = -1
            Next j
            
            If Not IsNull(rs("MEAS1")) Then iTmpMes(0) = rs("MEAS1")
            If Not IsNull(rs("MEAS2")) Then iTmpMes(1) = rs("MEAS2")
            If Not IsNull(rs("MEAS3")) Then iTmpMes(2) = rs("MEAS3")
            If Not IsNull(rs("MEAS4")) Then iTmpMes(3) = rs("MEAS4")
            If Not IsNull(rs("MEAS5")) Then iTmpMes(4) = rs("MEAS5")
            If Not IsNull(rs("MEAS6")) Then iTmpMes(5) = rs("MEAS6")
            If Not IsNull(rs("MEAS7")) Then iTmpMes(6) = rs("MEAS7")
            If Not IsNull(rs("MEAS8")) Then iTmpMes(7) = rs("MEAS8")
            If Not IsNull(rs("MEAS9")) Then iTmpMes(8) = rs("MEAS9")
            If Not IsNull(rs("MEAS10")) Then iTmpMes(9) = rs("MEAS10")
            
            '10点測定の場合
            If iOldFlg = 0 Then
                sql2 = "select HSXLTSPI from TBCME019"
                sql2 = sql2 & " where HINBAN = '" & ChkHin.hinban & "'"
                sql2 = sql2 & " and MNOREVNO = " & ChkHin.mnorevno
                sql2 = sql2 & " and FACTORY = '" & ChkHin.factory & "'"
                sql2 = sql2 & " and OPECOND = '" & ChkHin.opecond & "'"
                Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
                If rs2.RecordCount = 0 Then
                    Set rs2 = Nothing
                    GoTo proc_exit
                End If
                If Not IsNull(rs2("HSXLTSPI")) Then sIchi = rs2("HSXLTSPI") Else sIchi = ""
                Set rs2 = Nothing
            End If
            
            '計算結果取得
            iRet = KNS_CalculateMeasResult_LT(iCalcMeas, iTmpMes(), sIchi, iOldFlg)
            ''LT実績送信ﾃﾞｰﾀ変更　05/12/05 ooba END ====================================>
            
            'TBCMX001
            With recX001
                .Fields("SXLLT_SMPPOS").Value = rs("POSITION")          'LTサンプル測定位置(SXL位置情報)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_測定値 ピーク値
'                .Fields("SXLLT_CALCMEAS").Value = rs("CALCMEAS")        'SXLLT_計算結果
                .Fields("SXLLT_CALCMEAS").Value = iCalcMeas             'SXLLT_計算結果　05/12/05 ooba
            End With
                
            'TBCMX002
            With recX002
                .Fields("SXLT_SMPPOS").Value = rs("POSITION")           'SXLLTサンプル測定位置(SXL位置情報)
                .Fields("SXLLT_MEASPEAK").Value = rs("MEASPEAK")        'SXLLT_測定値 ピーク値
'                .Fields("SXLLT_MEAS1").Value = rs("MEAS1")              'SXLLT_測定値1
'                .Fields("SXLLT_MEAS2").Value = rs("MEAS2")              'SXLLT_測定値2
'                .Fields("SXLLT_MEAS3").Value = rs("MEAS3")              'SXLLT_測定値3
'                .Fields("SXLLT_MEAS4").Value = rs("MEAS4")              'SXLLT_測定値4
'                .Fields("SXLLT_MEAS5").Value = rs("MEAS5")              'SXLLT_測定値5

                ''LT実績登録変更　05/12/05 ooba START =====================================>
                '旧ﾃﾞｰﾀ
                If iOldFlg = 1 Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(3)           'SXLLT_測定値4
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(4)           'SXLLT_測定値5
                '3:CE,Inside3mm
                ElseIf sIchi = "3" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(7)           'SXLLT_測定値8
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(8)           'SXLLT_測定値9
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(9)           'SXLLT_測定値10
                '5:CE,Inside5mm
                ElseIf sIchi = "5" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(4)           'SXLLT_測定値5
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(5)           'SXLLT_測定値6
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(6)           'SXLLT_測定値7
                'A:CE,Inside10mm
                ElseIf sIchi = "A" Then
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_測定値4
                'その他
                Else
'                    Set rs = Nothing
'                    GoTo proc_exit
                    'その他の場合は｢A:CE,Inside10mm｣とする　05/12/21 ooba
                    .Fields("SXLLT_MEAS1").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS2").Value = iTmpMes(0)           'SXLLT_測定値1
                    .Fields("SXLLT_MEAS3").Value = iTmpMes(1)           'SXLLT_測定値2
                    .Fields("SXLLT_MEAS4").Value = iTmpMes(2)           'SXLLT_測定値3
                    .Fields("SXLLT_MEAS5").Value = iTmpMes(3)           'SXLLT_測定値4
                End If
                ''LT実績登録変更　05/12/05 ooba END =======================================>
            End With
            Set rs = Nothing
        End If
    End If

    getTBCMJ007 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ007 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFOi実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFOi実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFOi(recXSDCW As c_cmzcrec, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim WOi     As W_OI                     'WFOI構造体
    Dim HWFONKHN As String                  '品ＷＦ酸素濃度検査頻度＿抜　04/04/15 ooba
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFOi"
    
    getTBCMY013WFOi = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("WFOI_SMPPOS").Value = -1                   'WFOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFOI_NETSU").Value = ""                    'WFOI_熱処理条件
        .Fields("WFOI_ET").Value = ""                       'WFOI_エッチング条件
        .Fields("WFOI_MES").Value = ""                      'WFOI_計測方法
        .Fields("WFOI_MESDATA1").Value = -1                 'WFOI_測定データその１
        .Fields("WFOI_MESDATA2").Value = -1                 'WFOI_測定データその２
        .Fields("WFOI_MESDATA3").Value = -1                 'WFOI_測定データその３
        .Fields("WFOI_MESDATA4").Value = -1                 'WFOI_測定データその４
        .Fields("WFOI_MESDATA5").Value = -1                 'WFOI_測定データその５
        .Fields("WFOI_MESDATA6").Value = -1                 'WFOI_測定データその６
        .Fields("WFOI_MESDATA7").Value = -1                 'WFOI_測定データその７
        .Fields("WFOI_MESDATA8").Value = -1                 'WFOI_測定データその８
        .Fields("WFOI_MESDATA9").Value = -1                 'WFOI_測定データその９
        .Fields("WFOI_MESDATA10").Value = -1                'WFOI_測定データその１０
        .Fields("WFOI_ORG").Value = -1                      'WFOI_ORG計算結果
    
        '-------------------- TBCMY013の読み込み(WFOi) ----------------------------------------
        If (recXSDCW("WFINDOICW").Value <> "0") And (recXSDCW("WFRESOICW").Value <> "0") Then
            sql = "select * from TBCMY013 "
            sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDOICW").Value & "' and "
            sql = sql & "      SPEC = 'OI'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("WFOI_SMPPOS").Value = recXSDCW("INPOSCW").Value        'WFOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFOI_NETSU").Value = rs("NETSU")                       'WFOI_熱処理条件
            .Fields("WFOI_ET").Value = rs("ET")                             'WFOI_エッチング条件
            .Fields("WFOI_MES").Value = rs("MES")                           'WFOI_計測方法
            .Fields("WFOI_MESDATA1").Value = rs("MESDATA1")                 'WFOI_測定データその１
            .Fields("WFOI_MESDATA2").Value = rs("MESDATA2")                 'WFOI_測定データその２
            .Fields("WFOI_MESDATA3").Value = rs("MESDATA3")                 'WFOI_測定データその３
            .Fields("WFOI_MESDATA4").Value = rs("MESDATA4")                 'WFOI_測定データその４
            .Fields("WFOI_MESDATA5").Value = rs("MESDATA5")                 'WFOI_測定データその５
            .Fields("WFOI_MESDATA6").Value = rs("MESDATA6")                 'WFOI_測定データその６
            .Fields("WFOI_MESDATA7").Value = rs("MESDATA7")                 'WFOI_測定データその７
            .Fields("WFOI_MESDATA8").Value = rs("MESDATA8")                 'WFOI_測定データその８
            .Fields("WFOI_MESDATA9").Value = rs("MESDATA9")                 'WFOI_測定データその９
            .Fields("WFOI_MESDATA10").Value = rs("MESDATA10")               'WFOI_測定データその１０
            Set rs = Nothing
            
            'WFOi_ORG
            sql = "select E025.HWFONSPH, E019.HSXONSPT, E019.HSXONSPI, E025.HWFONHWT, E025.HWFONHWS, E025.HWFONMCL, E025.HWFONKHN "
            sql = sql & "from TBCME025 E025, TBCME019 E019 where "
            sql = sql & " E025.HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " E025.MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " E025.FACTORY = '" & HIN.factory & "' and "
            sql = sql & " E025.OPECOND = '" & HIN.opecond & "' and "
            sql = sql & " E019.HINBAN = E025.HINBAN and "
            sql = sql & " E019.MNOREVNO = E025.MNOREVNO and "
            sql = sql & " E019.FACTORY = E025.FACTORY and "
            sql = sql & " E019.OPECOND = E025.OPECOND"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            
            WOi.GuaranteeOi.cMeth = rs("HWFONSPH")                      '品ＷＦ酸素濃度測定位置＿方
            WOi.GuaranteeOi.cCount = rs("HSXONSPT")                     '品ＳＸ酸素濃度測定位置＿点
            WOi.GuaranteeOi.cPos = rs("HSXONSPI")                       '品ＳＸ酸素濃度測定位置＿位
            WOi.GuaranteeOi.cObj = rs("HWFONHWT")                       '品ＷＦ酸素濃度保証方法＿対
            WOi.GuaranteeOi.cJudg = rs("HWFONHWS")                      '品ＷＦ酸素濃度保証方法＿処
            WOi.GuaranteeCal = rs("HWFONMCL")                           '品ＷＦ酸素濃度面内計算
            If IsNull(rs("HWFONKHN")) = False Then HWFONKHN = rs("HWFONKHN")    '品ＷＦ酸素濃度検査頻度＿抜　04/04/15 ooba
            Set rs = Nothing
                
            WOi.Oi(0) = NtoZ2(.Fields("WFOI_MESDATA1").Value)           'Oi測定値1
            WOi.Oi(1) = NtoZ2(.Fields("WFOI_MESDATA2").Value)           'Oi測定値2
            WOi.Oi(2) = NtoZ2(.Fields("WFOI_MESDATA3").Value)           'Oi測定値3
            WOi.Oi(3) = NtoZ2(.Fields("WFOI_MESDATA4").Value)           'Oi測定値4
            WOi.Oi(4) = NtoZ2(.Fields("WFOI_MESDATA5").Value)           'Oi測定値5
            WOi.Oi(5) = NtoZ2(.Fields("WFOI_MESDATA6").Value)           'Oi測定値6
            WOi.Oi(6) = NtoZ2(.Fields("WFOI_MESDATA7").Value)           'Oi測定値7
            WOi.Oi(7) = NtoZ2(.Fields("WFOI_MESDATA8").Value)           'Oi測定値8
            WOi.Oi(8) = NtoZ2(.Fields("WFOI_MESDATA9").Value)           'Oi測定値9
            WOi.Oi(9) = NtoZ2(.Fields("WFOI_MESDATA10").Value)          'Oi測定値10
                
            .Fields("WFOI_ORG").Value = WFCORGCal(WOi.Oi(), WOi.GuaranteeOi, WOi.GuaranteeCal)      'WFOI_ORG計算結果
            
            '保証方法="H"、かつ、WFOI_ORG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
'            If (WOi.GuaranteeOi.cJudg = "H") And (.Fields("WFOI_ORG").Value = -1) Then GoTo proc_exit
            '保証方法ﾁｪｯｸの追加　04/04/15 ooba
            If ((WOi.GuaranteeOi.cJudg = "H") And CheckKHN(HWFONKHN, 2, sPos)) _
                And (.Fields("WFOI_ORG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMY013WFOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFRs実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFRs実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFRs(recXSDCW As c_cmzcrec, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim WRs     As W_RES                    'WFRs構造体
    Dim HWFRKHNN As String                  '品ＷＦ比抵抗検査頻度＿抜   04/04/15 ooba
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFRs"
    
    getTBCMY013WFRs = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX001
        .Fields("WFRS_SMPPOS").Value = -1                   'WFRSｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFRS_NETSU").Value = ""                    'WFRS_熱処理条件
        .Fields("WFRS_ET").Value = ""                       'WFRS_エッチング条件
        .Fields("WFRS_MES").Value = ""                      'WFRS_計測方法
        .Fields("WFRS_MESDATA1").Value = -1                 'WFRS_測定データその１
        .Fields("WFRS_MESDATA2").Value = -1                 'WFRS_測定データその２
        .Fields("WFRS_MESDATA3").Value = -1                 'WFRS_測定データその３
        .Fields("WFRS_MESDATA4").Value = -1                 'WFRS_測定データその４
        .Fields("WFRS_MESDATA5").Value = -1                 'WFRS_測定データその５
        .Fields("WFRS_RRG").Value = -1                      'WFRS_RRG計算結果
    
        '-------------------- TBCMY013の読み込み(WFRs) ----------------------------------------
        If (recXSDCW("WFINDRSCW").Value <> "0") And (recXSDCW("WFRESRS1CW").Value <> "0") Then
            sql = "select * from TBCMY013 "
            sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDRSCW").Value & "' and "
            sql = sql & "      SPEC = 'RES'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
            .Fields("WFRS_SMPPOS").Value = recXSDCW("INPOSCW").Value        'WFRSｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFRS_NETSU").Value = rs("NETSU")                       'WFRS_熱処理条件
            .Fields("WFRS_ET").Value = rs("ET")                             'WFRS_エッチング条件
            .Fields("WFRS_MES").Value = rs("MES")                           'WFRS_計測方法
            .Fields("WFRS_MESDATA1").Value = rs("MESDATA1")                 'WFRS_測定データその１
            .Fields("WFRS_MESDATA2").Value = rs("MESDATA2")                 'WFRS_測定データその２
            .Fields("WFRS_MESDATA3").Value = rs("MESDATA3")                 'WFRS_測定データその３
            .Fields("WFRS_MESDATA4").Value = rs("MESDATA4")                 'WFRS_測定データその４
            .Fields("WFRS_MESDATA5").Value = rs("MESDATA5")                 'WFRS_測定データその５
            Set rs = Nothing
                
            'WFRs_RRG
            sql = "select E021.HWFRSPOH, E018.HSXRSPOT, E018.HSXRSPOI, E021.HWFRHWYT, E021.HWFRHWYS, E021.HWFRMCAL, E021.HWFRKHNN "
            sql = sql & "from TBCME021 E021, TBCME018 E018 where "
            sql = sql & " E021.HINBAN = '" & HIN.hinban & "' and "
            sql = sql & " E021.MNOREVNO = " & HIN.mnorevno & " and "
            sql = sql & " E021.FACTORY = '" & HIN.factory & "' and "
            sql = sql & " E021.OPECOND = '" & HIN.opecond & "' and "
            sql = sql & " E018.HINBAN = E021.HINBAN and "
            sql = sql & " E018.MNOREVNO = E021.MNOREVNO and "
            sql = sql & " E018.FACTORY = E021.FACTORY and "
            sql = sql & " E018.OPECOND = E021.OPECOND"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                Set rs = Nothing
                GoTo proc_exit
            End If
                
            WRs.GuaranteeRes.cMeth = rs("HWFRSPOH")                     ' 品ＷＦ比抵抗測定位置＿方
            WRs.GuaranteeRes.cCount = rs("HSXRSPOT")                    ' 品ＳＸ比抵抗測定位置＿点
            WRs.GuaranteeRes.cPos = rs("HSXRSPOI")                      ' 品ＳＸ比抵抗測定位置＿位
            WRs.GuaranteeRes.cObj = rs("HWFRHWYT")                      ' 品ＷＦ比抵抗保証方法＿対
            WRs.GuaranteeRes.cJudg = rs("HWFRHWYS")                     ' 品ＷＦ比抵抗保証方法＿処
            WRs.GuaranteeCal = rs("HWFRMCAL")                           ' 品ＷＦ比抵抗面内計算
            If IsNull(rs("HWFRKHNN")) = False Then HWFRKHNN = rs("HWFRKHNN")    ' 品ＷＦ比抵抗検査頻度＿抜　04/04/15 ooba
            Set rs = Nothing
                
            WRs.Res(0) = NtoZ2(.Fields("WFRS_MESDATA1").Value)          'Rs測定値1
            WRs.Res(1) = NtoZ2(.Fields("WFRS_MESDATA2").Value)          'Rs測定値2
            WRs.Res(2) = NtoZ2(.Fields("WFRS_MESDATA3").Value)          'Rs測定値3
            WRs.Res(3) = NtoZ2(.Fields("WFRS_MESDATA4").Value)          'Rs測定値4
            WRs.Res(4) = NtoZ2(.Fields("WFRS_MESDATA5").Value)          'Rs測定値5
                
            .Fields("WFRS_RRG").Value = WFCRRGCal(WRs.Res(), WRs.GuaranteeRes, WRs.GuaranteeCal)        'WFRS_RRG計算結果
            
            '保証方法="H"、かつ、WFRS_RRG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
'            If (WRs.GuaranteeRes.cJudg = "H") And (.Fields("WFRS_RRG").Value = -1) Then GoTo proc_exit
            '保証方法ﾁｪｯｸの追加　04/04/15 ooba
            If ((WRs.GuaranteeRes.cJudg = "H") And CheckKHN(HWFRKHNN, 1, sPos)) _
                And (.Fields("WFRS_RRG").Value = -1) Then GoTo proc_exit
        
        End If
    End With

    getTBCMY013WFRs = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFRs = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFDOi実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :j               , I  ,Integer           , DOi No
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFDOi実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFDOi(recXSDCW As c_cmzcrec, j As Integer, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDOi"
    
    getTBCMY013WFDOi = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        If j = 1 Then
            .Fields("WFDOI_SMPPOS").Value = -1              'WFDOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        End If
        .Fields("WFDOI_NETU_" & j).Value = ""               'WFDOI_熱処理条件_x
        .Fields("WFDOI_MES_" & j).Value = ""                'WFDOI_計測方法_x
        .Fields("WFDOI_MESDATA1_" & j).Value = -1           'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)1_x
        .Fields("WFDOI_MESDATA2_" & j).Value = -1           'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)2_x
        .Fields("WFDOI_MESDATA3_" & j).Value = -1           'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)3_x
    End With
                
    'TBCMX002
    With recX002
        If j = 1 Then
            .Fields("WFDOI_SMPPOS").Value = -1              'WFDOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        End If
        .Fields("WFDOI" & j & "_NETSU").Value = " "         'WFDOI-x_熱処理条件
        .Fields("WFDOI" & j & "_MES").Value = " "           'WFDOI-x_計測方法
        .Fields("WFDOI" & j & "_MESDATA1").Value = " "      'WFDOI-x_測定値1
        .Fields("WFDOI" & j & "_MESDATA2").Value = " "      'WFDOI-x_測定値2
        .Fields("WFDOI" & j & "_MESDATA3").Value = " "      'WFDOI-x_測定値3
        .Fields("WFDOI" & j & "_MESDATA4").Value = " "      'WFDOI-x_測定値4
        .Fields("WFDOI" & j & "_MESDATA5").Value = " "      'WFDOI-x_測定値5
        .Fields("WFDOI" & j & "_MESDATA6").Value = " "      'WFDOI-x_測定値6
        .Fields("WFDOI" & j & "_MESDATA7").Value = " "      'WFDOI-x_測定値7
        .Fields("WFDOI" & j & "_MESDATA8").Value = " "      'WFDOI-x_測定値8
        .Fields("WFDOI" & j & "_MESDATA9").Value = " "      'WFDOI-x_測定値9
        .Fields("WFDOI" & j & "_MESDATA10").Value = " "     'WFDOI-x_測定値10
        .Fields("WFDOI" & j & "_MESDATA11").Value = " "     'WFDOI-x_測定値11
        .Fields("WFDOI" & j & "_MESDATA12").Value = " "     'WFDOI-x_測定値12
        .Fields("WFDOI" & j & "_MESDATA13").Value = " "     'WFDOI-x_測定値13
        .Fields("WFDOI" & j & "_MESDATA14").Value = " "     'WFDOI-x_測定値14
        .Fields("WFDOI" & j & "_MESDATA15").Value = " "     'WFDOI-x_測定値15
    End With
    
    '-------------------- TBCMY013の読み込み(WFDOi) ----------------------------------------
    If (recXSDCW("WFINDDO" & j & "CW").Value <> "0") And (recXSDCW("WFRESDO" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDO" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'DOI" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If .Fields("WFDOI_SMPPOS").Value = -1 Then
                .Fields("WFDOI_SMPPOS").Value = recXSDCW("INPOSCW").Value               'WFDOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            End If
            .Fields("WFDOI_NETU_" & j).Value = rs("NETSU")                              'WFDOI_熱処理条件_x
            .Fields("WFDOI_MES_" & j).Value = rs("MES")                                 'WFDOI_計測方法_x
            .Fields("WFDOI_MESDATA1_" & j).Value = rs("MESDATA1") - rs("MESDATA4")      'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)1_x
            .Fields("WFDOI_MESDATA2_" & j).Value = rs("MESDATA2") - rs("MESDATA5")      'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)2_x
            .Fields("WFDOI_MESDATA3_" & j).Value = rs("MESDATA3") - rs("MESDATA6")      'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)3_x
        End With
            
        'TBCMX002
        With recX002
            If .Fields("WFDOI_SMPPOS").Value = -1 Then
                .Fields("WFDOI_SMPPOS").Value = recXSDCW("INPOSCW").Value               'WFDOIｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            End If
            .Fields("WFDOI" & j & "_NETSU").Value = rs("NETSU")                         'WFDOI-x_熱処理条件
            .Fields("WFDOI" & j & "_MES").Value = rs("MES")                             'WFDOI-x_計測方法
            .Fields("WFDOI" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFDOI-x_測定値1
            .Fields("WFDOI" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFDOI-x_測定値2
            .Fields("WFDOI" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFDOI-x_測定値3
            .Fields("WFDOI" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFDOI-x_測定値4
            .Fields("WFDOI" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFDOI-x_測定値5
            .Fields("WFDOI" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFDOI-x_測定値6
            .Fields("WFDOI" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFDOI-x_測定値7
            .Fields("WFDOI" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFDOI-x_測定値8
            .Fields("WFDOI" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFDOI-x_測定値9
            .Fields("WFDOI" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFDOI-x_測定値10
            .Fields("WFDOI" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFDOI-x_測定値11
            .Fields("WFDOI" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFDOI-x_測定値12
            .Fields("WFDOI" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFDOI-x_測定値13
            .Fields("WFDOI" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFDOI-x_測定値14
            .Fields("WFDOI" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFDOI-x_測定値15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFDOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFOSF実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :j               , I  ,Integer           , OSF No
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'　　      :sPos  　　　    , I  ,String 　         , SXL位置(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'　　      :sTblName 　     , I  ,String 　         , テーブル名   11/06/24 Marushita　MIN値追加対応
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFOSF実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
'Private Function getTBCMY013WFOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMY013WFOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wos         As W_OSF                    'OSF構造体
    Dim keisu       As Double
    Dim k           As Integer
    Dim HWFOSFKN    As String                   '品ＷＦＯＳＦ検査頻度＿抜   04/04/15 ooba
    Dim nFlg        As Integer                  'MIN値セット判定用
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    '' 2006/09/25 SMP)kondoh Add -s-
    Const keisu7 As Double = 7.6923077
    '' 2006/09/25 SMP)kondoh Add -e-
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFOSF"
    
    getTBCMY013WFOSF = FUNCTION_RETURN_FAILURE

    '>>>>> 2011/06/27 SETsw)Marushita WFOSFx_判定時のMIN値_xのセット対応
    'MIN値の項目が存在する場合のみセット
    If FieldCheck(sTblName, "WFOSF" & j & "_MIN") = FUNCTION_RETURN_SUCCESS Then
        nFlg = 1
    Else
        nFlg = 0
    End If
    '<<<<< 2011/06/27 SETsw)Marushita WFOSFx_判定時のMIN値_xのセット対応
    
    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFOSF" & j & "_SMPPOS").Value = -1         'WFOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFOSF" & j & "_NETSU").Value = ""          'WFOSFx_熱処理条件
        .Fields("WFOSF" & j & "_ET").Value = ""             'WFOSFx_エッチング条件
        .Fields("WFOSF" & j & "_MES").Value = ""            'WFOSFx_計測方法
        .Fields("WFOSF" & j & "_MAX").Value = -1            'WFOSFx_判定時のMAX値_x
        .Fields("WFOSF" & j & "_AVE").Value = -1            'WFOSFx_判定時のAVE値_x
        If nFlg = 1 Then
            .Fields("WFOSF" & j & "_MIN").Value = -1        'WFOSFx_判定時のMIN値_x
            .Fields("WFOSF4_MIN").Value = -1                'WFOSF4_判定時のMIN値_4
        End If
        '↓Add SIRD評価対応　2010/04/19 Y.Hitomi
        .Fields("WFOSF4_SMPPOS").Value = -1                 'WFOSF4ｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFOSF4_NETSU").Value = ""                  'WFOSF4_熱処理条件
        .Fields("WFOSF4_ET").Value = ""                     'WFOSF4_エッチング条件
        .Fields("WFOSF4_MES").Value = ""                    'WFOSF4_計測方法
        .Fields("WFOSF4_MAX").Value = -1                    'WFOSF4_判定時のMAX値_4
        .Fields("WFOSF4_AVE").Value = -1                    'WFOSF4_判定時のAVE値_4
        '↑Add SIRD評価対応　2010/04/19 Y.Hitomi
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFOSF" & j & "_SMPPOS").Value = -1         'WFOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFOSF" & j & "_NETSU").Value = " "         'WFOSFx_熱処理条件
        .Fields("WFOSF" & j & "_ET").Value = " "            'WFOSFx_エッチング条件
        .Fields("WFOSF" & j & "_MES").Value = " "           'WFOSFx_計測方法
        .Fields("WFOSF" & j & "_DKAN").Value = " "          'WFOSFx_ＤＫアニール条件
        .Fields("WFOSF" & j & "_MESDATA1").Value = " "      'WFOSFx測定点1
        .Fields("WFOSF" & j & "_MESDATA2").Value = " "      'WFOSFx測定点2
        .Fields("WFOSF" & j & "_MESDATA3").Value = " "      'WFOSFx測定点3
        .Fields("WFOSF" & j & "_MESDATA4").Value = " "      'WFOSFx測定点4
        .Fields("WFOSF" & j & "_MESDATA5").Value = " "      'WFOSFx測定点5
        .Fields("WFOSF" & j & "_MESDATA6").Value = " "      'WFOSFx測定点6
        .Fields("WFOSF" & j & "_MESDATA7").Value = " "      'WFOSFx測定点7
        .Fields("WFOSF" & j & "_MESDATA8").Value = " "      'WFOSFx測定点8
        .Fields("WFOSF" & j & "_MESDATA9").Value = " "      'WFOSFx測定点9
        .Fields("WFOSF" & j & "_MESDATA10").Value = " "     'WFOSFx測定点10
        .Fields("WFOSF" & j & "_MESDATA11").Value = " "     'WFOSFx測定点11
        .Fields("WFOSF" & j & "_MESDATA12").Value = " "     'WFOSFx測定点12
        .Fields("WFOSF" & j & "_MESDATA13").Value = " "     'WFOSFx測定点13
        .Fields("WFOSF" & j & "_MESDATA14").Value = " "     'WFOSFx測定点14
        .Fields("WFOSF" & j & "_MESDATA15").Value = " "     'WFOSFx測定点15
        
        '↓Add SIRD評価対応　2010/04/19 Y.Hitomi
        .Fields("WFOSF4_SMPPOS").Value = -1         'WFOSF4ｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFOSF4_NETSU").Value = " "         'WFOSF4_熱処理条件
        .Fields("WFOSF4_ET").Value = " "            'WFOSF4_エッチング条件
        .Fields("WFOSF4_MES").Value = " "           'WFOSF4_計測方法
        .Fields("WFOSF4_DKAN").Value = " "          'WFOSF4_ＤＫアニール条件
        .Fields("WFOSF4_MESDATA1").Value = " "      'WFOSF4測定点1
        .Fields("WFOSF4_MESDATA2").Value = " "      'WFOSF4測定点2
        .Fields("WFOSF4_MESDATA3").Value = " "      'WFOSF4測定点3
        .Fields("WFOSF4_MESDATA4").Value = " "      'WFOSF4測定点4
        .Fields("WFOSF4_MESDATA5").Value = " "      'WFOSF4測定点5
        .Fields("WFOSF4_MESDATA6").Value = " "      'WFOSF4測定点6
        .Fields("WFOSF4_MESDATA7").Value = " "      'WFOSF4測定点7
        .Fields("WFOSF4_MESDATA8").Value = " "      'WFOSF4測定点8
        .Fields("WFOSF4_MESDATA9").Value = " "      'WFOSF4測定点9
        .Fields("WFOSF4_MESDATA10").Value = " "     'WFOSF4測定点10
        .Fields("WFOSF4_MESDATA11").Value = " "     'WFOSF4測定点11
        .Fields("WFOSF4_MESDATA12").Value = " "     'WFOSF4測定点12
        .Fields("WFOSF4_MESDATA13").Value = " "     'WFOSF4測定点13
        .Fields("WFOSF4_MESDATA14").Value = " "     'WFOSF4測定点14
        .Fields("WFOSF4_MESDATA15").Value = " "     'WFOSF4測定点15
        '↑Add SIRD評価対応　2010/04/19 Y.Hitomi
    End With
    
    '-------------------- TBCMY013の読み込み(WFOSF) ----------------------------------------
    If (recXSDCW("WFINDL" & j & "CW").Value <> "0") And (recXSDCW("WFRESL" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDL" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'OSF" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFOSF" & j & "_NETSU").Value = rs("NETSU")                         'WFOSFx_熱処理条件
            .Fields("WFOSF" & j & "_ET").Value = rs("ET")                               'WFOSFx_エッチング条件
            .Fields("WFOSF" & j & "_MES").Value = rs("MES")                             'WFOSFx_計測方法
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFOSF" & j & "_NETSU").Value = rs("NETSU")                         'WFOSFx_熱処理条件
            .Fields("WFOSF" & j & "_ET").Value = rs("ET")                               'WFOSFx_エッチング条件
            .Fields("WFOSF" & j & "_MES").Value = rs("MES")                             'WFOSFx_計測方法
            .Fields("WFOSF" & j & "_DKAN").Value = rs("DKAN")                           'WFOSFx_ＤＫアニール条件
            .Fields("WFOSF" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFOSFx測定点1
            .Fields("WFOSF" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFOSFx測定点2
            .Fields("WFOSF" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFOSFx測定点3
            .Fields("WFOSF" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFOSFx測定点4
            .Fields("WFOSF" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFOSFx測定点5
            .Fields("WFOSF" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFOSFx測定点6
            .Fields("WFOSF" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFOSFx測定点7
            .Fields("WFOSF" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFOSFx測定点8
            .Fields("WFOSF" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFOSFx測定点9
            .Fields("WFOSF" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFOSFx測定点10
            .Fields("WFOSF" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFOSFx測定点11
            .Fields("WFOSF" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFOSFx測定点12
            .Fields("WFOSF" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFOSFx測定点13
            .Fields("WFOSF" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFOSFx測定点14
            .Fields("WFOSF" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFOSFx測定点15
        End With
        Set rs = Nothing
        
        'WFOSF_MAX,AVE
        sql = "select HWFOF" & j & "SH, HWFOF" & j & "ST, HWFOF" & j & "SR, HWFOF" & j & "HT, "
        sql = sql & "HWFOF" & j & "HS, HWFOF" & j & "AX, HWFOF" & j & "MX, HWFOSF" & j & "PTK, "
        sql = sql & "HWFOF" & j & "KN "
        sql = sql & "from TBCME029 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
            
        If IsNull(rs("HWFOF" & j & "SH")) = False Then wos.GuaranteeOsf.cMeth = rs("HWFOF" & j & "SH")      '品ＷＦＯＳＦx測定位置＿方
        If IsNull(rs("HWFOF" & j & "ST")) = False Then wos.GuaranteeOsf.cCount = rs("HWFOF" & j & "ST")     '品ＷＦＯＳＦx測定位置＿点
        If IsNull(rs("HWFOF" & j & "SR")) = False Then wos.GuaranteeOsf.cPos = rs("HWFOF" & j & "SR")       '品ＷＦＯＳＦx測定位置＿領
        If IsNull(rs("HWFOF" & j & "HT")) = False Then wos.GuaranteeOsf.cObj = rs("HWFOF" & j & "HT")       '品ＷＦＯＳＦx保証方法＿対
        If IsNull(rs("HWFOF" & j & "HS")) = False Then wos.GuaranteeOsf.cJudg = rs("HWFOF" & j & "HS")      '品ＷＦＯＳＦx保証方法＿処
        If IsNull(rs("HWFOF" & j & "AX")) = False Then wos.SpecOsfAveMax = rs("HWFOF" & j & "AX")           '品ＷＦＯＳＦx平均上限
        If IsNull(rs("HWFOF" & j & "MX")) = False Then wos.SpecOsfMax = rs("HWFOF" & j & "MX")              '品ＷＦＯＳＦx上限
        If IsNull(rs("HWFOSF" & j & "PTK")) = False Then wos.JudgDataPTK = rs("HWFOSF" & j & "PTK")         '品ＷＦＯＳＦxパタン区分
        If IsNull(rs("HWFOF" & j & "KN")) = False Then HWFOSFKN = rs("HWFOF" & j & "KN")                    '品ＷＦＯＳＦ検査頻度＿抜　04/04/15 ooba
        Set rs = Nothing
            
        If wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "3" Then
            keisu = keisu1
        ElseIf wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "5" Then
            keisu = keisu2
        ElseIf wos.GuaranteeOsf.cMeth = "5" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu3
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "3" Then
            keisu = keisu4
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "5" Then
            keisu = keisu5
        ElseIf wos.GuaranteeOsf.cMeth = "6" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu6
        '' 2006/09/25 SMP)kondoh Add -s-
        ElseIf wos.GuaranteeOsf.cMeth = "E" And wos.GuaranteeOsf.cCount = "5" And wos.GuaranteeOsf.cPos = "A" Then
            keisu = keisu7
        '' 2006/09/25 SMP)kondoh Add -e-
        Else
            keisu = -1
'            GoTo proc_exit
        End If
            
        If keisu <> -1 Then
            With recX002
                wos.OSF(0) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA1").Value)                   'OSF測定値1
                wos.OSF(1) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA2").Value)                   'OSF測定値2
                wos.OSF(2) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA3").Value)                   'OSF測定値3
                wos.OSF(3) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA4").Value)                   'OSF測定値4
                wos.OSF(4) = NtoZ2(.Fields("WFOSF" & j & "_MESDATA5").Value)                   'OSF測定値5
                For k = 0 To 4
                    wos.OSF(k) = IIf(wos.OSF(k) <> -1, wos.OSF(k) * keisu, -1)
                Next
            End With
            
            recX001.Fields("WFOSF" & j & "_MAX").Value = JudgMax(wos.OSF())     'WFOSFx_判定時のMAX値_x
            recX001.Fields("WFOSF" & j & "_AVE").Value = JudgAve(wos.OSF())     'WFOSFx_判定時のAVE値_x
            '>>>>> 2011/06/24 SETsw)Marushita WFOSFx_判定時のMIN値_xのセット対応
            'MIN値の項目が存在する場合のみセット
            If nFlg = 1 Then
                recX001.Fields("WFOSF" & j & "_MIN").Value = JudgMin(wos.OSF())     'WFOSFx_判定時のMIN値_x
                recX001.Fields("WFOSF4_MIN").Value = -1     'WFOSFx_判定時のMIN値_4
            End If
            '<<<<< 2011/06/24 SETsw)Marushita WFOSFx_判定時のMIN値_xのセット対応
        End If
            
        '保証方法="H"、かつ、WFOSFのMAX,AVE値が-1の場合、エラーとする。2003/11/21 SystemBrain
        '保証方法ﾁｪｯｸの追加　04/04/15 ooba
        ''保証方法="H"、かつ、WFRS_RRG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
'        If (wos.GuaranteeOsf.cJudg = "H") And
        If ((wos.GuaranteeOsf.cJudg = "H") And CheckKHN(HWFOSFKN, j + 2, sPos)) And _
           (recX001.Fields("WFOSF" & j & "_MAX").Value = -1 Or _
            recX001.Fields("WFOSF" & j & "_AVE").Value = -1) Then GoTo proc_exit
        '？？？？？MIN値もチェック？？？？？
    End If

    getTBCMY013WFOSF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFOSF = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFBMD実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :j               , I  ,Integer           , BMD No
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFBMD実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFBMD(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim bm          As W_BMD                    'BMD構造体
    Dim k           As Integer
    Dim HWFBMDKN    As String                   '品ＷＦＢＭＤ検査頻度＿抜   04/04/15 ooba
    Dim JData(4)    As Double                   'MAX値算出用　06/09/06 ooba
    
    '' 2006/09/25 SMP)kondoh Del -s-
''    Const keisu As Double = 1        'BMDべき乗数変更対応　2003/05/19 osawa
    '' 2006/09/25 SMP)kondoh Del -e-
    '' 2006/09/25 SMP)kondoh Add -s-
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000
    '' 2006/09/25 SMP)kondoh Add -e-

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFBMD"
    
    getTBCMY013WFBMD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFBMD" & j & "_SMPPOS").Value = -1         'WFBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFBMD" & j & "_NETSU").Value = ""          'WFBMDx_熱処理条件
        .Fields("WFBMD" & j & "_ET").Value = ""             'WFBMDx_エッチング条件
        .Fields("WFBMD" & j & "_MES").Value = ""            'WFBMDx_計測方法
        .Fields("WFBMD" & j & "_MAX").Value = -1            'WFBMDx_判定時のMAX値_x
        .Fields("WFBMD" & j & "_AVE").Value = -1            'WFBMDx_判定時のAVE値_x
        .Fields("WFBMD" & j & "_MIN").Value = -1            'WFBMDx_判定時のMIN値_x
        .Fields("WFBMD" & j & "_MB").Value = -1             'WFBMDx_判定時の面内分布
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFBMD" & j & "_SMPPOS").Value = -1         'WFBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFBMD" & j & "_NETSU").Value = " "         'WFBMDx_熱処理条件
        .Fields("WFBMD" & j & "_ET").Value = " "            'WFBMDx_エッチング条件
        .Fields("WFBMD" & j & "_MES").Value = " "           'WFBMDx_計測方法
        .Fields("WFBMD" & j & "_DKAN").Value = " "          'WFBMDx_ＤＫアニール条件
        .Fields("WFBMD" & j & "_MESDATA1").Value = " "      'WFBMDx測定点1
        .Fields("WFBMD" & j & "_MESDATA2").Value = " "      'WFBMDx測定点2
        .Fields("WFBMD" & j & "_MESDATA3").Value = " "      'WFBMDx測定点3
        .Fields("WFBMD" & j & "_MESDATA4").Value = " "      'WFBMDx測定点4
        .Fields("WFBMD" & j & "_MESDATA5").Value = " "      'WFBMDx測定点5
        .Fields("WFBMD" & j & "_MESDATA6").Value = " "      'WFBMDx測定点6
        .Fields("WFBMD" & j & "_MESDATA7").Value = " "      'WFBMDx測定点7
        .Fields("WFBMD" & j & "_MESDATA8").Value = " "      'WFBMDx測定点8
        .Fields("WFBMD" & j & "_MESDATA9").Value = " "      'WFBMDx測定点9
        .Fields("WFBMD" & j & "_MESDATA10").Value = " "     'WFBMDx測定点10
        .Fields("WFBMD" & j & "_MESDATA11").Value = " "     'WFBMDx測定点11
        .Fields("WFBMD" & j & "_MESDATA12").Value = " "     'WFBMDx測定点12
        .Fields("WFBMD" & j & "_MESDATA13").Value = " "     'WFBMDx測定点13
        .Fields("WFBMD" & j & "_MESDATA14").Value = " "     'WFBMDx測定点14
        .Fields("WFBMD" & j & "_MESDATA15").Value = " "     'WFBMDx測定点15
    End With
    
    '-------------------- TBCMY013の読み込み(WFBMD) ----------------------------------------
    If (recXSDCW("WFINDB" & j & "CW").Value <> "0") And (recXSDCW("WFRESB" & j & "CW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDB" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'BMD" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFBMD" & j & "_NETSU").Value = rs("NETSU")                         'WFBMDx_熱処理条件
            .Fields("WFBMD" & j & "_ET").Value = rs("ET")                               'WFBMDx_エッチング条件
            .Fields("WFBMD" & j & "_MES").Value = rs("MES")                             'WFBMDx_計測方法
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFBMD" & j & "_NETSU").Value = rs("NETSU")                         'WFBMDx_熱処理条件
            .Fields("WFBMD" & j & "_ET").Value = rs("ET")                               'WFBMDx_エッチング条件
            .Fields("WFBMD" & j & "_MES").Value = rs("MES")                             'WFBMDx_計測方法
            .Fields("WFBMD" & j & "_DKAN").Value = rs("DKAN")                           'WFBMDx_ＤＫアニール条件
            .Fields("WFBMD" & j & "_MESDATA1").Value = rs("MESDATA1")                   'WFBMDx測定点1
            .Fields("WFBMD" & j & "_MESDATA2").Value = rs("MESDATA2")                   'WFBMDx測定点2
            .Fields("WFBMD" & j & "_MESDATA3").Value = rs("MESDATA3")                   'WFBMDx測定点3
            .Fields("WFBMD" & j & "_MESDATA4").Value = rs("MESDATA4")                   'WFBMDx測定点4
            .Fields("WFBMD" & j & "_MESDATA5").Value = rs("MESDATA5")                   'WFBMDx測定点5
            .Fields("WFBMD" & j & "_MESDATA6").Value = rs("MESDATA6")                   'WFBMDx測定点6
            .Fields("WFBMD" & j & "_MESDATA7").Value = rs("MESDATA7")                   'WFBMDx測定点7
            .Fields("WFBMD" & j & "_MESDATA8").Value = rs("MESDATA8")                   'WFBMDx測定点8
            .Fields("WFBMD" & j & "_MESDATA9").Value = rs("MESDATA9")                   'WFBMDx測定点9
            .Fields("WFBMD" & j & "_MESDATA10").Value = rs("MESDATA10")                 'WFBMDx測定点10
            .Fields("WFBMD" & j & "_MESDATA11").Value = rs("MESDATA11")                 'WFBMDx測定点11
            .Fields("WFBMD" & j & "_MESDATA12").Value = rs("MESDATA12")                 'WFBMDx測定点12
            .Fields("WFBMD" & j & "_MESDATA13").Value = rs("MESDATA13")                 'WFBMDx測定点13
            .Fields("WFBMD" & j & "_MESDATA14").Value = rs("MESDATA14")                 'WFBMDx測定点14
            .Fields("WFBMD" & j & "_MESDATA15").Value = rs("MESDATA15")                 'WFBMDx測定点15
        End With
        Set rs = Nothing
                    
        'WFBMD_MAX,MIN,AVE,MBP
        sql = "select HWFBM" & j & "SH, HWFBM" & j & "ST, HWFBM" & j & "SR, HWFBM" & j & "HT, "
        sql = sql & "HWFBM" & j & "HS, HWFBM" & j & "AN, HWFBM" & j & "AX, HWFBM" & j & "MBP, HWFBM" & j & "MCL, "
        sql = sql & "HWFBM" & j & "KN "
        sql = sql & "from TBCME029 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
                    
        If IsNull(rs("HWFBM" & j & "SH")) = False Then bm.GuaranteeBmd.cMeth = rs("HWFBM" & j & "SH")       '品ＷＦＢＭＤx測定位置＿方
        If IsNull(rs("HWFBM" & j & "ST")) = False Then bm.GuaranteeBmd.cCount = rs("HWFBM" & j & "ST")      '品ＷＦＢＭＤx測定位置＿点
        If IsNull(rs("HWFBM" & j & "SR")) = False Then bm.GuaranteeBmd.cPos = rs("HWFBM" & j & "SR")        '品ＷＦＢＭＤx測定位置＿領
        If IsNull(rs("HWFBM" & j & "HT")) = False Then bm.GuaranteeBmd.cObj = rs("HWFBM" & j & "HT")        '品ＷＦＢＭＤx保証方法＿対
        If IsNull(rs("HWFBM" & j & "HS")) = False Then bm.GuaranteeBmd.cJudg = rs("HWFBM" & j & "HS")       '品ＷＦＢＭＤx保証方法＿処
        If IsNull(rs("HWFBM" & j & "AN")) = False Then bm.SpecBmdAveMin = rs("HWFBM" & j & "AN")            '品ＷＦＢＭＤx平均下限
        If IsNull(rs("HWFBM" & j & "AX")) = False Then bm.SpecBmdAveMax = rs("HWFBM" & j & "AX")            '品ＷＦＢＭＤx平均上限
        If IsNull(rs("HWFBM" & j & "MBP")) = False Then bm.SpecBmdMBP = rs("HWFBM" & j & "MBP")             '品ＷＦＢＭＤx面内分布
        If IsNull(rs("HWFBM" & j & "MCL")) = False Then bm.SpecBmdMCL = rs("HWFBM" & j & "MCL")             '品ＷＦＢＭＤx面内計算
        If IsNull(rs("HWFBM" & j & "KN")) = False Then HWFBMDKN = rs("HWFBM" & j & "KN")                    '品ＷＦＢＭＤ検査頻度＿抜　04/04/15 ooba
        Set rs = Nothing

        '' 2006/09/25 SMP)kondoh Add -s-
        If bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "H" Then
            keisu = keisu1
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "H" Then
            keisu = keisu2
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu3
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu4
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "A" Then
            keisu = keisu5
        ElseIf bm.GuaranteeBmd.cMeth = "G" And bm.GuaranteeBmd.cCount = "3" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu6
        ElseIf bm.GuaranteeBmd.cMeth = "2" And bm.GuaranteeBmd.cCount = "5" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu7
        ElseIf bm.GuaranteeBmd.cMeth = "8" And bm.GuaranteeBmd.cCount = "4" And bm.GuaranteeBmd.cPos = "8" Then
            keisu = keisu8
        Else
            keisu = -1
        End If
        '' 2006/09/25 SMP)kondoh Add -e-

        '' 2006/09/25 SMP)kondoh Add -s-
        If keisu <> -1 Then
        '' 2006/09/25 SMP)kondoh Add -e-
                    
            With recX002
                bm.BMD(0) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA1").Value)     'BMD測定値1
                bm.BMD(1) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA2").Value)     'BMD測定値2
                bm.BMD(2) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA3").Value)     'BMD測定値3
                bm.BMD(3) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA4").Value)     'BMD測定値4
                bm.BMD(4) = NtoZ2(.Fields("WFBMD" & j & "_MESDATA5").Value)     'BMD測定値5
        
                For k = 0 To 4                                      ' 2003/05/20 ooba
                '   ' 2006/09/25 SMP)kondoh Add -s-
''                    bm.BMD(k) = IIf(bm.BMD(k) <> -1, bm.BMD(k) * keisu, -1)
                    bm.BMD(k) = IIf(bm.BMD(k) <> -1, bm.BMD(k) * CDbl(keisu / 10000), -1)
                    '' 2006/09/25 SMP)kondoh Add -e-
                Next
            End With
            
            ''06/09/06 ooba START =============================================================>
            '判定ｺｰﾄﾞが "F"：MAX(2,4点目)，"G"：MAX(2,3,4点目) の場合はMAX値にその値をｾｯﾄ
            If bm.GuaranteeBmd.cJudg = JudgCodeW01 And _
                (bm.GuaranteeBmd.cObj = ObjCode10 Or bm.GuaranteeBmd.cObj = ObjCode11) Then
            
                If GetWfJudgData(WFBMD_JUDG, bm.GuaranteeBmd, bm.BMD(), JData()) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                recX001.Fields("WFBMD" & j & "_MAX").Value = JData(0)
            Else
                recX001.Fields("WFBMD" & j & "_MAX").Value = JudgMax(bm.BMD())
            End If
            ''06/09/06 ooba END ===============================================================>
            
    '        recX001.Fields("WFBMD" & j & "_MAX").Value = JudgMax(bm.BMD())          'WFBMDx_判定時のMAX値_x
            recX001.Fields("WFBMD" & j & "_AVE").Value = JudgAve(bm.BMD())          'WFBMDx_判定時のAVE値_x
            recX001.Fields("WFBMD" & j & "_MIN").Value = JudgMin(bm.BMD())          'WFBMDx_判定時のMIN値_x
            If bm.SpecBmdMCL = "P " Then
                recX001.Fields("WFBMD" & j & "_MB").Value = JudgBmdMBP(bm.BMD())    'WFBMDx_判定時の面内分布
            Else
                recX001.Fields("WFBMD" & j & "_MB").Value = 0                       '面内分布が"P"以外の時は計算結果を0とする　2003/06/06 ooba
            End If
            
        '' 2006/09/25 SMP)kondoh Add -s-
        End If
        '' 2006/09/25 SMP)kondoh Add -e-
            
        '保証方法="H"、かつ、WFRS_RRG計算結果が-1の場合、エラーとする。2003/11/21 SystemBrain
'        If (bm.GuaranteeBmd.cJudg = "H") And (recX001.Fields("WFBMD" & j & "_MB").Value = -1) Then GoTo proc_exit
        '保証方法ﾁｪｯｸの追加　04/04/15 ooba
        If ((bm.GuaranteeBmd.cJudg = "H") And CheckKHN(HWFBMDKN, j + 6, sPos)) _
            And (recX001.Fields("WFBMD" & j & "_MB").Value = -1) Then GoTo proc_exit
        
    End If

    getTBCMY013WFBMD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFBMD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFDSOD実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFDSOD実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFDSOD(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDSOD"
    
    getTBCMY013WFDSOD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDSOD_SMPPOS").Value = -1         'WFDSODｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFDSOD_NETSU").Value = ""          'WFDSOD_熱処理条件
        .Fields("WFDSOD_ET").Value = ""             'WFDSOD_エッチング条件
        .Fields("WFDSOD_MES").Value = ""            'WFDSOD_計測方法
        .Fields("WFDSOD_TOTAL").Value = -1          'WFDSOD_判定時のTOTAL値
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFDSOD_SMPPOS").Value = -1         'WFDSODｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFDSOD_NETSU").Value = " "         'WFDSOD_熱処理条件
        .Fields("WFDSOD_ET").Value = " "            'WFDSOD_エッチング条件
        .Fields("WFDSOD_MES").Value = " "           'WFDSOD_計測方法
        .Fields("WFDSOD_DKAN").Value = " "          'WFDSOD_ＤＫアニール条件
        .Fields("WFDSOD_MESDATA1").Value = " "      'WFDSOD測定点1
        .Fields("WFDSOD_MESDATA2").Value = " "      'WFDSOD測定点2
        .Fields("WFDSOD_MESDATA3").Value = " "      'WFDSOD測定点3
        .Fields("WFDSOD_MESDATA4").Value = " "      'WFDSOD測定点4
        .Fields("WFDSOD_MESDATA5").Value = " "      'WFDSOD測定点5
        .Fields("WFDSOD_MESDATA6").Value = " "      'WFDSOD測定点6
        .Fields("WFDSOD_MESDATA7").Value = " "      'WFDSOD測定点7
        .Fields("WFDSOD_MESDATA8").Value = " "      'WFDSOD測定点8
        .Fields("WFDSOD_MESDATA9").Value = " "      'WFDSOD測定点9
        .Fields("WFDSOD_MESDATA10").Value = " "     'WFDSOD測定点10
        .Fields("WFDSOD_MESDATA11").Value = " "     'WFDSOD測定点11
        .Fields("WFDSOD_MESDATA12").Value = " "     'WFDSOD測定点12
        .Fields("WFDSOD_MESDATA13").Value = " "     'WFDSOD測定点13
        .Fields("WFDSOD_MESDATA14").Value = " "     'WFDSOD測定点14
        .Fields("WFDSOD_MESDATA15").Value = " "     'WFDSOD測定点15
    End With
    
    '-------------------- TBCMY013の読み込み(WFDSOD) ----------------------------------------
    If (recXSDCW("WFINDDSCW").Value <> "0") And (recXSDCW("WFRESDSCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDSCW").Value & "' and "
        sql = sql & "      SPEC = 'DSOD'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDSOD_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFDSODｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFDSOD_NETSU").Value = rs("NETSU")                         'WFDSOD_熱処理条件
            .Fields("WFDSOD_ET").Value = rs("ET")                               'WFDSOD_エッチング条件
            .Fields("WFDSOD_MES").Value = rs("MES")                             'WFDSOD_計測方法
            .Fields("WFDSOD_TOTAL").Value = rs("MESDATA1")                      'WFDSOD_判定時のTOTAL値
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDSOD_SMPPOS").Value = recXSDCW("INPOSCW").Value          'WFDSODｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFDSOD_NETSU").Value = rs("NETSU")                         'WFDSOD_熱処理条件
            .Fields("WFDSOD_ET").Value = rs("ET")                               'WFDSOD_エッチング条件
            .Fields("WFDSOD_MES").Value = rs("MES")                             'WFDSOD_計測方法
            .Fields("WFDSOD_DKAN").Value = rs("DKAN")                           'WFDSOD_ＤＫアニール条件
            .Fields("WFDSOD_MESDATA1").Value = rs("MESDATA1")                   'WFDSOD測定点1
            .Fields("WFDSOD_MESDATA2").Value = rs("MESDATA2")                   'WFDSOD測定点2
            .Fields("WFDSOD_MESDATA3").Value = rs("MESDATA3")                   'WFDSOD測定点3
            .Fields("WFDSOD_MESDATA4").Value = rs("MESDATA4")                   'WFDSOD測定点4
            .Fields("WFDSOD_MESDATA5").Value = rs("MESDATA5")                   'WFDSOD測定点5
            .Fields("WFDSOD_MESDATA6").Value = rs("MESDATA6")                   'WFDSOD測定点6
            .Fields("WFDSOD_MESDATA7").Value = rs("MESDATA7")                   'WFDSOD測定点7
            .Fields("WFDSOD_MESDATA8").Value = rs("MESDATA8")                   'WFDSOD測定点8
            .Fields("WFDSOD_MESDATA9").Value = rs("MESDATA9")                   'WFDSOD測定点9
            .Fields("WFDSOD_MESDATA10").Value = rs("MESDATA10")                 'WFDSOD測定点10
            .Fields("WFDSOD_MESDATA11").Value = rs("MESDATA11")                 'WFDSOD測定点11
            .Fields("WFDSOD_MESDATA12").Value = rs("MESDATA12")                 'WFDSOD測定点12
            .Fields("WFDSOD_MESDATA13").Value = rs("MESDATA13")                 'WFDSOD測定点13
            .Fields("WFDSOD_MESDATA14").Value = rs("MESDATA14")                 'WFDSOD測定点14
            .Fields("WFDSOD_MESDATA15").Value = rs("MESDATA15")                 'WFDSOD測定点15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFDSOD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDSOD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFSPV実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFSPV実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFSPV(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFSPV"
    
    getTBCMY013WFSPV = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFSPV_NETSU").Value = ""           'WFSPV_熱処理条件
        .Fields("WFSPV_ET").Value = ""              'WFSPV_エッチング条件
        .Fields("WFSPV_MES").Value = ""             'WFSPV_計測方法
        .Fields("WFSPV_KST_MAX").Value = -1         'WFSPV_拡散長判定時のMAX値
        .Fields("WFSPV_KST_AVE").Value = -1         'WFSPV_拡散長判定時のAVE値
        .Fields("WFSPV_KST_MIN").Value = -1         'WFSPV_拡散長判定時のMIN値
        .Fields("WFSPV_FE_MAX").Value = -1          'WFSPV_Fe濃度判定時のMAX値
        .Fields("WFSPV_FE_AVE").Value = -1          'WFSPV_Fe濃度判定時のAVE値
        .Fields("WFSPV_FE_MIN").Value = -1          'WFSPV_Fe濃度判定時のMIN値
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFSPV_SMPPOS").Value = -1         'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFSPV_NETSU").Value = " "         'WFSPV_熱処理条件
        .Fields("WFSPV_ET").Value = " "            'WFSPV_エッチング条件
        .Fields("WFSPV_MES").Value = " "           'WFSPV_計測方法
        .Fields("WFSPV_DKAN").Value = " "          'WFSPV_ＤＫアニール条件
        .Fields("WFSPV_MESDATA1").Value = " "      'WFSPV測定点1
        .Fields("WFSPV_MESDATA2").Value = " "      'WFSPV測定点2
        .Fields("WFSPV_MESDATA3").Value = " "      'WFSPV測定点3
        .Fields("WFSPV_MESDATA4").Value = " "      'WFSPV測定点4
        .Fields("WFSPV_MESDATA5").Value = " "      'WFSPV測定点5
        .Fields("WFSPV_MESDATA6").Value = " "      'WFSPV測定点6
        .Fields("WFSPV_MESDATA7").Value = " "      'WFSPV測定点7
        .Fields("WFSPV_MESDATA8").Value = " "      'WFSPV測定点8
        .Fields("WFSPV_MESDATA9").Value = " "      'WFSPV測定点9
        .Fields("WFSPV_MESDATA10").Value = " "     'WFSPV測定点10
        .Fields("WFSPV_MESDATA11").Value = " "     'WFSPV測定点11
        .Fields("WFSPV_MESDATA12").Value = " "     'WFSPV測定点12
        .Fields("WFSPV_MESDATA13").Value = " "     'WFSPV測定点13
        .Fields("WFSPV_MESDATA14").Value = " "     'WFSPV測定点14
        .Fields("WFSPV_MESDATA15").Value = " "     'WFSPV測定点15
    End With
    
    '-------------------- TBCMY013の読み込み(WFSPV) ----------------------------------------
    If (recXSDCW("WFINDSPCW").Value <> "0") And (recXSDCW("WFRESSPCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDSPCW").Value & "' and "
        sql = sql & "      SPEC = 'SPV'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFSPV_NETSU").Value = rs("NETSU")                      'WFSPV_熱処理条件
            .Fields("WFSPV_ET").Value = rs("ET")                            'WFSPV_エッチング条件
            .Fields("WFSPV_MES").Value = rs("MES")                          'WFSPV_計測方法
            .Fields("WFSPV_KST_MAX").Value = rs("MESDATA1")                 'WFSPV_拡散長判定時のMAX値
            .Fields("WFSPV_KST_AVE").Value = rs("MESDATA2")                 'WFSPV_拡散長判定時のAVE値
            .Fields("WFSPV_KST_MIN").Value = rs("MESDATA3")                 'WFSPV_拡散長判定時のMIN値
            .Fields("WFSPV_FE_MAX").Value = rs("MESDATA4")                  'WFSPV_Fe濃度判定時のMAX値
            .Fields("WFSPV_FE_AVE").Value = rs("MESDATA5")                  'WFSPV_Fe濃度判定時のAVE値
            .Fields("WFSPV_FE_MIN").Value = rs("MESDATA6")                  'WFSPV_Fe濃度判定時のMIN値
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFSPV_NETSU").Value = rs("NETSU")                      'WFSPV_熱処理条件
            .Fields("WFSPV_ET").Value = rs("ET")                            'WFSPV_エッチング条件
            .Fields("WFSPV_MES").Value = rs("MES")                          'WFSPV_計測方法
            .Fields("WFSPV_DKAN").Value = rs("DKAN")                        'WFSPV_ＤＫアニール条件
            .Fields("WFSPV_MESDATA1").Value = rs("MESDATA1")                'WFSPV測定点1
            .Fields("WFSPV_MESDATA2").Value = rs("MESDATA2")                'WFSPV測定点2
            .Fields("WFSPV_MESDATA3").Value = rs("MESDATA3")                'WFSPV測定点3
            .Fields("WFSPV_MESDATA4").Value = rs("MESDATA4")                'WFSPV測定点4
            .Fields("WFSPV_MESDATA5").Value = rs("MESDATA5")                'WFSPV測定点5
            .Fields("WFSPV_MESDATA6").Value = rs("MESDATA6")                'WFSPV測定点6
            .Fields("WFSPV_MESDATA7").Value = rs("MESDATA7")                'WFSPV測定点7
            .Fields("WFSPV_MESDATA8").Value = rs("MESDATA8")                'WFSPV測定点8
            .Fields("WFSPV_MESDATA9").Value = rs("MESDATA9")                'WFSPV測定点9
            .Fields("WFSPV_MESDATA10").Value = rs("MESDATA10")              'WFSPV測定点10
            .Fields("WFSPV_MESDATA11").Value = rs("MESDATA11")              'WFSPV測定点11
            .Fields("WFSPV_MESDATA12").Value = rs("MESDATA12")              'WFSPV測定点12
            .Fields("WFSPV_MESDATA13").Value = rs("MESDATA13")              'WFSPV測定点13
            .Fields("WFSPV_MESDATA14").Value = rs("MESDATA14")              'WFSPV測定点14
            .Fields("WFSPV_MESDATA15").Value = rs("MESDATA15")              'WFSPV測定点15
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFSPV = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFSPV = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFDZ実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFDZ実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMY013WFDZ(recXSDCW As c_cmzcrec, HIN As tFullHinban, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim DZ          As W_DZ                     'DZ構造体
    Dim JData(3)    As Double                   'MAX値算出用　06/09/06 ooba
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFDZ"
    
    getTBCMY013WFDZ = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDZ_SMPPOS").Value = -1           'WFDZｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFDZ_NETSU").Value = ""            'WFDZ_熱処理条件
        .Fields("WFDZ_ET").Value = ""               'WFDZ_エッチング条件
        .Fields("WFDZ_MES").Value = ""              'WFDZ_計測方法
        .Fields("WFDZ_MAX").Value = -1              'WFDZ_判定時のMAX値
        .Fields("WFDZ_AVE").Value = -1              'WFDZ_判定時のAVE値
        .Fields("WFDZ_MIN").Value = -1              'WFDZ_判定時のMIN値
    End With

    'TBCMX002
    With recX002
        .Fields("WFDZ_SMPPOS").Value = -1           'WFDZｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFDZ_NETSU").Value = " "           'WFDZ_熱処理条件
        .Fields("WFDZ_ET").Value = " "              'WFDZ_エッチング条件
        .Fields("WFDZ_MES").Value = " "             'WFDZ_計測方法
        .Fields("WFDZ_DKAN").Value = " "            'WFDZ_ＤＫアニール条件
        .Fields("WFDZ_MESDATA1").Value = " "        'WFDZ測定点1
        .Fields("WFDZ_MESDATA2").Value = " "        'WFDZ測定点2
        .Fields("WFDZ_MESDATA3").Value = " "        'WFDZ測定点3
        .Fields("WFDZ_MESDATA4").Value = " "        'WFDZ測定点4
        .Fields("WFDZ_MESDATA5").Value = " "        'WFDZ測定点5
        .Fields("WFDZ_MESDATA6").Value = " "        'WFDZ測定点6
        .Fields("WFDZ_MESDATA7").Value = " "        'WFDZ測定点7
        .Fields("WFDZ_MESDATA8").Value = " "        'WFDZ測定点8
        .Fields("WFDZ_MESDATA9").Value = " "        'WFDZ測定点9
        .Fields("WFDZ_MESDATA10").Value = " "       'WFDZ測定点10
        .Fields("WFDZ_MESDATA11").Value = " "       'WFDZ測定点11
        .Fields("WFDZ_MESDATA12").Value = " "       'WFDZ測定点12
        .Fields("WFDZ_MESDATA13").Value = " "       'WFDZ測定点13
        .Fields("WFDZ_MESDATA14").Value = " "       'WFDZ測定点14
        .Fields("WFDZ_MESDATA15").Value = " "       'WFDZ測定点15
    End With
    
    '-------------------- TBCMY013の読み込み(WFDZ) ----------------------------------------
    If (recXSDCW("WFINDDZCW").Value <> "0") And (recXSDCW("WFRESDZCW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDDZCW").Value & "' and "
        sql = sql & "      SPEC = 'DZ'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDZ_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFDZｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFDZ_NETSU").Value = rs("NETSU")                      'WFDZ_熱処理条件
            .Fields("WFDZ_ET").Value = rs("ET")                            'WFDZ_エッチング条件
            .Fields("WFDZ_MES").Value = rs("MES")                          'WFDZ_計測方法
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDZ_SMPPOS").Value = recXSDCW("INPOSCW").Value       'WFDZｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFDZ_NETSU").Value = rs("NETSU")                      'WFDZ_熱処理条件
            .Fields("WFDZ_ET").Value = rs("ET")                            'WFDZ_エッチング条件
            .Fields("WFDZ_MES").Value = rs("MES")                          'WFDZ_計測方法
            .Fields("WFDZ_DKAN").Value = rs("DKAN")                        'WFDZ_ＤＫアニール条件
            .Fields("WFDZ_MESDATA1").Value = rs("MESDATA1")                'WFDZ測定点1
            .Fields("WFDZ_MESDATA2").Value = rs("MESDATA2")                'WFDZ測定点2
            .Fields("WFDZ_MESDATA3").Value = rs("MESDATA3")                'WFDZ測定点3
            .Fields("WFDZ_MESDATA4").Value = rs("MESDATA4")                'WFDZ測定点4
            .Fields("WFDZ_MESDATA5").Value = rs("MESDATA5")                'WFDZ測定点5
            .Fields("WFDZ_MESDATA6").Value = rs("MESDATA6")                'WFDZ測定点6
            .Fields("WFDZ_MESDATA7").Value = rs("MESDATA7")                'WFDZ測定点7
            .Fields("WFDZ_MESDATA8").Value = rs("MESDATA8")                'WFDZ測定点8
            .Fields("WFDZ_MESDATA9").Value = rs("MESDATA9")                'WFDZ測定点9
            .Fields("WFDZ_MESDATA10").Value = rs("MESDATA10")              'WFDZ測定点10
            .Fields("WFDZ_MESDATA11").Value = rs("MESDATA11")              'WFDZ測定点11
            .Fields("WFDZ_MESDATA12").Value = rs("MESDATA12")              'WFDZ測定点12
            .Fields("WFDZ_MESDATA13").Value = rs("MESDATA13")              'WFDZ測定点13
            .Fields("WFDZ_MESDATA14").Value = rs("MESDATA14")              'WFDZ測定点14
            .Fields("WFDZ_MESDATA15").Value = rs("MESDATA15")              'WFDZ測定点15
        End With
        Set rs = Nothing
                
        'WFDZ_MAX,MIN,AVE
        With recX002
            DZ.DZ(0) = NtoZ2(.Fields("WFDZ_MESDATA1").Value)               'DZ測定値1
            DZ.DZ(1) = NtoZ2(.Fields("WFDZ_MESDATA2").Value)               'DZ測定値2
            DZ.DZ(2) = NtoZ2(.Fields("WFDZ_MESDATA3").Value)               'DZ測定値3
            DZ.DZ(3) = NtoZ2(.Fields("WFDZ_MESDATA4").Value)               'DZ測定値4
        End With
        
        ''06/09/06 ooba START =============================================================>
        'DZ仕様取得
        sql = "select HWFMKSPH, HWFMKSPT, HWFMKSPR, HWFMKHWT, HWFMKHWS, HWFMKMIN, HWFMKMAX "
        sql = sql & "from TBCME024 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFMKSPH")) = False Then DZ.GuaranteeDz.cMeth = rs("HWFMKSPH")    '品ＷＦ無欠陥層測定位置＿方
        If IsNull(rs("HWFMKSPT")) = False Then DZ.GuaranteeDz.cCount = rs("HWFMKSPT")   '品ＷＦ無欠陥層測定位置＿点
        If IsNull(rs("HWFMKSPR")) = False Then DZ.GuaranteeDz.cPos = rs("HWFMKSPR")     '品ＷＦ無欠陥層測定位置＿領
        If IsNull(rs("HWFMKHWT")) = False Then DZ.GuaranteeDz.cObj = rs("HWFMKHWT")     '品ＷＦ無欠陥層保証方法＿対
        If IsNull(rs("HWFMKHWS")) = False Then DZ.GuaranteeDz.cJudg = rs("HWFMKHWS")    '品ＷＦ無欠陥層保証方法＿処
        If IsNull(rs("HWFMKMIN")) = False Then DZ.SpecDzMin = rs("HWFMKMIN")            '品ＷＦ無欠陥層下限
        If IsNull(rs("HWFMKMAX")) = False Then DZ.SpecDzMax = rs("HWFMKMAX")            '品ＷＦ無欠陥層上限
        
        Set rs = Nothing
        
        '判定ｺｰﾄﾞが "F"：MAX(2,4点目)，"G"：MAX(2,3,4点目) の場合はMAX値にその値をｾｯﾄ
        If DZ.GuaranteeDz.cJudg = JudgCodeW01 And _
            (DZ.GuaranteeDz.cObj = ObjCode10 Or DZ.GuaranteeDz.cObj = ObjCode11) Then
        
            If GetWfJudgData(WFDZ_JUDG, DZ.GuaranteeDz, DZ.DZ(), JData()) = FUNCTION_RETURN_FAILURE Then
                GoTo proc_exit
            End If
            recX001.Fields("WFDZ_MAX").Value = JData(0)
        Else
            recX001.Fields("WFDZ_MAX").Value = JudgMax(DZ.DZ())
        End If
        ''06/09/06 ooba END ===============================================================>
        
'        recX001.Fields("WFDZ_MAX").Value = JudgMax(DZ.DZ())             'WFDZ_判定時のMAX値
        recX001.Fields("WFDZ_AVE").Value = JudgAve(DZ.DZ())             'WFDZ_判定時のAVE値
        recX001.Fields("WFDZ_MIN").Value = JudgMin(DZ.DZ())             'WFDZ_判定時のMIN値
    End If

    getTBCMY013WFDZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFDZ = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFAOi実績(TBCMY013)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFAOi実績(TBCMY013)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :03/12/19 ooba
Private Function getTBCMY013WFAOi(recXSDCW As c_cmzcrec, recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim DZ          As W_DZ                     'DZ構造体
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY013WFAOi"
    
    getTBCMY013WFAOi = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFDOI_NETU_3").Value = ""            'WFDOI_熱処理条件_３
        .Fields("WFDOI_MES_3").Value = ""             'WFDOI_計測方法_３
        .Fields("WFDOI_MESDATA1_3").Value = -1        'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)１_３
        .Fields("WFDOI_MESDATA2_3").Value = -1        'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_３
        .Fields("WFDOI_MESDATA3_3").Value = -1        'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_３
        .Fields("ZOIFLG").Value = ""                  '残存酸素存在ﾌﾗｸﾞ
    End With

    'TBCMX002
    With recX002
        .Fields("WFDOI3_NETSU").Value = " "           'WFDOI-3_熱処理条件
        .Fields("WFDOI3_MES").Value = " "             'WFDOI-3_計測方法
        .Fields("WFDOI3_MESDATA1").Value = " "        'WFDOI-3_測定点1
        .Fields("WFDOI3_MESDATA2").Value = " "        'WFDOI-3_測定点2
        .Fields("WFDOI3_MESDATA3").Value = " "        'WFDOI-3_測定点3
        .Fields("WFDOI3_MESDATA4").Value = " "        'WFDOI-3_測定点4
        .Fields("WFDOI3_MESDATA5").Value = " "        'WFDOI-3_測定点5
        .Fields("WFDOI3_MESDATA6").Value = " "        'WFDOI-3_測定点6
        .Fields("WFDOI3_MESDATA7").Value = " "        'WFDOI-3_測定点7
        .Fields("WFDOI3_MESDATA8").Value = " "        'WFDOI-3_測定点8
        .Fields("WFDOI3_MESDATA9").Value = " "        'WFDOI-3_測定点9
        .Fields("WFDOI3_MESDATA10").Value = " "       'WFDOI-3_測定点10
        .Fields("WFDOI3_MESDATA11").Value = " "       'WFDOI-3_測定点11
        .Fields("WFDOI3_MESDATA12").Value = " "       'WFDOI-3_測定点12
        .Fields("WFDOI3_MESDATA13").Value = " "       'WFDOI-3_測定点13
        .Fields("WFDOI3_MESDATA14").Value = " "       'WFDOI-3_測定点14
        .Fields("WFDOI3_MESDATA15").Value = " "       'WFDOI-3_測定点15
        .Fields("ZOIFLG").Value = " "                 '残存酸素存在ﾌﾗｸﾞ
    End With
    
    '-------------------- TBCMY013の読み込み(WFAOi) ----------------------------------------
    If (recXSDCW("WFINDAOICW").Value <> "0") And (recXSDCW("WFRESAOICW").Value <> "0") Then
        sql = "select * from TBCMY013 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("WFSMPLIDAOICW").Value & "' and "
        sql = sql & "      SPEC = 'AOI'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            .Fields("WFDOI_NETU_3").Value = rs("NETSU")                'WFDOI_熱処理条件_３
            .Fields("WFDOI_MES_3").Value = rs("MES")                   'WFDOI_計測方法_３
            .Fields("WFDOI_MESDATA1_3").Value = rs("MESDATA4")         'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)１_３
            .Fields("WFDOI_MESDATA2_3").Value = rs("MESDATA5")         'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)２_３
            .Fields("WFDOI_MESDATA3_3").Value = rs("MESDATA6")         'WFDOI_(ｲﾆｼｬﾙOi-AfterOi)３_３
            .Fields("ZOIFLG").Value = "1"                              '残存酸素存在ﾌﾗｸﾞ
        End With
            
        'TBCMX002
        With recX002
            .Fields("WFDOI3_NETSU").Value = rs("NETSU")                      'WFDOI-3_熱処理条件
            .Fields("WFDOI3_MES").Value = rs("MES")                          'WFDOI-3_計測方法
            .Fields("WFDOI3_MESDATA1").Value = rs("MESDATA1")                'WFDOI-3_測定点1
            .Fields("WFDOI3_MESDATA2").Value = rs("MESDATA2")                'WFDOI-3_測定点2
            .Fields("WFDOI3_MESDATA3").Value = rs("MESDATA3")                'WFDOI-3_測定点3
            .Fields("WFDOI3_MESDATA4").Value = rs("MESDATA4")                'WFDOI-3_測定点4
            .Fields("WFDOI3_MESDATA5").Value = rs("MESDATA5")                'WFDOI-3_測定点5
            .Fields("WFDOI3_MESDATA6").Value = rs("MESDATA6")                'WFDOI-3_測定点6
            .Fields("WFDOI3_MESDATA7").Value = rs("MESDATA7")                'WFDOI-3_測定点7
            .Fields("WFDOI3_MESDATA8").Value = rs("MESDATA8")                'WFDOI-3_測定点8
            .Fields("WFDOI3_MESDATA9").Value = rs("MESDATA9")                'WFDOI-3_測定点9
            .Fields("WFDOI3_MESDATA10").Value = rs("MESDATA10")              'WFDOI-3_測定点10
            .Fields("WFDOI3_MESDATA11").Value = rs("MESDATA11")              'WFDOI-3_測定点11
            .Fields("WFDOI3_MESDATA12").Value = rs("MESDATA12")              'WFDOI-3_測定点12
            .Fields("WFDOI3_MESDATA13").Value = rs("MESDATA13")              'WFDOI-3_測定点13
            .Fields("WFDOI3_MESDATA14").Value = rs("MESDATA14")              'WFDOI-3_測定点14
            .Fields("WFDOI3_MESDATA15").Value = rs("MESDATA15")              'WFDOI-3_測定点15
            .Fields("ZOIFLG").Value = "1"                                    '残存酸素存在ﾌﾗｸﾞ
        End With
        Set rs = Nothing
    End If

    getTBCMY013WFAOi = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY013WFAOi = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ｴﾋﾟOSF1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , 新ｻﾝﾌﾟﾙ管理(SXL)
'          :j               , I  ,Integer           , ｴﾋﾟOSF No
'          :HIN             , I  ,tFullHinban       , 品番
'　　      :sPos  　　　    , I  ,String 　         , SXL位置(TOP/BOT)
'          :recX004         , O  ,c_cmzcrec         , EP検査書
'          :recX005         , O  ,c_cmzcrec         , EP測定点ﾃﾞｰﾀ
'　　      :sTblName　　    , I  ,String 　         , テーブル名　11/06/24 Marushita  MIN値追加対応
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :EP先行評価結果(TBCMY022)から実績ﾃﾞｰﾀを取得し、EP検査書/EP測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :06/08/10 ooba
'Private Function getTBCMY022EPOSF(recXSDCW As c_cmzcrec, j As Integer, HIN As tFullHinban, sPos As String, recX004 As c_cmzcrec, recX005 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMY022EPOSF(recXSDCW As c_cmzcrec, j As Integer, _
                                  HIN As tFullHinban, sPos As String, _
                                  recX004 As c_cmzcrec, recX005 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
                                  
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim eosf        As W_OSF                    'ｴﾋﾟOSF構造体
    Dim keisu       As Double
    Dim k           As Integer
    Dim HEPOSFKN    As String                   '検査頻度_抜
    
    Const keisu1 As Double = 1.8248175
    Const keisu2 As Double = 1.8518519
    Const keisu3 As Double = 1.9230769
    Const keisu4 As Double = 3.649635
    Const keisu5 As Double = 3.7037037
    Const keisu6 As Double = 3.8461538
    Const keisu7 As Double = 7.6923077
    
    'ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY022EPOSF"
    
    getTBCMY022EPOSF = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX004
    With recX004
        .Fields("EPOSF" & j & "_SMPPOS").Value = vbNullString       'EPOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("EPOSF" & j & "_NETSU").Value = vbNullString        'EPOSFx_熱処理条件
        .Fields("EPOSF" & j & "_ET").Value = vbNullString           'EPOSFx_ｴｯﾁﾝｸﾞ条件
        .Fields("EPOSF" & j & "_MES").Value = vbNullString          'EPOSFx_計測方法
        .Fields("EPOSF" & j & "_MAX").Value = vbNullString          'EPOSFx_判定時のMAX値_x
        .Fields("EPOSF" & j & "_AVE").Value = vbNullString          'EPOSFx_判定時のAVE値_x
    End With
                
    'TBCMX005
    With recX005
        .Fields("EPOSF" & j & "_SMPPOS").Value = vbNullString       'EPOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("EPOSF" & j & "_NETSU").Value = vbNullString        'EPOSFx_熱処理条件
        .Fields("EPOSF" & j & "_ET").Value = vbNullString           'EPOSFx_ｴｯﾁﾝｸﾞ条件
        .Fields("EPOSF" & j & "_MES").Value = vbNullString          'EPOSFx_計測方法
        .Fields("EPOSF" & j & "_DKAN").Value = vbNullString         'EPOSFx_DKｱﾆｰﾙ条件
        .Fields("EPOSF" & j & "_MESDATA1").Value = vbNullString     'EPOSFx測定点1
        .Fields("EPOSF" & j & "_MESDATA2").Value = vbNullString     'EPOSFx測定点2
        .Fields("EPOSF" & j & "_MESDATA3").Value = vbNullString     'EPOSFx測定点3
        .Fields("EPOSF" & j & "_MESDATA4").Value = vbNullString     'EPOSFx測定点4
        .Fields("EPOSF" & j & "_MESDATA5").Value = vbNullString     'EPOSFx測定点5
        .Fields("EPOSF" & j & "_MESDATA6").Value = vbNullString     'EPOSFx測定点6
        .Fields("EPOSF" & j & "_MESDATA7").Value = vbNullString     'EPOSFx測定点7
        .Fields("EPOSF" & j & "_MESDATA8").Value = vbNullString     'EPOSFx測定点8
        .Fields("EPOSF" & j & "_MESDATA9").Value = vbNullString     'EPOSFx測定点9
        .Fields("EPOSF" & j & "_MESDATA10").Value = vbNullString    'EPOSFx測定点10
        .Fields("EPOSF" & j & "_MESDATA11").Value = vbNullString    'EPOSFx測定点11
        .Fields("EPOSF" & j & "_MESDATA12").Value = vbNullString    'EPOSFx測定点12
        .Fields("EPOSF" & j & "_MESDATA13").Value = vbNullString    'EPOSFx測定点13
        .Fields("EPOSF" & j & "_MESDATA14").Value = vbNullString    'EPOSFx測定点14
        .Fields("EPOSF" & j & "_MESDATA15").Value = vbNullString    'EPOSFx測定点15
    End With
    
    '-------------------- TBCMY022の読み込み(EPOSF) ----------------------------------------
    If (recXSDCW("EPINDL" & j & "CW").Value <> "0") And _
        (recXSDCW("EPRESL" & j & "CW").Value <> "0") Then
        
        sql = "select "
        sql = sql & "SAMPLEID "
        sql = sql & ",OSITEM "
        sql = sql & ",MAISU "
        sql = sql & ",SPEC "
        sql = sql & ",NETSU "
        sql = sql & ",ET "
        sql = sql & ",MES "
        sql = sql & ",DKAN "
        sql = sql & ",MESDATA1 "
        sql = sql & ",MESDATA2 "
        sql = sql & ",MESDATA3 "
        sql = sql & ",MESDATA4 "
        sql = sql & ",MESDATA5 "
        sql = sql & ",MESDATA6 "
        sql = sql & ",MESDATA7 "
        sql = sql & ",MESDATA8 "
        sql = sql & ",MESDATA9 "
        sql = sql & ",MESDATA10 "
        sql = sql & ",MESDATA11 "
        sql = sql & ",MESDATA12 "
        sql = sql & ",MESDATA13 "
        sql = sql & ",MESDATA14 "
        sql = sql & ",MESDATA15 "
        sql = sql & ",TXID "
        sql = sql & ",REGDATE "
        sql = sql & ",SENDFLAG "
        sql = sql & ",SENDDATE "
        sql = sql & "from TBCMY022 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("EPSMPLIDL" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'OSF" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX004
        With recX004
            .Fields("EPOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("EPOSF" & j & "_NETSU").Value = rs("NETSU")                 'EPOSFx_熱処理条件
            .Fields("EPOSF" & j & "_ET").Value = rs("ET")                       'EPOSFx_ｴｯﾁﾝｸﾞ条件
            .Fields("EPOSF" & j & "_MES").Value = rs("MES")                     'EPOSFx_計測方法
        End With
            
        'TBCMX005
        With recX005
            .Fields("EPOSF" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPOSFxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("EPOSF" & j & "_NETSU").Value = rs("NETSU")                 'EPOSFx_熱処理条件
            .Fields("EPOSF" & j & "_ET").Value = rs("ET")                       'EPOSFx_ｴｯﾁﾝｸﾞ条件
            .Fields("EPOSF" & j & "_MES").Value = rs("MES")                     'EPOSFx_計測方法
            .Fields("EPOSF" & j & "_DKAN").Value = rs("DKAN")                   'EPOSFx_DKｱﾆｰﾙ条件
            .Fields("EPOSF" & j & "_MESDATA1").Value = rs("MESDATA1")           'EPOSFx測定点1
            .Fields("EPOSF" & j & "_MESDATA2").Value = rs("MESDATA2")           'EPOSFx測定点2
            .Fields("EPOSF" & j & "_MESDATA3").Value = rs("MESDATA3")           'EPOSFx測定点3
            .Fields("EPOSF" & j & "_MESDATA4").Value = rs("MESDATA4")           'EPOSFx測定点4
            .Fields("EPOSF" & j & "_MESDATA5").Value = rs("MESDATA5")           'EPOSFx測定点5
            .Fields("EPOSF" & j & "_MESDATA6").Value = rs("MESDATA6")           'EPOSFx測定点6
            .Fields("EPOSF" & j & "_MESDATA7").Value = rs("MESDATA7")           'EPOSFx測定点7
            .Fields("EPOSF" & j & "_MESDATA8").Value = rs("MESDATA8")           'EPOSFx測定点8
            .Fields("EPOSF" & j & "_MESDATA9").Value = rs("MESDATA9")           'EPOSFx測定点9
            .Fields("EPOSF" & j & "_MESDATA10").Value = rs("MESDATA10")         'EPOSFx測定点10
            .Fields("EPOSF" & j & "_MESDATA11").Value = rs("MESDATA11")         'EPOSFx測定点11
            .Fields("EPOSF" & j & "_MESDATA12").Value = rs("MESDATA12")         'EPOSFx測定点12
            .Fields("EPOSF" & j & "_MESDATA13").Value = rs("MESDATA13")         'EPOSFx測定点13
            .Fields("EPOSF" & j & "_MESDATA14").Value = rs("MESDATA14")         'EPOSFx測定点14
            .Fields("EPOSF" & j & "_MESDATA15").Value = rs("MESDATA15")         'EPOSFx測定点15
        End With
        Set rs = Nothing
        
        'EPOSF_MAX,AVE
        sql = "select "
        sql = sql & "HEPOF" & j & "SH "
        sql = sql & ",HEPOF" & j & "ST "
        sql = sql & ",HEPOF" & j & "SR "
        sql = sql & ",HEPOF" & j & "HT "
        sql = sql & ",HEPOF" & j & "HS "
        sql = sql & ",HEPOF" & j & "KN "
        sql = sql & ",HEPOF" & j & "AX "
        sql = sql & ",HEPOF" & j & "MX "
        sql = sql & ",HEPOSF" & j & "PTK "
        sql = sql & "from TBCME050 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        If IsNull(rs("HEPOF" & j & "SH")) = False Then eosf.GuaranteeOsf.cMeth = rs("HEPOF" & j & "SH")     '品EPOSFx測定位置_方
        If IsNull(rs("HEPOF" & j & "ST")) = False Then eosf.GuaranteeOsf.cCount = rs("HEPOF" & j & "ST")    '品EPOSFx測定位置_点
        If IsNull(rs("HEPOF" & j & "SR")) = False Then eosf.GuaranteeOsf.cPos = rs("HEPOF" & j & "SR")      '品EPOSFx測定位置_領
        If IsNull(rs("HEPOF" & j & "HT")) = False Then eosf.GuaranteeOsf.cObj = rs("HEPOF" & j & "HT")      '品EPOSFx保証方法_対
        If IsNull(rs("HEPOF" & j & "HS")) = False Then eosf.GuaranteeOsf.cJudg = rs("HEPOF" & j & "HS")     '品EPOSFx保証方法_処
        If IsNull(rs("HEPOF" & j & "KN")) = False Then HEPOSFKN = rs("HEPOF" & j & "KN")                    '品EPOSFx検査頻度_抜
        If IsNull(rs("HEPOF" & j & "AX")) = False Then eosf.SpecOsfAveMax = rs("HEPOF" & j & "AX")          '品EPOSFx平均上限
        If IsNull(rs("HEPOF" & j & "MX")) = False Then eosf.SpecOsfMax = rs("HEPOF" & j & "MX")             '品EPOSFx上限
        If IsNull(rs("HEPOSF" & j & "PTK")) = False Then eosf.JudgDataPTK = rs("HEPOSF" & j & "PTK")        '品EPOSFxﾊﾟﾀﾝ区分
        Set rs = Nothing
            
        If eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "3" Then
            keisu = keisu1
        ElseIf eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "5" Then
            keisu = keisu2
        ElseIf eosf.GuaranteeOsf.cMeth = "5" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu3
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "3" Then
            keisu = keisu4
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "5" Then
            keisu = keisu5
        ElseIf eosf.GuaranteeOsf.cMeth = "6" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu6
        ElseIf eosf.GuaranteeOsf.cMeth = "E" And eosf.GuaranteeOsf.cCount = "5" And eosf.GuaranteeOsf.cPos = "A" Then
            keisu = keisu7
        Else
            keisu = -1
        End If
            
        If keisu <> -1 Then
            With recX005
                If IsNull(.Fields("EPOSF" & j & "_MESDATA1").Value) = False Then
                    eosf.OSF(0) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA1").Value)   'OSF測定値1
                Else
                    eosf.OSF(0) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA2").Value) = False Then
                    eosf.OSF(1) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA2").Value)   'OSF測定値2
                Else
                    eosf.OSF(1) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA3").Value) = False Then
                    eosf.OSF(2) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA3").Value)   'OSF測定値3
                Else
                    eosf.OSF(2) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA4").Value) = False Then
                    eosf.OSF(3) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA4").Value)   'OSF測定値4
                Else
                    eosf.OSF(3) = -1
                End If
                If IsNull(.Fields("EPOSF" & j & "_MESDATA5").Value) = False Then
                    eosf.OSF(4) = NtoZ2(.Fields("EPOSF" & j & "_MESDATA5").Value)   'OSF測定値5
                Else
                    eosf.OSF(4) = -1
                End If
                For k = 0 To 4
                    eosf.OSF(k) = IIf(eosf.OSF(k) <> -1, eosf.OSF(k) * keisu, -1)
                Next
            End With
            
            recX004.Fields("EPOSF" & j & "_MAX").Value = JudgMax(eosf.OSF())        'EPOSFx_判定時のMAX値_x
            recX004.Fields("EPOSF" & j & "_AVE").Value = JudgAve(eosf.OSF())        'EPOSFx_判定時のAVE値_x
            '>>>>> 2011/06/24 SETsw)Marushita MIN値セット追加対応
            If FieldCheck(sTblName, "EPOSF" & j & "_MIN") = FUNCTION_RETURN_SUCCESS Then
                recX004.Fields("EPOSF" & j & "_MIN").Value = JudgMin(eosf.OSF())        'EPOSFx_判定時のMIN値_x
            End If
            '<<<<< 2011/06/24 SETsw)Marushita MIN値セット追加対応
        End If
            
        '保証方法="H"かつMAX値/AVE値が-1の場合ｴﾗｰとする
        If ((eosf.GuaranteeOsf.cJudg = "H") And CheckKHN_EP(HEPOSFKN, j + 3, sPos)) And _
           (recX004.Fields("EPOSF" & j & "_MAX").Value = -1 Or _
            recX004.Fields("EPOSF" & j & "_AVE").Value = -1) Then GoTo proc_exit
        '？？？？？MIN値もチェック？？？？？
    End If

    getTBCMY022EPOSF = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY022EPOSF = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ｴﾋﾟBMD1～3実績(TBCMY022)ﾃﾞｰﾀ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :recXSDCW        , I  ,c_cmzcrec         , 新ｻﾝﾌﾟﾙ管理(SXL)
'          :j               , I  ,Integer           , ｴﾋﾟBMD No
'          :HIN             , I  ,tFullHinban       , 品番
'　　      :sPos  　　　    , I  ,String 　         , SXL位置(TOP/BOT)
'          :recX004         , O  ,c_cmzcrec         , EP検査書
'          :recX005         , O  ,c_cmzcrec         , EP測定点ﾃﾞｰﾀ
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :EP先行評価結果(TBCMY022)から実績ﾃﾞｰﾀを取得し、EP検査書/EP測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :06/08/10 ooba
Private Function getTBCMY022EPBMD(recXSDCW As c_cmzcrec, j As Integer, _
                                  HIN As tFullHinban, sPos As String, _
                                  recX004 As c_cmzcrec, recX005 As c_cmzcrec) As FUNCTION_RETURN
                                  
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim ebmd        As W_BMD                    'ｴﾋﾟBMD構造体
    Dim k           As Integer
    Dim HEPBMDKN    As String                   '検査頻度_抜
    Dim JData(4)    As Double
    
    Dim keisu As Double
    Const keisu1 As Double = 10000
    Const keisu2 As Double = 10000
    Const keisu3 As Double = 10000
    Const keisu4 As Double = 10000
    Const keisu5 As Double = 10000
    Const keisu6 As Double = 333000
    Const keisu7 As Double = 10000
    Const keisu8 As Double = 10000

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY022EPBMD"
    
    getTBCMY022EPBMD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX004
    With recX004
        .Fields("EPBMD" & j & "_SMPPOS").Value = vbNullString       'EPBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("EPBMD" & j & "_NETSU").Value = vbNullString        'EPBMDx_熱処理条件
        .Fields("EPBMD" & j & "_ET").Value = vbNullString           'EPBMDx_ｴｯﾁﾝｸﾞ条件
        .Fields("EPBMD" & j & "_MES").Value = vbNullString          'EPBMDx_計測方法
        .Fields("EPBMD" & j & "_MAX").Value = vbNullString          'EPBMDx_判定時のMAX値_x
        .Fields("EPBMD" & j & "_AVE").Value = vbNullString          'EPBMDx_判定時のAVE値_x
        .Fields("EPBMD" & j & "_MIN").Value = vbNullString          'EPBMDx_判定時のMIN値_x
        .Fields("EPBMD" & j & "_MBP").Value = vbNullString          'EPBMDx_判定時の面内分布
    End With
                
    'TBCMX005
    With recX005
        .Fields("EPBMD" & j & "_SMPPOS").Value = vbNullString       'EPBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("EPBMD" & j & "_NETSU").Value = vbNullString        'EPBMDx_熱処理条件
        .Fields("EPBMD" & j & "_ET").Value = vbNullString           'EPBMDx_ｴｯﾁﾝｸﾞ条件
        .Fields("EPBMD" & j & "_MES").Value = vbNullString          'EPBMDx_計測方法
        .Fields("EPBMD" & j & "_DKAN").Value = vbNullString         'EPBMDx_DKｱﾆｰﾙ条件
        .Fields("EPBMD" & j & "_MESDATA1").Value = vbNullString     'EPBMDx測定点1
        .Fields("EPBMD" & j & "_MESDATA2").Value = vbNullString     'EPBMDx測定点2
        .Fields("EPBMD" & j & "_MESDATA3").Value = vbNullString     'EPBMDx測定点3
        .Fields("EPBMD" & j & "_MESDATA4").Value = vbNullString     'EPBMDx測定点4
        .Fields("EPBMD" & j & "_MESDATA5").Value = vbNullString     'EPBMDx測定点5
        .Fields("EPBMD" & j & "_MESDATA6").Value = vbNullString     'EPBMDx測定点6
        .Fields("EPBMD" & j & "_MESDATA7").Value = vbNullString     'EPBMDx測定点7
        .Fields("EPBMD" & j & "_MESDATA8").Value = vbNullString     'EPBMDx測定点8
        .Fields("EPBMD" & j & "_MESDATA9").Value = vbNullString     'EPBMDx測定点9
        .Fields("EPBMD" & j & "_MESDATA10").Value = vbNullString    'EPBMDx測定点10
        .Fields("EPBMD" & j & "_MESDATA11").Value = vbNullString    'EPBMDx測定点11
        .Fields("EPBMD" & j & "_MESDATA12").Value = vbNullString    'EPBMDx測定点12
        .Fields("EPBMD" & j & "_MESDATA13").Value = vbNullString    'EPBMDx測定点13
        .Fields("EPBMD" & j & "_MESDATA14").Value = vbNullString    'EPBMDx測定点14
        .Fields("EPBMD" & j & "_MESDATA15").Value = vbNullString    'EPBMDx測定点15
    End With
    
    '-------------------- TBCMY022の読み込み(EPBMD) ----------------------------------------
    If (recXSDCW("EPINDB" & j & "CW").Value <> "0") And _
    (recXSDCW("EPRESB" & j & "CW").Value <> "0") Then
    
        sql = "select "
        sql = sql & "SAMPLEID "
        sql = sql & ",OSITEM "
        sql = sql & ",MAISU "
        sql = sql & ",SPEC "
        sql = sql & ",NETSU "
        sql = sql & ",ET "
        sql = sql & ",MES "
        sql = sql & ",DKAN "
        sql = sql & ",MESDATA1 "
        sql = sql & ",MESDATA2 "
        sql = sql & ",MESDATA3 "
        sql = sql & ",MESDATA4 "
        sql = sql & ",MESDATA5 "
        sql = sql & ",MESDATA6 "
        sql = sql & ",MESDATA7 "
        sql = sql & ",MESDATA8 "
        sql = sql & ",MESDATA9 "
        sql = sql & ",MESDATA10 "
        sql = sql & ",MESDATA11 "
        sql = sql & ",MESDATA12 "
        sql = sql & ",MESDATA13 "
        sql = sql & ",MESDATA14 "
        sql = sql & ",MESDATA15 "
        sql = sql & ",TXID "
        sql = sql & ",REGDATE "
        sql = sql & ",SENDFLAG "
        sql = sql & ",SENDDATE "
        sql = sql & "from TBCMY022 "
        sql = sql & "where SAMPLEID = '" & recXSDCW("EPSMPLIDB" & j & "CW").Value & "' and "
        sql = sql & "      SPEC = 'BMD" & j & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX004
        With recX004
            .Fields("EPBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("EPBMD" & j & "_NETSU").Value = rs("NETSU")                 'EPBMDx_熱処理条件
            .Fields("EPBMD" & j & "_ET").Value = rs("ET")                       'EPBMDx_ｴｯﾁﾝｸﾞ条件
            .Fields("EPBMD" & j & "_MES").Value = rs("MES")                     'EPBMDx_計測方法
        End With
            
        'TBCMX005
        With recX005
            .Fields("EPBMD" & j & "_SMPPOS").Value = recXSDCW("INPOSCW").Value  'EPBMDxｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("EPBMD" & j & "_NETSU").Value = rs("NETSU")                 'EPBMDx_熱処理条件
            .Fields("EPBMD" & j & "_ET").Value = rs("ET")                       'EPBMDx_ｴｯﾁﾝｸﾞ条件
            .Fields("EPBMD" & j & "_MES").Value = rs("MES")                     'EPBMDx_計測方法
            .Fields("EPBMD" & j & "_DKAN").Value = rs("DKAN")                   'EPBMDx_DKｱﾆｰﾙ条件
            .Fields("EPBMD" & j & "_MESDATA1").Value = rs("MESDATA1")           'EPBMDx測定点1
            .Fields("EPBMD" & j & "_MESDATA2").Value = rs("MESDATA2")           'EPBMDx測定点2
            .Fields("EPBMD" & j & "_MESDATA3").Value = rs("MESDATA3")           'EPBMDx測定点3
            .Fields("EPBMD" & j & "_MESDATA4").Value = rs("MESDATA4")           'EPBMDx測定点4
            .Fields("EPBMD" & j & "_MESDATA5").Value = rs("MESDATA5")           'EPBMDx測定点5
            .Fields("EPBMD" & j & "_MESDATA6").Value = rs("MESDATA6")           'EPBMDx測定点6
            .Fields("EPBMD" & j & "_MESDATA7").Value = rs("MESDATA7")           'EPBMDx測定点7
            .Fields("EPBMD" & j & "_MESDATA8").Value = rs("MESDATA8")           'EPBMDx測定点8
            .Fields("EPBMD" & j & "_MESDATA9").Value = rs("MESDATA9")           'EPBMDx測定点9
            .Fields("EPBMD" & j & "_MESDATA10").Value = rs("MESDATA10")         'EPBMDx測定点10
            .Fields("EPBMD" & j & "_MESDATA11").Value = rs("MESDATA11")         'EPBMDx測定点11
            .Fields("EPBMD" & j & "_MESDATA12").Value = rs("MESDATA12")         'EPBMDx測定点12
            .Fields("EPBMD" & j & "_MESDATA13").Value = rs("MESDATA13")         'EPBMDx測定点13
            .Fields("EPBMD" & j & "_MESDATA14").Value = rs("MESDATA14")         'EPBMDx測定点14
            .Fields("EPBMD" & j & "_MESDATA15").Value = rs("MESDATA15")         'EPBMDx測定点15
        End With
        Set rs = Nothing
                    
        'EPBMD_MAX,MIN,AVE,MBP
        sql = "select "
        sql = sql & "HEPBM" & j & "SH "
        sql = sql & ",HEPBM" & j & "ST "
        sql = sql & ",HEPBM" & j & "SR "
        sql = sql & ",HEPBM" & j & "HT "
        sql = sql & ",HEPBM" & j & "HS "
        sql = sql & ",HEPBM" & j & "KN "
        sql = sql & ",HEPBM" & j & "AN "
        sql = sql & ",HEPBM" & j & "AX "
        sql = sql & ",HEPBM" & j & "MBP "
        sql = sql & ",HEPBM" & j & "MCL "
        sql = sql & "from TBCME050 where "
        sql = sql & "HINBAN = '" & HIN.hinban & "' and "
        sql = sql & "MNOREVNO = " & HIN.mnorevno & " and "
        sql = sql & "FACTORY = '" & HIN.factory & "' and "
        sql = sql & "OPECOND = '" & HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
                    
        If IsNull(rs("HEPBM" & j & "SH")) = False Then ebmd.GuaranteeBmd.cMeth = rs("HEPBM" & j & "SH")     '品EPBMDx測定位置_方
        If IsNull(rs("HEPBM" & j & "ST")) = False Then ebmd.GuaranteeBmd.cCount = rs("HEPBM" & j & "ST")    '品EPBMDx測定位置_点
        If IsNull(rs("HEPBM" & j & "SR")) = False Then ebmd.GuaranteeBmd.cPos = rs("HEPBM" & j & "SR")      '品EPBMDx測定位置_領
        If IsNull(rs("HEPBM" & j & "HT")) = False Then ebmd.GuaranteeBmd.cObj = rs("HEPBM" & j & "HT")      '品EPBMDx保証方法_対
        If IsNull(rs("HEPBM" & j & "HS")) = False Then ebmd.GuaranteeBmd.cJudg = rs("HEPBM" & j & "HS")     '品EPBMDx保証方法_処
        If IsNull(rs("HEPBM" & j & "KN")) = False Then HEPBMDKN = rs("HEPBM" & j & "KN")                    '品EPBMDx検査頻度_抜
        If IsNull(rs("HEPBM" & j & "AN")) = False Then ebmd.SpecBmdAveMin = rs("HEPBM" & j & "AN")          '品EPBMDx平均下限
        If IsNull(rs("HEPBM" & j & "AX")) = False Then ebmd.SpecBmdAveMax = rs("HEPBM" & j & "AX")          '品EPBMDx平均上限
        If IsNull(rs("HEPBM" & j & "MBP")) = False Then ebmd.SpecBmdMBP = rs("HEPBM" & j & "MBP")           '品EPBMDx面内分布
        If IsNull(rs("HEPBM" & j & "MCL")) = False Then ebmd.SpecBmdMCL = rs("HEPBM" & j & "MCL")           '品EPBMDx面内計算
        Set rs = Nothing

        If ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "H" Then
            keisu = keisu1
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "H" Then
            keisu = keisu2
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu3
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu4
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "5" And ebmd.GuaranteeBmd.cPos = "A" Then
            keisu = keisu5
        ElseIf ebmd.GuaranteeBmd.cMeth = "G" And ebmd.GuaranteeBmd.cCount = "3" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu6
        ElseIf ebmd.GuaranteeBmd.cMeth = "2" And ebmd.GuaranteeBmd.cCount = "5" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu7
        ElseIf ebmd.GuaranteeBmd.cMeth = "8" And ebmd.GuaranteeBmd.cCount = "4" And ebmd.GuaranteeBmd.cPos = "8" Then
            keisu = keisu8
        Else
            keisu = -1
        End If

        If keisu <> -1 Then

            With recX005
                If IsNull(.Fields("EPBMD" & j & "_MESDATA1").Value) = False Then
                    ebmd.BMD(0) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA1").Value)   'BMD測定値1
                Else
                    ebmd.BMD(0) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA2").Value) = False Then
                    ebmd.BMD(1) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA2").Value)   'BMD測定値2
                Else
                    ebmd.BMD(1) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA3").Value) = False Then
                    ebmd.BMD(2) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA3").Value)   'BMD測定値3
                Else
                    ebmd.BMD(2) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA4").Value) = False Then
                    ebmd.BMD(3) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA4").Value)   'BMD測定値4
                Else
                    ebmd.BMD(3) = -1
                End If
                If IsNull(.Fields("EPBMD" & j & "_MESDATA5").Value) = False Then
                    ebmd.BMD(4) = NtoZ2(.Fields("EPBMD" & j & "_MESDATA5").Value)   'BMD測定値5
                Else
                    ebmd.BMD(4) = -1
                End If
                For k = 0 To 4
                    ebmd.BMD(k) = IIf(ebmd.BMD(k) <> -1, ebmd.BMD(k) * CDbl(keisu / 10000), -1)
                Next
            End With
                
            'EPBMDx_判定時のMAX値_x
            '判定ｺｰﾄﾞが "F"：MAX(2,4点目)，"G"：MAX(2,3,4点目) の場合はMAX値にその値をｾｯﾄ
            If ebmd.GuaranteeBmd.cJudg = JudgCodeW01 And _
                (ebmd.GuaranteeBmd.cObj = ObjCode10 Or ebmd.GuaranteeBmd.cObj = ObjCode11) Then
            
                If GetWfJudgData(WFBMD_JUDG, ebmd.GuaranteeBmd, ebmd.BMD(), JData()) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                recX004.Fields("EPBMD" & j & "_MAX").Value = JData(0)
            Else
                recX004.Fields("EPBMD" & j & "_MAX").Value = JudgMax(ebmd.BMD())
            End If
            recX004.Fields("EPBMD" & j & "_AVE").Value = JudgAve(ebmd.BMD())        'EPBMDx_判定時のAVE値_x
            recX004.Fields("EPBMD" & j & "_MIN").Value = JudgMin(ebmd.BMD())        'EPBMDx_判定時のMIN値_x
            If ebmd.SpecBmdMCL = "P " Then
                recX004.Fields("EPBMD" & j & "_MBP").Value = JudgBmdMBP(ebmd.BMD()) 'EPBMDx_判定時の面内分布
            Else
                recX004.Fields("EPBMD" & j & "_MBP").Value = 0                      '面内分布が"P"以外の時は計算結果を0とする
            End If
            
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech Start
            If ebmd.GuaranteeBmd.cObj = ObjCode18 Then
                recX004.Fields("EPBMD" & j & "_MBP").Value = ebmd.BMD(0)
            End If
'' 2008/10/20 BMD評価,外周1点保証機能追加 ADD By Systech End
            
        End If
        
        '保証方法="H"かつ面内分布が-1の場合ｴﾗｰとする
        If ((ebmd.GuaranteeBmd.cJudg = "H") And CheckKHN_EP(HEPBMDKN, j, sPos)) _
            And (recX004.Fields("EPBMD" & j & "_MBP").Value = -1) Then GoTo proc_exit
        
    End If

    getTBCMY022EPBMD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY022EPBMD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :共有サンプルチェック処理
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :inSXLID         , I  ,String            , SXL-ID
'          :inSMPLID        , I  ,String            , ｻﾝﾌﾟﾙID
'          :outSMPLID       , O  ,String            , 共有ｻﾝﾌﾟﾙID(共有でない場合、inSMPLIDを返す)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :指定されたｻﾝﾌﾟﾙIDが全共有かどうかをﾁｪｯｸし、全共有の場合、共有ｻﾝﾌﾟﾙIDを取得し返す
'履歴      :2003/11/19 SystemBrain 新規作成
Private Function chkComSAMPL(inSXLID As String, inSMPLID As String, outSMPLID As String) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim wXTALCW     As String
    Dim wINPOSCW    As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function chkComSAMPL"
    
    '-------------------- 初期ｸﾘｱ ----------------------------------------
    chkComSAMPL = FUNCTION_RETURN_SUCCESS
    outSMPLID = inSMPLID
    
    '-------------------- 全共有確認(XSDCW) ----------------------------------------
    sql = "select XTALCW, INPOSCW from XSDCW "
    sql = sql & "where SXLIDCW = '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW = '" & inSMPLID & "' and "
    sql = sql & "      (WFINDRSCW = '2' or WFINDRSCW = '0' or WFINDRSCW = ' ' or WFINDRSCW is null) and "
    sql = sql & "      (WFINDOICW = '2' or WFINDOICW = '0' or WFINDOICW = ' ' or WFINDOICW is null) and "
    sql = sql & "      (WFINDB1CW = '2' or WFINDB1CW = '0' or WFINDB1CW = ' ' or WFINDB1CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB2CW = '0' or WFINDB2CW = ' ' or WFINDB2CW is null) and "
    sql = sql & "      (WFINDB2CW = '2' or WFINDB3CW = '0' or WFINDB3CW = ' ' or WFINDB3CW is null) and "
    sql = sql & "      (WFINDL1CW = '2' or WFINDL1CW = '0' or WFINDL1CW = ' ' or WFINDL1CW is null) and "
    sql = sql & "      (WFINDL2CW = '2' or WFINDL2CW = '0' or WFINDL2CW = ' ' or WFINDL2CW is null) and "
    sql = sql & "      (WFINDL3CW = '2' or WFINDL3CW = '0' or WFINDL3CW = ' ' or WFINDL3CW is null) and "
    sql = sql & "      (WFINDL4CW = '2' or WFINDL4CW = '0' or WFINDL4CW = ' ' or WFINDL4CW is null) and "
    sql = sql & "      (WFINDDSCW = '2' or WFINDDSCW = '0' or WFINDDSCW = ' ' or WFINDDSCW is null) and "
    sql = sql & "      (WFINDDZCW = '2' or WFINDDZCW = '0' or WFINDDZCW = ' ' or WFINDDZCW is null) and "
    sql = sql & "      (WFINDSPCW = '2' or WFINDSPCW = '0' or WFINDSPCW = ' ' or WFINDSPCW is null) and "
    sql = sql & "      (WFINDDO1CW = '2' or WFINDDO1CW = '0' or WFINDDO1CW = ' ' or WFINDDO1CW is null) and "
    sql = sql & "      (WFINDDO2CW = '2' or WFINDDO2CW = '0' or WFINDDO2CW = ' ' or WFINDDO2CW is null) and "
    sql = sql & "      (WFINDDO3CW = '2' or WFINDDO3CW = '0' or WFINDDO3CW = ' ' or WFINDDO3CW is null) and "
    sql = sql & "      (WFINDOT1CW = '2' or WFINDOT1CW = '0' or WFINDOT1CW = ' ' or WFINDOT1CW is null) and "
    sql = sql & "      (WFINDOT2CW = '2' or WFINDOT2CW = '0' or WFINDOT2CW = ' ' or WFINDOT2CW is null) and "
    ''残存酸素追加　03/12/19 ooba
    sql = sql & "      (WFINDAOICW = '2' or WFINDAOICW = '0' or WFINDAOICW = ' ' or WFINDAOICW is null) and "
    ''GD追加　2005/02/17 ffc)tanabe
    sql = sql & "      (((WFINDGDCW = '2' or WFINDGDCW = '0' or WFINDGDCW = ' ' or WFINDGDCW is null) and WFHSGDCW = '0') or WFHSGDCW = '1') "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    sql = sql & "  and (EPINDB1CW = '2' or EPINDB1CW = '0' or EPINDB1CW = ' ' or EPINDB1CW is null) and "
    sql = sql & "      (EPINDB2CW = '2' or EPINDB2CW = '0' or EPINDB2CW = ' ' or EPINDB2CW is null) and "
    sql = sql & "      (EPINDB3CW = '2' or EPINDB3CW = '0' or EPINDB3CW = ' ' or EPINDB3CW is null) and "
    sql = sql & "      (EPINDL1CW = '2' or EPINDL1CW = '0' or EPINDL1CW = ' ' or EPINDL1CW is null) and "
    sql = sql & "      (EPINDL2CW = '2' or EPINDL2CW = '0' or EPINDL2CW = ' ' or EPINDL2CW is null) and "
    sql = sql & "      (EPINDL3CW = '2' or EPINDL3CW = '0' or EPINDL3CW = ' ' or EPINDL3CW is null) "
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    wXTALCW = rs("XTALCW")      '結晶番号
    wINPOSCW = rs("INPOSCW")    '結晶内位置
    Set rs = Nothing
    
    '-------------------- 共有ｻﾝﾌﾟﾙIDの取得(XSDCW) ----------------------------------------
    sql = "select REPSMPLIDCW from XSDCW "
    sql = sql & "where SXLIDCW like '" & left(wXTALCW, 9) & "%' and "       '09/05/26 ooba
    sql = sql & "      XTALCW = '" & wXTALCW & "' and "
    sql = sql & "      INPOSCW = '" & wINPOSCW & "' and "
    sql = sql & "      NUKISIFLGCW = '1' and "                              '09/05/26 ooba
    sql = sql & "      SXLIDCW != '" & inSXLID & "' and "
    sql = sql & "      REPSMPLIDCW != '" & inSMPLID & "' "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    outSMPLID = rs("REPSMPLIDCW")       '代表ｻﾝﾌﾟﾙID(共有)
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    chkComSAMPL = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :SXL確定指示(TBCMY007)ﾃｰﾌﾞﾙにｾｯﾄするSXLの比抵抗ﾃﾞｰﾀを取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO  ,型                :説明
'          :SXLID          ,I   ,String            ,SXLID
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)   04/04/15 ooba
'          :sPattern       ,I   ,String            ,比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
'                                                   ●ﾊﾟﾀｰﾝA : WF実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝB : 結晶実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝC : 取得ﾃﾞｰﾀなし
'          :mesdata()      ,O   ,String            ,比抵抗ﾃﾞｰﾀ
'          :戻り値          ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :04/02/12 ooba　作成
Public Function cmbc040_GetSxlRsData(SXLID As String, sPos As String, sPattern As String, mesdata() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B区分
    Dim i As Integer
    Dim j As Integer
    Dim sSql As String
    Dim rs As OraDynaset
    Dim dTmpData(10) As Double   '比抵抗(Rs)ﾃﾞｰﾀ
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsData"
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『A』の場合、WF実績ﾃﾞｰﾀ(TBCMY013)を取得する。
    If sPattern = "A" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のｻﾝﾌﾟﾙID_Rsを取得。
        'ｻﾝﾌﾟﾙID_Rsから、測定評価結果<TBCMY013>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
        sSql = "select MESDATA1, MESDATA2, MESDATA3, MESDATA4, MESDATA5 "
        sSql = sSql & "from TBCMY013 "
        sSql = sSql & "where OSITEM = 'RES' "
        sSql = sSql & "and SAMPLEID in ( "
        sSql = sSql & "         select WFSMPLIDRSCW from XSDCW "
        sSql = sSql & "         where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "         and SXLIDCW = '" & SXLID & "') "
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP側実績ﾃﾞｰﾀ
            If sTBkbn = "T" Then
                mesdata(1) = rs("MESDATA1")
                mesdata(2) = rs("MESDATA2")
                mesdata(3) = rs("MESDATA3")
                mesdata(4) = rs("MESDATA4")
                mesdata(5) = rs("MESDATA5")
            'BOT側実績ﾃﾞｰﾀ
            ElseIf sTBkbn = "B" Then
                mesdata(6) = rs("MESDATA1")
                mesdata(7) = rs("MESDATA2")
                mesdata(8) = rs("MESDATA3")
                mesdata(9) = rs("MESDATA4")
                mesdata(10) = rs("MESDATA5")
            End If
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『B』の場合、結晶実績ﾃﾞｰﾀ(TBCMJ002)を取得する。
    ElseIf sPattern = "B" Then
'''        For i = 1 To 2
'''            If i = 1 Then sTBkbn = "T" Else sTBkbn = "B"
        '該当SXLより、新ｻﾝﾌﾟﾙ管理-WF<XSDCW>のT/B区分、ｻﾝﾌﾟﾙﾌﾞﾛｯｸIDを取得。
        'T/B区分、ｻﾝﾌﾟﾙﾌﾞﾛｯｸIDから、新ｻﾝﾌﾟﾙ管理-ﾌﾞﾛｯｸ<XSDCS>の結晶番号、ｻﾝﾌﾟﾙID_Rsを取得。
        '結晶番号、ｻﾝﾌﾟﾙID_Rsから、結晶抵抗実績<TBCMJ002>の比抵抗実績ﾃﾞｰﾀ(TOP側/BOT側)を取得する。
        sSql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 "
        sSql = sSql & "from TBCMJ002 "
        sSql = sSql & "where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "         select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "         from XSDCS "
        sSql = sSql & "         where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                  select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                  from XSDCW "
        sSql = sSql & "                  where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                  and SXLIDCW = '" & SXLID & "')) "
        sSql = sSql & "and TRANCNT = ( "
        sSql = sSql & "         select max(TRANCNT) "
        sSql = sSql & "         from TBCMJ002 "
        sSql = sSql & "         where (CRYNUM, SMPLNO) in ( "
        sSql = sSql & "                  select XTALCS, CRYSMPLIDRSCS "
        sSql = sSql & "                  from XSDCS "
        sSql = sSql & "                  where (TBKBNCS, CRYNUMCS) in ( "
        sSql = sSql & "                           select TBKBNCW, SMCRYNUMCW "
        sSql = sSql & "                           from XSDCW "
        sSql = sSql & "                           where TBKBNCW = '" & sTBkbn & "' "
        sSql = sSql & "                           and SXLIDCW = '" & SXLID & "'))) "
    
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP側実績ﾃﾞｰﾀ
            If sTBkbn = "T" Then
                dTmpData(1) = rs("MEAS1")
                dTmpData(2) = rs("MEAS2")
                dTmpData(3) = rs("MEAS3")
                dTmpData(4) = rs("MEAS4")
                dTmpData(5) = rs("MEAS5")
                '型変換
                For j = 1 To 5
                    mesdata(j) = CStr(dTmpData(j))
                Next
            'BOT側実績ﾃﾞｰﾀ
            ElseIf sTBkbn = "B" Then
                dTmpData(6) = rs("MEAS1")
                dTmpData(7) = rs("MEAS2")
                dTmpData(8) = rs("MEAS3")
                dTmpData(9) = rs("MEAS4")
                dTmpData(10) = rs("MEAS5")
                '型変換
                For j = 6 To 10
                    mesdata(j) = CStr(dTmpData(j))
                Next
            End If
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
'''        Next
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『C』の場合、取得実績ﾃﾞｰﾀなし。
    ElseIf sPattern = "C" Then
    
    End If
    
    '取得ﾃﾞｰﾀが空白/-1/NULLの時はｽﾍﾟｰｽをｾｯﾄする。
'''    For i = 1 To 10
'''        If mesdata(i) = "" Or mesdata(i) = "-1" Or mesdata(i) = vbNullString Then
'''            mesdata(i) = " "
'''        End If
'''    Next
    For i = 1 To 5
        If sTBkbn = "T" Then j = i Else j = i + 5
        If mesdata(j) = "" Or mesdata(j) = "-1" Or mesdata(j) = vbNullString Then
            mesdata(j) = " "
        End If
    Next
    
    cmbc040_GetSxlRsData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :酸素析出と残存酸素の仕様チェック
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型              ,説明
'      　　:pHin　　    　,I  ,tFullHinban   　,品番
'      　　:戻り値        ,O  ,Integer       　,仕様チェック結果(-1:ｴﾗｰ，0:AOi仕様無，1:AOi仕様有)
'説明      :酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合エラーを返す
'          :≪s_cmzcF_cmkc001WF.bas≫内関数と同様
'履歴      :03/12/19 ooba

Public Function ChkAoiSiyou(pHIN As tFullHinban) As Integer

    Dim sSql As String
    Dim rs As OraDynaset
    Dim sDoiSiyou(2) As String  '検査有無(DOi1～3)
    Dim sAoiSiyou As String     '検査有無(AOi)
    Dim iCnt As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001WF.bas -- Function ChkAoiSiyou"

    sSql = "select HWFOS1HS, HWFOS2HS, HWFOS3HS, HWFZOHWS from TBCME025 "
    sSql = sSql & "where HINBAN = '" & pHIN.hinban & "' "
    sSql = sSql & "and MNOREVNO = " & pHIN.mnorevno & " "
    sSql = sSql & "and FACTORY = '" & pHIN.factory & "' "
    sSql = sSql & "and OPECOND = '" & pHIN.opecond & "' "

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        rs.Close
        ChkAoiSiyou = -1
        GoTo proc_exit
    End If
    
    If IsNull(rs("HWFOS1HS")) = False Then sDoiSiyou(0) = rs("HWFOS1HS") '品WF酸素析出1保証方法_処
    If IsNull(rs("HWFOS2HS")) = False Then sDoiSiyou(1) = rs("HWFOS2HS") '品WF酸素析出2保証方法_処
    If IsNull(rs("HWFOS3HS")) = False Then sDoiSiyou(2) = rs("HWFOS3HS") '品WF酸素析出3保証方法_処
    If IsNull(rs("HWFZOHWS")) = False Then sAoiSiyou = rs("HWFZOHWS")    '品WF残存酸素保証方法_処
    
    '酸素析出と残存酸素の仕様チェック
    ChkAoiSiyou = 0
    For iCnt = 0 To 2
        If sDoiSiyou(iCnt) = "H" Or sDoiSiyou(iCnt) = "S" Then
            '酸素析出(Δoi)と残存酸素の両方に仕様が立っていた場合はエラー
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = -1
                Exit For
            End If
        Else
            If sAoiSiyou = "H" Or sAoiSiyou = "S" Then
                ChkAoiSiyou = 1
            End If
        End If
    Next
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function
    
proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    ChkAoiSiyou = -1
    Resume proc_exit
    
End Function


'------------------------------------------------------------------------------------
'-概要    :ＷＦホールドコード＆理由取得処理
'-ﾊﾟﾗﾒｰﾀ  :変数名        ,IO  ,型                                     ,説明
'-        :sSXLID        ,I   ,string                                ,シングルID
'-        :sBLOCKID      ,I   ,string                                ,ブロックID
'-        :sHINBAN       ,I   ,string                                ,品番
'-        :sINGOTPOS     ,I   ,string                                ,結晶位置
'-        :sWFHOLDDATE   ,O   ,string                                ,ホールド日付
'-        :sUSER_ID      ,O   ,string                                ,WFホールド処理者ID
'-        :戻ﾘ値         ,O   ,FUNCTION_RETURN                       ,読み込み成否
'-説明    :TBCMY019[KEY:BLOCKID,TRANCNT]データを取得する。
'-         TBCMY019から取得したコードを元にKODA9[KEY:]から理由(日本語)を取得する。
'-履歴    :ＤＢ更新追加　2004/07/16 KOYAMA
'------------------------------------------------------------------------------------
Public Function DBDRV_s_cmbc040_SQL_Y019XSDCB(sSXLID As String, sBlockId As String, sHINBAN As String, _
                                                 sINGOTPOS As String, sWFHOLDDATE As String, _
                                                 sUSER_ID As String) As FUNCTION_RETURN



    Dim cbrs As OraDynaset         'XSDCB検索用カーソル
    Dim wfrs As OraDynaset         'TBCMY019検索用カーソル
    Dim sql As String
    Dim ksql As String
    Dim recCnt As Long
    Dim swfkbn As String           'WFホールド区分(1:WFホールド、0:ホールド解除)

    '変数初期化
    sql = ""
    ksql = ""
    recCnt = 0

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- DBDRV_s_cmbc040_SQL_Y019XSDCB"

    'XSDCB検索(WFホールド区分検索)
    sql = ""
    sql = "select "
    sql = sql & " WFHOLDFLGCB "                  ' WFホールド区分
    sql = sql & " from XSDCB "
    sql = sql & " where "
    sql = sql & " SXLIDCB = '" & sSXLID & "'"
    sql = sql & " and XTALCB = '" & sBlockId & "'"
    
    
 '   Debug.Print sql
    Set cbrs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'レコード0件時は正常
        
    If cbrs Is Nothing Then
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
        swfkbn = ""
        cbrs.Close
        GoTo proc_exit
    End If

    recCnt = cbrs.RecordCount
    'レコード0件時は処理日付、WFホールド処理者IDをスペースで返す。
    If recCnt = 0 Then
        swfkbn = ""
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    Else
        swfkbn = IIf(IsNull(cbrs("WFHOLDFLGCB")), "", cbrs("WFHOLDFLGCB"))  ' WFホールド区分
        cbrs.Close
    End If

    
    '**** WFホールド状態である場合(WFホールド区分=1)
    If swfkbn = "1" Then
    
    
        'TBCMY019検索(ＷＦホールド処理日、ホールド処理者ID取得)
        sql = ""
        sql = "select "
        sql = sql & " HOLDDT, "                  ' ホールド日付
        sql = sql & " USER_ID "                  ' WFホールド処理者ID
        sql = sql & " from TBCMY019 "
        sql = sql & " where "
        sql = sql & " TRANCNT = any(select MAX(TRANCNT)"
        sql = sql & " from TBCMY019 where BLOCKID ='" & sBlockId & "')"
        sql = sql & " and BLOCKID ='" & sBlockId & "'"
        
        Debug.Print sql
        Set wfrs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        'レコード0件時は正常
            
        If wfrs Is Nothing Then
            DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
            wfrs.Close
            GoTo proc_exit
        End If
    
        recCnt = wfrs.RecordCount
        'レコード0件時は処理日付、WFホールド処理者IDをスペースで返す。
        If recCnt = 0 Then
            sWFHOLDDATE = ""
            sUSER_ID = ""
        Else
            sWFHOLDDATE = IIf(IsNull(wfrs("HOLDDT")), "", wfrs("HOLDDT"))    ' ホールド日付
            sUSER_ID = IIf(IsNull(wfrs("USER_ID")), "", wfrs("USER_ID"))     ' WFホールド処理者ID
            wfrs.Close
        End If
    
    
    '**** WFホールド状態以外場合(WFホールド区分=0)
    Else
        DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_s_cmbc040_SQL_Y019XSDCB = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function

'概要      :WFGD実績(TBCMJ015)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003構造体(GD検査測定点ﾃﾞｰﾀ)
'          :sTblName        , I  ,String            , テーブル名　2011/06/23 Marushita GBG送信対応
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFGD実績(TBCMJ015)からﾃﾞｰﾀを取得し、GD検査測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2005/02/15 ffc)tanabe
'Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
Private Function getTBCMJ015WFGD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec, sTblName As String) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    Dim nFlg    As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ015WFGD"
    
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    '>>>>> 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
    'GBGのチェック
    If FieldCheck(sTblName, "SXLGD_HSFLG") = FUNCTION_RETURN_FAILURE Then
        nFlg = 1
    Else
        nFlg = 0
    End If
    '<<<<< 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
    'TBCMX003
    With recX003
        .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGDサンプル測定位置(SXL位置情報)
        .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_測定結果 Den
        .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_測定結果 L/DL
        .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_測定結果 DVD2
        '>>>>> 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
        If nFlg = 1 Then
        Else
        '<<<<< 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD測定結果保証フラグ
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_測定結果 DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD測定結果保証フラグ
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            .Fields("GD_PTNJUDGRES").Value = vbNullString                            'GDパターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
            
            For i = 1 To 15
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_測定値xx L/DL1
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_測定値xx L/DL2
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_測定値xx L/DL3
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_測定値xx L/DL4
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_測定値xx L/DL5
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_測定値xx Den1
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_測定値xx Den2
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_測定値xx Den3
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_測定値xx Den4
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_測定値xx Den5
            Next
            
            For i = 1 To 5
                .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_測定値xx DVD2
            Next
        End If
    End With
        
    '-------------------- TBCMJ015の読み込み(GD) ----------------------------------------
    sql = "select * from TBCMJ015 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDGDCW").Value & "'"
    sql = sql & " and   HSFLG = '1'"
    sql = sql & "order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        '>>>>> 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
        'GD実績がない場合はエラーにしない？
        If nFlg = 1 Then
            getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS
        End If
        '<<<<< 2011/06/24 SETsw)Marushita WFGD_判定時のセット対応
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("WFGD_SMPPOS").Value = rs("POSITION")                                                     'WFGDサンプル測定位置(SXL位置情報)
        .Fields("WFGD_MSRSDEN").Value = rs("MSRSDEN")                                                     'WFGD_測定結果 Den
        .Fields("WFGD_MSRSLDL").Value = rs("MSRSLDL")                                                     'WFGD_測定結果 L/DL
        .Fields("WFGD_MSRSDVD2").Value = rs("MSRSDVD2")                                                   'WFGD_測定結果 DVD2
        If nFlg = 1 Then
        Else
            .Fields("WFGD_HSFLG").Value = "1"                                                                 'WFGD測定結果保証フラグ
            
            For i = 1 To 15
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'WFGD_測定値xx Den1
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'WFGD_測定値xx Den2
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'WFGD_測定値xx Den3
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'WFGD_測定値xx Den4
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'WFGD_測定値xx Den5
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'WFGD_測定値xx L/DL1
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'WFGD_測定値xx L/DL2
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'WFGD_測定値xx L/DL3
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'WFGD_測定値xx L/DL4
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'WFGD_測定値xx L/DL5
            Next
            
            For i = 1 To 5
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")                                    'WFGD_測定値xx DVD2
            Next
        
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            'GDパターン判定結果
            If IsNull(.Fields("PTNJUDGRES")) = True Then
                .Fields("GD_PTNJUDGRES").Value = " "
            Else
                .Fields("GD_PTNJUDGRES").Value = rs("PTNJUDGRES")
            End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
        End If
    End With
    Set rs = Nothing

    getTBCMJ015WFGD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ015WFGD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶GD実績(TBCMJ006)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX003         , O  ,c_cmzcrec         , TBCMX003構造体(GD検査測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶GD実績(TBCMJ006)からﾃﾞｰﾀを取得し、GD検査測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'          :結晶GD実績(TBCMJ006)の測定データの初期値である-1をNULLに変更してTBCMX003に登録する。
'履歴      :2005/02/15 ffc)tanabe
Private Function getTBCMJ006GD(CRYNUM As String, recXSDCW As c_cmzcrec, recX003 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ006GD"
    
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
        
    'TBCMX003
    With recX003
            .Fields("SXLGD_HSFLG").Value = vbNullString                              'SXLGDGD測定結果保証フラグ
            .Fields("SXLGD_SMPPOS").Value = vbNullString                             'SXLGDGDサンプル測定位置(SXL位置情報)
            .Fields("SXLGD_MSRSDEN").Value = vbNullString                            'SXLGDGD_測定結果 Den
            .Fields("SXLGD_MSRSLDL").Value = vbNullString                            'SXLGDGD_測定結果 L/DL
            .Fields("SXLGD_MSRSDVD2").Value = vbNullString                           'SXLGDGD_測定結果 DVD2
            .Fields("WFGD_HSFLG").Value = vbNullString                               'WFGD測定結果保証フラグ
            .Fields("WFGD_SMPPOS").Value = vbNullString                              'WFGDサンプル測定位置(SXL位置情報)
            .Fields("WFGD_MSRSDEN").Value = vbNullString                             'WFGD_測定結果 Den
            .Fields("WFGD_MSRSLDL").Value = vbNullString                             'WFGD_測定結果 L/DL
            .Fields("WFGD_MSRSDVD2").Value = vbNullString                            'WFGD_測定結果 DVD2
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
            .Fields("GD_PTNJUDGRES").Value = vbNullString                            'GDパターン判定結果
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End
            
        For i = 1 To 15
            .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = vbNullString       'WFGD_測定値xx L/DL1
            .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = vbNullString       'WFGD_測定値xx L/DL2
            .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = vbNullString       'WFGD_測定値xx L/DL3
            .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = vbNullString       'WFGD_測定値xx L/DL4
            .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = vbNullString       'WFGD_測定値xx L/DL5
            .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = vbNullString       'WFGD_測定値xx Den1
            .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = vbNullString       'WFGD_測定値xx Den2
            .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = vbNullString       'WFGD_測定値xx Den3
            .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = vbNullString       'WFGD_測定値xx Den4
            .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = vbNullString       'WFGD_測定値xx Den5
        Next
        
        For i = 1 To 5
            .Fields("WFGD_MS01DVD2" & i).Value = vbNullString                        'WFGD_測定値xx DVD2
        Next
        
    End With
        
    '-------------------- TBCMJ006の読み込み(GD) ----------------------------------------
    sql = "select * from TBCMJ006 "
    sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
    sql = sql & "      SMPLNO = " & Trim(recXSDCW("WFSMPLIDGDCW").Value)
    sql = sql & " order by TRANCNT desc"
    sql = "select * from (" & sql & ") where rownum = 1"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    'TBCMX003
    With recX003
        .Fields("SXLGD_HSFLG").Value = "1"                          'SXLGD測定結果保証フラグ
        .Fields("SXLGD_SMPPOS").Value = rs("POSITION")              'SXLGDサンプル測定位置(SXL位置情報)
        If rs("MSRSDEN") <> -1 Then
            .Fields("SXLGD_MSRSDEN").Value = rs("MSRSDEN")          'SXLGD_測定結果 Den
        End If
        If rs("MSRSLDL") <> -1 Then
            .Fields("SXLGD_MSRSLDL").Value = rs("MSRSLDL")          'SXLGD_測定結果 L/DL
        End If
        If rs("MSRSDVD2") <> -1 Then
            .Fields("SXLGD_MSRSDVD2").Value = rs("MSRSDVD2")        'SXLGD_測定結果 DVD2
        End If
        
        For i = 1 To 15
            If rs("MS" & Format(i, "00") & "DEN1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN1").Value = rs("MS" & Format(i, "00") & "DEN1")      'SXLGD_測定値xx Den1
            End If
            If rs("MS" & Format(i, "00") & "DEN2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN2").Value = rs("MS" & Format(i, "00") & "DEN2")      'SXLGD_測定値xx Den2
            End If
            If rs("MS" & Format(i, "00") & "DEN3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN3").Value = rs("MS" & Format(i, "00") & "DEN3")      'SXLGD_測定値xx Den3
            End If
            If rs("MS" & Format(i, "00") & "DEN4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN4").Value = rs("MS" & Format(i, "00") & "DEN4")      'SXLGD_測定値xx Den4
            End If
            If rs("MS" & Format(i, "00") & "DEN5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "DEN5").Value = rs("MS" & Format(i, "00") & "DEN5")      'SXLGD_測定値xx Den5
            End If
            If rs("MS" & Format(i, "00") & "LDL1") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL1").Value = rs("MS" & Format(i, "00") & "LDL1")      'SXLGD_測定値xx L/DL1
            End If
            If rs("MS" & Format(i, "00") & "LDL2") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL2").Value = rs("MS" & Format(i, "00") & "LDL2")      'SXLGD_測定値xx L/DL2
            End If
            If rs("MS" & Format(i, "00") & "LDL3") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL3").Value = rs("MS" & Format(i, "00") & "LDL3")      'SXLGD_測定値xx L/DL3
            End If
            If rs("MS" & Format(i, "00") & "LDL4") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL4").Value = rs("MS" & Format(i, "00") & "LDL4")      'SXLGD_測定値xx L/DL4
            End If
            If rs("MS" & Format(i, "00") & "LDL5") <> -1 Then
                .Fields("WFGD_MS" & Format(i, "00") & "LDL5").Value = rs("MS" & Format(i, "00") & "LDL5")      'SXLGD_測定値xx L/DL5
            End If
        Next
        
        For i = 1 To 5
            If rs("MS0" & i & "DVD2") <> -1 Then
                .Fields("WFGD_MS01DVD2" & i).Value = rs("MS0" & i & "DVD2")         'SXLGD_測定値xx DVD2
            End If
        Next
        
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech Start
        'GDパターン判定結果
        If IsNull(rs("PTNJUDGRES")) = True Then
            .Fields("GD_PTNJUDGRES").Value = " "
        Else
            .Fields("GD_PTNJUDGRES").Value = rs("PTNJUDGRES")
        End If
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 ADD By Systech End

    End With
    Set rs = Nothing

    getTBCMJ006GD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ006GD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

''Upd start 2005/06/23 (TCS)t.terauchi  SPV9点対応
'概要      :WFSPV実績(TBCMJ016)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :HIN             , I  ,tFullHinban       , 品番(全品番構造体)
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'          :recX002         , O  ,c_cmzcrec         , TBCMX002構造体(測定点ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFSPV実績(TBCMJ016)からﾃﾞｰﾀを取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2005/06/23  新規作成　(TCS)t.terauchi
Private Function getTBCMJ016WFSPV(CRYNUM As String, recXSDCW As c_cmzcrec, HIN As tFullHinban, _
                                  recX001 As c_cmzcrec, recX002 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset   'TBCMJ016用
    Dim rs2         As OraDynaset   'TBCME028用
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV"
    
    getTBCMJ016WFSPV = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
    With recX001
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFSPV_NETSU").Value = ""           'WFSPV_熱処理条件
        .Fields("WFSPV_ET").Value = ""              'WFSPV_エッチング条件
        .Fields("WFSPV_MES").Value = ""             'WFSPV_計測方法
        .Fields("WFSPV_KST_MAX").Value = -1         'WFSPV_拡散長判定時のMAX値
        .Fields("WFSPV_KST_AVE").Value = -1         'WFSPV_拡散長判定時のAVE値
        .Fields("WFSPV_KST_MIN").Value = -1         'WFSPV_拡散長判定時のMIN値
        .Fields("WFSPV_FE_MAX").Value = -1          'WFSPV_Fe濃度判定時のMAX値
        .Fields("WFSPV_FE_AVE").Value = -1          'WFSPV_Fe濃度判定時のAVE値
        .Fields("WFSPV_FE_MIN").Value = -1          'WFSPV_Fe濃度判定時のMIN値
        
        ''>>=====SPV判定　20060529 SMP桜井
        .Fields("WFSPV_FE_PUA").Value = -1           ''SPV_Fe PUA値
        .Fields("WFSPV_FE_PUAP").Value = -1          ''SPV_Fe PUA％値
        .Fields("WFSPV_FE_STD").Value = -1           ''SPV_Fe STD
        .Fields("WFSPV_DIFF_PUA").Value = -1         ''SPV_拡散長 PUA値
        .Fields("WFSPV_DIFF_PUAP").Value = -1        ''SPV_拡散長 PUA％値
        .Fields("WFSPV_NR_MAX").Value = -1           ''SPV_OtherRecords_MAX
        .Fields("WFSPV_NR_AVE").Value = -1           ''SPV_OtherRecords_AVE
        .Fields("WFSPV_NR_STD").Value = -1           ''SPV_OtherRecords_STD
        .Fields("WFSPV_NR_PUA").Value = -1           ''SPV_OtherRecords_PUA値
        .Fields("WFSPV_NR_PUAP").Value = -1          ''SPV_OtherRecords_PUA％値
        ''==============================<<
    End With
                
    'TBCMX002
    With recX002
        .Fields("WFSPV_SMPPOS").Value = -1          'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
        .Fields("WFSPV_NETSU").Value = " "          'WFSPV_熱処理条件
        .Fields("WFSPV_ET").Value = " "             'WFSPV_エッチング条件
        .Fields("WFSPV_MES").Value = " "            'WFSPV_計測方法
        .Fields("WFSPV_DKAN").Value = " "           'WFSPV_ＤＫアニール条件
        .Fields("WFSPV_MESDATA1").Value = " "       'WFSPV測定点1
        .Fields("WFSPV_MESDATA2").Value = " "       'WFSPV測定点2
        .Fields("WFSPV_MESDATA3").Value = " "       'WFSPV測定点3
        .Fields("WFSPV_MESDATA4").Value = " "       'WFSPV測定点4
        .Fields("WFSPV_MESDATA5").Value = " "       'WFSPV測定点5
        .Fields("WFSPV_MESDATA6").Value = " "       'WFSPV測定点6
        .Fields("WFSPV_MESDATA7").Value = " "       'WFSPV測定点7
        .Fields("WFSPV_MESDATA8").Value = " "       'WFSPV測定点8
        .Fields("WFSPV_MESDATA9").Value = " "       'WFSPV測定点9
        .Fields("WFSPV_MESDATA10").Value = " "      'WFSPV測定点10
        .Fields("WFSPV_MESDATA11").Value = " "      'WFSPV測定点11
        .Fields("WFSPV_MESDATA12").Value = " "      'WFSPV測定点12
        .Fields("WFSPV_MESDATA13").Value = " "      'WFSPV測定点13
        .Fields("WFSPV_MESDATA14").Value = " "      'WFSPV測定点14
        .Fields("WFSPV_MESDATA15").Value = " "      'WFSPV測定点15
            
'        .Fields("WFSPV_SMPPOS2").Value = -1         'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
'        .Fields("WFSPV_NETSU2").Value = " "         'WFSPV_熱処理条件
'        .Fields("WFSPV_ET2").Value = " "            'WFSPV_エッチング条件
'        .Fields("WFSPV_MES2").Value = " "           'WFSPV_計測方法
'        .Fields("WFSPV_DKAN2").Value = " "          'WFSPV_ＤＫアニール条件
'
'        .Fields("WFSPV_FE_MAX").Value = " "         'WFSPV_Fe_MAX
'        .Fields("WFSPV_FE_AVE").Value = " "         'WFSPV_Fe_AVE
'        .Fields("WFSPV_FE_MIN").Value = " "         'WFSPV_Fe_MIN
'        .Fields("WFSPV_F_MESDATA1").Value = " "     'WFSPV測定点1   SPV_Fe
'        .Fields("WFSPV_F_MESDATA2").Value = " "     'WFSPV測定点2   SPV_Fe
'        .Fields("WFSPV_F_MESDATA3").Value = " "     'WFSPV測定点3   SPV_Fe
'        .Fields("WFSPV_F_MESDATA4").Value = " "     'WFSPV測定点4   SPV_Fe
'        .Fields("WFSPV_F_MESDATA5").Value = " "     'WFSPV測定点5   SPV_Fe
'        .Fields("WFSPV_F_MESDATA6").Value = " "     'WFSPV測定点6   SPV_Fe
'        .Fields("WFSPV_F_MESDATA7").Value = " "     'WFSPV測定点7   SPV_Fe
'        .Fields("WFSPV_F_MESDATA8").Value = " "     'WFSPV測定点8   SPV_Fe
'        .Fields("WFSPV_F_MESDATA9").Value = " "     'WFSPV測定点9   SPV_Fe
'
'        .Fields("WFSPV_DIFF_MAX").Value = " "       'WFSPV_拡散長_MAX
'        .Fields("WFSPV_DIFF_AVE").Value = " "       'WFSPV_拡散長_AVE
'        .Fields("WFSPV_DIFF_MIN").Value = " "       'WFSPV_拡散長_MIN
'        .Fields("WFSPV_D_MESDATA1").Value = " "     'WFSPV測定点1   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA2").Value = " "     'WFSPV測定点2   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA3").Value = " "     'WFSPV測定点3   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA4").Value = " "     'WFSPV測定点4   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA5").Value = " "     'WFSPV測定点5   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA6").Value = " "     'WFSPV測定点6   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA7").Value = " "     'WFSPV測定点7   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA8").Value = " "     'WFSPV測定点8   SPV_拡散長
'        .Fields("WFSPV_D_MESDATA9").Value = " "     'WFSPV測定点9   SPV_拡散長
    
''        ''>>=====SPV判定　20060529 SMP桜井
''        .Fields("WFSPV_FE_PUA").Value = -1           ''SPV_Fe PUA値
''        .Fields("WFSPV_FE_PUAP").Value = -1          ''SPV_Fe PUA％値
''        .Fields("WFSPV_FE_STD").Value = -1           ''SPV_Fe STD
''        .Fields("WFSPV_DIFF_PUA").Value = -1         ''SPV_拡散長 PUA値
''        .Fields("WFSPV_DIFF_PUAP").Value = -1        ''SPV_拡散長 PUA％値
''        .Fields("WFSPV_NR_MAX").Value = -1           ''SPV_OtherRecords_MAX
''        .Fields("WFSPV_NR_AVE").Value = -1           ''SPV_OtherRecords_AVE
''        .Fields("WFSPV_NR_STD").Value = -1           ''SPV_OtherRecords_STD
''        .Fields("WFSPV_NR_PUA").Value = -1           ''SPV_OtherRecords_PUA値
''        .Fields("WFSPV_NR_PUAP").Value = -1          ''SPV_OtherRecords_PUA％値
''        ''==============================<<
    End With
    
    If (recXSDCW("WFINDSPCW").Value <> "0") And (recXSDCW("WFRESSPCW").Value <> "0") Then
        
    '-------------------- TBCMJ016の読み込み(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " select *"
        sql = sql & " from   tbcmj016 " & vbLf
        sql = sql & " where  crynum = '" & CRYNUM & "'" & vbLf
        sql = sql & " and    smplno = '" & recXSDCW("WFSMPLIDSPCW").Value & "'" & vbLf
        sql = sql & " and    hsflg = '1'" & vbLf
        sql = sql & " and    trancnt = ( select   max(trancnt) from tbcmj016 " & vbLf
        sql = sql & "                    where    crynum = '" & CRYNUM & "'" & vbLf
        sql = sql & "                    and      smplno = '" & recXSDCW("WFSMPLIDSPCW").Value & "'" & vbLf
        sql = sql & "                    and      hsflg = '1'" & vbLf
        sql = sql & "                   )" & vbLf
                
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            GoTo proc_exit
        End If
                
    '-------------------- TBCME028の読み込み(SPV仕様) ----------------------------------------
        sql = ""
        sql = sql & " select HWFSPVSH,HWFSPVST,HWFSPVSI,HWFDLSPH,HWFDLSPT,HWFDLSPI"
        sql = sql & " from   TBCME028"
        sql = sql & " where  HINBAN = '" & HIN.hinban & "'"
        sql = sql & " and    MNOREVNO = " & HIN.mnorevno
        sql = sql & " and    FACTORY = '" & HIN.factory & "'"
        sql = sql & " and    OPECOND = '" & HIN.opecond & "'"
        
        Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs2.RecordCount = 0 Then
            GoTo proc_exit
        End If
                
        'TBCMX001
        With recX001
            If Not IsNull(recXSDCW("INPOSCW").Value) Then .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value     'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            If Not IsNull(rs("NETSU")) Then .Fields("WFSPV_NETSU").Value = rs("NETSU")                                  'WFSPV_熱処理条件
            If Not IsNull(rs("ET")) Then .Fields("WFSPV_ET").Value = rs("ET")                                           'WFSPV_エッチング条件
            If Not IsNull(rs("MES")) Then .Fields("WFSPV_MES").Value = rs("MES")                                        'WFSPV_計測方法
            
            'MAP測定の場合
            If rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "AMX" Then
                If Not IsNull(rs("SPV_FE_MAX")) Then .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                   'WFSPV_Fe濃度判定時のMAX値
                If Not IsNull(rs("SPV_FE_AVE")) Then .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                   'WFSPV_Fe濃度判定時のAVE値
                If Not IsNull(rs("SPV_FE_MIN")) Then .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                   'WFSPV_Fe濃度判定時のMIN値
                
            '9点測定の場合
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "V9T" Then
            
                ''Fe濃度のMAX,MIN,AVEを取得
                If getTBCMJ016WFSPV_Fe(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            'いずれでもない場合は拡散長の測定方法で判断
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "AMX" Then
                If Not IsNull(rs("SPV_FE_MAX")) Then .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                   'WFSPV_Fe濃度判定時のMAX値
                If Not IsNull(rs("SPV_FE_AVE")) Then .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                   'WFSPV_Fe濃度判定時のAVE値
                If Not IsNull(rs("SPV_FE_MIN")) Then .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                   'WFSPV_Fe濃度判定時のMIN値
            
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "V9T" And _
                    IsNull(rs("MS01_SPV_FE")) = False Then
                
                ''Fe濃度のMAX,MIN,AVEを取得
                If getTBCMJ016WFSPV_Fe(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            End If
            
            'MAP測定の場合
            If rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "AMX" Then
                If Not IsNull(rs("SPV_DIFF_MAX")) Then .Fields("WFSPV_KST_MAX").Value = rs("SPV_DIFF_MAX")              'WFSPV_拡散長判定時のMAX値
                If Not IsNull(rs("SPV_DIFF_AVE")) Then .Fields("WFSPV_KST_AVE").Value = rs("SPV_DIFF_AVE")              'WFSPV_拡散長判定時のAVE値
                If Not IsNull(rs("SPV_DIFF_MIN")) Then .Fields("WFSPV_KST_MIN").Value = rs("SPV_DIFF_MIN")              'WFSPV_拡散長判定時のMIN値
            
            '9点測定の場合
            ElseIf rs2.Fields("HWFDLSPH") & rs2.Fields("HWFDLSPT") & rs2.Fields("HWFDLSPI") = "V9T" Then
                
                ''拡散長のMAX,MIN,AVEを取得
                If getTBCMJ016WFSPV_Diff(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
            
            'いずれでもない場合はFe濃度の測定方法で判断
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "AMX" Then
                If Not IsNull(rs("SPV_DIFF_MAX")) Then .Fields("WFSPV_KST_MAX").Value = rs("SPV_DIFF_MAX")              'WFSPV_拡散長判定時のMAX値
                If Not IsNull(rs("SPV_DIFF_AVE")) Then .Fields("WFSPV_KST_AVE").Value = rs("SPV_DIFF_AVE")              'WFSPV_拡散長判定時のAVE値
                If Not IsNull(rs("SPV_DIFF_MIN")) Then .Fields("WFSPV_KST_MIN").Value = rs("SPV_DIFF_MIN")              'WFSPV_拡散長判定時のMIN値
            
            ElseIf rs2.Fields("HWFSPVSH") & rs2.Fields("HWFSPVST") & rs2.Fields("HWFSPVSI") = "V9T" And _
                    IsNull(rs("MS01_SPV_DIFF")) = False Then
                
                ''拡散長のMAX,MIN,AVEを取得
                If getTBCMJ016WFSPV_Diff(CRYNUM, rs("SMPLNO"), rs("TRANCNT"), _
                                         recX001) = FUNCTION_RETURN_FAILURE Then
                    GoTo proc_exit
                End If
                
            End If
            ''>>>===SPV判定　20060529 SMP桜井 ==
            If Not IsNull(rs("SPV_Fe_PUA")) Then .Fields("WFSPV_FE_PUA").Value = rs("SPV_Fe_PUA")            ''SPV_Fe PUA値
            If Not IsNull(rs("SPV_Fe_PUAP")) Then .Fields("WFSPV_FE_PUAP").Value = rs("SPV_Fe_PUAP")         ''SPV_Fe PUA％値
            If Not IsNull(rs("SPV_Fe_STD")) Then .Fields("WFSPV_FE_STD").Value = rs("SPV_Fe_STD")          ''SPV_Fe STD
            If Not IsNull(rs("SPV_Diff_PUA")) Then .Fields("WFSPV_DIFF_PUA").Value = rs("SPV_Diff_PUA")      ''SPV_拡散長 PUA値
            If Not IsNull(rs("SPV_Diff_PUAP")) Then .Fields("WFSPV_DIFF_PUAP").Value = rs("SPV_Diff_PUAP")     ''SPV_拡散長 PUA％値
            If Not IsNull(rs("SPV_Nr_MAX")) Then .Fields("WFSPV_NR_MAX").Value = rs("SPV_Nr_MAX")           ''SPV_OtherRecords_MAX
            If Not IsNull(rs("SPV_Nr_AVE")) Then .Fields("WFSPV_NR_AVE").Value = rs("SPV_Nr_AVE")           ''SPV_OtherRecords_AVE
            If Not IsNull(rs("SPV_Nr_STD")) Then .Fields("WFSPV_NR_STD").Value = rs("SPV_Nr_STD")          ''SPV_OtherRecords_STD
            If Not IsNull(rs("SPV_Nr_PUA")) Then .Fields("WFSPV_NR_PUA").Value = rs("SPV_Nr_PUA")          ''SPV_OtherRecords_PUA値
            If Not IsNull(rs("SPV_Nr_PUAP")) Then .Fields("WFSPV_NR_PUAP").Value = rs("SPV_Nr_PUAP")         ''SPV_OtherRecords_PUA％値
            ''==================================<<
        End With
            
        'TBCMX002
        With recX002
            
            If Not IsNull(recXSDCW("INPOSCW").Value) Then .Fields("WFSPV_SMPPOS").Value = recXSDCW("INPOSCW").Value     'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            If Not IsNull(rs("NETSU")) Then .Fields("WFSPV_NETSU").Value = rs("NETSU")                                  'WFSPV_熱処理条件
            If Not IsNull(rs("ET")) Then .Fields("WFSPV_ET").Value = rs("ET")                                           'WFSPV_エッチング条件
            If Not IsNull(rs("MES")) Then .Fields("WFSPV_MES").Value = rs("MES")                                        'WFSPV_計測方法
            If Not IsNull(rs("DKAN")) Then .Fields("WFSPV_DKAN").Value = rs("DKAN")                                     'WFSPV_ＤＫアニール条件

            .Fields("WFSPV_SMPPOS2").Value = recXSDCW("INPOSCW").Value          'WFSPVｻﾝﾌﾟﾙ-ID測定位置(SXL位置情報)
            .Fields("WFSPV_NETSU2").Value = rs("NETSU")                         'WFSPV_熱処理条件
            .Fields("WFSPV_ET2").Value = rs("ET")                               'WFSPV_エッチング条件
            .Fields("WFSPV_MES2").Value = rs("MES")                             'WFSPV_計測方法
            .Fields("WFSPV_DKAN2").Value = rs("DKAN")                           'WFSPV_ＤＫアニール条件

            .Fields("WFSPV_FE_MAX").Value = rs("SPV_FE_MAX")                    'WFSPV_Fe_MAX
            .Fields("WFSPV_FE_AVE").Value = rs("SPV_FE_AVE")                    'WFSPV_Fe_AVE
            .Fields("WFSPV_FE_MIN").Value = rs("SPV_FE_MIN")                    'WFSPV_Fe_MIN
            .Fields("WFSPV_F_MESDATA1").Value = rs("MS01_SPV_FE")               'WFSPV測定点1   SPV_Fe
            .Fields("WFSPV_F_MESDATA2").Value = rs("MS02_SPV_FE")               'WFSPV測定点2   SPV_Fe
            .Fields("WFSPV_F_MESDATA3").Value = rs("MS03_SPV_FE")               'WFSPV測定点3   SPV_Fe
            .Fields("WFSPV_F_MESDATA4").Value = rs("MS04_SPV_FE")               'WFSPV測定点4   SPV_Fe
            .Fields("WFSPV_F_MESDATA5").Value = rs("MS05_SPV_FE")               'WFSPV測定点5   SPV_Fe
            .Fields("WFSPV_F_MESDATA6").Value = rs("MS06_SPV_FE")               'WFSPV測定点6   SPV_Fe
            .Fields("WFSPV_F_MESDATA7").Value = rs("MS07_SPV_FE")               'WFSPV測定点7   SPV_Fe
            .Fields("WFSPV_F_MESDATA8").Value = rs("MS08_SPV_FE")               'WFSPV測定点8   SPV_Fe
            .Fields("WFSPV_F_MESDATA9").Value = rs("MS09_SPV_FE")               'WFSPV測定点9   SPV_Fe
            .Fields("WFSPV_DIFF_MAX").Value = rs("SPV_DIFF_MAX")                'WFSPV_拡散長_MAX
            .Fields("WFSPV_DIFF_AVE").Value = rs("SPV_DIFF_AVE")                'WFSPV_拡散長_AVE
            .Fields("WFSPV_DIFF_MIN").Value = rs("SPV_DIFF_MIN")                'WFSPV_拡散長_MIN
            .Fields("WFSPV_D_MESDATA1").Value = rs("MS01_SPV_DIFF")             'WFSPV測定点1   SPV_拡散長
            .Fields("WFSPV_D_MESDATA2").Value = rs("MS02_SPV_DIFF")             'WFSPV測定点2   SPV_拡散長
            .Fields("WFSPV_D_MESDATA3").Value = rs("MS03_SPV_DIFF")             'WFSPV測定点3   SPV_拡散長
            .Fields("WFSPV_D_MESDATA4").Value = rs("MS04_SPV_DIFF")             'WFSPV測定点4   SPV_拡散長
            .Fields("WFSPV_D_MESDATA5").Value = rs("MS05_SPV_DIFF")             'WFSPV測定点5   SPV_拡散長
            .Fields("WFSPV_D_MESDATA6").Value = rs("MS06_SPV_DIFF")             'WFSPV測定点6   SPV_拡散長
            .Fields("WFSPV_D_MESDATA7").Value = rs("MS07_SPV_DIFF")             'WFSPV測定点7   SPV_拡散長
            .Fields("WFSPV_D_MESDATA8").Value = rs("MS08_SPV_DIFF")             'WFSPV測定点8   SPV_拡散長
            .Fields("WFSPV_D_MESDATA9").Value = rs("MS09_SPV_DIFF")             'WFSPV測定点9   SPV_拡散長
            
            
''            ''>>>===SPV判定　20060529 SMP桜井 ==
''            If Not IsNull(rs("SPV_Fe_PUA")) Then .Fields("WFSPV_FE_PUA").Value = rs("SPV_Fe_PUA")            ''SPV_Fe PUA値
''            If Not IsNull(rs("SPV_Fe_PUAP")) Then .Fields("WFSPV_FE_PUAP").Value = rs("SPV_Fe_PUAP")         ''SPV_Fe PUA％値
''            If Not IsNull(rs("SPV_Fe_STD")) Then .Fields("WFSPV_FE_STD").Value = rs("SPV_Fe_STD")          ''SPV_Fe STD
''            If Not IsNull(rs("SPV_Diff_PUA")) Then .Fields("WFSPV_DIFF_PUA").Value = rs("SPV_Diff_PUA")      ''SPV_拡散長 PUA値
''            If Not IsNull(rs("SPV_Diff_PUAP")) Then .Fields("WFSPV_DIFF_PUAP").Value = rs("SPV_Diff_PUAP")     ''SPV_拡散長 PUA％値
''            If Not IsNull(rs("SPV_Nr_MAX")) Then .Fields("WFSPV_NR_MAX").Value = rs("SPV_Nr_MAX")           ''SPV_OtherRecords_MAX
''            If Not IsNull(rs("SPV_Nr_AVE")) Then .Fields("WFSPV_NR_AVE").Value = rs("SPV_Nr_AVE")           ''SPV_OtherRecords_AVE
''            If Not IsNull(rs("SPV_Nr_STD")) Then .Fields("WFSPV_NR_STD").Value = rs("SPV_Nr_STD")          ''SPV_OtherRecords_STD
''            If Not IsNull(rs("SPV_Nr_PUA")) Then .Fields("WFSPV_NR_PUA").Value = rs("SPV_Nr_PUA")          ''SPV_OtherRecords_PUA値
''            If Not IsNull(rs("SPV_Nr_PUAP")) Then .Fields("WFSPV_NR_PUAP").Value = rs("SPV_Nr_PUAP")         ''SPV_OtherRecords_PUA％値
''            ''==================================<<

        End With
        
        Set rs = Nothing
        Set rs2 = Nothing
    
    End If

    getTBCMJ016WFSPV = FUNCTION_RETURN_SUCCESS

proc_exit:
    
    Set rs = Nothing
    Set rs2 = Nothing
    
    '終了
    gErr.Pop
    Exit Function

proc_err:
    
    Set rs = Nothing
    Set rs2 = Nothing
    
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFSPV実績(TBCMJ016) Fe濃度のMAX/AVE/MINデータ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :sSmplID         , I  ,String            , ｻﾝﾌﾟﾙID
'          :iTranCnt        , I  ,Integer           , 処理回数
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFSPV実績(TBCMJ016)からFe濃度のMAX/AVE/MINﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2005/06/23  新規作成　(TCS)t.terauchi
Private Function getTBCMJ016WFSPV_Fe(CRYNUM As String, sSmplID As String, iTrancnt As Integer, _
                                    recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV_Fe"
    
    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_FAILURE
                    
    '-------------------- TBCMJ016の読み込み(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " SELECT  MAX(SPV_FE) AS MAX_FE,MIN(SPV_FE) AS MIN_FE,AVG(SPV_FE) AS AVE_FE" & vbLf
        sql = sql & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_FE AS SPV_FE" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "        )" & vbLf
                
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_MAX").Value = rs("MAX_FE")   'WFSPV_Fe濃度判定時のMAX値
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_AVE").Value = rs("AVE_FE")   'WFSPV_Fe濃度判定時のAVE値
            If Not IsNull(rs("MAX_FE")) Then .Fields("WFSPV_FE_MIN").Value = rs("MIN_FE")   'WFSPV_Fe濃度判定時のMIN値
        End With
        Set rs = Nothing

    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV_Fe = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :WFSPV実績(TBCMJ016) 拡散長のMAX/AVE/MINデータ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :sSmplID         , I  ,String            , ｻﾝﾌﾟﾙID
'          :iTranCnt        , I  ,Integer           , 処理回数
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :WFSPV実績(TBCMJ016)から拡散長のMAX/AVE/MINﾃﾞｰﾀを取得し、SXL検査書構造体にｾｯﾄする
'履歴      :2005/06/23  新規作成　(TCS)t.terauchi
Private Function getTBCMJ016WFSPV_Diff(CRYNUM As String, sSmplID As String, iTrancnt As Integer, _
                                    recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ016WFSPV_Diff"
    
    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_FAILURE
                    
    '-------------------- TBCMJ016の読み込み(WFSPV) ----------------------------------------
        sql = ""
        sql = sql & " SELECT  MAX(SPV_DIFF) AS MAX_DIFF,MIN(SPV_DIFF) AS MIN_DIFF,AVG(SPV_DIFF) AS AVE_DIFF" & vbLf
        sql = sql & " FROM   (SELECT  CRYNUM,SMPLNO,TRANCNT,ms01_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms02_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms03_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms04_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms05_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms06_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms07_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms08_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "             UNION ALL" & vbLf
        sql = sql & "         SELECT  CRYNUM,SMPLNO,TRANCNT,ms09_SPV_DIFF AS SPV_DIFF" & vbLf
        sql = sql & "         FROM    TBCMJ016" & vbLf
        sql = sql & "         WHERE   CRYNUM = '" & CRYNUM & "'" & vbLf
        sql = sql & "         AND     SMPLNO = '" & sSmplID & "'" & vbLf
        sql = sql & "         AND     TRANCNT = " & iTrancnt & vbLf
        sql = sql & "         AND     HSFLG = '1'" & vbLf
        sql = sql & "        )" & vbLf
               
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            GoTo proc_exit
        End If
        
        'TBCMX001
        With recX001
            If Not IsNull(rs("MAX_DIFF")) Then .Fields("WFSPV_KST_MAX").Value = rs("MAX_DIFF")  'WFSPV_拡散長判定時のMAX値
            If Not IsNull(rs("AVE_DIFF")) Then .Fields("WFSPV_KST_AVE").Value = rs("AVE_DIFF")  'WFSPV_拡散長判定時のAVE値
            If Not IsNull(rs("MIN_DIFF")) Then .Fields("WFSPV_KST_MIN").Value = rs("MIN_DIFF")  'WFSPV_拡散長判定時のMIN値
        End With
        Set rs = Nothing

    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ016WFSPV_Diff = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'概要      :標準測定(TBCMY018)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :sBlkId          , I  ,String            , ｻﾝﾌﾟﾙﾌﾞﾛｯｸID
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SXL検査書)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :標準測定(TBCMY018)からWarp実績を取得し、SXL検査書・測定点ﾃﾞｰﾀ構造体にｾｯﾄする
'履歴      :2005/06/23  新規作成  (TCS)T.terauchi
Private Function getTBCMY018WARP(sBlkId As String, recXSDCW As c_cmzcrec, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMY018WARP"
    
    getTBCMY018WARP = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    'TBCMX001
'    With recX001
'        .Fields("WARP_1").Value = vbNullString            '反りWarp-1
'        .Fields("WARP_2").Value = vbNullString            '反りWarp-2
'        .Fields("WARP_3").Value = vbNullString            '反りWarp-3
'    End With
                    
    '-------------------- TBCMY018の読み込み ----------------------------------------
    sql = ""
    sql = sql & " select max(to_number(measdata)) as warp_1" & vbLf
    sql = sql & "        ,avg(to_number(measdata)) as warp_2" & vbLf
    sql = sql & "        ,min(to_number(measdata)) as warp_3" & vbLf
    sql = sql & " from   tbcmy018" & vbLf
    sql = sql & " where  sublotid = '" & sBlkId & "'" & vbLf
    sql = sql & " and    measitem = 'MSL04WARPU'" & vbLf
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCMY018WARP = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    'TBCMX001
    With recX001
        .Fields("WARP_1").Value = rs("warp_1")          '反りWarp-1(最大値)
        .Fields("WARP_2").Value = rs("warp_2")          '反りWarp-2(平均)
        .Fields("WARP_3").Value = rs("warp_3")          '反りWarp-3(最小値)
    End With
        
    Set rs = Nothing
    
    getTBCMY018WARP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMY018WARP = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
''Upd end   2005/06/23 (TCS)T.terauchi      SPV9点対応

'概要      :SXLIDをｷｰにXSDCAからﾌﾞﾛｯｸIDを取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       :説明
'          :sSXLID        ,I  ,String                   :SXLID
'          :戻り値        ,O  ,FUNCTION_RETURN          :抽出の成否
'説明      :
'履歴      :06/01/20 ooba
Public Function GetCaBlockID(sSXLID As String) As FUNCTION_RETURN

    Dim i, m        As Integer
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmec060_1.frm -- Function GetCaBlockID"
    
    
    sql = "select CRYNUMCA from XSDCA "
    sql = sql & "where SXLIDCA = '" & sSXLID & "' "
    sql = sql & "group by CRYNUMCA "
    sql = sql & "order by CRYNUMCA"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount < 1 Then
        GetCaBlockID = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    m = rs.RecordCount
    ReDim sReBlkID(m)
    ''抽出結果を格納する
    For i = 1 To m
        sReBlkID(i) = rs("CRYNUMCA")        'ﾌﾞﾛｯｸID
        rs.MoveNext
    Next i
        
    Set rs = Nothing
    
    GetCaBlockID = FUNCTION_RETURN_SUCCESS
  

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GetCaBlockID = FUNCTION_RETURN_FAILURE
    Resume proc_exit
    
End Function

'■■EDI情報ﾘﾝｸ対応 2009/12/4 Add Start SPK habuki　↓↓↓
'---------------------------------------------------------------------------
'概要      :SXLID上７桁より、XODC6_1を検索し、送信対象のEDI情報を返す
'---------------------------------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :pSXLID      ,I  　　,String                 ,SXLID
'          :pPN         ,O  　　,Double                 ,不純物濃度(P:ﾘﾝ)
'          :pBN         ,O  　　,Double                 ,不純物濃度(B:ﾎﾞﾛﾝ)
'          :pASN        ,O  　　,Double                 ,不純物濃度(AS:砒素)
'          :pCN         ,O  　　,Double                 ,不純物濃度(C:炭素)
'          :pflgEDI     ,O  　　,Boolean                ,EDI情報有無判定用(True:有、False：無)
'          :戻り値      ,O      ,Boolean                ,[True:OK／False:NG]
'---------------------------------------------------------------------------
Public Function fncGetEdiInfo(ByVal pSXLID As String, _
                              ByRef pPN As Double, _
                              ByRef pBN As Double, _
                              ByRef pASN As Double, _
                              ByRef pCN As Double, _
                              ByRef pflgEDI As Boolean _
                             ) As Boolean
    Dim i, m        As Integer
    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet
    
    '--ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function fncGetEdiInfo"
    
    '--初期化
    pPN = 0: pBN = 0: pASN = 0: pCN = 0
    fncGetEdiInfo = False: pflgEDI = False
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL文生成
    '<投入重量優先版>
''    sql = "select" & vbCrLf
''    sql = sql & "    c61.PNC6         PN" & vbCrLf                      '不純物濃度(P:ﾘﾝ)
''    sql = sql & "   ,c61.BNC6         BN" & vbCrLf                      '不純物濃度(B:ﾎﾞﾛﾝ)
''    sql = sql & "   ,c61.ASNC6        ASN" & vbCrLf                     '不純物濃度(AS:砒素)
''    sql = sql & "   ,c61.CNC6         CN" & vbCrLf                      '不純物濃度(C:炭素)
''    sql = sql & "from" & vbCrLf
''    sql = sql & "   (" & vbCrLf
''    sql = sql & "     select" & vbCrLf
''    sql = sql & "         max(MATERYOC6)    MATERYOC6" & vbCrLf         '重量／枚数
''    sql = sql & "     from" & vbCrLf
''    sql = sql & "         XODC6_1" & vbCrLf                             '副材料
''    sql = sql & "     where" & vbCrLf
''    sql = sql & "          substr(XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf    '引上結晶番号=SXLID(7桁)
''    sql = sql & "      and MATEKC6  = '1'" & vbCrLf                     '区分(ﾎﾟﾘｼﾘｺﾝ)
''    sql = sql & "      and EDIFLGC6 = '2'" & vbCrLf                     'EDIﾌﾗｸﾞ(送信対象)
''    sql = sql & "   ) tkey" & vbCrLf
''    sql = sql & "   ,XODC6_1  c61" & vbCrLf                             '副材料
''    sql = sql & "   ,TBCMH001 t01" & vbCrLf                             '引上指示実績
''    sql = sql & "where" & vbCrLf
''    sql = sql & "     substr(c61.XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf      '引上結晶番号=SXLID(7桁)
''    sql = sql & " and c61.MATEKC6   = '1'" & vbCrLf                     '区分(ﾎﾟﾘｼﾘｺﾝ)
''    sql = sql & " and c61.EDIFLGC6  = '2'" & vbCrLf                     'EDIﾌﾗｸﾞ(送信対象)
''    sql = sql & " and c61.MATERYOC6 = tkey.MATERYOC6" & vbCrLf          '重量／枚数
''    sql = sql & " and substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)" & vbCrLf          '引上結晶番号7桁＋9桁目
''    sql = sql & " and t01.CODE > '4'" & vbCrLf                          '仕掛ｺｰﾄﾞ
''    sql = sql & " and rownum = 1" & vbCrLf                              '1ﾚｺｰﾄﾞ目
''    sql = sql & "order by" & vbCrLf
''    sql = sql & "    substr(c61.XTALC6,9,1)" & vbCrLf                   '引上結晶番号9桁目(0:通常、1～4：ﾘﾁｬｰｼﾞ、A～C：AB取り)
''    sql = sql & "   ,c61.MATESYUC6" & vbCrLf                            '種類(原料ｺｰﾄﾞ)

    '<結晶番号優先版>
    sql = "select" & vbCrLf
    sql = sql & "    c61.PNC6         PN" & vbCrLf                      '不純物濃度(P:ﾘﾝ)
    sql = sql & "   ,c61.BNC6         BN" & vbCrLf                      '不純物濃度(B:ﾎﾞﾛﾝ)
    sql = sql & "   ,c61.ASNC6        ASN" & vbCrLf                     '不純物濃度(AS:砒素)
    sql = sql & "   ,c61.CNC6         CN" & vbCrLf                      '不純物濃度(C:炭素)
    sql = sql & "from" & vbCrLf
    sql = sql & "    XODC6_1  c61" & vbCrLf                             '副材料
    sql = sql & "   ,TBCMH001 t01" & vbCrLf                             '引上指示実績
    sql = sql & "where" & vbCrLf
    sql = sql & "     substr(c61.XTALC6,1,7) = '" & left(pSXLID, 7) & "'" & vbCrLf      '引上結晶番号=SXLID(7桁)
    sql = sql & " and c61.MATEKC6   = '1'" & vbCrLf                     '区分(ﾎﾟﾘｼﾘｺﾝ)
    sql = sql & " and c61.EDIFLGC6  = '2'" & vbCrLf                     'EDIﾌﾗｸﾞ(送信対象)
    sql = sql & " and substr(c61.XTALC6,1,7)||substr(c61.XTALC6,9,1) = substr(t01.UPINDNO,1,7)||substr(t01.UPINDNO,9,1)" & vbCrLf          '引上結晶番号7桁＋9桁目
    sql = sql & " and t01.CODE > '4'" & vbCrLf                          '仕掛ｺｰﾄﾞ
    sql = sql & " and rownum = 1" & vbCrLf                              '1ﾚｺｰﾄﾞ目
    sql = sql & "order by" & vbCrLf
    sql = sql & "    substr(c61.XTALC6,9,1)" & vbCrLf                   '引上結晶番号9桁目(0:通常、1～4：ﾘﾁｬｰｼﾞ、A～C：AB取り)
    sql = sql & "   ,c61.MATESYUC6" & vbCrLf                            '種類(原料ｺｰﾄﾞ)

    '--ﾃﾞｰﾀを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--抽出結果参照
    If rs.EOF Then
        '<< ﾃﾞｰﾀ無し >>
        fncGetEdiInfo = True
        GoTo proc_exit
    Else
        '<< ﾃﾞｰﾀ有り >>
        rs.MoveFirst
        pPN = IIf(IsNull(rs("PN")), 0, rs("PN"))                '不純物濃度(P:ﾘﾝ)
        pBN = IIf(IsNull(rs("BN")), 0, rs("BN"))                '不純物濃度(B:ﾎﾞﾛﾝ)
        pASN = IIf(IsNull(rs("ASN")), 0, rs("ASN"))             '不純物濃度(AS:砒素)
        pCN = IIf(IsNull(rs("CN")), 0, rs("CN"))                '不純物濃度(C:炭素)
        pflgEDI = True                                          'EDI情報有無判定用(True:有、False：無)
    End If
    
    fncGetEdiInfo = True

proc_exit:
    '<< 終了 >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< ｴﾗｰﾊﾝﾄﾞﾗ >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit
    
End Function
'■■EDI情報ﾘﾝｸ対応 2009/12/4 Add End   SPK habuki　↑↑↑
    
'概要      :WFSIRD実績(TBCMJ022)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCW        , I  ,c_cmzcrec         , XSDCW構造体   (新ｻﾝﾌﾟﾙ管理(SXL))
'          :recX001         , O  ,c_cmzcrec         , TBCMX001構造体(SIRD評価実績ﾃﾞｰﾀ)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :SIRD評価実績(TBCMJ022)からﾃﾞｰﾀを取得し、構造体にｾｯﾄする
'履歴      :2010/04/19 Y.Hitomi
Private Function getTBCMJ022SIRD(CRYNUM As String, recXSDCW As c_cmzcrec, recX001 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ022SIRD"
    
    getTBCMJ022SIRD = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
        
    'TBCMX001
    With recX001
            .Fields("SIRD_POS").Value = vbNullString           'FSIRDサンプル測定位置(SXL位置情報)
            .Fields("SIRD_TOTAL").Value = vbNullString         'WFSIRD_判定時のTOTAL値
    End With
        
    '-------------------- TBCMJ022の読み込み(SIRD) ----------------------------------------
    sql = "select * from TBCMJ022 "
    sql = sql & " where CRYNUM = '" & CRYNUM & "'"
'DEL 2010/05/20 Y.Hitomi
'    sql = sql & " and   SMPLNO = '" & recXSDCW("WFSMPLIDL4CW").Value & "'"
    sql = sql & " and   TRANCNT = 0"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCMJ022SIRD = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    'TBCMX001
    With recX001
        .Fields("SIRD_POS").Value = rs("POSITION")              'WFSIRDサンプル測定位置(SXL位置情報)
        .Fields("SIRD_TOTAL").Value = rs("SIRDCNT")             'WFSIRD_判定時のTOTAL値
        
    End With
    Set rs = Nothing

    getTBCMJ022SIRD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ022SIRD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'=================================================================================
' 2011/01/17 tkimura ADD START
'概要      :ρTop位置,ρBot位置を取得する。
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO  ,型                :説明
'          :SXLID           ,I   ,String            ,SXLID OR ブロックID
'　　      :sPos  　　　    ,I   ,String 　         ,SXL位置(TOP/BOT)
'          :sPattern        ,I   ,String            ,比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
'                                                   ●ﾊﾟﾀｰﾝA : WF実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝB : 結晶実績ﾃﾞｰﾀ取得
'                                                   ●ﾊﾟﾀｰﾝC : 取得ﾃﾞｰﾀなし
'          :sRsPos()        ,O   ,String            ,比抵抗位置(TOP/BOT)
'          :戻り値          ,O   ,FUNCTION_RETURN   ,抽出の成否
'説明      :
'履歴      :2011/01/17 tkimura 作成
''Public Function cmbc040_GetSxlRsPos(ByVal data As String, _
''                                    ByVal sPos As String, _
''                                    ByVal sSmpId As String, _
''                                    ByVal sPattern As String, _
''                                    ByRef sRsPos() As String) As FUNCTION_RETURN
Public Function cmbc040_GetSxlRsPos(ByVal data As String, _
                                    ByVal sPos As String, _
                                    ByVal sPattern As String, _
                                    ByRef sRsPos() As String) As FUNCTION_RETURN
    
    Dim sTBkbn As String        'T/B区分
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer           'Topのとき1,Botのとき2を代入する。
    Dim sSql As String
    Dim rs As OraDynaset
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function cmbc040_GetSxlRsPos"
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_FAILURE
    
    If sPos = "TOP" Then sTBkbn = "T" Else sTBkbn = "B"  '04/04/15 ooba
    
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『A』の場合、XSDCW[新サンプル管理(SXL)]より位置を取得する。
    If sPattern = "A" Then
        sSql = ""
        sSql = sSql & "SELECT" & vbCrLf
        sSql = sSql & " INPOSCW " & vbCrLf                      '結晶内位置
        sSql = sSql & "FROM" & vbCrLf
        sSql = sSql & " XSDCW " & vbCrLf
        sSql = sSql & "WHERE" & vbCrLf
        sSql = sSql & " SXLIDCW = '" & data & "' AND" & vbCrLf  'SXLID
        sSql = sSql & " TBKBNCW = '" & sTBkbn & "'" & vbCrLf    'TB区分
        ''sSQL = sSQL & " REPSMPLIDCW = '" & sSmpId & "'"
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount > 0 Then
            'TOP側位置
            If sTBkbn = "T" Then
                k = 1
            'BOT側位置
            ElseIf sTBkbn = "B" Then
                k = 2
            End If
            sRsPos(k) = rs("INPOSCW")       '結晶内位置(TOPまたはBOT)
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『B』の場合、XSDCS[新サンプル管理(ブロック)]より位置を取得する。
    'サンプルIDがないときにこのパターンがくるのでsql文にサンプルIDを条件に入れない。
    ElseIf sPattern = "B" Then
        sSql = ""
        sSql = sSql & "SELECT" & vbCrLf
        sSql = sSql & " INPOSCS" & vbCrLf                        '結晶内位置
        sSql = sSql & "FROM" & vbCrLf
        sSql = sSql & " XSDCS" & vbCrLf
        sSql = sSql & "WHERE" & vbCrLf
        sSql = sSql & " CRYNUMCS = '" & data & "' AND" & vbCrLf  'ブロックID
        sSql = sSql & " TBKBNCS = '" & sTBkbn & "'" & vbCrLf     'TB区分
        'sSQL = sSQL & " REPSMPLIDCS = '" & sSmpId & "'"
        
        Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_NO_BLANKSTRIP)
    
        If rs.RecordCount > 0 Then
            'TOP側位置
            If sTBkbn = "T" Then
                k = 1
            'BOT側位置
            ElseIf sTBkbn = "B" Then
                k = 2
            End If
            sRsPos(k) = rs("INPOSCS")       '結晶内位置(TOPまたはBOT)
        Else
            '実績ﾃﾞｰﾀがない場合はｴﾗｰ
            Set rs = Nothing
            GoTo proc_exit
        End If
        Set rs = Nothing
    '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝが『C』の場合、取得実績ﾃﾞｰﾀなし。
    ElseIf sPattern = "C" Then
    
    End If
    
'''    '取得ﾃﾞｰﾀが空白/-1/NULLの時はｽﾍﾟｰｽをｾｯﾄする。
    If sRsPos(k) = "" Or sRsPos(k) = "-1" Or sRsPos(k) = vbNullString Then
        sRsPos(k) = " "
    End If
        
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    cmbc040_GetSxlRsPos = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要      :ウェハ－センタ－入庫情報(TBCMY011)の送信フラグを更新する。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :SXLID           , I  ,String            , シングルID
'          :hin             , I  ,tFullHinban       , 品番(全品番構造体)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :TBCME018.HSXRKHNMの値を調べ、その値が0の場合はTBCMEY011.QA1SNDFLGの値を3とする。それ以外のときは2とする。
'履歴      :2011/01/17 tkimura
Private Function UpdateTBCMY011SendFlag(ByVal SXLID As String, _
                                        ByRef HIN As tFullHinban) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim hindo   As String       '検査頻度
    Dim sndFlg  As String       '送信フラグ
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function UpdateTBCMY011SendFlag"
    
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE

'    Del Start 2011/07/05 Y.Hitomi
'    '製品仕様SXLデータ１から検査頻度を取得する。
'    Set rs = Nothing
'    sql = ""
'    sql = sql & "SELECT" & vbCrLf
'    sql = sql & " HSXRKHNM" & vbCrLf        '検査頻度
'    sql = sql & "FROM" & vbCrLf
'    sql = sql & " TBCME018" & vbCrLf
'    sql = sql & "WHERE" & vbCrLf
'    sql = sql & " HINBAN ='" & HIN.hinban & "' AND" & vbCrLf
'    sql = sql & " MNOREVNO =" & HIN.mnorevno & " AND" & vbCrLf
'    sql = sql & " FACTORY ='" & HIN.factory & "' AND" & vbCrLf
'    sql = sql & " OPECOND ='" & HIN.opecond & "'" & vbCrLf
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    If rs.RecordCount = 0 Then
'        Set rs = Nothing
'        UpdateTBCMY011SendFlag = FUNCTION_RETURN_SUCCESS
'        GoTo proc_exit
'    End If
'
'    hindo = rs("HSXRKHNM")
'
'    '検査頻度が「0」の場合、品質システム送信フラグ（G52)='3'[送信予約]
'    'それ以外では品質システム送信フラグ（G52)='2'[送信済み]とする。
'    If hindo = "0" Then sndFlg = "3" Else sndFlg = "2"
'    Del End 2011/07/05 Y.Hitomi
    
'Add Start 2011/07/05 Y.Hitomi 全数送信対応
    sndFlg = "3"
'Add End   2011/07/05 Y.Hitomi
    
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMY011" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " TBCMY011.QA1SNDFLG='" & sndFlg & "'," & vbCrLf                         '品質システム送信フラグ（G52)
    sql = sql & " UPDPROC='CW800'," & vbCrLf                                             '更新工程　2011/01/31 tkimura
    sql = sql & " UPDDATE=to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbCrLf '更新日付　2011/01/31 tkimura
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " MSXLID='" & SXLID & "'" & vbCrLf     'SXLID
        
    If 0 >= OraDB.ExecuteSQL(sql) Then
        UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    Set rs = Nothing
    
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    UpdateTBCMY011SendFlag = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要      :ウェハ－センタ－入庫情報(TBCMY011)の
'           インゴット位置引上率,枚葉推定抵抗値(Center)を更新する。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :blockId         , I  ,String            , ブロックID
'          :blockSeq        , I  ,String            , ブロック内連番
'          :up_Ratio        , I  ,String            , インゴット位置引上率
'          :rs_Meas         , I  ,String            , 枚葉推定抵抗値(Center)
'          :dSXL_Pos        , I  ,Double            , SXL位置(Intel)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :
'履歴      :2011/01/17 tkimura
'       　　2011/04/25 Marushita 引数SXL位置追加（Micronｲﾝｺﾞｯﾄ位置管理追加対応）
'Private Function UpdateTBCMY011SuiteiResData(ByVal BLOCKID As String, _
'                                             ByVal BLOCKSEQ As Integer, _
'                                             ByVal up_Ratio As String, _
'                                             ByVal rs_Meas As String) As FUNCTION_RETURN
        
Private Function UpdateTBCMY011SuiteiResData(ByVal BLOCKID As String, _
                                             ByVal BLOCKSEQ As Integer, _
                                             ByVal up_Ratio As String, _
                                             ByVal rs_Meas As String, _
                                             ByVal dSXL_Pos As Double) As FUNCTION_RETURN
    Dim sql     As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function UpdateTBCMY011SuiteiResData"
    
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
        
    sql = ""
    sql = sql & "UPDATE" & vbCrLf
    sql = sql & " TBCMY011" & vbCrLf
    sql = sql & "SET" & vbCrLf
    sql = sql & " UP_RATIO='" & up_Ratio & "'," & vbCrLf    'インゴット位置引上率
    sql = sql & " RS_MEAS='" & rs_Meas & "'" & vbCrLf       '枚葉推定抵抗値(Center)
    sql = sql & ",HTOP_POS='" & dSXL_Pos & "'" & vbCrLf     '補正結晶長(SXL位置(Intel)) 2011/04/25 ADD Marushita
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " LOTID='" & BLOCKID & "' AND" & vbCrLf     'ブロックID
    sql = sql & " BLOCKSEQ=" & BLOCKSEQ & "" & vbCrLf       'ブロック内連番
    
    If 0 >= OraDB.ExecuteSQL(sql) Then
        UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
''    Debug.Print "BLOCK ", BLOCKID, " SEQ ", BLOCKSEQ
''    Debug.Print "UP_RATIO ", up_Ratio, " RS_MEAS ", rs_Meas, " HTOP_POS ", dSXL_Pos
            
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    UpdateTBCMY011SuiteiResData = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/18 tkimura ADD START
'概要      :ρTop位置引上率,ρBot位置引上率,実効偏析,基準抵抗値をもとめる。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型                       , 説明
'          :CRYNUM          , I  ,String                    , 結晶番号
'          :sRsData         , I  ,String                    , 抵抗データ(Top/Bot)
'          :sRsPos          , I  ,String                    , 抵抗位置データ(Top/Bot)
'          :d               , O  ,type_Coefficient_new2     , 推定抵抗,推定引上率計算構造体
'      　　:戻り値          , O  ,FUNCTION_RETURN　         , 成否
'説明      :
'履歴      :2011/01/18 tkimura
Private Function GetStandardPosRes(ByVal CRYNUM As String, _
                                   ByRef sRsData() As String, _
                                   ByRef sRsPos() As String, _
                                   ByRef d As type_Coefficient_new2) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    
    Dim wgtCharge   As Long                 '偏析計算用パラメータ(チャージ量)
    Dim wgtChargeA  As Long                 '偏析計算用パラメータ(STARTチャージ量)
    Dim wgtTop      As Double               '偏析計算用パラメータ(Top取量)
    Dim wgtTopCut   As Double               '偏析計算用パラメータ(肩重量)
    Dim DM          As Double               '偏析計算用パラメータ(直径平均)
    Dim HIKIFLG     As Integer              '引上げフラグ(1=通常、2=BC結晶)
    
    '>>>>> 2011/04/25 ADD Marushita ☆Micronｲﾝｺﾞｯﾄ位置管理追加対応
    Dim p_CRYNUM    As String               '前結晶番号取得用
    Dim p_wgtTop    As Double               '前引上情報取得用(前Top取量)
    Dim p_DM        As Double               '前引上情報取得用(前直径平均)
    Dim p_wgtTA     As Double               '前引上情報取得用(前テイル重量)
    Dim p_LENTK     As Long                 '前引上情報取得用(前引上長)
    Dim dMaeBatLen  As Double               '前バッチ結晶長　Add 2011/09/27
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function GetStandardPosRes"
    
    GetStandardPosRes = FUNCTION_RETURN_FAILURE
        
    If GetCoeffParams_new2(CRYNUM, wgtCharge, wgtChargeA, wgtTop, _
                           wgtTopCut, DM, HIKIFLG) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
        
    d.DUNMENSEKI = AreaOfCircle(DM)     '断面積(直径平均より計算)
    d.TOPSMPLPOS = sRsPos(1)            'TOP位置
    d.BOTSMPLPOS = sRsPos(2)            'BOT位置
    d.CHARGEWEIGHT = wgtCharge          'チャージ量
    d.TOPWEIGHT = wgtTop + wgtTopCut    'トップ重量(Top取量＋肩重量)
    d.TOPRES = sRsData(1)               'TOP位置抵抗値
    d.BOTRES = sRsData(6)               'BOT位置抵抗値
    d.CHARGEWEIGHTA = wgtChargeA        'STARTチャージ量
    d.HIKIFLG = HIKIFLG                 '引上フラグ(1=通常、2=BC結晶)
    
    'ρTop位置引上げ率を求める。
    d.SMPLPOS = d.TOPSMPLPOS
    d.GT = HikiageCalculation(d)
    'ρBot位置引上げ率を求める。
    d.SMPLPOS = d.BOTSMPLPOS
    d.GB = HikiageCalculation(d)

    '実行偏析を求める。
    d.Henseki = CoefficientCalculation_new2(d)
    If d.Henseki = -9999 Then
        d.Henseki = 0       'SXLRS_実効偏析
    End If
    
    '基準抵抗値を求める。
    d.KIJUNTEIKOU = StandardResCalculation(d)
    If d.KIJUNTEIKOU = -9999 Then
        d.KIJUNTEIKOU = 0
    End If
    
    '>>>>> 2011/04/25 ADD Marushita ☆Micronｲﾝｺﾞｯﾄ位置管理追加対応
    d.HOSEICHO = 0
    '補正結晶長を求める
    If HIKIFLG = "1" Then       '通常引上の場合
        'Top取量/(断面積*0.00233)
        d.HOSEICHO = wgtTop / (d.DUNMENSEKI * HIJU_SILICONE)
    ElseIf HIKIFLG = "2" Then   'B、C結晶(残量引)の場合
        '>>>>> 2011/09/27 ADD Marushita C結晶以上対応
        '前バッチ結晶長を取得する
        dMaeBatLen = GetMaeBatLen(CRYNUM)
        '現在結晶の補正(Top取量/(断面積*0.00233))を足して補正長を求める
        d.HOSEICHO = dMaeBatLen + wgtTop / (d.DUNMENSEKI * HIJU_SILICONE)
'        '前引上の結晶番号を取得(結晶番号)
'        If GetPreCrynum(CRYNUM, p_CRYNUM) = FUNCTION_RETURN_FAILURE Then
'        Else
'            '前引上の結晶情報を取得(前Top取量、前直径平均、前引上長、前テイル重量)
'            If GetPreXSDC1(p_CRYNUM, p_wgtTop, p_DM, p_LENTK, p_wgtTA) = FUNCTION_RETURN_FAILURE Then
'            Else
'                '補正結晶長の計算(前バッチの結晶長を足す)
'                '補正結晶長 = (前Top取量/(前断面積*0.00233))+前引上長+((前テイル重量+Top取量)/(断面積*0.00233))
'                d.HOSEICHO = (p_wgtTop / (AreaOfCircle(p_DM) * HIJU_SILICONE)) + p_LENTK + _
'                       ((p_wgtTA + wgtTop) / (d.DUNMENSEKI * HIJU_SILICONE))
'            End If
'        End If
        '<<<<< 2011/09/27 ADD Marushita C結晶以上対応
    End If
    '<<<<< 2011/04/25 ADD Marushita ☆Micronｲﾝｺﾞｯﾄ位置管理追加対応
    
''    Debug.Print "Top引上げ率", d.GT
''    Debug.Print "Bot引上げ率", d.GB
''    Debug.Print "直径", DM
''    Debug.Print "断面積", d.DUNMENSEKI
''    Debug.Print "トップカット重量", wgtTopCut
''    Debug.Print "トップ重量", wgtTop
''    Debug.Print "推定チャージ量", d.CHARGEWEIGHT
''    Debug.Print "推定チャージ量A", d.CHARGEWEIGHTA
''    Debug.Print "トップ抵抗値", d.TOPRES
''    Debug.Print "ボトム抵抗値", d.BOTRES
''    Debug.Print "トップ位置", d.TOPSMPLPOS
''    Debug.Print "ボトム位置", d.BOTSMPLPOS
''    Debug.Print "実効偏析", d.Henseki
''    Debug.Print "基準抵抗値", d.KIJUNTEIKOU
''    Debug.Print "補正結晶長", d.HOSEICHO
    
    '2011/01/19 kimura デバッグ情報取得(後で消す)
''    Debug.Print "Top引上げ率"
''    Debug.Print d.GT
''    Debug.Print "Bot引上げ率"
''    Debug.Print d.GB
''    Debug.Print "直径"
''    Debug.Print DM
''    Debug.Print "断面積"
''    Debug.Print d.DUNMENSEKI
''    Debug.Print "トップカット重量"
''    Debug.Print wgtTopCut
''    Debug.Print "トップ重量"
''    Debug.Print wgtTop
''    Debug.Print "推定チャージ量(前バッチも含む)"
''    Debug.Print d.CHARGEWEIGHT
''    Debug.Print "推定チャージ量A"
''    Debug.Print d.CHARGEWEIGHTA
''    Debug.Print "トップ抵抗値"
''    Debug.Print d.TOPRES
''    Debug.Print "ボトム抵抗値"
''    Debug.Print d.BOTRES
''    Debug.Print "トップ位置"
''    Debug.Print d.TOPSMPLPOS
''    Debug.Print "ボトム位置"
''    Debug.Print d.BOTSMPLPOS
''    Debug.Print "実効偏析"
''    Debug.Print d.Henseki
''    Debug.Print "基準抵抗値"
''    Debug.Print d.KIJUNTEIKOU
    
    GetStandardPosRes = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    GetStandardPosRes = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/18 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/18 tkimura ADD START
'概要      :SXLIDの枚葉ごとに
'           インゴット位置引上率,枚葉推定抵抗値(Center)を計算する。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型                       , 説明
'          :SXLID           , I  ,String                    , SXLID
'          :d               , I  ,type_Coefficient_new2     , 推定抵抗,推定引上率計算構造体
'      　　:戻り値          , O  ,FUNCTION_RETURN　         , 成否
'説明      :
'履歴      :2011/01/18 tkimura
Private Function SuiteiResDataCalculation(ByVal SXLID As String, _
                                          ByRef d As type_Coefficient_new2, sUP_RATIO() As String) As FUNCTION_RETURN
'Private Function SuiteiResDataCalculation(ByVal SXLID As String, _
                                          ByRef d As type_Coefficient_new2) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim i       As Integer
    Dim recY011() As typ_Y011               '' Y011のデータをまとめた構造体
    Dim y011Cnt As Integer
    Dim Index   As Integer
    Dim suiteiHiki As String                '推定対象引上げ率
    Dim suiteiTei  As String                '推定位置比抵抗値
    Dim dSXL_Pos   As Double                'SXL位置(Intel)
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function SuiteiResDataCalculation"
    
    SuiteiResDataCalculation = FUNCTION_RETURN_FAILURE
        
    '推定位置を取得する。
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT " & vbCrLf
    sql = sql & " LOTID," & vbCrLf                  'ブロックID
    sql = sql & " BLOCKSEQ," & vbCrLf               'ブロック内連番
    sql = sql & " RITOP_POS " & vbCrLf              '理論結晶内位置
    sql = sql & "FROM " & vbCrLf
    sql = sql & " TBCMY011 " & vbCrLf
    sql = sql & "WHERE " & vbCrLf
    sql = sql & " MSXLID='" & SXLID & "'" & vbCrLf  'SXLID
    sql = sql & "ORDER BY " & vbCrLf
    sql = sql & " LOTID," & vbCrLf                  'ブロックID
    sql = sql & " BLOCKSEQ" & vbCrLf                'ブロック内連番
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            
    ''Debug.Print "ループするところ(Y011)のSQL文"
    ''Debug.Print sql
    
    y011Cnt = rs.RecordCount - 1
    ReDim recY011(y011Cnt)      '☆配列にセットする意味は？
    Index = 0
    Do While Not rs.EOF
        'ﾌﾞﾛｯｸID
        If IsNull(rs("LOTID")) Then     '☆NULLはありえない？
            recY011(Index).LOTID = ""
        Else
            recY011(Index).LOTID = rs("LOTID")
        End If
        'ﾌﾞﾛｯｸ内連番
        If IsNull(rs("BLOCKSEQ")) Then  '☆NULLはありえない？
            recY011(Index).BLOCKSEQ = 0
        Else
            recY011(Index).BLOCKSEQ = CInt(rs("BLOCKSEQ"))
        End If
        '推定位置(理論結晶内位置)
        'If IsNull(rs.Fields("RTOP_POS")) Then
        If IsNull(rs.Fields("RITOP_POS")) Then
            recY011(Index).RITOP_POS = vbNullString     '☆NULLの場合はNULLをセット(不要？)
        Else
            'recY011(Index).RTOP_POS = rs.Fields("RTOP_POS")
            recY011(Index).RITOP_POS = rs.Fields("RITOP_POS")
        End If
                
        '推定対象引上げ率を求める。
        d.SMPLPOS = recY011(Index).RITOP_POS        '☆推定位置(理論結晶内位置)のセット
        d.SUITEIHIKIRITU = HikiageCalculation(d)
                
        '推定位置比抵抗値を求める。
        d.SUITEITEIKOU = SuiteiResCalculation(d)

        
'④TBCMY011テーブルのインゴット位置引上率,枚葉推定抵抗値を更新する。
        '2011/01/19 tkimura 小数点2桁から4桁に増やした。
        suiteiHiki = RoundDown(d.SUITEIHIKIRITU, 4)
        suiteiTei = RoundDown(d.SUITEITEIKOU, 5)
        '2011/04/25 ADD Marushita ☆Micronｲﾝｺﾞｯﾄ位置管理追加対応
        '☆SXL位置(Intel)を計算し小数点2桁で切り捨て(補正結晶長+理論結晶内位置)
        dSXL_Pos = RoundDown(d.HOSEICHO + recY011(Index).RITOP_POS, 2)
        
        'Add Start 2011/05/31 Y.Hitomi
        If Index = 0 Then
            sUP_RATIO(0) = suiteiHiki
        End If
        'Add End   2011/05/31 Y.Hitomi
        
        '2011/04/25 MOD Marushita ☆Micronｲﾝｺﾞｯﾄ位置管理追加対応
        '☆SXL位置(Intel)を引数に追加
'        If UpdateTBCMY011SuiteiResData(recY011(Index).LOTID, _
'                                       recY011(Index).BLOCKSEQ, _
'                                       suiteiHiki, _
'                                       suiteiTei) = FUNCTION_RETURN_FAILURE Then
        If UpdateTBCMY011SuiteiResData(recY011(Index).LOTID, _
                                       recY011(Index).BLOCKSEQ, _
                                       suiteiHiki, suiteiTei, dSXL_Pos) _
                                       = FUNCTION_RETURN_FAILURE Then
            GoTo proc_exit
        End If
            
        Index = Index + 1
        rs.MoveNext
    Loop
    
    'Add Start 2011/05/31 Y.Hitomi
    sUP_RATIO(1) = suiteiHiki
    'Add End   2011/05/31 Y.Hitomi
    
    SuiteiResDataCalculation = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SuiteiResDataCalculation = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/01/18 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/02/14 tkimura ADD START
'概要      :CLESTA評価実績(TBCMJ023)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX006         , O  ,c_cmzcrec         , TBCMX006構造体(Cu-Deco構造体)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :CLESTA評価実績(TBCMJ023)からﾃﾞｰﾀを取得し、Cu-Deco構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ023(CRYNUM As String, recXSDCS As c_cmzcrec, recX006 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ023"
    
    getTBCMJ023 = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX006
        '.Fields("SXLC_HOUDATTIM").Value =                 'C-測定日付
        .Fields("SXLC_SGYCD").Value = " "                  'C-測定作業者
        .Fields("SXLC_PTN").Value = " "                    'C-ﾊﾟﾀｰﾝ
        .Fields("SXLC_DHANKEI").Value = -1                 'C-Disk径(半径)
        .Fields("SXLC_RNAIKEI").Value = -1                 'C-Ring内径
        .Fields("SXLC_RGAIKEI").Value = -1                 'C-Ring外径
        .Fields("SXLC_SZ").Value = " "                     'C-測定条件
        '.Fields("SXLCJ_HOUDATTIM").Value =                 'CJ-測定日付
        .Fields("SXLCJ_SGYCD").Value = " "                 'CJ-測定作業者
        .Fields("SXLCJ_PTN").Value = " "                   'CJ-ﾊﾟﾀｰﾝ
        .Fields("SXLCJ_DHANKEI").Value = -1                'CJ-Disk径(半径)
        .Fields("SXLCJ_RNAIKEI").Value = -1                'CJ-Ring内径
        .Fields("SXLCJ_RGAIKEI").Value = -1                'CJ-Ring外径
        .Fields("SXLCJ_BNAIKEI").Value = -1                'CJ-Band内径
        .Fields("SXLCJ_BGAIKEI").Value = -1                'CJ-Band外径
        .Fields("SXLCJ_RHABA").Value = -1                  'CJ-Ring幅
        .Fields("SXLCJ_PIHABA").Value = -1                 'CJ-Pi幅
        .Fields("SXLCJ_NETU").Value = " "                  'CJ-熱処理法
        .Fields("SXLCJ_JUDGE").Value = " "                 'CJ-部位別判定結果
        '.Fields("SXLCJLT_HOUDATTIM").Value =               'CJLT-測定日付
        .Fields("SXLCJLT_SGYCD").Value = " "               'CJLT-測定作業者
        .Fields("SXLCJLT_PTN").Value = " "                 'CJLT-パターン
        .Fields("SXLCJLT_DHANKEI").Value = -1              'CJLT-Disk径（半径）
        .Fields("SXLCJLT_RNAIKEI").Value = -1              'CJLT-Ring内径
        .Fields("SXLCJLT_RGAIKEI").Value = -1              'CJLT-Ring外径
        .Fields("SXLCJLT_BNAIKEI").Value = -1              'CJLT-Band内径
        .Fields("SXLCJLT_BGAIKEI").Value = -1              'CJLT-Band外径
        .Fields("SXLCJLT_RHABA").Value = -1                'CJLT-Ring幅
        .Fields("SXLCJLT_PIHABA").Value = -1               'CJLT-Pi 幅
        .Fields("SXLCJLT_BHABA").Value = -1                'CJLT-Band幅
        .Fields("SXLCJLT_NETU").Value = " "                'CJLT-熱処理法
        '.Fields("SXLCJ2_HOUDATTIM").Value =                'CJ2-測定日付
        .Fields("SXLCJ2_SGYCD").Value = " "                'CJ2-測定作業者
        .Fields("SXLCJ2_PTN").Value = " "                  'CJ2-ﾊﾟﾀｰﾝ
        .Fields("SXLCJ2_DHANKEI").Value = -1               'CJ2-Disk径(半径)
        .Fields("SXLCJ2_RNAIKEI").Value = -1               'CJ2-Ring内径
        .Fields("SXLCJ2_RGAIKEI").Value = -1               'CJ2-Ring外径
        .Fields("SXLCJ2_PIHABA").Value = -1                'CJ2-Pi幅
        .Fields("SXLCJ2_NETU").Value = " "                 'CJ2-熱処理法
        .Fields("SXLCJ2_JUDGE").Value = " "                'CJ2-部位別判定結果

        '-------------------- TBCMJ023の読み込み(CJ) ----------------------------------------
        sql = "select * from TBCMJ023 "
        sql = sql & "where CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "      SMPLNO = " & recXSDCS("CRYSMPLIDL4CS").Value
        sql = sql & "order by TRANCNT desc"
        'sql = "select J023.*,to_char(J023.REGDATEC,'YYYY/MM/DD HH24:MI:SS') AS REGDATE from (" & sql & ") J023 where rownum = 1"
        sql = "select * from (" & sql & ") where rownum = 1"
        Debug.Print (sql)
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            getTBCMJ023 = FUNCTION_RETURN_SUCCESS                     '2011/02/28 tkimura
            GoTo proc_exit
        End If

        'TBCMX006
        .Fields("SXLC_HOUDATTIM").Value = rs("REGDATEC")                'C-測定日付
        '.Fields("SXLC_HOUDATTIM").Value = "to_date( " & rs("REGDATEC") & ",'yyyy/mm/dd hh24:mi:ss')"                'C-測定日付"
        .Fields("SXLC_SGYCD").Value = rs("TSTAFFIDC")                    'C-測定作業者
        .Fields("SXLC_PTN").Value = rs("CPTNJSK")                        'C-ﾊﾟﾀｰﾝ
        If IsNull(rs("CDISKJSK").Value) = False Then
            .Fields("SXLC_DHANKEI").Value = rs("CDISKJSK")               'C-Disk径(半径)
        End If
        If IsNull(rs("CRINGNKJSK").Value) = False Then
            .Fields("SXLC_RNAIKEI").Value = rs("CRINGNKJSK")             'C-Ring内径
        End If
        If IsNull(rs("CRINGGKJSK").Value) = False Then
            .Fields("SXLC_RGAIKEI").Value = rs("CRINGGKJSK")             'C-Ring外径
        End If
        .Fields("SXLC_SZ").Value = rs("C_SZ")                            'C-測定条件
        .Fields("SXLCJ_HOUDATTIM").Value = rs("REGDATECJ")               'CJ-測定日付
        'Format(Now, "yyyy/mm/dd hh:mm:ss")
        '.Fields("SXLCJ_HOUDATTIM").Value = Format(rs("REGDATECJ"), "yyyy/mm/dd hh:mm:ss")              'CJ-測定日付
        .Fields("SXLCJ_SGYCD").Value = rs("TSTAFFIDCJ")                  'CJ-測定作業者
        .Fields("SXLCJ_PTN").Value = rs("CJPTNJSK")                      'CJ-ﾊﾟﾀｰﾝ
        If IsNull(rs("CJDISKJSK").Value) = False Then
            .Fields("SXLCJ_DHANKEI").Value = rs("CJDISKJSK")             'CJ-Disk径(半径)
        End If
        If IsNull(rs("CJRINGNKJSK").Value) = False Then
            .Fields("SXLCJ_RNAIKEI").Value = rs("CJRINGNKJSK")           'CJ-Ring内径
        End If
        If IsNull(rs("CJRINGGKJSK").Value) = False Then
            .Fields("SXLCJ_RGAIKEI").Value = rs("CJRINGGKJSK")           'CJ-Ring外径
        End If
        If IsNull(rs("CJBANDNKJSK").Value) = False Then
            .Fields("SXLCJ_BNAIKEI").Value = rs("CJBANDNKJSK")           'CJ-Band内径
        End If
        If IsNull(rs("CJBANDGKJSK").Value) = False Then
            .Fields("SXLCJ_BGAIKEI").Value = rs("CJBANDGKJSK")           'CJ-Band外径
        End If
        If IsNull(rs("CJRINGCALC").Value) = False Then
            .Fields("SXLCJ_RHABA").Value = rs("CJRINGCALC")              'CJ-Ring幅
        End If
        If IsNull(rs("CJPICALC").Value) = False Then
            .Fields("SXLCJ_PIHABA").Value = rs("CJPICALC")               'CJ-Pi幅
        End If
        .Fields("SXLCJ_NETU").Value = rs("CJ_NETU")                      'CJ-熱処理法
        .Fields("SXLCJ_JUDGE").Value = rs("CJHANTEI")                    'CJ-部位別判定結果
        .Fields("SXLCJLT_HOUDATTIM").Value = rs("REGDATECJLT")           'CJLT-測定日付
        .Fields("SXLCJLT_SGYCD").Value = rs("TSTAFFIDCJLT")              'CJLT-測定作業者
        .Fields("SXLCJLT_PTN").Value = rs("CJLTPTNJSK")                  'CJLT-パターン
        If IsNull(rs("CJLTDISKJSK").Value) = False Then
            .Fields("SXLCJLT_DHANKEI").Value = rs("CJLTDISKJSK")         'CJLT-Disk径（半径）
        End If
        If IsNull(rs("CJLTRINGNKJSK").Value) = False Then
            .Fields("SXLCJLT_RNAIKEI").Value = rs("CJLTRINGNKJSK")       'CJLT-Ring内径
        End If
        If IsNull(rs("CJLTRINGGKJSK").Value) = False Then
            .Fields("SXLCJLT_RGAIKEI").Value = rs("CJLTRINGGKJSK")       'CJLT-Ring外径
        End If
        If IsNull(rs("CJLTBANDNKJSK").Value) = False Then
            .Fields("SXLCJLT_BNAIKEI").Value = rs("CJLTBANDNKJSK")       'CJLT-Band内径
        End If
        If IsNull(rs("CJLTBANDGKJSK").Value) = False Then
            .Fields("SXLCJLT_BGAIKEI").Value = rs("CJLTBANDGKJSK")       'CJLT-Band外径
        End If
        If IsNull(rs("CJLTRINGCALC").Value) = False Then
            .Fields("SXLCJLT_RHABA").Value = rs("CJLTRINGCALC")          'CJLT-Ring幅
        End If
        If IsNull(rs("CJLTPICALC").Value) = False Then
            .Fields("SXLCJLT_PIHABA").Value = rs("CJLTPICALC")           'CJLT-Pi 幅
        End If
        If IsNull(rs("CJLTPICALC").Value) = False Then
            .Fields("SXLCJLT_BHABA").Value = rs("CJLTPICALC")            'CJLT-Band幅
        End If
        .Fields("SXLCJLT_NETU").Value = rs("CJLT_NETU")                  'CJLT-熱処理法
        .Fields("SXLCJ2_HOUDATTIM").Value = rs("REGDATECJ2")             'CJ2-測定日付
        '.Fields("SXLCJ2_HOUDATTIM").Value = CDate(rs("REGDATECJ2"))             'CJ2-測定日付
        .Fields("SXLCJ2_SGYCD").Value = rs("TSTAFFIDCJ2")                'CJ2-測定作業者
        .Fields("SXLCJ2_PTN").Value = rs("CJ2PTNJSK")                    'CJ2-ﾊﾟﾀｰﾝ
        If IsNull(rs("CJ2DISKJSK").Value) = False Then
            .Fields("SXLCJ2_DHANKEI").Value = rs("CJ2DISKJSK")           'CJ2-Disk径(半径)
        End If
        If IsNull(rs("CJ2RINGNKJSK").Value) = False Then
            .Fields("SXLCJ2_RNAIKEI").Value = rs("CJ2RINGNKJSK")         'CJ2-Ring内径
        End If
        If IsNull(rs("CJ2RINGGKJSK").Value) = False Then
            .Fields("SXLCJ2_RGAIKEI").Value = rs("CJ2RINGGKJSK")         'CJ2-Ring外径
        End If
        If IsNull(rs("CJ2PICALC").Value) = False Then
            .Fields("SXLCJ2_PIHABA").Value = rs("CJ2PICALC")             'CJ2-Pi幅
        End If
        .Fields("SXLCJ2_NETU").Value = rs("CJ2_NETU")                    'CJ2-熱処理法
        .Fields("SXLCJ2_JUDGE").Value = rs("CJ2HANTEI")                  'CJ2-部位別判定結果

        Set rs = Nothing
    End With

    getTBCMJ023 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ023 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/02/14 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/02/14 tkimura ADD START
'概要      :結晶OSF実績データ取得設定(TBCMJ005)データ取得設定
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :CRYNUM          , I  ,String            , 結晶番号
'          :recXSDCS        , I  ,c_cmzcrec         , XSDCS構造体   (新ｻﾝﾌﾟﾙ管理(ﾌﾞﾛｯｸ))
'          :recX006         , O  ,c_cmzcrec         , TBCMX006構造体(Cu-Deco構造体)
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :結晶OSF実績(TBCMJ005)からﾃﾞｰﾀを取得し、Cu-Deco構造体にｾｯﾄする
'履歴      :2003/10/18 SystemBrain 新規作成
Private Function getTBCMJ005CuDeco(CRYNUM As String, recXSDCS As c_cmzcrec, recX006 As c_cmzcrec) As FUNCTION_RETURN
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCMJ005CuDeco"
    
    getTBCMJ005CuDeco = FUNCTION_RETURN_FAILURE

    '-------------------- 初期ｸﾘｱ ----------------------------------------
    With recX006
        '.Fields("SXLCOSF3_HOUDATTIM").Value =                  'C-OSF3測定日付
        .Fields("SXLCOSF3_SGYCD").Value = " "                   'C-OSF3測定作業者
        .Fields("SXLCOSF3_PTN").Value = " "                     'C-OSF3ﾊﾟﾀｰﾝ
        .Fields("SXLCOSF3_DHANKEI").Value = -1                  'C-OSF3Disk径(半径)
        .Fields("SXLCOSF3_RNAIKEI").Value = -1                  'C-OSF3Ring内径
        .Fields("SXLCOSF3_RGAIKEI").Value = -1                  'C-OSF3Ring外径
        .Fields("SXLCOSF3_RHABA").Value = -1                    'C-OSF3Ring幅
        .Fields("SXLCOSF3_NETU").Value = " "                    'C-OSF3熱処理法
        .Fields("SXLCOSF3_JUDGE").Value = " "                   'C-OSF3部位別判定結果

        '-------------------- TBCMJ005の読み込み(OSF1,2) ----------------------------------------
        sql = "select"
        sql = sql & "    REGDATE,"
        sql = sql & "    TSTAFFID,"
        sql = sql & "    SXLCOSF3_PTN,"
        sql = sql & "    SXLCOSF3_DHANKEI,"
        sql = sql & "    SXLCOSF3_RNAIKEI,"
        sql = sql & "    SXLCOSF3_RGAIKEI,"
        sql = sql & "    SXLCOSF3_RHABA,"
        sql = sql & "    HTPRC,"
        sql = sql & "    PTNJUDGRES"
        sql = sql & " from"
        sql = sql & "    (select "
        sql = sql & "        TSTAFFID,"
        sql = sql & "        REGDATE,"
        sql = sql & "        case when(OSFRD1 = '-' and OSFRD2 = '-') then '0' "
        sql = sql & "             when(OSFRD1 = 'D' and OSFRD2 = '-') then '2' "
        sql = sql & "             when(OSFRD1 = 'R' and OSFRD2 = '-') then '1' "
        sql = sql & "             when(OSFRD1 = 'D' and OSFRD2 = 'R' or OSFRD1 = 'R' and OSFRD2 = 'D') then '3' "
'Cng Start 2011/04/12 Y.Hitomi
'        sql = sql & "             else '-1' "
        sql = sql & "             when(OSFRD1 = 'R' and OSFRD2 = 'R') then '1' "
        sql = sql & "             else ' ' "
'Cng Start 2011/04/12 Y.Hitomi
        sql = sql & "        end SXLCOSF3_PTN,"
        sql = sql & "        case when(OSFRD1 = 'D') then OSFWID1 "
        sql = sql & "             when(OSFRD2 = 'D') then OSFWID2 "
        sql = sql & "        end SXLCOSF3_DHANKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then (150-OSFPOS1-OSFWID1) "
        sql = sql & "             when(OSFRD2 = 'R') then (150-OSFPOS2-OSFWID2) "
        sql = sql & "        end SXLCOSF3_RNAIKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then (150-OSFPOS1) "
        sql = sql & "             when(OSFRD2 = 'R') then (150-OSFPOS2) "
        sql = sql & "        end SXLCOSF3_RGAIKEI,"
        sql = sql & "        case when(OSFRD1 = 'R') then OSFWID1 "
        sql = sql & "             when(OSFRD2 = 'R') then OSFWID2  "
        sql = sql & "        end SXLCOSF3_RHABA,"
        sql = sql & "        HTPRC,"
        sql = sql & "        PTNJUDGRES "
        sql = sql & "    from "
        sql = sql & "        TBCMJ005"
        sql = sql & "    where "
        sql = sql & "        CRYNUM = '" & CRYNUM & "' and "
        sql = sql & "        SMPLNO = " & recXSDCS("CRYSMPLIDL4CS").Value
        sql = sql & "        order by TRANCNT desc"
        sql = sql & "    ) "
        sql = sql & " where rownum = 1"
        Debug.Print (sql)
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            getTBCMJ005CuDeco = FUNCTION_RETURN_SUCCESS                     '2011/02/28 tkimura
            GoTo proc_exit
        End If

        'TBCMX006(数値にNULLが入ることを防いでおく。)
        .Fields("SXLCOSF3_HOUDATTIM").Value = rs("REGDATE")                  'C-OSF3測定日付
        .Fields("SXLCOSF3_SGYCD").Value = rs("TSTAFFID")                     'C-OSF3測定作業者
        .Fields("SXLCOSF3_PTN").Value = rs("SXLCOSF3_PTN")                   'C-OSF3ﾊﾟﾀｰﾝ
        If IsNull(rs("SXLCOSF3_DHANKEI").Value) = False Then
            .Fields("SXLCOSF3_DHANKEI").Value = rs("SXLCOSF3_DHANKEI")       'C-OSF3Disk径(半径)
        End If
        If IsNull(rs("SXLCOSF3_RNAIKEI").Value) = False Then
            .Fields("SXLCOSF3_RNAIKEI").Value = rs("SXLCOSF3_RNAIKEI")       'C-OSF3Ring内径
        End If
        If IsNull(rs("SXLCOSF3_RGAIKEI").Value) = False Then
            .Fields("SXLCOSF3_RGAIKEI").Value = rs("SXLCOSF3_RGAIKEI")       'C-OSF3Ring外径
        End If
        If IsNull(rs("SXLCOSF3_RHABA").Value) = False Then
            .Fields("SXLCOSF3_RHABA").Value = rs("SXLCOSF3_RHABA")           'C-OSF3Ring幅
        End If
        .Fields("SXLCOSF3_NETU").Value = rs("HTPRC")                         'C-OSF3熱処理法
        .Fields("SXLCOSF3_JUDGE").Value = rs("PTNJUDGRES")                   'C-OSF3部位別判定結果

        Set rs = Nothing
    End With

    getTBCMJ005CuDeco = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCMJ005CuDeco = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
' 2011/02/14 tkimura ADD END
'=================================================================================

''2011/01/17 tkimura ADD START ==========================================================>
'概要      :偏析計算に必要な各合計重量実績を取得する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :CRYNUM        ,I  ,String    ,結晶番号
'          :wgtCharge     ,O  ,Long      ,炉内量
'          :wgtChargeA    ,O  ,Long      ,A結晶の炉内量
'          :wgtTop        ,O  ,Double    ,トップ重量実績値
'          :wgtTopCut     ,O  ,Double    ,トップカット重量実績値
'          :DM            ,O  ,Double    ,直径１～３の平均
'          :hikiFlg       ,O  ,Integer   ,引上げフラグ(1=通常、2=BC結晶)
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :【マルチ引上対応】 全量引き､残量引き､RC引きにあわせて実績データを取得する
'履歴      :2008/04/21 作成  SETsw Nakada
'          :2011/01/17 参照作成  tkimura
'          :2011/04/28 Marushita （\cmmc001\s_cmmc001z.bas から移動
Public Function GetCoeffParams_new2(ByVal CRYNUM$, _
                                    wgtCharge As Long, _
                                    wgtChargeA As Long, _
                                    wgtTop As Double, _
                                    wgtTopCut As Double, _
                                    DM As Double, _
                                    HIKIFLG As Integer) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim cryNumA As String       'BC結晶処理でのA結晶を格納する。

    On Error GoTo Err
    GetCoeffParams_new2 = FUNCTION_RETURN_FAILURE
    wgtCharge = 0
    wgtChargeA = 0
    wgtTop = 0#
    wgtTopCut = 0#
    DM = 0#
    
    '' 推定チャージ、重量（TOP）、トップカット重量、直胴直径の平均値 取得
    sql = " SELECT C1.SUICHARGE, C1.WGHTTOC1, C1.PUTCUTWC1, "
    sql = sql & " (C1.DIA1C1 + C1.DIA2C1 + C1.DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 C1 "
    sql = sql & " WHERE C1.XTALC1 = '" & CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        wgtCharge = rs("SUICHARGE")       ''推定チャージ
        wgtTop = rs("WGHTTOC1")           ''重量（TOP）
        wgtTopCut = rs("PUTCUTWC1")       ''トップカット重量
        DM = rs("DM")                     ''直胴直径(平均値)
    End If
    rs.Close
    
    '結晶番号の9桁がBorCならばBC結晶となる。
    If Mid(CRYNUM, 9, 1) = "B" Or Mid(CRYNUM, 9, 1) = "C" Then
        HIKIFLG = "2"       'BC結晶
    Else
        HIKIFLG = "1"       '通常
    End If
    
    'このあとにwgtChargeAを求める必要がある。(HIKIFLG="2"のときのみ)
    If HIKIFLG = "2" Then
        cryNumA = Mid(CRYNUM, 1, 8) & "A" & Mid(CRYNUM, 10, 3)      '結晶番号の9桁目をAにする。
        sql = " SELECT C1.SUICHARGE "
        sql = sql & " FROM XSDC1 C1 "
        sql = sql & " WHERE C1.XTALC1 = '" & cryNumA & "'"

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        If rs.RecordCount > 0 Then
            wgtChargeA = rs("SUICHARGE")       ''推定チャージ
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    GetCoeffParams_new2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    On Error GoTo 0
    Exit Function

Err:
    Resume proc_exit
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要      :偏析係数を求める。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型                       ,説明
'          :d             ,I ,type_Coefficient_new2     ,推定抵抗,推定引上率計算構造体
'          :戻り値        ,O  ,Double                   ,偏析係数
'説明      :
'履歴      :2001/06/23　佐野 信哉　作成
'          :2011/01/17  参照作成  tkimura
'          :2011/04/28  Marushita （\cmmc001\s_cmmc001z.bas から移動
Public Function CoefficientCalculation_new2(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
    
    CoefficientCalculation_new2 = Log(d.BOTRES / (d.TOPRES * 1)) / Log((1 - d.GT) / (1 - d.GB)) + 1
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    CoefficientCalculation_new2 = -9999
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :引上げ率を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O  ,Double                     ,位置引上率
'説明       :
'履歴       :2011/01/17 tkimura
'           :2011/04/28 Marushita （\cmmc001\s_cmmc001z.bas から移動
Public Function HikiageCalculation(ByRef d As type_Coefficient_new2) As Double
    Dim result As Double

    '通常
    If d.HIKIFLG = "1" Then
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT) / (d.CHARGEWEIGHT)
    'BC結晶
    Else
        result = (d.DUNMENSEKI * d.SMPLPOS * HIJU_SILICONE + d.TOPWEIGHT + d.CHARGEWEIGHTA - d.CHARGEWEIGHT) / (d.CHARGEWEIGHTA)
    End If
    
    HikiageCalculation = result
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :基準抵抗値を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O ,Double                      ,基準抵抗値
'説明       :
'履歴       :2011/01/17 tkimura
'           :2011/04/28 Marushita （\cmmc001\s_cmmc001z.bas から移動
Public Function StandardResCalculation(d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    StandardResCalculation = d.TOPRES * (1 - d.GT) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    StandardResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
' 2011/01/17 tkimura ADD START
'概要       :推定位置比抵抗値を計算する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :d              ,I  ,type_Coefficient_new2      ,推定抵抗,推定引上率計算構造体
'           :戻り値         ,O ,Double                      ,推定位置比抵抗値
'説明       :
'履歴       :2011/01/17 tkimura
'           :2011/04/28 Marushita （\cmmc001\s_cmmc001z.bas から移動
Public Function SuiteiResCalculation(ByRef d As type_Coefficient_new2) As Double
    
    On Error GoTo Err
        
    SuiteiResCalculation = d.KIJUNTEIKOU / (1 - d.SUITEIHIKIRITU) ^ (d.Henseki - 1)
        
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    SuiteiResCalculation = -9999
    
End Function
' 2011/01/17 tkimura ADD END
'=================================================================================

'=================================================================================
'概要       :前引上の結晶番号を取得する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sCrynum        ,I  ,String             　      ,現在結晶番号
'           :sP_Crynum      ,O  ,String             　      ,前引上の結晶番号
'           :戻り値         ,O  ,FUNCTION_RETURN            ,
'説明       :
'履歴       :2011/04/25 Marushita
Public Function GetPreCrynum(ByVal sCryNum As String, ByRef sP_Crynum As String) As FUNCTION_RETURN
    
    On Error GoTo Err
    GetPreCrynum = FUNCTION_RETURN_FAILURE
                
    Dim sCrynum9 As String
    Dim iCrynum9 As Integer
            
    sP_Crynum = ""
    
    'sCrynumの9桁目をセット
    sCrynum9 = Mid(sCryNum, 9, 1)
    
    '"B"より小さい記号はエラーで返す
    If sCrynum9 < "B" Then
        Exit Function
    Else
        iCrynum9 = Asc(sCrynum9)
        sP_Crynum = Mid(sCryNum, 1, 8) & Chr(iCrynum9 - 1) & Mid(sCryNum, 10, 3)
    End If
    
    GetPreCrynum = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function
Err:
    On Error GoTo 0
    
End Function
'=================================================================================

'=================================================================================
'概要       :前引上の結晶情報を取得する。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sCRYNUM        ,I  ,String    ,前引上結晶番号
'           :dWgtTop        ,O  ,Double    ,トップ重量実績値
'           :dDM            ,O  ,Double    ,直径１～３の平均
'           :lLentk         ,O  ,Long      ,引上長
'           :dwgtTA         ,O  ,Double    ,テイル重量実績値
'           :戻り値         ,O  ,FUNCTION_RETURN            ,
'説明       :
'履歴       :2011/04/25 Marushita
Public Function GetPreXSDC1(ByVal sCryNum As String, _
                            ByRef dwgtTop As Double, _
                            ByRef dDM As Double, _
                            ByRef lLenTK As Long, _
                            ByRef dwgtTA As Double) As FUNCTION_RETURN
    
    On Error GoTo proc_err
    
    Dim sql As String
    Dim rs As OraDynaset
        
    GetPreXSDC1 = FUNCTION_RETURN_FAILURE
    
    sql = " SELECT WGHTTOC1, LENTKC1, WGHTTAC1, "
    sql = sql & " (DIA1C1 + DIA2C1 + DIA3C1) / 3.0 AS DM "
    sql = sql & " FROM XSDC1 "
    sql = sql & " WHERE XTALC1 = '" & sCryNum & "'"
        
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount > 0 Then
        dwgtTop = rs("WGHTTOC1")           ''重量（TOP）
        dDM = rs("DM")                     ''直胴直径(平均値)
        lLenTK = rs("LENTKC1")             ''引上長
        dwgtTA = rs("WGHTTAC1")            ''重量（TAIL）
    End If
    rs.Close
    
    GetPreXSDC1 = FUNCTION_RETURN_SUCCESS
    
    On Error GoTo 0
    Exit Function

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    GetPreXSDC1 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
    
End Function
'=================================================================================

'=================================================================================
'概要       :テーブルに項目名が存在するかチェックする。
'ﾊﾟﾗﾒｰﾀ     :変数名         ,IO ,型                         ,説明
'           :sTblName       ,I  ,String             　      ,現在結晶番号
'           :sFldName       ,I  ,String             　      ,前引上の結晶番号
'           :戻り値         ,O  ,FUNCTION_RETURN            ,
'説明       :
'履歴       :2011/04/25 Marushita
Public Function FieldCheck(ByVal sTblName As String, ByVal sFldName As String) As FUNCTION_RETURN
Dim rs As OraDynaset
Dim fld As OraField

    FieldCheck = FUNCTION_RETURN_FAILURE
    'エラーハンドラの設定
    On Error GoTo proc_err

    Set rs = OraDB.CreateDynaset("select * from " & sTblName, ORADYN_NO_BLANKSTRIP)
    For Each fld In rs.Fields
        ''そのフィールドが未登録なら、既定値で登録する
        If fld.Name = sFldName Then
            FieldCheck = FUNCTION_RETURN_SUCCESS
            rs.Close
            Exit Function
        End If
    Next
    rs.Close

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    FieldCheck = FUNCTION_RETURN_FAILURE
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit

End Function

'=================================================================================
'概要      :品番マスタ情報(TBCME036)の中間抜試単位を取得する。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :HIN             , I  ,tFullHinban       , 品番(全品番構造体)
'          :iUnit           , I  ,Integer           , 中間抜試単位
'      　　:戻り値          , O  ,FUNCTION_RETURN　 , 成否
'説明      :TBCME036.MCUTUNITの値を取得する。
'履歴      :2011/06/24 Marushita
Private Function getTBCME036(ByRef HIN As tFullHinban, _
                                        ByRef iUnit As Integer, ByRef sFlg As Integer) As FUNCTION_RETURN
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sndFlg  As String       '送信フラグ
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getTBCME036"
    
    getTBCME036 = FUNCTION_RETURN_FAILURE

    iUnit = 0
    'から中間抜試単位を取得する。
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT" & vbCrLf
    sql = sql & " NVL(MSMPTANIMAI,0) as MSMPTANIMAI," & vbCrLf     '中間抜試単位
    sql = sql & " NVL(MSMPFLG,'0')   as MSMPFLG" & vbCrLf          '中間抜試フラグ
    sql = sql & "FROM" & vbCrLf
    sql = sql & " TBCME036" & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " HINBAN ='" & HIN.hinban & "' AND" & vbCrLf
    sql = sql & " MNOREVNO =" & HIN.mnorevno & " AND" & vbCrLf
    sql = sql & " FACTORY ='" & HIN.factory & "' AND" & vbCrLf
    sql = sql & " OPECOND ='" & HIN.opecond & "'" & vbCrLf
            
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        getTBCME036 = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    iUnit = rs("MSMPTANIMAI")
    sFlg = rs("MSMPFLG")
        
    getTBCME036 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getTBCME036 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'=================================================================================
'概要      :XSDC2のブロック開始位置を取得する。
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :sCrynum         , I  ,String            , 品番(全品番構造体)
'      　　:戻り値          , O  ,Integer        　 , ブロック開始位置
'説明      :XSDC2.INPOSの値を取得する。
'履歴      :2011/07/11 Marushita
Private Function getXSDC2Pos(ByVal sCryNum As String) As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sndFlg  As String       '送信フラグ
    Dim i       As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc040_SQL.bas -- Function getXSDC2Pos"
    
    getXSDC2Pos = 0

    'XSDC2からINPOSを取得する。
    Set rs = Nothing
    sql = ""
    sql = sql & "SELECT" & vbCrLf
    sql = sql & "NVL(INPOSC2,0) INPOSC2" & vbCrLf        '中間抜試単位
    sql = sql & "FROM" & vbCrLf
    sql = sql & " XSDC2" & vbCrLf
    sql = sql & "WHERE" & vbCrLf
    sql = sql & " CRYNUMC2 ='" & sCryNum & "' AND" & vbCrLf
    sql = sql & " LIVKC2 <> '1' " & vbCrLf
            
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        GoTo proc_exit
    End If
    
    getXSDC2Pos = CInt(rs.Fields("INPOSC2"))
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    getXSDC2Pos = 0
    gErr.HandleError
    Resume proc_exit
End Function

