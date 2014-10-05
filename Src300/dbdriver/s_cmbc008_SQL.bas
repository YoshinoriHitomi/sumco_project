Attribute VB_Name = "s_cmbc008_SQL"
Option Explicit

Public Type type_DBDRV_cmgc001f1
    CRYNUM As String * 12   ' 結晶番号
    INGOTPOS As Integer     ' 開始位置
    HINBAN As String * 8    ' Top品番＝品番
    REVNUM As Integer       ' 製品番号改訂番号
    factory As String * 1   ' 工場
    opecond As String * 1   ' 操業条件
End Type

Public Type type_DBDRV_cmgc001f2
    PALTNUM As String * 4   ' パレット番号
    BDCODE As String * 3    ' 不良理由コード＝格下区分
    BLOCKID As String * 12  ' ブロックID
End Type

Public Type type_DBDRV_cmgc001f3
    PGID As String * 8      ' ＰＧ－ＩＤ
    CRYNUM As String * 12   ' 結晶番号
End Type

Public Type type_DBDRV_cmgc001f4
    DMTOP1 As Double        ' 直径
    DMTOP2 As Double        ' 直径
    DMTAIL1 As Double       ' 直径
    DMTAIL2 As Double       ' 直径
    INGOTPOS As Integer     ' 開始位置
    TRANCNT As Integer      ' 処理回数
    NCHPOS As String * 2    ' ノッチ位置
    CRYNUM As String * 12   ' 結晶番号
End Type

Public Type type_DBDRV_cmgc001f5
    CRYNUM As String * 12   ' 結晶番号
    MAGTYPE As String * 2   ' 磁場タイプ＝製造法
End Type

Public Type type_DBDRV_cmgc001f6
'2002/04/25 S.Sano    TYPE As String * 1      ' 品ＳＸタイプ＝タイプ
    HSXCDIR As String * 1   ' 品ＳＸ結晶面方位＝方位
    HINBAN As String * 8    ' Top品番＝品番
End Type

Public Type type_DBDRV_cmgc001f7
    DPNTCLS As String * 7   ' ドーパント種類＝ドーパント　結晶番号前7桁+"00"＝引上げ指示№
    CRYNUM As String * 12   ' 結晶番号
    TYPE As String * 2      ' タイプ'2002/04/25 S.Sano
End Type

Public Type type_DBDRV_cmgc001f8
    SMPLNO As Integer       ' サンプル№
    CRYNUM As String * 12   ' 結晶番号
    INGOTPOS As Integer     ' 開始位置
    SMPKBN As String * 1    ' サンプル区分
End Type

Public Type type_DBDRV_cmgc001f9
    CRYNUM As String * 12   ' 結晶番号
    TRANCNT As Integer      ' 処理回数
    TRANCOND As String * 1  ' 処理条件
    INGOTPOS As Integer     ' 開始位置
    SMPLNO As Long          ' サンプル№        Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    SMPKBN As String * 1    ' サンプル区分
    JudgData As Double      ' Topサンプル№に対応する検索対象値
End Type

Public Type type_DBDRV_cmgc001f10
    CRYNUM As String * 12   ' 結晶番号
    PRCMCN As String * 1    ' 研削機
    SEED As String * 4      ' SEED
End Type

' クリスタルカタログ検索格上
Public Type type_DBDRV_scmzc_fcmgc001f_Hinban
    ' 制作条件
    MAGTYPE As String * 2   ' 磁場タイプ
    ' 製品仕様SXLデータ1
    HSXD1MIN As Double      ' 品ＳＸ直径１下限
    HSXD1MAX As Double      ' 品ＳＸ直径１上限
    ' 製品仕様SXLデータ1
    HSXTYPE As String * 1   ' 品ＳＸタイプ
    HSXDOP As String * 1    ' 品ＳＸドーパント
    HSXCDIR As String * 1   ' 品ＳＸ結晶面方位
    HSXRMIN As Double       ' 品ＳＸ比抵抗下限
    HSXRMAX As Double       ' 品ＳＸ比抵抗上限
    ' 製品仕様SXLデータ2
    HSXONMIN As Double      ' 品ＳＸ酸素濃度下限
    HSXONMAX As Double      ' 品ＳＸ酸素濃度上限
    HSXCNMIN As Double      ' 品ＳＸ炭素濃度下限
    HSXCNMAX As Double      ' 品ＳＸ炭素濃度上限
    ' 製品仕様SXLデータ1
    HSXDPDIR As String * 2  ' 品ＳＸ溝位置方位
    ' 製品仕様SXLデータ3
    HSXDVDMX As Integer     ' 品ＳＸＤＶＤ２上限
    HSXDVDMN As Integer     ' 品ＳＸＤＶＤ２下限
    HSXOS1AX As Double      ' 品ＳＸＯＳＦ１平均上限
    HSXOS1MX As Double      ' 品ＳＸＯＳＦ１上限
    HSXOS2AX As Double      ' 品ＳＸＯＳＦ２平均上限
    HSXOS2MX As Double      ' 品ＳＸＯＳＦ２上限
    HSXOS3AX As Double      ' 品ＳＸＯＳＦ３平均上限
    HSXOS3MX As Double      ' 品ＳＸＯＳＦ３上限
    HSXOS4AX As Double      ' 品ＳＸＯＳＦ４平均上限
    HSXOS4MX As Double      ' 品ＳＸＯＳＦ４上限
    HSXBM1AN As Double      ' 品ＳＸＢＭＤ１平均下限
    HSXBM1AX As Double      ' 品ＳＸＢＭＤ１平均上限
    HSXBM2AN As Double      ' 品ＳＸＢＭＤ２平均下限
    HSXBM2AX As Double      ' 品ＳＸＢＭＤ２平均上限
    HSXBM3AN As Double      ' 品ＳＸＢＭＤ３平均下限
    HSXBM3AX As Double      ' 品ＳＸＢＭＤ３平均上限
    HSXLTMIN As Integer     ' 品ＳＸＬタイム下限
    HSXLTMAX As Integer     ' 品ＳＸＬタイム上限

    SGLENGTH As Integer     ' 最低合格長さ
End Type

Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
    ' 検索用情報
    CRYNUM As String * 12   ' 結晶番号
    INGOTPOS As Integer     ' 開始位置
    TOPSMPKBN As String * 1 ' Topサンプル区分
    BOTSMPKBN As String * 1 ' Botサンプル区分
    REVNUM As Integer       ' 製品番号改訂番号
    factory As String * 1   ' 工場
    opecond As String * 1   ' 操業条件
    ' ブロック管理
    BLOCKID As String * 12  ' ブロックID
    LENGTH As Integer       ' 長さ＝ブロック長
    ' 品番管理
    HINBAN As String * 8    ' Top品番＝品番
    ' クリスタルカタログ受入実績
    PALTNUM As String * 4   ' パレット番号
    BDCODE As String * 3    ' 不良理由コード＝格下区分
    ' 結晶情報
    DIAMETER As Double      ' 直径
    PGID As String * 8      ' ＰＧ－ＩＤ
    ' 結晶抵抗実績
    TOPRES As Double        ' Topサンプル№に対応する検索対象値＝Top推定ρ
    BOTRES As Double        ' Botサンプル№に対応する検索対象値＝Bot推定ρ
    TOPRESSMP As Integer        ' Topサンプル№
    BOTRESSMP As Integer        ' Botサンプル№
    TOPIND As String        '状態区分
    BOTIND As String        '状態区分
    ' Oi実績
    TOPOI   As Double       ' Topサンプル№に対応する検索対象値＝TopOi
    BOTOI   As Double       ' Topサンプル№に対応する検索対象値＝BotOi
    TOPOISMP As Long        ' Topサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTOISMP As Long        ' Botサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' Cs実績
    BOTCS   As Double       ' Cs実測値＝Cs
    BOTCSSMP As Long        ' Botサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' 加工払い出し実績
    PRCMCN  As String * 1   ' 研削機
    SEED    As String * 4   ' SEED
    ' 研削加工実績
    NCHPOS As String * 2    ' ノッチ位置
    DMTOP1 As Double        ' 直径
    DMTOP2 As Double        ' 直径
    DMTAIL1 As Double       ' 直径
    DMTAIL2 As Double       ' 直径
    ' 制作条件
    MAGTYPE As String * 2   ' 磁場タイプ＝製造法
    ' 製品仕様SXLﾃﾞｰﾀ１
    TYPE As String * 1      ' 品ＳＸタイプ＝タイプ
    HSXCDIR As String * 3   ' 品ＳＸ結晶面方位＝方位
    ' 引上げ投入実績
    DPNTCLS As String * 7   ' ドーパント種類＝ドーパント　結晶番号前7桁+"00"＝引上げ指示№
    ' ブロック管理
    TOPPOS As Integer       ' 結晶内開始位置＝Top部位
    BOTPOS As Integer       ' 結晶内開始位置＋長さ＝Bot部位
    UPDDATE As Date         ' 更新日付＝受入日付
    ' 結晶サンプル管理
    TOPSMPLNO As Long       ' Topサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTSMPLNO As Long       ' Botサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' GD実績
    TOPDVD2 As Integer      ' 測定結果 DVD2＝DVD2(Top)
    BOTDVD2 As Integer      ' 測定結果 DVD2＝DVD2(Bot)
    TOPDVD2SMP As Long        ' Topサンプル№   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTDVD2SMP As Long        ' Botサンプル№   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' OSF実績
    ' HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
    ' が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
    TOPOSF(3) As Double     ' 計算結果 Max＝OSF(Top)
    BOTOSF(3) As Double     ' 計算結果 Max＝OSF(Bot)
    TOPOSFSMP(3) As Long        ' Topサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTOSFSMP(3) As Long        ' Botサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' BMD実績
    ' HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
    ' が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
    TOPBMD(2) As Double     ' Max＝OSF(Top)
    BOTBMD(2) As Double     ' Max＝OSF(Bot)
    TOPBMDSMP(2) As Long        ' Topサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTBMDSMP(2) As Long        ' Botサンプル№     Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    ' ライフタイム
    TOPLT As Integer        ' 計算結果＝ライフタイム(Top)
    BOTLT As Integer        ' 計算結果＝ライフタイム(Bot)
    TOPLTSMP As Long        ' Topサンプル№         Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    BOTLTSMP As Long        ' Botサンプル№         Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
    '---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi START ----
    ' 引上パターン
    HIKIAGEPTRN As String   '引上パターン＝グループ指示数＋最終引上本数＋グループ内引上順番
'---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi END ----
    'ホールドデータ追加    2006/03
    HOLDKT As String
    BIKOU As String
    HLDCMNT As String
    HLDTRCLS As String
    HLDCAUSE As String
    AGRSTATUS           As String           ' 承認確認区分      add SETkimizuka
    STOP                As String           ' 停止      add SETkimizuka
    CAUSE               As String           ' 停止理由  add SETkimizuka
    PRINTNO             As String           ' 先行評価  add SETkimizuka
End Type

'---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi START ----
Public Type type_DBDRV_xsdc1
    CRYNUM As String * 12   ' 結晶番号
    HIKIAGEPTRN As String   ' 引上パターン
End Type
'---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi END ----

' 品番入力時
Public Function DBDRV_scmzc_fcmgc001f_Hinban(HINBAN As String, _
                                             Zyouken As type_DBDRV_scmzc_fcmgc001f_Hinban) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim fullHinban As tFullHinban

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Hinban"

    DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_SUCCESS

    '12桁品番を求める
    If GetLastHinban(HINBAN, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '検索条件を取得
    sql = "select "
    sql = sql & " T.MAGTYPE, "          ' 磁場タイプ
    sql = sql & " V.E018HSXD1MIN, "     ' 品ＳＸ直径１下限
    sql = sql & " V.E018HSXD1MAX, "     ' 品ＳＸ直径１上限
    sql = sql & " V.E018HSXTYPE, "      ' 品ＳＸタイプ
    sql = sql & " V.E018HSXDOP, "       ' 品ＳＸドーパント
    sql = sql & " V.E018HSXCDIR, "      ' 品ＳＸ結晶面方位
    sql = sql & " V.E018HSXRMIN, "      ' 品ＳＸ比抵抗下限
    sql = sql & " V.E018HSXRMAX, "      ' 品ＳＸ比抵抗上限
    sql = sql & " V.E018HSXDPDIR, "     ' 品ＳＸ溝位置方位
    sql = sql & " V.E019HSXONMIN, "     ' 品ＳＸ酸素濃度下限
    sql = sql & " V.E019HSXONMAX, "     ' 品ＳＸ酸素濃度上限
    sql = sql & " V.E019HSXCNMIN, "     ' 品ＳＸ炭素濃度下限
    sql = sql & " V.E019HSXCNMAX, "     ' 品ＳＸ炭素濃度上限
    sql = sql & " V.E019HSXLTMIN, "     ' 品ＳＸＬタイム下限
    sql = sql & " V.E019HSXLTMAX, "     ' 品ＳＸＬタイム上限
    sql = sql & " V.E020HSXDVDMXN, "     ' 品ＳＸＤＶＤ２上限   ＷＦサンプル処理変更 2003.05.20 yakimura
    sql = sql & " V.E020HSXDVDMNN, "     ' 品ＳＸＤＶＤ２下限   ＷＦサンプル処理変更 2003.05.20 yakimura
    sql = sql & " V.E020HSXOF1AX, "     ' 品ＳＸＯＳＦ１平均上限
    sql = sql & " V.E020HSXOF1MX, "     ' 品ＳＸＯＳＦ１上限
    sql = sql & " V.E020HSXOF2AX, "     ' 品ＳＸＯＳＦ２平均上限
    sql = sql & " V.E020HSXOF2MX, "     ' 品ＳＸＯＳＦ２上限
    sql = sql & " V.E020HSXOF3AX, "     ' 品ＳＸＯＳＦ３平均上限
    sql = sql & " V.E020HSXOF3MX, "     ' 品ＳＸＯＳＦ３上限
    sql = sql & " V.E020HSXOF4AX, "     ' 品ＳＸＯＳＦ４平均上限
    sql = sql & " V.E020HSXOF4MX, "     ' 品ＳＸＯＳＦ４上限
    sql = sql & " V.E020HSXBM1AN, "     ' 品ＳＸＢＭＤ１平均下限
    sql = sql & " V.E020HSXBM1AX, "     ' 品ＳＸＢＭＤ１平均上限
    sql = sql & " V.E020HSXBM2AN, "     ' 品ＳＸＢＭＤ２平均下限
    sql = sql & " V.E020HSXBM2AX, "     ' 品ＳＸＢＭＤ２平均上限
    sql = sql & " V.E020HSXBM3AN, "     ' 品ＳＸＢＭＤ３平均下限
    sql = sql & " V.E020HSXBM3AX  "     ' 品ＳＸＢＭＤ３平均上限
    sql = sql & " from VECME001 V, TBCMB012 T, TBCME018 S"
    With fullHinban
        sql = sql & " where V.E018HINBAN='" & .HINBAN & "' "
        sql = sql & " and V.E018MNOREVNO=" & .mnorevno & " "
        sql = sql & " and V.E018FACTORY='" & .factory & "' "
        sql = sql & " and V.E018OPECOND='" & .opecond & "' "
        sql = sql & " and S.HINBAN='" & .HINBAN & "' "
        sql = sql & " and S.MNOREVNO=" & .mnorevno & " "
        sql = sql & " and S.FACTORY='" & .factory & "' "
        sql = sql & " and S.OPECOND='" & .opecond & "' "
        sql = sql & " and trim(T.MKCONDNO)=trim(S.MCNO) "
    End With
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_scmzc_fcmgc001f_Hinban = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    'NULL対応 ----- START ----- 2003/12/09
    With Zyouken
        .MAGTYPE = rs("MAGTYPE")        ' 磁場タイプ
        .HSXD1MIN = fncNullCheck(rs("E018HSXD1MIN"))  ' 品ＳＸ直径１下限
        .HSXD1MAX = fncNullCheck(rs("E018HSXD1MAX"))  ' 品ＳＸ直径１上限
        .HSXTYPE = rs("E018HSXTYPE")    ' 品ＳＸタイプ
        .HSXDOP = rs("E018HSXDOP")      ' 品ＳＸドーパント
        .HSXCDIR = rs("E018HSXCDIR")    ' 品ＳＸ結晶面方位
        .HSXRMIN = fncNullCheck(rs("E018HSXRMIN"))    ' 品ＳＸ比抵抗下限
        .HSXRMAX = fncNullCheck(rs("E018HSXRMAX"))    ' 品ＳＸ比抵抗上限
        .HSXDPDIR = rs("E018HSXDPDIR")  ' 品ＳＸ溝位置方位
        .HSXONMIN = fncNullCheck(rs("E019HSXONMIN"))  ' 品ＳＸ酸素濃度上限
        .HSXONMAX = fncNullCheck(rs("E019HSXONMAX"))  ' 品ＳＸ炭素濃度下限
        .HSXCNMIN = fncNullCheck(rs("E019HSXCNMIN"))  ' 品ＳＸ炭素濃度上限
        .HSXCNMAX = fncNullCheck(rs("E019HSXCNMAX"))  ' 品ＳＸ溝位置方位
        .HSXLTMIN = fncNullCheck(rs("E019HSXLTMIN"))  ' 品ＳＸＬタイム下限
        .HSXLTMAX = fncNullCheck(rs("E019HSXLTMAX"))  ' 品ＳＸＬタイム上限
        .HSXDVDMX = fncNullCheck(rs("E020HSXDVDMXN"))  ' 品ＳＸＤＶＤ２上限   ＷＦサンプル処理変更 2003.05.20 yakimura
        .HSXDVDMN = fncNullCheck(rs("E020HSXDVDMNN"))  ' 品ＳＸＤＶＤ２下限   ＷＦサンプル処理変更 2003.05.20 yakimura
        .HSXOS1AX = fncNullCheck(rs("E020HSXOF1AX"))  ' 品ＳＸＯＳＦ１平均上限
        .HSXOS1MX = fncNullCheck(rs("E020HSXOF1MX"))  ' 品ＳＸＯＳＦ１上限
        .HSXOS2AX = fncNullCheck(rs("E020HSXOF2AX"))  ' 品ＳＸＯＳＦ２平均上限
        .HSXOS2MX = fncNullCheck(rs("E020HSXOF2MX"))  ' 品ＳＸＯＳＦ２上限
        .HSXOS3AX = fncNullCheck(rs("E020HSXOF3AX"))  ' 品ＳＸＯＳＦ３平均上限
        .HSXOS3MX = fncNullCheck(rs("E020HSXOF3MX"))  ' 品ＳＸＯＳＦ３上限
        .HSXOS4AX = fncNullCheck(rs("E020HSXOF4AX"))  ' 品ＳＸＯＳＦ４平均上限
        .HSXOS4MX = fncNullCheck(rs("E020HSXOF4MX"))  ' 品ＳＸＯＳＦ４上限 HSXOS4AX→HSXOS4MXに変更 2003/12/09
        .HSXBM1AN = fncNullCheck(rs("E020HSXBM1AN"))  ' 品ＳＸＢＭＤ１平均下限
        .HSXBM1AX = fncNullCheck(rs("E020HSXBM1AX"))  ' 品ＳＸＢＭＤ１平均上限
        .HSXBM2AN = fncNullCheck(rs("E020HSXBM2AN"))  ' 品ＳＸＢＭＤ２平均下限
        .HSXBM2AX = fncNullCheck(rs("E020HSXBM2AX"))  ' 品ＳＸＢＭＤ２平均上限
        .HSXBM3AN = fncNullCheck(rs("E020HSXBM3AN"))  ' 品ＳＸＢＭＤ３平均下限
        .HSXBM3AX = fncNullCheck(rs("E020HSXBM3AX"))  ' 品ＳＸＢＭＤ３平均上限
    End With
    'NULL対応 -----  END  ----- 2003/12/09
    rs.Close

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

' 初期表示
Public Function DBDRV_scmzc_fcmgc001f_INITDISP(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim c0 As Integer
    Dim c1 As Integer
    Dim recCount As Integer
    Dim temp0 As String
    Dim CodeData As String
    Dim MaxRec As Integer
    Dim i As Integer
    Dim BlockIdBuf  As String
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_INITDISP"

    DBDRV_scmzc_fcmgc001f_INITDISP = FUNCTION_RETURN_SUCCESS

    '初期表示ブロックIDを取得
    'sql = "select distinct BLOCKID, CRYNUM, INGOTPOS, LENGTH, REALLEN, UPDDATE "    ' ブロックID
    'sql = sql & "from TBCME040 where "
    'sql = sql & "DELCLS = '0' and HOLDCLS = '0' "
    'sql = sql & "and RSTATCLS = 'G'"
    sql = " SELECT DISTINCT CRYNUMC2, INPOSC2, GNLC2, KDAYC2, XTALC2,"
    sql = sql & "H2.PGID, H2.TYPE, H2.DPNTCLS,"
    sql = sql & "HINBCA, C.PALTNUM, C.BDCODE, PUPTNC1, "
    sql = sql & "A.REPSMPLIDCS AS ATOP,  "
    sql = sql & "B.REPSMPLIDCS AS BBOT "
    sql = sql & ",HOLDBC2, HOLDCC2, HOLDKTC2 "
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
    ' 流動停止項目追加 add SETkimizuka Start  09/03/18
'    sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
'    sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
'    sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
'    sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
   ' 流動停止項目追加 add SETkimizuka End    09/03/18
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    sql = sql & "FROM XSDC2, TBCMH002 H2, XSDCA, XSDCS A, XSDCS B, TBCMG007 C, XSDC1 "
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
    sql = sql & "    ,XODY3,XODY4 Y4,KODA9  "
    '' 流動停止項目追加 add SETkimizuka Start  09/03/18
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' 流動停止項目追加 add SETkimizuka Start  09/03/18
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    sql = sql & "WHERE "
    sql = sql & "LIVKC2  <> '1' AND "
    sql = sql & "GNWKNTC2 = 'CB320' AND "
    sql = sql & "SUBSTR(CRYNUMC2,1,9) = H2.UPINDNO AND "
    sql = sql & "CRYNUMC2 = CRYNUMCA AND "
    sql = sql & "INPOSC2 = INPOSCA AND "
    sql = sql & "LIVKCA <> '1' AND "
    sql = sql & "CRYNUMC2 = A.CRYNUMCS AND "
    sql = sql & "CRYNUMC2 = B.CRYNUMCS AND "
    sql = sql & "A.SMPKBNCS = 'T' AND "
    sql = sql & "B.SMPKBNCS = 'B' AND "
    sql = sql & "CRYNUMC2 = C.CRYNUM AND "
    sql = sql & "C.TRANCNT=(SELECT MAX(TRANCNT) FROM TBCMG007 WHERE CRYNUM=C.CRYNUM) AND "
    sql = sql & "XTALC2 = XTALC1 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    'sql = sql & "   AND CRYNUMCA     = Y4.XTALNO(+) "            'add 09/03/18 SETkimizuka
    sql = sql & " AND CRYNUMCA = XTALNOY3(+) "
    sql = sql & " AND LIVKY3(+) = '0' "
    sql = sql & " AND LIVKY4(+) = '0' "
    sql = sql & " AND XTALNOY3 = XTALNOY4(+) "
    sql = sql & " AND RCNTY3 = RCNTY4(+) "
    sql = sql & " AND SYSCA9(+) = 'X' AND SHUCA9(+) = '30' AND CAUSEY4 = CODEA9(+) "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    
    Debug.Print sql
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount = 0 Then
        DBDRV_scmzc_fcmgc001f_INITDISP = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '流動停止項目追加に伴う修正 upd 09/04/23 Start SETkimizuka
    '抽出結果を格納する
    ReDim Res(0) As type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DataInit Res()
    TotalLength = 0
    For c0 = 1 To recCount
        If rs("CRYNUMC2") <> BlockIdBuf Then
            i = i + 1
            ReDim Preserve Res(i)
            DataInit2 Res(), i
            
            With Res(i)
                .BLOCKID = rs("CRYNUMC2")    '
                .LENGTH = rs("GNLC2")
                .CRYNUM = rs("XTALC2")
                .INGOTPOS = rs("INPOSC2")
                .UPDDATE = rs("KDAYC2")
                .TOPPOS = .INGOTPOS
                .BOTPOS = .INGOTPOS + .LENGTH
                TotalLength = TotalLength + .LENGTH
            
                .DPNTCLS = rs("DPNTCLS")
                .TYPE = rs("TYPE")
                .PGID = rs("PGID")
                
                .HINBAN = rs("HINBCA")
            
                .TOPSMPLNO = rs("ATOP")
                .BOTSMPLNO = rs("BBOT")
                .BDCODE = rs("BDCODE")
                .PALTNUM = rs("PALTNUM")
            
                .HIKIAGEPTRN = rs("PUPTNC1")
                
                If IsNull(rs("HOLDBC2")) = False Then .HLDTRCLS = rs("HOLDBC2")    '2006/03
                If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")  '2006/03
                If IsNull(rs("HOLDCC2")) = False Then .HLDCAUSE = rs("HOLDCC2")    '2006/03
                
                ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
                '.AGRSTATUS = rs("AGRSTATUS")
                '.STOP = rs("STOP")
                'If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                '    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                'End If
                If rs("STOP") <> "2" And rs("WKKTY4") = "CB320" Then
                    .AGRSTATUS = rs("AGRSTATUS")
                    .STOP = rs("STOP")
                    If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                        Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                    End If
                End If
                ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
                If Trim(rs("PRINTNO")) <> "" And InStr(Res(i).PRINTNO, rs("PRINTNO")) = 0 Then
                    Res(i).PRINTNO = Res(i).PRINTNO & rs("PRINTNO") & vbTab
                End If
                
                BlockIdBuf = rs("CRYNUMC2")
            End With
        Else
            ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
            'If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
            '    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
            'End If
            If rs("STOP") <> "2" And rs("WKKTY4") = "CB320" Then
                If Trim(Res(i).AGRSTATUS) = "" Or rs("AGRSTATUS") < Res(i).AGRSTATUS Then
                    Res(i).AGRSTATUS = rs("AGRSTATUS")
                    Res(i).STOP = rs("STOP")
                End If
                If Trim(rs("CAUSE")) <> "" And InStr(Res(i).CAUSE, rs("CAUSE")) = 0 Then
                    Res(i).CAUSE = Res(i).CAUSE & rs("CAUSE") & vbTab
                End If
            End If
            ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
            If Trim(rs("PRINTNO")) <> "" And InStr(Res(i).PRINTNO, rs("PRINTNO")) = 0 Then
                Res(i).PRINTNO = Res(i).PRINTNO & rs("PRINTNO") & vbTab
            End If
        End If
        
        rs.MoveNext
    Next
    rs.Close
'    '抽出結果を格納する
'    ReDim Res(recCount) As type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DataInit Res()
'    TotalLength = 0
'    For c0 = 1 To recCount
'        With Res(c0)
'            '.BLOCKID = rs("BLOCKID")    ' ブロックID
'            .BLOCKID = rs("CRYNUMC2")    '
'            '.LENGTH = rs("REALLEN")
'            .LENGTH = rs("GNLC2")
'            '.CRYNUM = rs("CRYNUM")
'            .CRYNUM = rs("XTALC2")
'            '.INGOTPOS = rs("INGOTPOS")
'            .INGOTPOS = rs("INPOSC2")
'            '.UPDDATE = rs("UPDDATE")
'            .UPDDATE = rs("KDAYC2")
'            .TOPPOS = .INGOTPOS
'            .BOTPOS = .INGOTPOS + .LENGTH
'            TotalLength = TotalLength + .LENGTH
'
'            .DPNTCLS = rs("DPNTCLS")
'            .TYPE = rs("TYPE")
'            .PGID = rs("PGID")
'
'            .hinban = rs("HINBCA")
'
'            .TOPSMPLNO = rs("ATOP")
'            .BOTSMPLNO = rs("BBOT")
'            .BDCODE = rs("BDCODE")
'            .PALTNUM = rs("PALTNUM")
'
'            .HIKIAGEPTRN = rs("PUPTNC1")
'
'            If IsNull(rs("HOLDBC2")) = False Then .HLDTRCLS = rs("HOLDBC2")    '2006/03
'            If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")  '2006/03
'            If IsNull(rs("HOLDCC2")) = False Then .HLDCAUSE = rs("HOLDCC2")    '2006/03
'        End With
'
'        rs.MoveNext
'    Next
'    rs.Close
    '流動停止項目追加に伴う修正 upd 09/04/23 End SETkimizuka
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

Public Sub DataInit(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)

    Dim c0 As Integer
    Dim c1 As Integer
    Dim recCount As Integer

    recCount = UBound(Res())
    For c0 = 1 To recCount
        With Res(c0)
            .BLOCKID = ""           ' ブロックID
            .LENGTH = -1            ' 長さ＝ブロック長
            '品番管理
            .HINBAN = ""            ' Top品番＝品番
            'クリスタルカタログ受入実績
            .PALTNUM = ""           ' パレット番号
            .BDCODE = ""            ' 不良理由コード＝格下区分
            '結晶情報
            .DIAMETER = -1          ' 直径
            .PGID = ""              ' ＰＧ－ＩＤ
            '結晶抵抗実績
            .TOPRES = -1            ' Topサンプル№に対応する検索対象値＝Top推定ρ
            .BOTRES = -1            ' Botサンプル№に対応する検索対象値＝Bot推定ρ
            'Oi実績
            .TOPOI = -1             ' Topサンプル№に対応する検索対象値＝TopOi
            .BOTOI = -1             ' Botサンプル№に対応する検索対象値＝BotOi
            'Cs実績
            .BOTCS = -1             ' Cs実測値＝Cs
            '研削加工実績
            .NCHPOS = ""            ' ノッチ位置
            '制作条件
            .MAGTYPE = ""           ' 磁場タイプ＝製造法
            '製品仕様SXLﾃﾞｰﾀ１
            .TYPE = ""              ' 品ＳＸタイプ＝タイプ
            .HSXCDIR = ""           ' 品ＳＸ結晶面方位＝方位
            ' 引上げ投入実績
            .DPNTCLS = ""           ' ドーパント種類＝ドーパント　結晶番号前7桁+"00"＝引上げ指示№
            'ブロック管理
            .TOPPOS = -1            ' 結晶内開始位置＝Top部位
            .BOTPOS = -1            ' 結晶内開始位置＋長さ＝Bot部位
            .UPDDATE = -1           ' 更新日付＝受入日付
            '結晶サンプル管理
            .TOPSMPLNO = -1         ' サンプル№
            .BOTSMPLNO = -1         ' サンプル№
            'GD実績
            .TOPDVD2 = -1           ' 測定結果 DVD2＝DVD2(Top)
            .BOTDVD2 = -1           ' 測定結果 DVD2＝DVD2(Bot)
            'OSF実績
            'HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
            'が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
            For c1 = 0 To 3
                .TOPOSF(c1) = -1    ' 計算結果 Max＝OSF(Top)
                .BOTOSF(c1) = -1    ' 計算結果 Max＝OSF(Bot)
            Next
            'BMD実績
            'HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
            'が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
            For c1 = 0 To 2
                .TOPBMD(c1) = -1    ' Max＝OSF(Top)
                .BOTBMD(c1) = -1    ' Max＝OSF(Bot)
            Next
            'ライフタイム
            .TOPLT = -1             ' 計算結果＝ライフタイム(Top)
            .BOTLT = -1             ' 計算結果＝ライフタイム(Bot)
        End With
    Next

End Sub

Public Sub DataInit2(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku, Initnum As Integer)

    Dim c1 As Integer

        With Res(Initnum)
            .BLOCKID = ""           ' ブロックID
            .LENGTH = -1            ' 長さ＝ブロック長
            '品番管理
            .HINBAN = ""            ' Top品番＝品番
            'クリスタルカタログ受入実績
            .PALTNUM = ""           ' パレット番号
            .BDCODE = ""            ' 不良理由コード＝格下区分
            '結晶情報
            .DIAMETER = -1          ' 直径
            .PGID = ""              ' ＰＧ－ＩＤ
            '結晶抵抗実績
            .TOPRES = -1            ' Topサンプル№に対応する検索対象値＝Top推定ρ
            .BOTRES = -1            ' Botサンプル№に対応する検索対象値＝Bot推定ρ
            'Oi実績
            .TOPOI = -1             ' Topサンプル№に対応する検索対象値＝TopOi
            .BOTOI = -1             ' Botサンプル№に対応する検索対象値＝BotOi
            'Cs実績
            .BOTCS = -1             ' Cs実測値＝Cs
            '研削加工実績
            .NCHPOS = ""            ' ノッチ位置
            '制作条件
            .MAGTYPE = ""           ' 磁場タイプ＝製造法
            '製品仕様SXLﾃﾞｰﾀ１
            .TYPE = ""              ' 品ＳＸタイプ＝タイプ
            .HSXCDIR = ""           ' 品ＳＸ結晶面方位＝方位
            ' 引上げ投入実績
            .DPNTCLS = ""           ' ドーパント種類＝ドーパント　結晶番号前7桁+"00"＝引上げ指示№
            'ブロック管理
            .TOPPOS = -1            ' 結晶内開始位置＝Top部位
            .BOTPOS = -1            ' 結晶内開始位置＋長さ＝Bot部位
            .UPDDATE = -1           ' 更新日付＝受入日付
            '結晶サンプル管理
            .TOPSMPLNO = -1         ' サンプル№
            .BOTSMPLNO = -1         ' サンプル№
            'GD実績
            .TOPDVD2 = -1           ' 測定結果 DVD2＝DVD2(Top)
            .BOTDVD2 = -1           ' 測定結果 DVD2＝DVD2(Bot)
            'OSF実績
            'HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
            'が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
            For c1 = 0 To 3
                .TOPOSF(c1) = -1    ' 計算結果 Max＝OSF(Top)
                .BOTOSF(c1) = -1    ' 計算結果 Max＝OSF(Bot)
            Next
            'BMD実績
            'HTPRC = 熱処理方法､KKSP = 結晶欠陥測定位置､KKSET = 結晶欠陥測定条件 + 選択ET代
            'が、同じ仕様を探し、その仕様の番号(OSF1とかOSF2)を求め、対応する場所へ格納する。
            For c1 = 0 To 2
                .TOPBMD(c1) = -1    ' Max＝OSF(Top)
                .BOTBMD(c1) = -1    ' Max＝OSF(Bot)
            Next
            'ライフタイム
            .TOPLT = -1             ' 計算結果＝ライフタイム(Top)
            .BOTLT = -1             ' 計算結果＝ライフタイム(Bot)
        End With

End Sub

' 品番管理
Public Sub GETTBCME041(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f1
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '結晶番号と結晶位置から品番を求める
    sql = "select blk.CRYNUM, blk.INGOTPOS, hin.HINBAN, hin.REVNUM, hin.FACTORY, hin.OPECOND "
    sql = sql & "from TBCME040 blk, TBCME041 hin "
    sql = sql & "where (blk.CRYNUM = hin.CRYNUM) "
    sql = sql & "and ((blk.INGOTPOS >= hin.INGOTPOS) and (blk.INGOTPOS < (hin.INGOTPOS + hin.LENGTH))) "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f1
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("CRYNUM")
            buf(c0).factory = rs("FACTORY")
            buf(c0).HINBAN = rs("HINBAN")
            buf(c0).INGOTPOS = rs("INGOTPOS")
            buf(c0).opecond = rs("OPECOND")
            buf(c0).REVNUM = rs("REVNUM")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) Then
                    Res(c0).HINBAN = buf(c1).HINBAN
                    Res(c0).REVNUM = buf(c1).REVNUM     ' 製品番号改訂番号
                    Res(c0).factory = buf(c1).factory   ' 工場
                    Res(c0).opecond = buf(c1).opecond   ' 操業条件
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).HINBAN = " "
                Res(c0).REVNUM = -1     ' 製品番号改訂番号
                Res(c0).factory = " "   ' 工場
                Res(c0).opecond = " "   ' 操業条件
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' クリスタルカタログ受入実績
Public Sub GETTBCMG007(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f2
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    'ブロックIDからパレット番号、不良理由コード(格下区分)を求める
    MaxRec = UBound(Res())
    For c0 = 1 To MaxRec
        sql = "select CRYNUM, PALTNUM, BDCODE from TBCMG007 G"
        sql = sql & " where TRANCNT=(select max(TRANCNT) from TBCMG007 where CRYNUM=G.CRYNUM)"
        sql = sql & " and crynum = '" & Res(c0).BLOCKID & "' "
        DoEvents
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCount = rs.RecordCount
        If recCount <> 0 Then
            'ReDim buf(recCount) As type_DBDRV_cmgc001f2
            'For c0 = 1 To recCount
            '    buf(c0).BDCODE = rs("BDCODE")
            '    buf(c0).BLOCKID = rs("CRYNUM")
            '    buf(c0).PALTNUM = rs("PALTNUM")
            '    rs.MoveNext
            'Next
            'rs.Close
            'MaxRec = UBound(Res())
            'For c0 = 1 To MaxRec
            '    For c1 = 1 To recCount
            '        If (Res(c0).BLOCKID = buf(c1).BLOCKID) Then
            '            Res(c0).BDCODE = buf(c1).BDCODE
            '            Res(c0).PALTNUM = buf(c1).PALTNUM
            '            OKFlag = True
            '            Exit For
            '        End If
            '    Next
            '    If Not OKFlag Then
            '        Res(c0).BDCODE = " "
            '        Res(c0).PALTNUM = " "
            '    End If
            'Next
            Res(c0).BDCODE = rs("BDCODE")
            Res(c0).PALTNUM = rs("PALTNUM")
        End If
    Next
    rs.Close
    On Error GoTo 0

End Sub

' 結晶情報
Public Sub GETTBCME037(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f3
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '直径、PG-IDを求める
    sql = "select CRYNUM, DIAMETER, PGID from TBCME037"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f3
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("CRYNUM")
            buf(c0).PGID = rs("PGID")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) Then
                    Res(c0).PGID = buf(c1).PGID
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).PGID = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' 制作条件
Public Sub GETTBCMB012(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim sMKCONDNO As String * 12         ' 製作条件No.
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '結晶番号から製作条件を求める
'        sql1 = "select trim(PRODCOND) from TBCME037 where "
'        sql1 = sql1 & "CRYNUM = '" & Res(c0).CRYNUM & "'"
'
'        '製作条件から磁場タイプ(製造法)を求める
'        sql = "select MAGTYPE from TBCMB012 where "
'        sql = sql & "trim(MKCONDNO) = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).MAGTYPE = " "
'        Else
'            Res(c0).MAGTYPE = rs("MAGTYPE")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' 結晶番号
'    MAGTYPE As String * 2   '磁場タイプ＝製造法
'End Type
    Dim buf() As type_DBDRV_cmgc001f5
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '結晶番号から製作条件を求める
    '製作条件から磁場タイプ(製造法)を求める
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     A.MAGTYPE "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMB012 A , TBCME037 B "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = B.CRYNUM AND"
    sql = sql & "     TRIM(A.MKCONDNO) = TRIM(B.PRODCOND) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                Res(c0).MAGTYPE = rs("MAGTYPE")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0
End Sub

' 製品仕様SXLﾃﾞｰﾀ１
Public Sub GETTBCME018(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f6
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '品番から品SXタイプ、品ＳＸ結晶面方位を求める
    sql = "select HINBAN, HSXTYPE, HSXCDIR from TBCME018"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f6
        For c0 = 1 To recCount
            buf(c0).HINBAN = rs("HINBAN")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).HINBAN = buf(c1).HINBAN) Then
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).TYPE = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' 引上げ投入実績
Public Sub GETTBCMH002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        'ドーパント種類(ドーパント)を求める
'        sql = "select DPNTCLS from TBCMH002 where "
'        sql = sql & "UPINDNO = '" & Left(Res(c0).CRYNUM, 7) & "00' "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).DPNTCLS = " "
'        Else
'            Res(c0).DPNTCLS = rs("DPNTCLS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    DPNTCLS As String * 7   'ドーパント種類＝ドーパント　結晶番号前7桁+"00"＝引上げ指示No.
'    CRYNUM As String * 12         ' 結晶番号
'End Type
    Dim buf() As type_DBDRV_cmgc001f7
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    'ドーパント種類(ドーパント)を求める
    sql = "select UPINDNO, DPNTCLS, TYPE from TBCMH002"
    'sql = sql & "UPINDNO = '" & Left(Res(c0).CRYNUM, 7) & "00' "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_cmgc001f7
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("UPINDNO")
            buf(c0).DPNTCLS = rs("DPNTCLS")
            buf(c0).TYPE = rs("TYPE") '2002/04/25 S.Sano
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
'2004.09.10 Y.K リチャージ指示No９桁対応
'                If (Left(Res(c0).CRYNUM, 8) & "0" = Trim(buf(c1).CRYNUM)) Then
                If (Left(Res(c0).CRYNUM, 9) = Trim(buf(c1).CRYNUM)) Then
                    Res(c0).DPNTCLS = buf(c1).DPNTCLS
                    Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
                    OKFlag = True
                    Exit For
                End If
'2004.09.10 Y.K リチャージ指示No９桁対応
'残量引きでも表示したい場合はロジックとなる（サンプル）
'''                If (Left(Res(c0).CRYNUM, 8) = Left(Trim(buf(c1).CRYNUM), 8)) Then
'''                    If (IsNumeric(Mid(Res(c0).CRYNUM, 9, 1)) = True) Then
'''                        If (Mid(Res(c0).CRYNUM, 9, 1) = Mid(Trim(buf(c1).CRYNUM), 9, 1)) Then
'''                            Res(c0).DPNTCLS = buf(c1).DPNTCLS
'''                            Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
'''                            OKFlag = True
'''                            Exit For
'''                        End If
'''                    Else
'''                        If ("A" = Mid(Trim(buf(c1).CRYNUM), 9, 1)) Then
'''                            Res(c0).DPNTCLS = buf(c1).DPNTCLS
'''                            Res(c0).TYPE = buf(c1).TYPE '2002/04/25 S.Sano
'''                            OKFlag = True
'''                            Exit For
'''                        End If
'''                    End If
'''                End If
            Next
            If Not OKFlag Then
                Res(c0).DIAMETER = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0

End Sub

' 結晶サンプル管理
Public Sub GETTBCME043(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        'サンプルNoを求める
'        sql = "select SMPLNO, SMPKBN from TBCME043 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and INGOTPOS = '" & Res(c0).INGOTPOS & "' "
'        sql = sql & "and ((SMPKBN = 'T') or "
'        sql = sql & "((SMPKBN = 'B') and "
'        sql = sql & "(CRYINDRS='3' or CRYINDOI='3' or CRYINDB1='3' or "
'        sql = sql & "CRYINDB2='3' or CRYINDB3='3' or CRYINDL1='3' or "
'        sql = sql & "CRYINDL2='3' or CRYINDL3='3' or CRYINDL4='3' or "
'        sql = sql & "CRYINDCS='3' or CRYINDGD='3' or CRYINDT='3' or CRYINDEP='3')))"
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPSMPLNO = -1
'            Res(c0).TOPSMPKBN = " "
'        Else
'            Res(c0).TOPSMPLNO = rs("SMPLNO")
'            Res(c0).TOPSMPKBN = rs("SMPKBN")
'        End If
'        rs.Close
'
'        'サンプルNoを求める
'        sql = "select SMPLNO, SMPKBN from TBCME043 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and INGOTPOS = '" & Res(c0).INGOTPOS + Res(c0).LENGTH & "' "
'        sql = sql & "and ((SMPKBN = 'B') or "
'        sql = sql & "((SMPKBN = 'T') and "
'        sql = sql & "(CRYINDRS='3' or CRYINDOI='3' or CRYINDB1='3' or "
'        sql = sql & "CRYINDB2='3' or CRYINDB3='3' or CRYINDL1='3' or "
'        sql = sql & "CRYINDL2='3' or CRYINDL3='3' or CRYINDL4='3' or "
'        sql = sql & "CRYINDCS='3' or CRYINDGD='3' or CRYINDT='3' or CRYINDEP='3')))"
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTSMPLNO = -1
'            Res(c0).BOTSMPKBN = " "
'        Else
'            Res(c0).BOTSMPLNO = rs("SMPLNO")
'            Res(c0).BOTSMPKBN = rs("SMPKBN")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    SMPLNO As Integer    'サンプルNo
'    CRYNUM As String * 12         ' 結晶番号
'    INGOTPOS As Integer         ' 開始位置
'    SMPKBN As String * 1       ' サンプル区分
'End Type
    Dim buf() As type_DBDRV_cmgc001f8
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next
    'サンプルNoを求める
'    sql = "select E040CRYNUM, E043INGOTPOS, E043SMPKBN, E043SMPLNO from VECME010 order by E040CRYNUM, E043INGOTPOS"
    sql = "select E040CRYNUM, E043INPOSCS, E043SMPKBNCS, E043REPSMPLIDCS from VECME010 order by E040CRYNUM, E043INPOSCS"

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
            Res(c0).TOPSMPLNO = rs("E043REPSMPLIDCS")
            Res(c0).TOPSMPKBN = rs("E043SMPKBNCS")
        ReDim buf(recCount) As type_DBDRV_cmgc001f8
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("E040CRYNUM")
'            buf(c0).INGOTPOS = rs("E043INGOTPOS")
'            buf(c0).SMPLNO = rs("E043SMPLNO")
'            buf(c0).SMPKBN = rs("E043SMPKBN")
            buf(c0).INGOTPOS = rs("E043INPOSCS")
            buf(c0).SMPLNO = rs("E043REPSMPLIDCS")
            buf(c0).SMPKBN = rs("E043SMPKBNCS")
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And (Res(c0).INGOTPOS = buf(c1).INGOTPOS) Then
                    Res(c0).TOPSMPLNO = buf(c1).SMPLNO
                    Res(c0).TOPSMPKBN = buf(c1).SMPKBN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).TOPSMPLNO = -1
                Res(c0).TOPSMPKBN = " "
            End If
            For c1 = 1 To recCount
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) Then
                    Res(c0).BOTSMPLNO = buf(c1).SMPLNO
                    Res(c0).BOTSMPKBN = buf(c1).SMPKBN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).BOTSMPLNO = -1
                Res(c0).BOTSMPKBN = " "
            End If
        Next
    End If
    rs.Close
    On Error GoTo 0
End Sub

' 結晶抵抗実績
Public Sub GETTBCMJ002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ002 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select JUDGDATA from TBCMJ002 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPRES = -1
'        Else
'            Res(c0).TOPRES = rs("JUDGDATA")
'        End If
'        rs.Close
'
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ002 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select JUDGDATA from TBCMJ002 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTRES = -1
'        Else
'            Res(c0).BOTRES = rs("JUDGDATA")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' 結晶番号
'    TRANCNT As Integer              ' 処理回数
'    INGOTPOS As Integer         ' 開始位置
'    SMPLNO As Integer    'サンプルNo
'    SMPKBN As String * 1       ' サンプル区分
'    JUDGDATA As Double        'TopサンプルNo.に対応する検索対象値
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    ''判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, JUDGDATA from TBCMJ002 RS "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=RS.CRYNUM and POSITION=RS.POSITION and SMPKBN=RS.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("JUDGDATA")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c0).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then

    '                iTRANCNT = buf(c0).TRANCNT
    '                Res(c0).TOPRES = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPRES = -1
    '        End If

    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then

    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTRES = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTRES = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDRSCS AS TRS,"
    sql = sql & "     CRYINDRSCS,"
    sql = sql & "     A.MEAS1 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ002 A"
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDRSCS = A.SMPLNO AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ002"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO )"
    sql = sql & " UNION  "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDRSCS AS TRS,"
    sql = sql & "     CRYINDRSCS,"
    sql = sql & "     A.MEAS1 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ002 A"
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDRSCS = A.SMPLNO AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ002"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") And rs("CRYINDRSCS") <> "3" Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPRES = rs("MEAS1")
                    Res(c0).TOPRESSMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTRES = rs("MEAS1")
                    Res(c0).BOTRESSMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' Oi実績
Public Sub GETTBCMJ003(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ003 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select JUDGDATA from TBCMJ003 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPOI = -1
'        Else
'            Res(c0).TOPOI = rs("JUDGDATA")
'        End If
'        rs.Close
'
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ003 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select JUDGDATA from TBCMJ003 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTOI = -1
'        Else
'            Res(c0).BOTOI = rs("JUDGDATA")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' 結晶番号
'    TRANCNT As Integer              ' 処理回数
'    INGOTPOS As Integer         ' 開始位置
'    SMPLNO As Integer    'サンプルNo
'    SMPKBN As String * 1       ' サンプル区分
'    JUDGDATA As Double        'TopサンプルNo.に対応する検索対象値
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    ''判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, JUDGDATA from TBCMJ003 OI "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=OI.CRYNUM and POSITION=OI.POSITION and SMPKBN=OI.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("JUDGDATA")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPOI = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPOI = -1
    '        End If
            
   '         iTRANCNT = 0
   '         For c1 = 1 To recCount
   '             If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
   '                 iTRANCNT = buf(c1).TRANCNT
   '                 Res(c0).BOTOI = buf(c1).JudgData
   '                 OKFlag = True
   '             End If
   '         Next
   '         If Not OKFlag Then
   '             Res(c0).BOTOI = -1
   '         End If
   '     Next
   ' End If
   ' rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDOICS AS TRS,"
    sql = sql & "     A.OIMEAS1"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ003 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDOICS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ003 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " AND A.TRANCOND = 0 "  'GFAのFTIR換算値取得異常対応 2011/02/28 SETsw kubota
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDOICS AS TRS,"
    sql = sql & "     A.OIMEAS1"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ003 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDOICS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ003 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " AND A.TRANCOND = 0 "  'GFAのFTIR換算値取得異常対応 2011/02/28 SETsw kubota
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    rs.MoveFirst
    recCount = UBound(Res)
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    'Res(c0).TOPOI = rs("OIMEAS1")
                    If IsNull(rs("OIMEAS1")) = False Then Res(c0).TOPOI = rs("OIMEAS1") Else Res(c0).TOPOI = -1   ' OI_NULL対応　2005/03/08 TUKU
                    Res(c0).TOPOISMP = rs("TRS")
                    Exit For
                Else
                    'Res(c0).BOTOI = rs("OIMEAS1")
                    If IsNull(rs("OIMEAS1")) = False Then Res(c0).BOTOI = rs("OIMEAS1") Else Res(c0).BOTOI = -1   ' OI_NULL対応　2005/03/08 TUKU
                    Res(c0).BOTOISMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' Cs実績
Public Sub GETTBCMJ004(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ004 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select CSMEAS from TBCMJ004 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
''        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTOI = -1
'        Else
'            Res(c0).BOTOI = rs("CSMEAS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' 結晶番号
'    TRANCNT As Integer              ' 処理回数
'    INGOTPOS As Integer         ' 開始位置
'    SMPLNO As Integer    'サンプルNo
'    SMPKBN As String * 1       ' サンプル区分
'    JUDGDATA As Double        'TopサンプルNo.に対応する検索対象値
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    ''判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CSMEAS from TBCMJ004 CS "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=CS.CRYNUM and POSITION=CS.POSITION and SMPKBN=CS.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CSMEAS")
    '        rs.MoveNext
    '    Next
    '    rs.Close
   '     MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTCS = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTCS = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT "
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDCSCS AS TRS,"
    sql = sql & "     A.CSMEAS"
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ004 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDCSCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ004 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                'Res(c0).BOTCS = rs("CSMEAS")
                If IsNull(rs("CSMEAS")) = False Then Res(c0).BOTCS = rs("CSMEAS") Else Res(c0).BOTCS = -1  ' OI_NULL対応　2005/03/08 TUKU
                Res(c0).BOTCSSMP = rs("TRS")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' OSF実績
Public Sub GETTBCMJ005(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim c1 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        For c1 = 0 To 3
'            '処理回数のもっとも大きい値を求める
'            sql1 = "select max(TRANCNT) from TBCMJ005 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '判定データを求める
'            sql = "select CALCMAX from TBCMJ005 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).TOPOSF(c1) = -1
'            Else
'                Res(c0).TOPOSF(c1) = rs("CALCMAX")
'            End If
'            rs.Close
'
'            '処理回数のもっとも大きい値を求める
'            sql1 = "select max(TRANCNT) from TBCMJ005 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '判定データを求める
'            sql = "select CALCMAX from TBCMJ005 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).BOTOSF(c1) = -1
'            Else
'                Res(c0).BOTOSF(c1) = rs("CALCMAX")
'            End If
'            rs.Close
'        Next
'    Next
'    On Error GoTo 0
'Public Type type_DBDRV_scmzc_fcmgc001f_Kensaku
'    CRYNUM As String * 12         ' 結晶番号
'    TRANCNT As Integer              ' 処理回数
'    INGOTPOS As Integer         ' 開始位置
'    SMPLNO As Integer    'サンプルNo
'    SMPKBN As String * 1       ' サンプル区分
'    JUDGDATA As Double        'TopサンプルNo.に対応する検索対象値
'End Type
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CALCMAX from TBCMJ005 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).TRANCOND = rs("TRANCOND")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CALCMAX")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c2 = 0 To 3
    '        For c0 = 1 To MaxRec
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c0).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                       (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                        
   '                     iTRANCNT = buf(c1).TRANCNT
   '                     Res(c0).TOPOSF(c2) = buf(c1).JudgData
   '                     OKFlag = True
   '                 End If
   '             Next
   '             If Not OKFlag Then
   '                 Res(c0).TOPOSF(c2) = -1
   '             End If
                
   '             iTRANCNT = 0
   '             For c1 = 1 To recCount
   '                 If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                        
    ''                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).BOTOSF(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTOSF(c2) = -1
    '            End If
    '        Next
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDL1CS AS TRS1,CRYSMPLIDL2CS AS TRS2,CRYSMPLIDL3CS AS TRS3,CRYSMPLIDL4CS AS TRS4,"
    sql = sql & "     A.CALCMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ005 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ005"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDL1CS OR A.SMPLNO = CRYSMPLIDL2CS OR A.SMPLNO = CRYSMPLIDL3CS OR A.SMPLNO = CRYSMPLIDL4CS) "
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDL1CS AS TRS1,CRYSMPLIDL2CS AS TRS2,CRYSMPLIDL3CS AS TRS3,CRYSMPLIDL4CS AS TRS4,"
    sql = sql & "     A.CALCMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ005 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ005"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDL1CS OR A.SMPLNO = CRYSMPLIDL2CS OR A.SMPLNO = CRYSMPLIDL3CS OR A.SMPLNO = CRYSMPLIDL4CS) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).TOPOSF(0) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).TOPOSF(1) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).TOPOSF(2) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(2) = rs("TRS3")
                        Case "4"
                            Res(c0).TOPOSF(3) = rs("CALCMAX")
                            Res(c0).TOPOSFSMP(3) = rs("TRS4")
                    End Select
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).BOTOSF(0) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).BOTOSF(1) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).BOTOSF(2) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(2) = rs("TRS3")
                        Case "4"
                            Res(c0).BOTOSF(3) = rs("CALCMAX")
                            Res(c0).BOTOSFSMP(3) = rs("TRS4")
                    End Select
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' GD実績
Public Sub GETTBCMJ006(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ006 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select MSRSDVD2 from TBCMJ006 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPDVD2 = -1
'        Else
'            Res(c0).TOPDVD2 = rs("MSRSDVD2")
'        End If
'        rs.Close
'
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ006 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select MSRSDVD2 from TBCMJ006 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTDVD2 = -1
'        Else
'            Res(c0).BOTDVD2 = rs("MSRSDVD2")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, MSRSDVD2 from TBCMJ006 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("MSRSDVD2")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPDVD2 = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPDVD2 = -1
    '        End If
            
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTDVD2 = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTDVD2 = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDGDCS AS TRS,"
    sql = sql & "     A.MSRSDVD2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ006 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDGDCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ006 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDGDCS AS TRS,"
    sql = sql & "     A.MSRSDVD2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ006 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDGDCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND"
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ006 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPDVD2 = rs("MSRSDVD2")
                    Res(c0).TOPDVD2SMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTDVD2 = rs("MSRSDVD2")
                    Res(c0).BOTDVD2SMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' ライフタイム
Public Sub GETTBCMJ007(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ007 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select CALCMEAS from TBCMJ007 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).TOPLT = -1
'        Else
'            Res(c0).TOPLT = rs("CALCMEAS")
'        End If
'        rs.Close
'
'        '処理回数のもっとも大きい値を求める
'        sql1 = "select max(TRANCNT) from TBCMJ007 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql1 = sql1 & "and TRANCOND = '0' "
'
'        '判定データを求める
'        sql = "select CALCMEAS from TBCMJ007 where "
'        sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'        sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'        sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'        sql = sql & "and TRANCOND = '0' "
'        sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'        DoEvents
'        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'        RecCount = rs.RecordCount
'        If RecCount = 0 Then
'            Res(c0).BOTLT = -1
'        Else
'            Res(c0).BOTLT = rs("CALCMEAS")
'        End If
'        rs.Close
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, CALCMEAS from TBCMJ007 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("CALCMEAS")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c0 = 1 To MaxRec
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                   (Res(c0).TOPSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).TOPSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).TOPLT = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).TOPLT = -1
    '        End If
            
    '        iTRANCNT = 0
    '        For c1 = 1 To recCount
    '            If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                   (buf(c1).TRANCNT > iTRANCNT) And _
                   (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                   (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                   (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                    
    '                iTRANCNT = buf(c1).TRANCNT
    '                Res(c0).BOTLT = buf(c1).JudgData
    '                OKFlag = True
    '            End If
    '        Next
    '        If Not OKFlag Then
    '            Res(c0).BOTLT = -1
    '        End If
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDTCS AS TRS,"
    sql = sql & "     A.CALCMEAS "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ007 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     CRYSMPLIDTCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND "
    sql = sql & " A.TRANCOND = '0' AND "
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ007 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO)"
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDTCS AS TRS,"
    sql = sql & "     A.CALCMEAS "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, XSDCS, TBCMJ007 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     CRYSMPLIDTCS = A.SMPLNO AND"
    sql = sql & " A.CRYNUM = XTALC2 AND "
    sql = sql & " A.TRANCOND = '0' AND "
    sql = sql & " A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ007 WHERE CRYNUM=A.CRYNUM AND SMPLNO=A.SMPLNO) "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Res(c0).TOPLT = rs("CALCMEAS")
                    Res(c0).TOPLTSMP = rs("TRS")
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Res(c0).BOTLT = rs("CALCMEAS")
                    Res(c0).BOTLTSMP = rs("TRS")
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' BMD実績
Public Sub GETTBCMJ008(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
'    Dim sql As String
'    Dim sql1 As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim RecCount As Integer
'    Dim MaxRec As Integer
'    Dim c0 As Integer
'    Dim c1 As Integer
'
'    'エラーハンドラの設定
'    On Error Resume Next
'
'    MaxRec = UBound(Res())
'    For c0 = 1 To MaxRec
'        For c1 = 0 To 2
'            '処理回数のもっとも大きい値を求める
'            sql1 = "select max(TRANCNT) from TBCMJ008 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '判定データを求める
'            sql = "select MEASMAX from TBCMJ008 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).TOPSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).TOPBMD(c1) = -1
'            Else
'                Res(c0).TOPBMD(c1) = rs("MEASMAX")
'            End If
'            rs.Close
'
'            '処理回数のもっとも大きい値を求める
'            sql1 = "select max(TRANCNT) from TBCMJ008 where CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql1 = sql1 & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql1 = sql1 & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql1 = sql1 & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'
'            '判定データを求める
'            sql = "select MEASMAX from TBCMJ008 where "
'            sql = sql & "CRYNUM = '" & Res(c0).CRYNUM & "' "
'            sql = sql & "and POSITION = " & Res(c0).INGOTPOS + Res(c0).LENGTH & " "
'            sql = sql & "and SMPKBN = '" & Res(c0).BOTSMPKBN & "' "
'            sql = sql & "and TRANCOND = '" & Trim(Str(c1 + 1)) & "' "
'            sql = sql & "and TRANCNT = any(" & sql1 & ") "
'
'            DoEvents
'            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'            RecCount = rs.RecordCount
'            If RecCount = 0 Then
'                Res(c0).BOTBMD(c1) = -1
'            Else
'                Res(c0).BOTBMD(c1) = rs("MEASMAX")
'            End If
'            rs.Close
'        Next
'    Next
'    On Error GoTo 0
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f9
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '判定データを求める
    'sql = "select CRYNUM, POSITION, SMPLNO, SMPKBN, TRANCNT, TRANCOND, MEASMAX from TBCMJ008 J "
    'sql = sql & "where TRANCNT=(select max(TRANCNT) from TBCMJ002 where CRYNUM=J.CRYNUM and POSITION=J.POSITION and SMPKBN=J.SMPKBN)"

    'DoEvents
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCount = rs.RecordCount
    'If recCount <> 0 Then
    '    ReDim buf(recCount) As type_DBDRV_cmgc001f9
    '    For c0 = 1 To recCount
    '        buf(c0).CRYNUM = rs("CRYNUM")
    '        buf(c0).INGOTPOS = rs("POSITION")
    '        buf(c0).TRANCNT = rs("TRANCNT")
    '        buf(c0).TRANCOND = rs("TRANCOND")
    '        buf(c0).SMPLNO = rs("SMPLNO")
    '        buf(c0).SMPKBN = rs("SMPKBN")
    '        buf(c0).JudgData = rs("MEASMAX")
    '        rs.MoveNext
    '    Next
    '    rs.Close
    '    MaxRec = UBound(Res())
    '    For c2 = 0 To 2
    '        For c0 = 1 To MaxRec
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
                        
    '                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).TOPBMD(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTBMD(c2) = -1
    '            End If
                
    '            iTRANCNT = 0
    '            For c1 = 1 To recCount
    '                If (Res(c0).CRYNUM = buf(c1).CRYNUM) And _
                       (buf(c1).TRANCNT > iTRANCNT) And _
                       (buf(c1).TRANCOND = Trim(Str(c2 + 1))) And _
                       (Res(c0).INGOTPOS + Res(c0).LENGTH = buf(c1).INGOTPOS) And _
                       (Res(c0).BOTSMPLNO = buf(c1).SMPLNO) And _
                       (Res(c0).BOTSMPKBN = buf(c1).SMPKBN) Then
    '
    '                    iTRANCNT = buf(c1).TRANCNT
    '                    Res(c0).BOTBMD(c2) = buf(c1).JudgData
    '                    OKFlag = True
    '                End If
    '            Next
    '            If Not OKFlag Then
    '                Res(c0).BOTBMD(c2) = -1
    '            End If
    '        Next
    '    Next
    'End If
    'rs.Close
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDB1CS AS TRS1,CRYSMPLIDB2CS AS TRS2,CRYSMPLIDB3CS AS TRS3,"
    sql = sql & "     A.MEASMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ008 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'T' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ008"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDB1CS OR A.SMPLNO = CRYSMPLIDB2CS OR A.SMPLNO = CRYSMPLIDB3CS )  "
    sql = sql & " UNION "
    sql = sql & " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     SMPKBNCS,"
    sql = sql & "     CRYSMPLIDB1CS AS TRS1,CRYSMPLIDB2CS AS TRS2,CRYSMPLIDB3CS AS TRS3,"
    sql = sql & "     A.MEASMAX, A.TRANCOND,A.TRANCNT"
    sql = sql & " FROM"
    sql = sql & "     XSDC2,"
    sql = sql & "     XSDCS,"
    sql = sql & "     TBCMJ008 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     CRYNUMC2 = CRYNUMCS AND"
    sql = sql & "     SMPKBNCS = 'B' AND"
    sql = sql & "     A.CRYNUM = XTALC2 AND"
    sql = sql & "     A.TRANCNT = (SELECT"
    sql = sql & "                     MAX(TRANCNT)"
    sql = sql & "                 FROM"
    sql = sql & "                     TBCMJ008"
    sql = sql & "                 WHERE"
    sql = sql & "                     CRYNUM=A.CRYNUM AND"
    sql = sql & "                     SMPLNO=A.SMPLNO  AND"
    sql = sql & "                     TRANCOND=A.TRANCOND ) AND"
    sql = sql & " (A.SMPLNO = CRYSMPLIDB1CS OR A.SMPLNO = CRYSMPLIDB2CS OR A.SMPLNO = CRYSMPLIDB3CS )  "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If rs("SMPKBNCS") = "T" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).TOPBMD(0) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).TOPBMD(1) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).TOPBMD(2) = rs("MEASMAX")
                            Res(c0).TOPBMDSMP(2) = rs("TRS3")
                    End Select
                    Exit For
                ElseIf rs("SMPKBNCS") = "B" Then
                    Select Case rs("TRANCOND")
                        Case "1"
                            Res(c0).BOTBMD(0) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(0) = rs("TRS1")
                        Case "2"
                            Res(c0).BOTBMD(1) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(1) = rs("TRS2")
                        Case "3"
                            Res(c0).BOTBMD(2) = rs("MEASMAX")
                            Res(c0).BOTBMDSMP(2) = rs("TRS3")
                    End Select
                    Exit For
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' 加工払出実績実績
Public Sub GETTBCMI001(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim iTRANCNT As Integer
    Dim buf() As type_DBDRV_cmgc001f10
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim MaxRecCode As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim c2 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
    On Error Resume Next
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     PRCMCN,"
    sql = sql & "     SEED "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMI001 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = CRYNUM AND"
    sql = sql & "     A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMI001 WHERE CRYNUM=A.CRYNUM)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                Res(c0).PRCMCN = rs("PRCMCN")
                Res(c0).SEED = rs("SEED")
                Res(c0).HSXCDIR = GetCodeField("SC", "28", Left(Res(c0).SEED, 1), "INFO2")
                Exit For
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub
' 研削加工実績
Public Sub GETTBCMI002(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)
    Dim buf() As type_DBDRV_cmgc001f4
    Dim sql As String
    Dim rs As OraDynaset    'RecordSet
    Dim recCount As Integer
    Dim MaxRec As Integer
    Dim c0 As Integer
    Dim c1 As Integer
    Dim OKFlag As Boolean

    'エラーハンドラの設定
'    On Error Resume Next

    'ノッチ位置を求める
    sql = " SELECT DISTINCT"
    sql = sql & "     CRYNUMC2,"
    sql = sql & "     INPOSC2,"
    sql = sql & "     XTALC2,"
    sql = sql & "     INGOTPOS,NCHPOS, DMTOP1, DMTOP2, DMTAIL1, DMTAIL2 "
    sql = sql & " FROM"
    sql = sql & "     XSDC2, TBCMI002 A "
    sql = sql & " WHERE"
    sql = sql & "     LIVKC2  <> '1' AND"
    sql = sql & "     GNWKNTC2 = 'CB320' AND"
    sql = sql & "     XTALC2 = CRYNUM AND"
    sql = sql & "     A.TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMI002 WHERE CRYNUM=A.CRYNUM)"
    sql = sql & "order by INGOTPOS "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = UBound(Res)
    rs.MoveFirst
    Do Until rs.EOF
        For c0 = 1 To recCount
            If Res(c0).BLOCKID = rs("CRYNUMC2") Then
                If Res(c0).PRCMCN = "A" Then
                            Res(c0).NCHPOS = rs("NCHPOS")
                            Res(c0).DMTOP1 = rs("DMTOP1")
                            Res(c0).DMTOP2 = rs("DMTOP2")
                            Res(c0).DMTAIL1 = rs("DMTAIL1")
                            Res(c0).DMTAIL2 = rs("DMTAIL2")
                            Res(c0).DIAMETER = (Res(c0).DMTOP1 + Res(c0).DMTOP2 + Res(c0).DMTAIL1 + Res(c0).DMTAIL2) / 4
                            Exit For
                ElseIf Res(c0).PRCMCN = "M" Then
                        If Res(c0).INGOTPOS >= rs("INGOTPOS") Then
                            Res(c0).NCHPOS = rs("NCHPOS")
                            Res(c0).DMTOP1 = rs("DMTOP1")
                            Res(c0).DMTOP2 = rs("DMTOP2")
                            Res(c0).DMTAIL1 = rs("DMTAIL1")
                            Res(c0).DMTAIL2 = rs("DMTAIL2")
                            Res(c0).DIAMETER = (Res(c0).DMTOP1 + Res(c0).DMTOP2 + Res(c0).DMTAIL1 + Res(c0).DMTAIL2) / 4
                            Exit For
                        End If
                End If
            End If
        Next
        rs.MoveNext
    Loop
    rs.Close
    On Error GoTo 0

End Sub

' 廃棄処理を行う
Public Function DBDRV_scmzc_fcmgc001f_Haiki(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim BlockMng As typ_TBCME040
    Dim sql As String

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Haiki"

    DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_FAILURE
    
    ' kuramoto変更
    'ブロック管理を更新する
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KAKUAGE & "', "             ' 現在管理工程
    sql = sql & "NOWPROC='" & PROCD_KAKUAGE & "', "               ' 現在工程
    sql = sql & "DELCLS='1', "
    sql = sql & "LSTATCLS = 'H', "
    sql = sql & "RSTATCLS = 'T', "
    sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
    sql = sql & "SENDFLAG='0' "                         ' 送信フラグ
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'ブロック管理を更新する
'    With BlockMng
'        .CryNum = rec.CryNum                    ' 結晶番号
'        .INGOTPOS = rec.INGOTPOS                ' 結晶内開始位置
'        .LENGTH = rec.LENGTH                    ' 長さ
'        .REALLEN = rec.LENGTH                   ' 実長さ
'        .BLOCKID = rec.BLOCKID                  ' ブロックID
'        .KRPROCCD = MGPRCD_KAKUAGE              ' 現在管理工程
'        .NOWPROC = PROCD_KAKUAGE                ' 現在工程
'        '.LPKRPROCCD = MGPRCD_KAKUAGE            ' 最終通過管理工程  --- 最終通過工程は、Gに落とした工程を残す
'        '.LASTPASS = PROCD_KAKUAGE               ' 最終通過工程      --- 最終通過工程は、Gに落とした工程を残す
'        .DELCLS = "1"                           ' 削除区分
'        .LSTATCLS = "H"                         ' 最終状態区分
'        .RSTATCLS = "T"                         ' 流動状態区分
'        .HOLDCLS = "0"                          ' ホールド区分
'        .BDCAUS = "   "                         ' 不良理由
'    End With
'
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        Exit Function
'    End If

    DBDRV_scmzc_fcmgc001f_Haiki = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

' リメルト処理を行う
Public Function DBDRV_scmzc_fcmgc001f_Remerto(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim BlockMng As typ_TBCME040
    Dim HIN As tFullHinban
    Dim sql As String
    Dim sWhere      As String       'ADD 2004/10/22 TCS)R.Kawaguchi
    Dim rec_xodc2() As typ_XSDC2    'ADD 2004/10/22 TCS)R.Kawaguchi
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Remerto"
    
    DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_FAILURE

    ' kuramoto変更
    'ブロック管理を更新する
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_RIMERUTO_UKEIRE & "', "     ' 現在管理工程
    sql = sql & "NOWPROC='" & PROCD_RIMERUTO_UKEIRE & "', "       ' 現在工程
    sql = sql & "LSTATCLS = 'T', "
    sql = sql & "RSTATCLS = 'M', "
    sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
    sql = sql & "SENDFLAG='0' "                         ' 送信フラグ
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'ブロック管理を更新する
'    With BlockMng
'        .CryNum = rec.CryNum                    ' 結晶番号
'        .INGOTPOS = rec.INGOTPOS                ' 結晶内開始位置
'        .LENGTH = rec.LENGTH                    ' 長さ
'        .REALLEN = rec.LENGTH                   ' 実長さ
'        .BLOCKID = rec.BLOCKID                  ' ブロックID
'        .KRPROCCD = MGPRCD_RIMERUTO_UKEIRE      ' 現在管理工程
'        .NOWPROC = PROCD_RIMERUTO_UKEIRE        ' 現在工程
'        '.LPKRPROCCD = MGPRCD_KAKUAGE            ' 最終通過管理工程  --- 最終通過工程は、Gに落とした工程を残す
'        '.LASTPASS = PROCD_KAKUAGE               ' 最終通過工程      --- 最終通過工程は、Gに落とした工程を残す
'        .DELCLS = "0"                           ' 削除区分
'        .LSTATCLS = "T"                         ' 最終状態区分
'        .RSTATCLS = "M"                         ' 流動状態区分
'        .HOLDCLS = "0"                          ' ホールド区分
'        .BDCAUS = "   "                         ' 不良理由
'    End With
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        Exit Function
'    End If
    
    '品番を'Z'に変える
    With rec
        HIN.HINBAN = "Z"
        HIN.mnorevno = 0
        HIN.factory = "Y"
        HIN.opecond = "1"
        If ChangeAreaHinban(.CRYNUM, .INGOTPOS, .LENGTH, HIN) = FUNCTION_RETURN_FAILURE Then
            Exit Function
        End If
    End With

'---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi START ----
    ''ブロック管理の情報(重量、結晶内開始位置)を取得
    'WHERE条件式
    sWhere = "WHERE CRYNUMC2 = '" & rec.BLOCKID & "'"
    Call DBDRV_GetXSDC2(rec_xodc2(), sWhere)
    '該当データ無しの場合
    If UBound(rec_xodc2) = 0 Then
        GoTo proc_exit
    End If

    ''精製原料管理(XODCX)作成処理
    If InsXODCX(rec, rec_xodc2(1)) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    ''原料工程実績(XODB3)作成処理
    If InsXODB3(rec, rec_xodc2(1)) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    '*** UPDATE START T.TERAUCHI 2004/12/06 精製原料ﾗﾍﾞﾙ発行用処理追加対応
        '精製原料ﾗﾍﾞﾙ発行処理
        
    '*** UPDATE START T.TERAUCHI 2005/01/18 精製原料ﾗﾍﾞﾙ発行用工程ｺｰﾄﾞ変更対応
    '   If Ins_TBCMC001_New("RP10", "81", f_cmbc008_1.txtStaffID.Text, rec.BLOCKID & "0", gsSysdate) = False Then
        If Ins_TBCMC001_New(Right(PROCD_KAKUAGE, 4), "81", f_cmbc008_1.txtStaffID.Text, rec.BLOCKID & "0", gsSysdate) = False Then
    '*** UPDATE END   T.TERAUCHI 2005/01/18
                
            GoTo proc_exit
        End If
    '*** UPDATE END   T.TERAUCHI 2004/12/06
    
'---- ADD [精製原料システム対応] 2004/10/22 TCS)R.Kawaguchi END ----

    DBDRV_scmzc_fcmgc001f_Remerto = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'ブロックの理論長さを得る
Private Function GetBlockLength(blkID$) As Integer
Dim sql$
Dim rs As OraDynaset

    sql = "select LENGTH from TBCME040 where BLOCKID='" & blkID & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
    If rs.RecordCount = 0 Then
        GetBlockLength = 0
    Else
        GetBlockLength = rs("LENGTH")
    End If
    rs.Close
    Set rs = Nothing
End Function

' 格上げ処理を行う
Public Function DBDRV_scmzc_fcmgc001f_Kakuage(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    Dim BlockMng As typ_TBCME040
    Dim CC() As typ_TBCMG008

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_scmzc_fcmgc001f_Kakuage"
    DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_FAILURE

    '12桁品番を求める
    If GetLastHinban(NewHinBan, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    'XSDCSを更新する
    If ChangeXSDCSHinban(rec.BLOCKID, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    '品番管理を更新する
    If ChangeAreaHinban(rec.CRYNUM, rec.INGOTPOS, GetBlockLength(rec.BLOCKID), fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    
    ' kuramoto変更
    'ブロック管理を更新する
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "   ' 現在管理工程
    sql = sql & "NOWPROC='" & PROCD_KESSYOU_SOUGOUHANTEI & "', "     ' 現在工程
    sql = sql & "LPKRPROCCD='" & MGPRCD_KAKUAGE & "', "              ' 最終通過管理工程
    sql = sql & "LASTPASS='" & PROCD_KAKUAGE & "', "                 ' 最終通過工程
    sql = sql & "LSTATCLS = 'T', "
    sql = sql & "RSTATCLS = 'T', "
    sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
    sql = sql & "SUMMITSENDFLAG='0', "                  ' SUMMIT送信フラグ
    sql = sql & "SENDFLAG='0' "                         ' 送信フラグ
    sql = sql & "where "
    sql = sql & "CRYNUM='" & rec.CRYNUM & "' "
    sql = sql & "and INGOTPOS=" & rec.INGOTPOS & " "
    
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    'ブロック管理を更新する
'    With BlockMng
'        .CryNum = rec.CryNum                    ' 結晶番号
'        .INGOTPOS = rec.INGOTPOS                ' 結晶内開始位置
'        .LENGTH = rec.LENGTH                    ' 長さ
'        .REALLEN = rec.LENGTH                   ' 実長さ
'        .BLOCKID = rec.BLOCKID                  ' ブロックID
'        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI ' 現在管理工程
'        .NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI   ' 現在工程
'        .LPKRPROCCD = MGPRCD_KAKUAGE            ' 最終通過管理工程
'        .LASTPASS = PROCD_KAKUAGE               ' 最終通過工程
'        .DELCLS = "0"                           ' 削除区分
'        .LSTATCLS = "T"                         ' 最終状態区分
'        .RSTATCLS = "T"                         ' 流動状態区分
'        .HOLDCLS = "0"                          ' ホールド区分
'        .BDCAUS = "   "                         ' 不良理由
'    End With
'    If DBDRV_BlockMng_Upd(BlockMng) = FUNCTION_RETURN_FAILURE Then
'        GoTo proc_exit
'    End If

    If DBDRV_PutTBCMG008(rec, fullHinban) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmgc001f_Kakuage = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

Public Function DBDRV_PutTBCMG008(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, fullHinban As tFullHinban) As FUNCTION_RETURN

    Dim rs As OraDynaset    'RecordSet
    Dim sql As String
    Dim CC() As typ_TBCMG008
    Dim InsertFlag As Boolean

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function DBDRV_PutTBCMG008"
    DBDRV_PutTBCMG008 = FUNCTION_RETURN_FAILURE

    '処理回数のもっとも大きい値を求める
    sql = "where "
    sql = sql & "CRYNUM = '" & rec.BLOCKID & "' "
    sql = sql & "and TRANCNT = any("
    sql = sql & "select max(TRANCNT) from TBCMG008 where CRYNUM = '" & rec.BLOCKID & "'"
    sql = sql & ") "
    If DBDRV_GetTBCMG008(CC(), sql) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If

    InsertFlag = False
    If UBound(CC()) = 0 Then
        ReDim CC(1) As typ_TBCMG008
        InsertFlag = True
    End If
    With CC(1)
        .CRYNUM = rec.BLOCKID
        .KRPROCCD = MGPRCD_KAKUAGE          ' 管理工程コード
        .PROCCODE = PROCD_KAKUAGE           ' 工程コード
'        .OHINBAN = rec.hinban              ' 旧品番
'        .OMNOREVNO = rec.REVNUM            ' 旧製品番号改訂番号
'        .OFACTORY = rec.factory            ' 旧工場
'        .OOPECOND = rec.opecond            ' 旧操業条件
        
        .OHINBAN = "G       "               ' 旧品番
        .OMNOREVNO = 0                      ' 旧製品番号改訂番号
        .OFACTORY = " "                     ' 旧工場
        .OOPECOND = " "                     ' 旧操業条件
        
        .NHINBAN = fullHinban.HINBAN        ' 新品番
        .NMNOREVNO = fullHinban.mnorevno    ' 新製品番号改訂番号
        .NFACTORY = fullHinban.factory      ' 新工場
        .NOPECOND = fullHinban.opecond      ' 新操業条件
    End With

    sql = " insert into TBCMG008 ( "
    sql = sql & "CRYNUM, "      ' 結晶番号
    sql = sql & "KRPROCCD, "    ' 管理工程コード
    sql = sql & "PROCCODE, "    ' 工程コード
    sql = sql & "NHINBAN, "     ' 新品番
    sql = sql & "NMNOREVNO, "   ' 新製品番号改訂番号
    sql = sql & "NFACTORY, "    ' 新工場
    sql = sql & "NOPECOND, "    ' 新操業条件
    sql = sql & "OHINBAN, "     ' 旧品番
    sql = sql & "OMNOREVNO, "   ' 旧製品番号改訂番号
    sql = sql & "OFACTORY, "    ' 旧工場
    sql = sql & "OOPECOND, "    ' 旧操業条件
    sql = sql & "KSTAFFID, "    ' 更新社員ＩＤ
    sql = sql & "UPDDATE, "     ' 更新日付
    sql = sql & "SENDFLAG, "    ' 送信フラグ
    sql = sql & "SENDDATE,"      ' 送信日付
    sql = sql & "TRANCNT, "     ' 処理回数
    sql = sql & "TSTAFFID, "    ' 登録社員ID
    sql = sql & "REGDATE "     ' 登録日付
    sql = sql & ")"
    With CC(1)
        sql = sql & " values ( "
        sql = sql & "'" & .CRYNUM & "'," ' 結晶番号
        sql = sql & "'" & .KRPROCCD & "'," ' 管理工程コード
        sql = sql & "'" & .PROCCODE & "'," ' 工程コード
        sql = sql & "'" & .NHINBAN & "'," ' 新品番
        sql = sql & .NMNOREVNO & ","  ' 新製品番号改訂番号
        sql = sql & "'" & .NFACTORY & "'," ' 新工場
        sql = sql & "'" & .NOPECOND & "'," ' 新操業条件
        sql = sql & "'" & .OHINBAN & "'," ' 旧品番
        sql = sql & .OMNOREVNO & "," ' 旧製品番号改訂番号
        sql = sql & "'" & .OFACTORY & "'," ' 旧工場
        sql = sql & "'" & .OOPECOND & "'," ' 旧操業条件
        sql = sql & "'" & STAFFIDBUFF & "'," ' 更新社員ＩＤ
        sql = sql & "sysdate," ' 更新日付
        sql = sql & "'0'," ' 送信フラグ
        sql = sql & "sysdate," ' 送信日付
        If InsertFlag Then
            sql = sql & "1," ' 処理回数
            sql = sql & "'" & STAFFIDBUFF & "'," ' 登録社員ID
            sql = sql & "sysdate" ' 登録日付
        Else
            sql = sql & .TRANCNT + 1 & "," ' 処理回数
            sql = sql & "'" & .TSTAFFID & "',"  ' 登録社員ID
            sql = sql & "sysdate"  ' 登録日付
        End If
        sql = sql & ")"
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        GoTo proc_exit
    End If

    DBDRV_PutTBCMG008 = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMG008」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMG008 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMG008_SQL.basより移動)
Public Function DBDRV_GetTBCMG008(records() As typ_TBCMG008, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, NHINBAN, NMNOREVNO, NFACTORY, NOPECOND, OHINBAN, OMNOREVNO, OFACTORY," & _
              " OOPECOND, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMG008"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMG008 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号（格上げ）
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .NHINBAN = rs("NHINBAN")         ' 新品番
            .NMNOREVNO = rs("NMNOREVNO")     ' 新製品番号改訂番号
            .NFACTORY = rs("NFACTORY")       ' 新工場
            .NOPECOND = rs("NOPECOND")       ' 新操業条件
            .OHINBAN = rs("OHINBAN")         ' 旧品番
            .OMNOREVNO = rs("OMNOREVNO")     ' 旧製品番号改訂番号
            .OFACTORY = rs("OFACTORY")       ' 旧工場
            .OOPECOND = rs("OOPECOND")       ' 旧操業条件
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

    DBDRV_GetTBCMG008 = FUNCTION_RETURN_SUCCESS
End Function


'---- ADD [精製原料管理作成対応] 2004/10/22 TCS)R.Kawaguchi START ----

' @(f)
' 機能      : XSDC1検索関数
'
' 返り値    : なし
'
' 引き数    : 検索結果格納構造体
'
' 機能説明  : 各結晶番号の引上げパターンを取得する
Public Sub GETXSDC1(Res() As type_DBDRV_scmzc_fcmgc001f_Kensaku)

    Dim buf()       As type_DBDRV_xsdc1
    Dim sql         As String
    Dim rs          As OraDynaset    'RecordSet
    Dim recCount    As Integer
    Dim MaxRec      As Integer
    Dim c0          As Integer
    Dim c1          As Integer
    Dim OKFlag      As Boolean

    'エラーハンドラの設定
    On Error Resume Next

    '判定データを求める
    sql = ""
    sql = "select XTALC1, PUPTNC1 from XSDC1 "

    DoEvents
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCount = rs.RecordCount
    If recCount <> 0 Then
        ReDim buf(recCount) As type_DBDRV_xsdc1
        For c0 = 1 To recCount
            buf(c0).CRYNUM = rs("XTALC1")
            buf(c0).HIKIAGEPTRN = NulltoStr(rs("PUPTNC1"))
            rs.MoveNext
        Next
        rs.Close
        MaxRec = UBound(Res())
        For c0 = 1 To MaxRec
            For c1 = 1 To recCount
                
            '*** UPDATE START T.TERAUCHI 2004/12/06 条件をブロックIDから結晶番号に変更
            '    If (Res(c0).BLOCKID = buf(c1).CRYNUM) Then
                If (Res(c0).CRYNUM = buf(c1).CRYNUM) Then
            '*** UPDATE END   T.TERAUCHI 2004/12/06
                     
                    Res(c0).HIKIAGEPTRN = buf(c1).HIKIAGEPTRN
                    OKFlag = True
                    Exit For
                End If
            Next
            If Not OKFlag Then
                Res(c0).HIKIAGEPTRN = " "
            End If
            OKFlag = False
        Next

    End If
    On Error GoTo 0

End Sub

' @(f)
' 機能      : XODCX作成関数
'
' 返り値    : FUNCTION_RETURN_FAILURE：異常
'             FUNCTION_RETURN_SUCCESS：正常
'
' 引き数    : 検索結果格納構造体
'
' 機能説明  : 該当データの精製原料管理を作成する
Private Function InsXODCX(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, _
                            rec_xodc2 As typ_XSDC2) As FUNCTION_RETURN

    Dim objDS       As Object
    Dim sSql        As String
    Dim sDopType    As String
    Dim sCSDop      As String       'CSドープ有無
    Dim sNDop       As String       '窒素ドープ有無
    Dim sWhere      As String
    Dim sUserID     As String
    Dim sSCNTRL     As String       '識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ ADD 2011/03/24 TSMC品識別対応
    
'*** UPDATE START T.TERAUCHI 2004/12/07 ﾗｲﾌﾀｲﾑ仕様有無
    Dim sLTUmu      As String
'*** UPDATE END   T.TERAUCHI 2004/12/07
'*** UPDATE START TAGAWA 2004/12/16
    Dim sFlag       As String
'*** UPDATE END  TAGAWA 2004/12/16

On Error GoTo PROC_ERR

    InsXODCX = FUNCTION_RETURN_FAILURE
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function InsXODCX"
    
    '登録社員ＩＤ
    sUserID = f_cmbc008_1.txtStaffID.Text
    
    With rec
        
        '///精製原料基本情報取得SQL作成
        Call GetAssistSQL_300(sSql, .CRYNUM)
        If DynSet2(objDS, sSql) = False Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If
        '該当データ無しの場合
        If objDS.EOF = True Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If

        '///CSドープ有無、窒素ドープ有無設定
'*** UPDATE START T.TERAUCHI 2004/12/07 小文字→大文字変換対応
'        sDopType = NulltoStr(objDS.Fields("DTYPEC1").Value)
'        If sDopType = " " Or sDopType = "p" Then
'            sCSDop = "2"
'            sNDop = "1"
'        ElseIf sDopType = "n" Then
'            sCSDop = "1"
'            sNDop = "2"
'        Else
'            sCSDop = " "
'            sNDop = " "
'        End If
        sDopType = UCase(NulltoStr(objDS.Fields("DTYPEC1").Value))
'*** UPDATE START TAGAWA 2004/12/16**************************
''        If sDopType = " " Or sDopType = "P" Then
''            sCSDop = "2"
''            sNDop = "1"
''        ElseIf sDopType = "N" Then
''            sCSDop = "1"
''            sNDop = "2"
''        Else
''            sCSDop = " "
''            sNDop = " "
''        End If
        ''結晶ドープ取得
        sFlag = UCase(Trim(NulltoStr(objDS.Fields("DPNTCLS").Value)))
        ''Csドープの時
        If sFlag = "C" Then
            sCSDop = "2"
            sNDop = "1"
        ''窒素ドープの時
        ElseIf sFlag = "N" Then
            sCSDop = "1"
            sNDop = "2"
        ''Ｗドープの時
        ElseIf sFlag = "M" Then
            sCSDop = "2"
            sNDop = "2"
        ''その他
        Else
            sCSDop = "1"
            sNDop = "1"
        End If
'*** UPDATE END TAGAWA 2004/12/16**************************
'*** UPDATE END   T.TERAUCHI 2004/12/07

    '*** UPDATE START T.TERAUCHI 2004/12/07 ﾗｲﾌﾀｲﾑ仕様有無判別
        ''ライフタイム仕様有無
        If objDS.Fields("HSXLTHWS").Value = "H" Then
            ''有り
            sLTUmu = "2"
        Else
            ''無し
            sLTUmu = "1"
        End If
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
        ''精製原料チェックフラグの判断
        If NulltoStr(objDS.Fields("MTRLCHKFLG").Value) = "1" Then
            ''品番NULL時の識別コントロールコードセット(空白3桁)
            If NulltoStr(objDS.Fields("HINBCX").Value) = "" Then
                sSCNTRL = "   "
            Else
                ''識別コントロールコードセット(品番3桁)
                sSCNTRL = Left(objDS.Fields("HINBCX").Value, 3)
            End If
        Else
            ''識別コントロールコードセット(空白3桁)
            sSCNTRL = "   "
        End If
    '*** UPDATE END   Marushita 2011/03/24
        '///精製原料管理作成
        sSql = ""
        sSql = sSql & "INSERT INTO xodcx(" & vbLf
        sSql = sSql & "crynumcx" & vbLf    ''ブロックID
        sSql = sSql & ",mtrlnumcx" & vbLf   ''原料№
        sSql = sSql & ",wkktcx" & vbLf      ''工程コード
        sSql = sSql & ",workcx" & vbLf      ''工場コード
        sSql = sSql & ",hdaycx" & vbLf      ''発生日時
        sSql = sSql & ",weightcx" & vbLf    ''重量
        sSql = sSql & ",htkbncx" & vbLf     ''廃棄/適合区分
        sSql = sSql & ",divumucx" & vbLf    ''分割有無
        sSql = sSql & ",toworkcx" & vbLf    ''払出先工場コード
        sSql = sSql & ",frworkcx" & vbLf    ''発生工場コード
        sSql = sSql & ",hinbcx" & vbLf      ''品番
        sSql = sSql & ",typecx" & vbLf      ''タイプ
        sSql = sSql & ",dptypecx" & vbLf    ''ドープタイプ
        sSql = sSql & ",tposcx" & vbLf      ''位置L(トップ側)
        sSql = sSql & ",lencx" & vbLf       ''ブロック長さ
        sSql = sSql & ",siweightcx" & vbLf  ''仕込み重量
        sSql = sSql & ",updmcx" & vbLf      ''引上AV径
        sSql = sSql & ",prodmcx" & vbLf     ''製品径
        sSql = sSql & ",tdopposcx" & vbLf   ''追加ドープ投入位置L
        sSql = sSql & ",wdopumucx" & vbLf   ''Wドープ(P/N混合)有無
        sSql = sSql & ",csdopumucx" & vbLf  ''CSドープ有無
        sSql = sSql & ",ndopumucx" & vbLf   ''窒素ドープ有無
        sSql = sSql & ",ltspecumucx" & vbLf ''ライフタイム仕様有無
        sSql = sSql & ",csspecumucx" & vbLf ''CS仕様有無
        sSql = sSql & ",topwcx" & vbLf      ''トップWT
        sSql = sSql & ",dmkcx" & vbLf       ''直径区分
        sSql = sSql & ",xtalcx" & vbLf      ''結晶番号
        sSql = sSql & ",livkcx" & vbLf      ''生死区分
        sSql = sSql & ",unifgcx" & vbLf     ''結合FLG
        sSql = sSql & ",twarifgcx" & vbLf   ''縦割FLG
        sSql = sSql & ",refusefgcx" & vbLf  ''受入可否FLG
        sSql = sSql & ",tstafidcx" & vbLf   ''登録社員ID
        sSql = sSql & ",tdaycx" & vbLf      ''登録日付
    '*** UPDATE START T.TERAUCHI 2004/12/07 更新者、更新日時を追加
        sSql = sSql & ",kstafidcx" & vbLf   ''更新者
        sSql = sSql & ",kdaycx" & vbLf      ''更新日時
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        sSql = sSql & ",crydopcx" & vbLf    ''結晶ドープ
        sSql = sSql & ",crydopvlcx" & vbLf  ''結晶ドープ量
        sSql = sSql & ",bkformcx" & vbLf    ''ブロック形状
        sSql = sSql & ",pgidcx" & vbLf      ''PG-ID
        sSql = sSql & ",blktypcx" & vbLf    ''ブロック種別
        sSql = sSql & ",tkacutwcx" & vbLf   ''Tサンプル前重量
    '*** UPDATE START TAGAWA 2004/12/16***************
        sSql = sSql & ",denflgcx" & vbLf     ''電極材フラグ
    '*** UPDATE END   TAGAWA 2004/12/16***************
    '*** UPDATE START T.TERAUCHI 2004/12/07 ﾄｯﾌﾟ取出しWT追加
        sSql = sSql & ",toptwcx" & vbLf     ''ﾄｯﾌﾟ取出しWT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
        sSql = sSql & ",scntrlcx" & vbLf    ''識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ
    '*** UPDATE END   Marushita 2011/03/24

        sSql = sSql & ")values(" & vbLf
        sSql = sSql & "'" & .BLOCKID & "0" & "'" & vbLf                         ''ﾌﾞﾛｯｸID
        sSql = sSql & ",' '" & vbLf                                             ''原料No
        sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf    ''工程ｺｰﾄﾞB410
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''工場ｺｰﾄﾞ
        sSql = sSql & ",sysdate" & vbLf                                         ''発生日時
        sSql = sSql & "," & rec_xodc2.GNWC2 & vbLf                              ''分割結晶(ﾌﾞﾛｯｸ)･現在重量
        sSql = sSql & ",'1'" & vbLf                                             ''廃棄・適合区分
        sSql = sSql & ",'1'" & vbLf                                             ''分割有無
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 払出工場ｺｰﾄﾞを設定
    '    sSQL = sSQL & ",' '" & vbLf                                             ''払出先工場ｺｰﾄﾞ
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''払出先工場ｺｰﾄﾞ
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    
        sSql = sSql & ",'" & gsFactryCd & "'" & vbLf                            ''発生工場ｺｰﾄﾞ
        sSql = sSql & ",'" & objDS.Fields("HINBCX").Value & "'" & vbLf          ''品番
        sSql = sSql & ",'" & objDS.Fields("HSXTYPE").Value & "'  " & vbLf       ''タイプ
        sSql = sSql & ",'" & sDopType & "'" & vbLf                              ''ドープタイプ
        sSql = sSql & ", " & rec_xodc2.INPOSC2 & vbLf                           ''分割結晶(ﾌﾞﾛｯｸ)･結晶内開始位置
        sSql = sSql & ", " & .LENGTH & vbLf                                     ''ブロック長さ
        sSql = sSql & ", " & ConvNum(objDS.Fields("SUICHARGE").Value) & vbLf    ''仕込み重量
        sSql = sSql & ", " & ConvNum(objDS.Fields("UPDMCX").Value) & vbLf       ''引上AV径
        sSql = sSql & ", " & ConvNum(objDS.Fields("PRODMCX").Value) & vbLf      ''製品径
        sSql = sSql & ", " & ConvNum(objDS.Fields("ADDOPPC1").Value) & vbLf     ''追加ドープ投入位置L
        sSql = sSql & ",'1'" & vbLf                                             ''Wドープ(P/N混合)有無
        sSql = sSql & ",'" & sCSDop & "'" & vbLf                                ''CSドープ有無
        sSql = sSql & ",'" & sNDop & "'" & vbLf                                 ''窒素ドープ有無
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 ﾗｲﾌﾀｲﾑ仕様有無は、判定結果より判別
    '    sSQL = sSQL & ",'" & objDS.Fields("HSXLTHWS").Value & "'" & vbLf        ''ライフタイム使用有無
        sSql = sSql & ",'" & sLTUmu & "'" & vbLf                                ''ライフタイム使用有無
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        
        sSql = sSql & ",'2'" & vbLf                                             ''CS使用有無
        
    '*** UPDATE START T.TERAUCHI 2004/12/07 TOPWTを肩重量に変更
    '    sSQL = sSQL & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''トップWT
        sSql = sSql & "," & ConvNum(objDS.Fields("CTR01A9").Value) & vbLf       '’トップWT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        
        sSql = sSql & ",'300'" & vbLf                                           ''直径区分
        sSql = sSql & ",'" & .CRYNUM & "'" & vbLf                               ''引上結晶番号
        sSql = sSql & ",'0'" & vbLf                                             ''生死区分
        sSql = sSql & ",'0'" & vbLf                                             ''結合FLG
        sSql = sSql & ",'0'" & vbLf                                             ''縦割FLG
        sSql = sSql & ",'0'" & vbLf                                             ''受入可否FLG
        sSql = sSql & ",'" & sUserID & "'" & vbLf                               ''登録社員ID
        sSql = sSql & ",sysdate" & vbLf                                         ''登録日付
    '*** UPDATE START T.TERAUCHI 2004/12/07 更新者、更新日時追加対応
        sSql = sSql & ",'" & sUserID & "'" & vbLf                               ''更新社員ID
        sSql = sSql & ",sysdate" & vbLf                                         ''更新日付
    '*** UPDATE END   T.TERAUCHI 2004/12/07
        sSql = sSql & ",'" & objDS.Fields("DPNTCLS").Value & "'" & vbLf         ''結晶ドープ
        sSql = sSql & "," & ConvNum(objDS.Fields("DOPANT").Value) & vbLf        ''結晶ドープ量
        sSql = sSql & ",'3'" & vbLf                                             ''ブロック形状
        sSql = sSql & ",'" & objDS.Fields("PGID").Value & "'" & vbLf            ''PG-ID
        sSql = sSql & ",'A'" & vbLf                                             ''ブロック種別
        sSql = sSql & ",0" & vbLf                                               ''Tサンプル前重量
    '*** UPDATE START TAGAWA 2004/12/16***************
        sSql = sSql & ",'1'" & vbLf                                             ''電極材フラグ
    '*** UPDATE END   TAGAWA 2004/12/16******************
    '*** UPDATE START T.TERAUCHI 2004/12/07
        sSql = sSql & "," & ConvNum(objDS.Fields("PUTCUTWC1").Value) & vbLf     ''ﾄｯﾌﾟ取出しWT
    '*** UPDATE END   T.TERAUCHI 2004/12/07
    '*** UPDATE START Marushita 2011/03/24 TSMC品識別対応
        sSql = sSql & ",'" & sSCNTRL & "'" & vbLf                               ''識別ｺﾝﾄﾛｰﾙｺｰﾄﾞ
    '*** UPDATE END   Marushita 2011/03/24
    
        sSql = sSql & ")"
        
        If SqlExec2(sSql) = -1 Then
            If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
            GoTo proc_exit
        End If
        
    End With
    
    InsXODCX = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' 機能      : XODB3作成関数
'
' 返り値    : FUNCTION_RETURN_FAILURE：異常
'             FUNCTION_RETURN_SUCCESS：正常
'
' 引き数    : 検索結果格納構造体
'
' 機能説明  : 該当データの精製原料管理を作成する
Private Function InsXODB3(rec As type_DBDRV_scmzc_fcmgc001f_Kensaku, _
                            rec_xodc2 As typ_XSDC2) As FUNCTION_RETURN

    Dim objDS       As Object
    Dim sSql        As String
    Dim sUserID     As String
    Dim sUserName   As String
    Dim iIdx        As Integer
    Dim iRenban     As Integer
    Dim sYear       As String
    Dim sMonth      As String
    Dim sDay        As String
    Dim sHour       As String
    Dim sMin        As String
    Dim sNowdate    As String
    Dim sCyoku      As String

On Error GoTo PROC_ERR

    InsXODB3 = FUNCTION_RETURN_FAILURE
    gErr.Push "s_cmzcf_cmbc008_1_SQL.bas -- Function InsXODB3"
    
    '****** 登録情報の作成 *****
    '' 登録社員ＩＤ、社員名
    sUserID = f_cmbc008_1.txtStaffID.Text
    sUserName = f_cmbc008_1.txtJfName.Text
    
    '' システム日付、実績日付等の設定
    If Not GetSysdate Then
        GoTo proc_exit
    End If
    sNowdate = gsSysdate
    'サーバーシステム日付を実績日に変更
    sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
    '実績日より直区分を判定
    sCyoku = GetCYOKU(gsSysdate)
    '実績日から切り取り
    sYear = Mid(sNowdate, 1, 4)     '年
    sMonth = Mid(sNowdate, 5, 2)    '月
    sDay = Mid(sNowdate, 7, 2)      '日
    sHour = Mid(sNowdate, 9, 2)     '時
    sMin = Mid(sNowdate, 11, 2)     '分
        
    '' 工程連番の取得
    iRenban = 0
    sSql = ""
    sSql = sSql & " SELECT NVL(MAX(kcntb3),0) maxcnt     " & vbLf   '工程連番
    sSql = sSql & " FROM   xodb3                         " & vbLf
    
'*** UPDATE START T.TERAUCHI 2004/12/06 現品ロットNoを13桁とする
'    sSQL = sSQL & " WHERE  polnob3 = '" & rec.BLOCKID & "'" & vbLf
    sSql = sSql & " WHERE  polnob3 = '" & rec.BLOCKID & "0" & "'" & vbLf
'*** UPDATE END   T.TERAUCHI 2004/12/06
    
    'SQL文実行
    If DynSet2(objDS, sSql) = False Then
        If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        GoTo proc_exit
    End If
    '取得したデータを格納
    iRenban = objDS.Fields("maxcnt").Value + 1
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing
        
    With rec
        
        '****** 原料工程実績作成 ******
        sSql = ""
        sSql = sSql & "insert into XODB3(                       " & vbLf
        sSql = sSql & "             POLNOB3                     " & vbLf   '原料番号
        sSql = sSql & "            ,KCNTB3                      " & vbLf   '工程連番
        sSql = sSql & "            ,CRSEQB3                     " & vbLf   '処理連番
        sSql = sSql & "            ,TDAYB3                      " & vbLf   '登録日付
        sSql = sSql & "            ,RDAYB3                      " & vbLf   '修正日付
        sSql = sSql & "            ,SDAYB3                      " & vbLf   '送信日付
        sSql = sSql & "            ,SNDKB3                      " & vbLf   '送信区分
        sSql = sSql & "            ,SAKJB3                      " & vbLf   '削除区分
        sSql = sSql & "            ,POKUBB3                     " & vbLf   '原料区分
        sSql = sSql & "            ,POKIDCB3                    " & vbLf   '原料種類コード
        sSql = sSql & "            ,POLTNB3                     " & vbLf   '原料ロットNo
        sSql = sSql & "            ,MODKBB3                     " & vbLf   '赤黒区分
        sSql = sSql & "            ,SUMKBB3                     " & vbLf   '集計区分
        sSql = sSql & "            ,WKKTB3                      " & vbLf   '工程コード
        sSql = sSql & "            ,PLACB3                      " & vbLf   'ラインコード
        sSql = sSql & "            ,FRWB3                       " & vbLf   '受入重量
        sSql = sSql & "            ,TOWB3                       " & vbLf   '払出重量
        sSql = sSql & "            ,LOSWB3                      " & vbLf   'ロス重量
        sSql = sSql & "            ,FRWKKTB3                    " & vbLf   '受入工程コード
        sSql = sSql & "            ,TOWKKTB3                    " & vbLf   '払出工程コード
        sSql = sSql & "            ,TOWKKBB3                    " & vbLf   '払出区分
        sSql = sSql & "            ,TOWORKB3                    " & vbLf   '払出工場コード
        sSql = sSql & "            ,TOPLACB3                    " & vbLf   '払出ラインコード
        sSql = sSql & "            ,CHGNB3                      " & vbLf   'チャージNo
        sSql = sSql & "            ,EYYB3                       " & vbLf   '実績日付(年)
        sSql = sSql & "            ,EMMB3                       " & vbLf   '実績日付(月)
        sSql = sSql & "            ,EDDB3                       " & vbLf   '実績日付(日)
        sSql = sSql & "            ,ECYOKB3                     " & vbLf   '直区分
        sSql = sSql & "            ,EHHB3                       " & vbLf   '実績時間(時)
        sSql = sSql & "            ,EMIB3　                     " & vbLf   '実績時間(分)
        sSql = sSql & "            ,MANB3                       " & vbLf   '担当者
        sSql = sSql & "            ,MANJB3                      " & vbLf   '担当者名
        sSql = sSql & "            ,DENKB3                      " & vbLf   '濃度区分
        sSql = sSql & "            ,DENSITYB3                   " & vbLf   '濃度値
        sSql = sSql & "            ,GSNDFLGB3                   " & vbLf   '原料送信フラグ
        sSql = sSql & "            ,HFLGB3                      " & vbLf   '発生フラグ
        sSql = sSql & "            ,htkbnb3                     " & vbLf   '現品区分
        sSql = sSql & "            ,plworkb3                    " & vbLf   '使用予定工場
        sSql = sSql & "            ,mdensityb3                  " & vbLf   '元濃度値
        sSql = sSql & "            ,gsdayb3                     " & vbLf
        sSql = sSql & ")VALUES(                                 " & vbLf
        
    '*** UPDATE START T.TERAUCHI 2004/12/06 現品ロットNoを13桁とする
    '    sSQL = sSQL & " '" & .BLOCKID & "'                      " & vbLf   '分割結晶番号
        sSql = sSql & " '" & .BLOCKID & "0" & "'                      " & vbLf  '分割結晶番号
    '*** UPDATE END   T.TERAUCHI 2004/12/06
        
        sSql = sSql & "," & iRenban & "                         " & vbLf   '工程連番
        sSql = sSql & ",1                                       " & vbLf   '処理連番
        sSql = sSql & ",to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '登録日付
        sSql = sSql & ",null                                    " & vbLf   '修正日付
        sSql = sSql & ",null                                    " & vbLf   '送信日付
        sSql = sSql & ",' '                                     " & vbLf   '送信区分
        sSql = sSql & ",'0'                                     " & vbLf   '削除区分
        sSql = sSql & ",'2'                                     " & vbLf   '原料区分
        sSql = sSql & ",'888'                                   " & vbLf   '原料種類コード
        sSql = sSql & ",' '                                     " & vbLf   '原料ロット番号
        sSql = sSql & ",' '                                     " & vbLf   '赤黒区分
        sSql = sSql & ",' '                                     " & vbLf   '集計区分
        
    '*** UPDATE START T.TERAUCHI 2004/12/09 工程変更
    '    sSQL = sSQL & ",'" & Right(PROCD_RIMERUTO_UKEIRE, 4) & "'" & vbLf  '工程コード
        sSql = sSql & ",'" & Right(PROCD_KAKUAGE, 4) & "'" & vbLf  '工程コード  B320
    '*** UPDATE END   T.TERAUCHI 2004/12/09
        
        sSql = sSql & ",' '                                     " & vbLf   'ラインコード
        sSql = sSql & "," & rec_xodc2.GNWC2 & "                 " & vbLf   '画面仕掛重量
        sSql = sSql & "," & rec_xodc2.GNWC2 & "                 " & vbLf   '画面仕掛重量
        sSql = sSql & ",0                                       " & vbLf   'ロス重量
        
    '*** UPDATE START T.TERAUCHI 2004/12/09 工程変更
    '    sSQL = sSQL & ",'" & Right(PROCD_RIMERUTO_UKEIRE, 4) & "'" & vbLf  '受入工程コード
        sSql = sSql & ",'" & Right(PROCD_KAKUAGE, 4) & "'" & vbLf  '受入工程コード　B320
    '*** UPDATE END   T.TERAUCHI 2004/12/09
        
        sSql = sSql & ",'" & Right(PROCD_KOUNYU_TAN_KESSYOU, 4) & "'" & vbLf '払出工程コード('B410')
        sSql = sSql & ",' '                                     " & vbLf   '払出区分
        sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '払出工場コード
        sSql = sSql & ",' '                                     " & vbLf   '払出ラインコード
        sSql = sSql & ",' '                                     " & vbLf   'チャージNo
        sSql = sSql & ",'" & sYear & "'                         " & vbLf   '実績日付(年)
        sSql = sSql & ",'" & sMonth & "'                        " & vbLf   '実績日付(月)
        sSql = sSql & ",'" & sDay & "'                          " & vbLf   '実績日付(日)
        sSql = sSql & ",'" & sCyoku & "'                        " & vbLf   '直区分
        sSql = sSql & ",'" & sHour & "'                         " & vbLf   '実績時間(時)
        sSql = sSql & ",'" & sMin & "'                          " & vbLf   '実績時間(分)
        sSql = sSql & ",'" & sUserID & "'                       " & vbLf   '担当者
        sSql = sSql & ",'" & sUserName & "'                     " & vbLf   '担当者名
        sSql = sSql & ",' '                                     " & vbLf   '濃度区分
        sSql = sSql & ",NULL                                    " & vbLf   '濃度値
        sSql = sSql & ",'7'                                     " & vbLf   '原料送信フラグ
        sSql = sSql & ",'0'                                     " & vbLf   '発生フラグ
        sSql = sSql & ",'1'                                     " & vbLf   '現品区分
        sSql = sSql & ",'" & gsFactryCd & "'                    " & vbLf   '使用予定工場
        sSql = sSql & ",NULL                                    " & vbLf   '元濃度
        sSql = sSql & ",NULL                                    " & vbLf '
        sSql = sSql & ")"
        
        If SqlExec2(sSql) = -1 Then
            GoTo proc_exit
        End If
        
    End With
    
    InsXODB3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    
    If Not objDS Is Nothing Then objDS.Close: Set objDS = Nothing

    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit

End Function

' @(f)
' 機能      : SQL数値変換関数
'
' 返り値    : <入力数値> or NULL
'
' 引き数    : 変換対象数値
'
' 機能説明  : 渡された数値がNULLであれば"NULL"をそうでなければそのまま出力する
Private Function ConvNum(vinput) As String
    If IsNull(vinput) Or vinput = "NULL" Then
        vinput = ""
    End If
    
    If vinput = "" Then
        ConvNum = "NULL"
    Else
        ConvNum = vinput
    End If
End Function

' @(f)
' 機能      : 直区分判定
'
' 返り値    : 直区分
'
' 引き数    : nowdate  -  日付データ
'
' 機能説明  : 渡された日付データの時刻から直区分を判定する
'
Private Function GetCYOKU(nowdate As String) As String
    
    Dim jitutime As String     '判定用の時刻を格納する変数

    '判定用に時刻を切り出す
    jitutime = Format(nowdate, "hhnnss")

    '直区分を設定する
    '3直 00:00から07:59
    If jitutime >= "000000" And jitutime < "080000" Then
        GetCYOKU = "3"
    '1直 08:00から15:59
    ElseIf jitutime >= "080000" And jitutime < "160000" Then
        GetCYOKU = "1"
    '2直 16:00から23:59
    ElseIf jitutime >= "160000" And jitutime < "240000" Then
        GetCYOKU = "2"
    End If
End Function
Public Function DBDRV_SELECT_HOLD(pTblDispData As type_DBDRV_scmzc_fcmgc001f_Kensaku) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim sCryNum As String
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "s_XSDC1_SQL.bas -- Function DBDRV_SELECT_HOLD"

    With pTblDispData
        sCryNum = Left(.BLOCKID, 9) & "000"
        ''SQLを組み立てる
        sql = "SELECT HLDCMNT FROM TBCMJ012 "
        sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"
        sql = sql & " ORDER BY TRANCNT"
        'データを抽出する
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        
        If rs Is Nothing Then
            DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        If rs.RecordCount > 0 Then
           rs.MoveLast
            If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
        End If
    End With
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

' 2007/09/18 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      'レコード数
    Dim i  As Long
    
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "f_cmbc008_0.frm -- Function Getstaffauthority"
    
    GetMukeCode = FUNCTION_RETURN_FAILURE
    
    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
'    sql = sql & "or CODEA9 = 'ALL') "
    sql = sql & "or CODEA9 = 'ZX' "         '08/07/01 ooba
    sql = sql & "or CODEA9 = 'ZZ') "        '08/07/01 ooba
    sql = sql & "order by CODEA9 "          '08/07/01 ooba

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim s_Mukesaki(recCnt)
    
    If recCnt = 0 Then
        Exit Function
    End If
    
    For i = 1 To recCnt
        With s_Mukesaki(i)
            If IsNull(rs.Fields("CODEA9")) = False Then
                .sMukeCode = rs.Fields("CODEA9")    ' 向先コード
            End If
            
            If IsNull(rs.Fields("NAMEJA9")) = False Then
                .sMukeName = rs.Fields("NAMEJA9")  ' 向先名
            End If
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

Public Function ChgMukesaki(sZuban As String) As String
    Dim lLp As Long
    Dim sBuf As String
    Dim rs As OraDynaset
    Dim sql As String
    Dim gsMuke4 As String
    Dim gsMuke5 As String
    Dim gsMuke6 As String
    Dim sCScode As String           '顧客ｺｰﾄﾞ　08/07/01 ooba
    Dim sTECHcode As String         'TECHXIV品顧客ｺｰﾄﾞ　08/07/01 ooba
    
    sBuf = ""
    
    sql = "Select hinban,MAX(MNOREVNO), SUM(NVL(TRIM(E1.KFCTFLAG1),'')) FLAG1, SUM(NVL(TRIM(E1.KFCTFLAG2),'')) FLAG2, SUM(NVL(TRIM(E1.KFCTFLAG3),'')) FLAG3 "
    sql = sql & ", MAX(E1.KMGCSCOD) CSCODE "                '08/07/01 ooba
    sql = sql & "from TBCME001 E1 "
    sql = sql & "where E1.HINBAN = '" & Trim(sZuban) & "' "
    sql = sql & "and E1.OPECOND = '1' "
    sql = sql & "group by hinban"
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    If IsNull(rs("FLAG1")) = False Then gsMuke4 = CStr(rs("FLAG1"))   '向先４棟
    If IsNull(rs("FLAG2")) = False Then gsMuke5 = CStr(rs("FLAG2"))   '向先５棟
    If IsNull(rs("FLAG3")) = False Then gsMuke6 = CStr(rs("FLAG3"))   '向先６棟
    sCScode = rs("CSCODE")                                            '購管理顧客ｺｰﾄﾞ　08/07/01 ooba
    
    rs.Close

    'TECHXIV品顧客ｺｰﾄﾞ取得
    sTECHcode = GetSSComboStrA9("X", "21", "CODEA9")
    
    For lLp = 1 To UBound(s_Mukesaki)
        Select Case lLp
            Case 1
                If gsMuke4 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            Case 2
                If gsMuke5 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            Case 3
                If gsMuke6 <> "" Then
                    sBuf = s_Mukesaki(lLp).sMukeCode
                    Exit For
                End If
            'TECHXIV品　08/07/01 ooba
            Case 4
                If gsMuke4 = "" And gsMuke5 = "" And gsMuke6 = "" Then
                    'TECHXIV品ﾁｪｯｸ　08/07/01 ooba
                    If InStr(1, sTECHcode, sCScode) > 0 Then
                        sBuf = s_Mukesaki(lLp).sMukeCode
                        Exit For
                    End If
                End If
            'Bar出荷品　08/07/01 ooba
            Case Else
                If gsMuke4 = "" And gsMuke5 = "" And gsMuke6 = "" Then
                    'TECHXIV品ﾁｪｯｸ　08/07/01 ooba
                    If InStr(1, sTECHcode, sCScode) > 0 Then
                    Else
                        sBuf = s_Mukesaki(lLp).sMukeCode
                        Exit For
                    End If
                End If
        End Select
    Next lLp
    
    If sBuf = "" Then
' 2007/10/10 SPK Tsutsumi Add Start
        '４棟・５棟・６棟に何もフラグがたっていない場合、Bar出荷
        ChgMukesaki = "ZZ"
'        f_cmbc008_1.lblMsg.Caption = "向先取得エラー TBCME001"
' 2007/10/10 SPK Tsutsumi Add End
    Else
        ChgMukesaki = sBuf
    End If
            
proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
' 2007/09/18 SPK Tsutsumi Add End
