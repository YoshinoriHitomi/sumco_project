Attribute VB_Name = "s_cmbc037_SQL"
Option Explicit
'購入単結晶


Public Type PURCHASE_CRYSTAL_SPECIFICATION
    Res(4)      As Double   '比抵抗 Top1～5
    RRG         As Double   '比抵抗 Top RRG
    Oi(4)       As Double   'Oi Top 1～5
    ORG         As Double   'Oi Top ORG
    Cs          As Double   'Cs Top
    LD1(1)      As Double   'LD-1 Top Max, LD-1 Top Ave
    LD2(1)      As Double   'LD-2 Top Max, LD-2 Top Ave
    BMD(1)      As Double   'BMD Top Max, LD-2 Top Ave
    GD(3)       As Double   'GD1 Top,GD2 Top, DIA1 Top, DIA2 Top
    Lt          As Double   'LifeTime From Top
    EPD         As Double   'EPD
End Type


Public Type PURCHASE_CRYSTAL
    DELETE      As String       '受入取消区分
    KRPROCCD    As String       '管理工程コード
    PROCCODE    As String       '工程コード
    TSTAFFID    As String       '社員ID
    HCNO        As String       '発注NO
    RBATCHNO    As String       '炉パッチNo
    blkID       As String       'ブロックID (IN)
    hinban      As String       '品番
    DMTOP(1)    As Double       '直径Top1～2
    DMTAIL(1)   As Double       '直径Bot1～2
    NCHDPTH(1)  As Double       'ノッチ深さ1～2
    NCHWIDTH(1) As Double       'ノッチ巾1～2
    NCHPOS      As String * 2   'ノッチ位置
    SEEDDEG     As String * 1   'シード傾き
    UPLENGTH    As Double       '引上長
    SXLPOS      As Double       'SXL位置
    BlkLen      As Double       'ブロック長さ
    BLKWGHT     As Double       'ブロック重量
    Spec(1)     As PURCHASE_CRYSTAL_SPECIFICATION
End Type
' 製品仕様
Public Type typ_HinSpec1
    HIN As tFullHinban          ' 品番
    HSXTYPE As String * 1       ' タイプ
    HSXCDIR As String * 1       ' 方位
    HSXD1CEN As Double          ' 直径
    HSXDOP As String * 1        ' 結晶ドープ
    HSXDPDIR As String * 2      ' ノッチ位置
    HSXDDMIN As Double          ' ノッチ深さ（ＭＩＮ）
    HSXDDMAX As Double          ' ノッチ深さ（ＭＡＸ）
    HSXSDSLP As Integer         ' シード傾き
' 払出規制項目追加対応 yakimura 2002.12.01 start
    TOPREG As Integer           ' TOP規制
    TAILREG As Double           ' TAIL規制
    BTMSPRT As Integer          ' ボトム析出規制
' 払出規制項目追加対応 yakimura 2002.12.01 end
End Type
' 切断指示
Public Type typ_CutInd
    INGOTPOS As Integer         ' カット位置
    TRANCNT As Integer          ' 処理回数
    LENGTH As Integer           ' 長さ
    PROCCODE As String * 5      ' 工程コード
    BDCAUS As String * 3        ' 区分
    HINUP As tFullHinban        ' 上品番
    HINDN As tFullHinban        ' 下品番
    BLOCKID As String * 12      ' ブロックID
    SMP As typ_SXLSample        ' 検査項目
    PALTNUM As String * 4       ' パレット番号
    ERRUPFLG As Boolean         ' 上品番エラーフラグ
    ERRDNFLG As Boolean         ' 下品番エラーフラグ
    RECOMMEND(1 To 13) As String * 1    'お勧め検査(Rs～EPD)
End Type





'ブロックID入力時
'概要    :購入単結晶 表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                    ,説明
'        :record       ,IO  ,PURCHASE_CRYSTAL                      ,購入単結晶取得用
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                       ,読み込み成否
'説明    :
'履歴    :2001/06/18 蔵本 作成
Public Function DBDRV_s_cmbc037_Disp(record As PURCHASE_CRYSTAL) As FUNCTION_RETURN
    
    
    Dim sql As String
    Dim rs As OraDynaset
    Dim cdc As OraFields

    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_s_cmbc037_Disp"

    DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_SUCCESS

    sql = "select "
    sql = sql & "CRYNUM, "           ' 結晶番号
    sql = sql & "TRANCNT, "          ' 処理回数
    sql = sql & "KRPROCCD, "         ' 管理工程コード
    sql = sql & "PROCCODE, "         ' 工程コード
    sql = sql & "HINBAN, "           ' 品番
    sql = sql & "MNOREVNO, "         ' 製品番号改訂番号
    sql = sql & "FACTORY, "          ' 工場
    sql = sql & "OPECOND, "          ' 操業条件
    sql = sql & "REPCCL, "           ' 受入取消区分
    sql = sql & "RBATCHNO, "         ' 炉バッチＮｏ
    sql = sql & "DMTOP1, "           ' 直径ＴＯＰ１
    sql = sql & "DMTOP2, "           ' 直径ＴＯＰ２
    sql = sql & "DMTAIL1, "          ' 直径ＴＡＩＬ１
    sql = sql & "DMTAIL2, "          ' 直径ＴＡＩＬ２
    sql = sql & "NCHPOS, "           ' ノッチ位置
    sql = sql & "NCHWID1, "          ' ノッチ巾１
    sql = sql & "NCHWID2, "          ' ノッチ巾２
    sql = sql & "SEEDDEG, "          ' シード傾き
    sql = sql & "NCHDPTH1, "         ' ノッチ深さ１
    sql = sql & "NCHDPTH2, "         ' ノッチ深さ２
    sql = sql & "UPLENGTH, "         ' 引上げ長
    sql = sql & "SXLPOS, "           ' ＳＸＬ位置
    sql = sql & "BLKLEN, "           ' ブロック長さ
    sql = sql & "BLKWGHT, "          ' ブロック重量
    sql = sql & "CMPTOP1, "          ' 比抵抗TOP　１
    sql = sql & "CMPTOP2, "          ' 比抵抗TOP　２
    sql = sql & "CMPTOP3, "          ' 比抵抗TOP　３
    sql = sql & "CMPTOP4, "          ' 比抵抗TOP　４
    sql = sql & "CMPTOP5, "          ' 比抵抗TOP　５
    sql = sql & "CMPTOPR, "          ' 比抵抗TOP　RRG
    sql = sql & "CMPTAIL1, "         ' 比抵抗TAIL　１
    sql = sql & "CMPTAIL2, "         ' 比抵抗TAIL　２
    sql = sql & "CMPTAIL3, "         ' 比抵抗TAIL　３
    sql = sql & "CMPTAIL4, "         ' 比抵抗TAIL　４
    sql = sql & "CMPTAIL5, "         ' 比抵抗TAIL　５
    sql = sql & "CMPTAILR, "         ' 比抵抗TAIL　RRG
    sql = sql & "OITOP1, "           ' Oi　TOP　１
    sql = sql & "OITOP2, "           ' Oi　TOP　２
    sql = sql & "OITOP3, "           ' Oi　TOP　３
    sql = sql & "OITOP4, "           ' Oi　TOP　４
    sql = sql & "OITOP5, "           ' Oi　TOP　５
    sql = sql & "OITOPR, "           ' Oi　TOP　ROG
    sql = sql & "OITAIL1, "          ' Oi　TAIL　１
    sql = sql & "OITAIL2, "          ' Oi　TAIL　２
    sql = sql & "OITAIL3, "          ' Oi　TAIL　３
    sql = sql & "OITAIL4, "          ' Oi　TAIL　４
    sql = sql & "OITAIL5, "          ' Oi　TAIL　５
    sql = sql & "OITAILR, "          ' Oi　TAIL　ROG
    sql = sql & "CSTOP, "            ' Cs　TOP
    sql = sql & "CSTAIL, "           ' Cs　TAIL
    sql = sql & "LD1TOPMX, "         ' LD-1　TOP　MAX
    sql = sql & "LD1TOPAV, "         ' LD-1　TOP　AVE
    sql = sql & "LD1TAILM, "         ' LD-1　TAIL　MAX
    sql = sql & "LD1TAILA, "         ' LD-1　TAIL　AVE
    sql = sql & "LD2TOPMM, "         ' LD-2　TOP　MAX
    sql = sql & "LD2TOPAV, "         ' LD-2　TOP　AVE
    sql = sql & "LD2TAILM, "         ' LD-2　TAIL　MAX
    sql = sql & "LD2TAILA, "         ' LD-2　TAIL　AVE
    sql = sql & "BMDTOPMX, "         ' BMD　TOP　MAX
    sql = sql & "BMDTOPAV, "         ' BMD　TOP　AVE
    sql = sql & "BMDTAILM, "         ' BMD　TAIL　MAX
    sql = sql & "BMDTAILA, "         ' BMD　TAIL　AVE
    sql = sql & "GD1TOP, "           ' GD1 TOP
    sql = sql & "GD1TAIL, "          ' GD1 TAIL
    sql = sql & "GD2TOP, "           ' GD2 TOP
    sql = sql & "GD2TAIL, "          ' GD2 TAIL
    sql = sql & "DIA1TOP, "          ' DIA1 TOP
    sql = sql & "DIA1TAIL, "         ' DIA1 TAIL
    sql = sql & "DIA2TOP, "          ' DIA2 TOP
    sql = sql & "DIA2TAIL, "         ' DIA2 TAIL
    sql = sql & "LTFTOP, "           ' LIFETIME from TOP
    sql = sql & "LTFTAIL, "          ' LIFETIME from TAIL
    sql = sql & "EPD, "              ' EPD
    sql = sql & "HCNO, "             ' 発注No
    sql = sql & "TSTAFFID, "         ' 登録社員ID
    sql = sql & "REGDATE, "          ' 登録日付
    sql = sql & "KSTAFFID, "         ' 更新社員ID
    sql = sql & "UPDDATE, "          ' 更新日付
    sql = sql & "SENDFLAG, "         ' 送信フラグ
    sql = sql & "SENDDATE "          ' 送信日付
    sql = sql & " From TBCMG002 "
    sql = sql & " where TRANCNT=ANY(select MAX(TRANCNT) from TBCMG002 Where CRYNUM='" & record.blkID & "') and  CRYNUM='" & record.blkID & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    'レコード0件時はエラー
    If rs.RecordCount = 0 Then
        DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    Set cdc = rs.Fields
    With record
        .DELETE = cdc("REPCCL").Value                   '受入取消区分
        .RBATCHNO = cdc("RBATCHNO").Value               ' 炉バッチＮｏ
        .hinban = cdc("HINBAN").Value & String(2 - Len(Trim(cdc("MNOREVNO").Value)), "0") & Trim(cdc("MNOREVNO").Value) ' 品番
        .DMTOP(0) = cdc("DMTOP1").Value                 ' 直径ＴＯＰ１
        .DMTOP(1) = cdc("DMTOP2").Value                 ' 直径ＴＯＰ２
        .DMTAIL(0) = cdc("DMTAIL1").Value               ' 直径ＴＡＩＬ１
        .DMTAIL(1) = cdc("DMTAIL2").Value               ' 直径ＴＡＩＬ２
        .NCHPOS = cdc("NCHPOS").Value                   ' ノッチ位置
        .NCHWIDTH(0) = cdc("NCHWID1").Value             ' ノッチ巾１
        .NCHWIDTH(1) = cdc("NCHWID2").Value             ' ノッチ巾２
        .SEEDDEG = cdc("SEEDDEG").Value                 ' シード傾き
        .NCHDPTH(0) = cdc("NCHDPTH1").Value             ' ノッチ深さ１
        .NCHDPTH(1) = cdc("NCHDPTH2").Value             ' ノッチ深さ２
        .UPLENGTH = cdc("UPLENGTH").Value               ' 引上げ長
        .SXLPOS = cdc("SXLPOS").Value                   ' ＳＸＬ位置
        .BlkLen = cdc("BLKLEN").Value                   ' ブロック長さ
        .BLKWGHT = cdc("BLKWGHT").Value                 ' ブロック重量
        .Spec(0).Res(0) = cdc("CMPTOP1").Value          ' 比抵抗TOP　１
        .Spec(0).Res(1) = cdc("CMPTOP2").Value          ' 比抵抗TOP　２
        .Spec(0).Res(2) = cdc("CMPTOP3").Value          ' 比抵抗TOP　３
        .Spec(0).Res(3) = cdc("CMPTOP4").Value          ' 比抵抗TOP　４
        .Spec(0).Res(4) = cdc("CMPTOP5").Value          ' 比抵抗TOP　５
        .Spec(1).Res(0) = cdc("CMPTAIL1").Value         ' 比抵抗TAIL　１
        .Spec(1).Res(1) = cdc("CMPTAIL2").Value         ' 比抵抗TAIL　２
        .Spec(1).Res(2) = cdc("CMPTAIL3").Value         ' 比抵抗TAIL　３
        .Spec(1).Res(3) = cdc("CMPTAIL4").Value         ' 比抵抗TAIL　４
        .Spec(1).Res(4) = cdc("CMPTAIL5").Value         ' 比抵抗TAIL　５
        .Spec(0).RRG = cdc("CMPTOPR").Value             ' 比抵抗TOP　RRG
        .Spec(1).RRG = cdc("CMPTAILR").Value            ' 比抵抗TAIL　RRG
        .Spec(0).Oi(0) = cdc("OITOP1").Value            ' Oi　TOP　１
        .Spec(0).Oi(1) = cdc("OITOP2").Value            ' Oi　TOP　２
        .Spec(0).Oi(2) = cdc("OITOP3").Value            ' Oi　TOP　３
        .Spec(0).Oi(3) = cdc("OITOP4").Value            ' Oi　TOP　４
        .Spec(0).Oi(4) = cdc("OITOP5").Value            ' Oi　TOP　５
        .Spec(1).Oi(0) = cdc("OITAIL1").Value           ' Oi　TAIL　１
        .Spec(1).Oi(1) = cdc("OITAIL2").Value           ' Oi　TAIL　２
        .Spec(1).Oi(2) = cdc("OITAIL3").Value           ' Oi　TAIL　３
        .Spec(1).Oi(3) = cdc("OITAIL4").Value           ' Oi　TAIL　４
        .Spec(1).Oi(4) = cdc("OITAIL5").Value           ' Oi　TAIL　５
        .Spec(0).ORG = cdc("OITOPR").Value              ' Oi　TOP　ROG
        .Spec(1).ORG = cdc("OITAILR").Value             ' Oi　TAIL　ROG
        .Spec(0).Cs = cdc("CSTOP").Value                ' Cs　TOP
        .Spec(1).Cs = cdc("CSTAIL").Value               ' Cs　TAIL
        .Spec(0).LD1(0) = cdc("LD1TOPMX").Value         ' LD-1　TOP　MAX
        .Spec(0).LD1(1) = cdc("LD1TOPAV").Value         ' LD-1　TOP　AVE
        .Spec(1).LD1(0) = cdc("LD1TAILM").Value         ' LD-1　TAIL　MAX
        .Spec(1).LD1(1) = cdc("LD1TAILA").Value         ' LD-1　TAIL　AVE
        .Spec(0).LD2(0) = cdc("LD2TOPMM").Value         ' LD-2　TOP　MAX
        .Spec(0).LD2(1) = cdc("LD2TOPAV").Value         ' LD-2　TOP　AVE
        .Spec(1).LD2(0) = cdc("LD2TAILM").Value         ' LD-2　TAIL　MAX
        .Spec(1).LD2(1) = cdc("LD2TAILA").Value         ' LD-2　TAIL　AVE
        .Spec(0).BMD(0) = cdc("BMDTOPMX").Value         ' BMD　TOP　MAX
        .Spec(0).BMD(1) = cdc("BMDTOPAV").Value         ' BMD　TOP　AVE
        .Spec(1).BMD(0) = cdc("BMDTAILM").Value         ' BMD　TAIL　MAX
        .Spec(1).BMD(1) = cdc("BMDTAILA").Value         ' BMD　TAIL　AVE
        .Spec(0).GD(0) = cdc("GD1TOP").Value            ' GD1 TOP
        .Spec(0).GD(1) = cdc("GD2TOP").Value            ' GD1 TAIL
        .Spec(0).GD(2) = cdc("DIA1TOP").Value           ' GD2 TOP
        .Spec(0).GD(3) = cdc("DIA2TOP").Value           ' GD2 TAIL
        .Spec(1).GD(0) = cdc("GD1TAIL").Value           ' DIA1 TOP
        .Spec(1).GD(1) = cdc("GD2TAIL").Value           ' DIA1 TAIL
        .Spec(1).GD(2) = cdc("DIA1TAIL").Value          ' DIA2 TOP
        .Spec(1).GD(3) = cdc("DIA2TAIL").Value          ' DIA2 TAIL
        .Spec(0).Lt = cdc("LTFTOP").Value               ' LIFETIME from TOP
        .Spec(1).Lt = cdc("LTFTAIL").Value              ' LIFETIME from TAIL
        .Spec(1).EPD = cdc("EPD").Value                 ' EPD
        .HCNO = cdc("HCNO").Value                       ' 発注No
    End With
    rs.Close
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_s_cmbc037_Disp = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'実行時
'概要    :購入単結晶 更新、挿入用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO  ,型                                ,説明
'        :record       ,I   ,PURCHASE_CRYSTAL                 ,購入単結晶挿入用
'        :sCmd         ,I   ,String                           ,関数呼出コマンド　2003/10/31 ooba
'        :UpDateFlag   ,I   ,Boolean                          ,更新挿入フラグ
'        :戻ﾘ値        ,O   ,FUNCTION_RETURN                   ,読み込み成否
'説明    :
'履歴    :2001/06/18 蔵本 作成
'　　    :2001/07/19 Sano 改造
'        :受入取消/更新登録の場合[delete]して[insert]するように変更　2003/10/31 ooba

Public Function DBDRV_scmzc_fcmec001b_Exec(record As PURCHASE_CRYSTAL, _
                                            sCmd As String, _
                                            UpDateFlag As Boolean, _
                                            pCryOld() As typ_XSDCS, _
                                            pCrySmp() As typ_XSDCS _
                                            ) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    Dim CryInf As typ_TBCME037
    Dim BlockMng As typ_TBCME040
    Dim hinban As typ_TBCME041
    Dim recCnt As Integer
    Dim i As Long
    Dim sDbName As String
    Dim sDelSql As String    ''delete用SQL　2003/10/31 ooba
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_scmzc_fcmec001b_Exec"

    DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_SUCCESS

    '12桁品番を求める
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '購入単結晶テーブルの挿入、更新 TBCMG002
    
    If DBDRV_KCryTbl_Exec(record, UpDateFlag) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'TBCMI002　研削加工実績
    If UpDateFlag Then
        sDelSql = "delete from TBCMI002 "
        sDelSql = sDelSql & "where CRYNUM = '" & Left(record.blkID, 9) & "000' "
        
        If DBDRV_DeleteTable(sDelSql, "TBCMI002") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:実行、[cancel]:受入取消
    If sCmd = "execution" Then
        If DBDRV_TBCMI002_Exec(record, UpDateFlag) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    
    '結晶情報への挿入
    With CryInf
        .CRYNUM = Left(record.blkID, 9) & "000"             ' 結晶番号
        .DELCLS = record.DELETE                             ' 削除区分
        .LPKRPROCCD = MGPRCD_KOUNYU_TAN_KESSYOU             ' 最終通過管理工程
        .LASTPASS = PROCD_KOUNYU_TAN_KESSYOU                ' 最終通過工程
        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI             ' 管理工程コード
        .PROCCD = PROCD_KESSYOU_SOUGOUHANTEI                ' 工程コード
        .RPHINBAN = fullHinban.hinban                       ' ねらい品番
        .RPREVNUM = fullHinban.mnorevno                     ' ねらい品番製品番号改訂番号
        .RPFACT = fullHinban.factory                        ' ねらい品番工場
        .RPOPCOND = fullHinban.opecond                      ' ねらい品番操業条件
        .PRODCOND = ""                                      ' 製作条件
        .PGID = ""                                          ' ＰＧ－ＩＤ
        .UPLENGTH = record.UPLENGTH                         ' 引上げ長さ
        .TOPLENG = 0                                        ' ＴＯＰ長さ
        .BODYLENG = record.UPLENGTH                         ' 直胴長さ
        .BOTLENG = 0                                        ' ＢＯＴ長さ
        .FREELENG = record.UPLENGTH                         ' フリー長
        .DIAMETER = (record.DMTOP(0) + record.DMTOP(1)) / 2 ' 直径
        .CHARGE = 0                                         ' チャージ量
        .SEED = ""                                          ' シード
        .ADDDPCLS = ""                                      ' 追加ドープ種類
        .ADDDPPOS = 0                                       ' 追加ドープ位置
        .ADDDPVAL = 0                                       ' 追加ドープ量
'        .REGDATE                                            ' 登録日付
'        .UPDDATE                                            ' 更新日付
'        .SENDFLAG                                           ' 送信フラグ
'        .SENDDATE                                           ' 送信日付
    End With
    If UpDateFlag Then
'        If DBDRV_CryInf_Upd(CryInf) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        sDelSql = "delete from TBCME037 "
        sDelSql = sDelSql & "where CRYNUM = '" & CryInf.CRYNUM & "' "
        
        If DBDRV_DeleteTable(sDelSql, "TBCME037") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:実行、[cancel]:受入取消
    If sCmd = "execution" Then
        If DBDRV_CryInf_Ins(CryInf) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    
    'ブロック管理への挿入
    With BlockMng
        .CRYNUM = Left(record.blkID, 9) & "000"             ' 結晶番号
        .INGOTPOS = record.SXLPOS                           ' 結晶内開始位置
        .REALLEN = record.BlkLen                            ' 実長さ
        .LENGTH = record.BlkLen                             ' 長さ
        .BLOCKID = record.blkID                             ' ブロックID
        .KRPROCCD = MGPRCD_KESSYOU_SOUGOUHANTEI             ' 現在管理工程
        .NOWPROC = PROCD_KESSYOU_SOUGOUHANTEI               ' 現在工程
        .LPKRPROCCD = MGPRCD_KOUNYU_TAN_KESSYOU             ' 最終通過管理工程
        .LASTPASS = PROCD_KOUNYU_TAN_KESSYOU                ' 最終通過工程
        .DELCLS = record.DELETE                             ' 削除区分
        .LSTATCLS = "T"                                     ' 最終状態区分
        .RSTATCLS = "T"                                     ' 流動状態区分
        .HOLDCLS = "0"                                      ' ホールド区分
        .BDCAUS = ""                                        ' 不良理由
'        .REGDATE                                            ' 登録日付
'        .UPDDATE                                            ' 更新日付
        .SUMMITSENDFLAG = ""                                ' SUMMIT送信フラグ
'        .SENDFLAG                                           ' 送信フラグ
'        .SENDDATE                                           ' 送信日付
    End With
    If UpDateFlag Then
'        If DBDRV_BlockMng_Upd_SS(BlockMng) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        sDelSql = "delete from TBCME040 "
        sDelSql = sDelSql & "where CRYNUM = '" & BlockMng.CRYNUM & "' "
'        sDelSql = sDelSql & "and INGOTPOS = " & BlockMng.INGOTPOS
        
        If DBDRV_DeleteTable(sDelSql, "TBCME040") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:実行、[cancel]:受入取消
    If sCmd = "execution" Then
        If DBDRV_BlockMng_Ins(BlockMng) = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    With hinban
        .CRYNUM = BlockMng.CRYNUM       ' 結晶番号
        .INGOTPOS = record.SXLPOS       ' 結晶内開始位置
        .hinban = fullHinban.hinban     ' 品番
        .REVNUM = fullHinban.mnorevno   ' 製品番号改訂番号
        .factory = fullHinban.factory   ' 工場
        .opecond = fullHinban.opecond   ' 操業条件
        .LENGTH = record.BlkLen         ' 長さ
    End With
    If UpDateFlag Then
'        If record.DELETE = "1" Then
''            sql = "delete from TBCME041 where CRYNUM = '" & HINBAN.CryNum & "' and INGOTPOS = " & HINBAN.IngotPos
''            If 0 >= OraDB.ExecuteSQL(sql) Then
''                DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
''                GoTo proc_exit
''            End If
'        Else
'            With hinban
'            sql = "update TBCME041 set "
'            sql = sql & "INGOTPOS=" & .INGOTPOS & ", "
'            sql = sql & "HINBAN='" & .hinban & "', "           ' 品番
'            sql = sql & "REVNUM='" & .REVNUM & "', "           ' 製品番号改訂番号
'            sql = sql & "FACTORY='" & .factory & "', "         ' 工場
'            sql = sql & "OPECOND='" & .opecond & "', "         ' 操業条件
'            sql = sql & "LENGTH='" & .LENGTH & "', "           ' 長さ
'            sql = sql & " UPDDATE=sysdate, "
'            sql = sql & " SENDFLAG='0' "
'            sql = sql & " where CRYNUM='" & .CRYNUM & "' "
'            sql = sql & " and INGOTPOS=" & key.POSITION & " "
'            End With
'            If 0 >= OraDB.ExecuteSQL(sql) Then
'                DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
'                GoTo proc_exit
'            End If
'        End If
'    Else
        sDelSql = "delete from TBCME041 "
        sDelSql = sDelSql & "where CRYNUM = '" & hinban.CRYNUM & "' "
'        sDelSql = sDelSql & "and INGOTPOS = " & hinban.INGOTPOS
        
        If DBDRV_DeleteTable(sDelSql, "TBCME041") = FUNCTION_RETURN_FAILURE Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    ''[execution]:実行、[cancel]:受入取消
    If sCmd = "execution" Then
        '品番管理の挿入
        sql = "insert into TBCME041 ( "
        sql = sql & "CRYNUM, "            ' 結晶番号
        sql = sql & "INGOTPOS, "          ' 結晶内開始位置
        sql = sql & "HINBAN, "            ' 品番
        sql = sql & "REVNUM, "            ' 製品番号改訂番号
        sql = sql & "FACTORY, "           ' 工場
        sql = sql & "OPECOND, "           ' 操業条件
        sql = sql & "LENGTH, "            ' 長さ
        sql = sql & "REGDATE, "           ' 登録日付
        sql = sql & "UPDDATE, "           ' 更新日付
        sql = sql & "SENDFLAG, "          ' 送信フラグ
        sql = sql & "SENDDATE  ) "          ' 送信日付
        With hinban
        sql = sql & "values ("
        sql = sql & " '" & .CRYNUM & "', "          ' 結晶番号
        sql = sql & " " & .INGOTPOS & ", "          ' 結晶内開始位置
        sql = sql & " '" & .hinban & "', "          ' 品番
        sql = sql & " " & .REVNUM & ", "            ' 製品番号改訂番号
        sql = sql & " '" & .factory & "', "         ' 工場
        sql = sql & " '" & .opecond & "', "         ' 操業条件
        sql = sql & " " & .LENGTH & ", "            ' 長さ
        End With
        sql = sql & " sysdate, "                    ' 登録日付
        sql = sql & " sysdate, "                    ' 更新日付
        sql = sql & " '0', "                        ' 送信フラグ
        sql = sql & " sysdate ) "                   ' 送信日付
        If 0 >= OraDB.ExecuteSQL(sql) Then
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    
    'XSDC1 追加/更新
            
    'XSDC2 追加/更新
    
    'XSDC3 追加/更新
    
    'XSDCA 追加/更新
    
    
    sDbName = "E043"
'    If record.DELETE = "1" Then
    If UpDateFlag Then
        '' 結晶サンプル管理の削除
        If DBDRV_CrySmp_Del(pCryOld()) = FUNCTION_RETURN_FAILURE Then
            'sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    ''[execution]:実行、[cancel]:受入取消
    If sCmd = "execution" Then
        '' サンプル№の取得
        recCnt = UBound(pCrySmp)
        For i = 1 To recCnt
            If pCrySmp(i).REPSMPLIDCS = 0 Then
                pCrySmp(i).REPSMPLIDCS = GetNewID_SampleNo()
            End If
        Next i

        '' 結晶サンプル管理の挿入／更新
        If DBDRV_CrySmp_UpdIns037Only(pCryOld(), pCrySmp()) = FUNCTION_RETURN_FAILURE Then
            'sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_scmzc_fcmec001b_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_GetCryCheck(CRYNUM As String, CryFlag As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_GetCryCheck"

    DBDRV_GetCryCheck = FUNCTION_RETURN_FAILURE
    
    sql = "select CRYNUM, DELCLS from TBCME037 where CRYNUM='" & CRYNUM & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    CryFlag = True
    If rs.RecordCount = 0 Then
        CryFlag = False
    Else
        If rs("DELCLS") = "1" Then
            CryFlag = False
        End If
    End If
    rs.Close
    
    DBDRV_GetCryCheck = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_GetCryCheck = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

Public Function DBDRV_GetCryInBlk(CRYNUM As String, blkPos As Integer, BlkLen As Integer, OkNgFlag As Boolean) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_GetCryInBlk"

    DBDRV_GetCryInBlk = FUNCTION_RETURN_FAILURE
    
   '指定範囲に入っているブロック の検索

    sql = "select CRYNUM from TBCME040 where CRYNUM='" & CRYNUM & "' and ("
'    sql = "select count(CRYNUM) from TBCME040 where CRYNUM='" & CryNum & "' and ("
    sql = sql & "(INGOTPOS > " & blkPos & " and INGOTPOS < " & blkPos + BlkLen & ") "
    sql = sql & "or (INGOTPOS + LENGTH > " & blkPos & " and INGOTPOS + LENGTH <  " & blkPos + BlkLen & ") "
    sql = sql & "or (" & blkPos & " > INGOTPOS and " & blkPos + BlkLen & " < INGOTPOS + LENGTH ) "
'    sql = sql & "and (" & BlkPos + BlkLen & " > INGOTPOS and " & BlkPos & " < INGOTPOS + LENGTH)"
    sql = sql & ")"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    OkNgFlag = False
    If rs.RecordCount = 0 Then
        rs.Close
        OkNgFlag = True
    End If
    rs.Close
    
    DBDRV_GetCryInBlk = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_GetCryInBlk = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


Public Function DBDRV_KCryTbl_Exec(record As PURCHASE_CRYSTAL, UpDateFlag As Boolean) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_KCryTbl_Exec"

    DBDRV_KCryTbl_Exec = FUNCTION_RETURN_SUCCESS

    '12桁品番を求める
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'If Not UpDateFlag Then
        
        '購入単結晶テーブルへ値の挿入
        sql = "insert into TBCMG002 ( "
        sql = sql & "CRYNUM, "           ' 結晶番号
        sql = sql & "TRANCNT, "          ' 処理回数
        sql = sql & "KRPROCCD, "         ' 管理工程コード
        sql = sql & "PROCCODE, "         ' 工程コード
        sql = sql & "HINBAN, "           ' 品番
        sql = sql & "MNOREVNO, "         ' 製品番号改訂番号
        sql = sql & "FACTORY, "          ' 工場
        sql = sql & "OPECOND, "          ' 操業条件
        sql = sql & "REPCCL, "           ' 受入取消区分
        sql = sql & "RBATCHNO, "         ' 炉バッチＮｏ
        sql = sql & "DMTOP1, "           ' 直径ＴＯＰ１
        sql = sql & "DMTOP2, "           ' 直径ＴＯＰ２
        sql = sql & "DMTAIL1, "          ' 直径ＴＡＩＬ１
        sql = sql & "DMTAIL2, "          ' 直径ＴＡＩＬ２
        sql = sql & "NCHPOS, "           ' ノッチ位置
        sql = sql & "NCHWID1, "          ' ノッチ巾１
        sql = sql & "NCHWID2, "          ' ノッチ巾２
        sql = sql & "SEEDDEG, "          ' シード傾き
        sql = sql & "NCHDPTH1, "         ' ノッチ深さ１
        sql = sql & "NCHDPTH2, "         ' ノッチ深さ２
        sql = sql & "UPLENGTH, "         ' 引上げ長
        sql = sql & "SXLPOS, "           ' ＳＸＬ位置
        sql = sql & "BLKLEN, "           ' ブロック長さ
        sql = sql & "BLKWGHT, "          ' ブロック重量
        sql = sql & "CMPTOP1, "          ' 比抵抗TOP　１
        sql = sql & "CMPTOP2, "          ' 比抵抗TOP　２
        sql = sql & "CMPTOP3, "          ' 比抵抗TOP　３
        sql = sql & "CMPTOP4, "          ' 比抵抗TOP　４
        sql = sql & "CMPTOP5, "          ' 比抵抗TOP　５
        sql = sql & "CMPTOPR, "          ' 比抵抗TOP　RRG
        sql = sql & "CMPTAIL1, "         ' 比抵抗TAIL　１
        sql = sql & "CMPTAIL2, "         ' 比抵抗TAIL　２
        sql = sql & "CMPTAIL3, "         ' 比抵抗TAIL　３
        sql = sql & "CMPTAIL4, "         ' 比抵抗TAIL　４
        sql = sql & "CMPTAIL5, "         ' 比抵抗TAIL　５
        sql = sql & "CMPTAILR, "         ' 比抵抗TAIL　RRG
        sql = sql & "OITOP1, "           ' Oi　TOP　１
        sql = sql & "OITOP2, "           ' Oi　TOP　２
        sql = sql & "OITOP3, "           ' Oi　TOP　３
        sql = sql & "OITOP4, "           ' Oi　TOP　４
        sql = sql & "OITOP5, "           ' Oi　TOP　５
        sql = sql & "OITOPR, "           ' Oi　TOP　ROG
        sql = sql & "OITAIL1, "          ' Oi　TAIL　１
        sql = sql & "OITAIL2, "          ' Oi　TAIL　２
        sql = sql & "OITAIL3, "          ' Oi　TAIL　３
        sql = sql & "OITAIL4, "          ' Oi　TAIL　４
        sql = sql & "OITAIL5, "          ' Oi　TAIL　５
        sql = sql & "OITAILR, "          ' Oi　TAIL　ROG
        sql = sql & "CSTOP, "            ' Cs　TOP
        sql = sql & "CSTAIL, "           ' Cs　TAIL
        sql = sql & "LD1TOPMX, "         ' LD-1　TOP　MAX
        sql = sql & "LD1TOPAV, "         ' LD-1　TOP　AVE
        sql = sql & "LD1TAILM, "         ' LD-1　TAIL　MAX
        sql = sql & "LD1TAILA, "         ' LD-1　TAIL　AVE
        sql = sql & "LD2TOPMM, "         ' LD-2　TOP　MAX
        sql = sql & "LD2TOPAV, "         ' LD-2　TOP　AVE
        sql = sql & "LD2TAILM, "         ' LD-2　TAIL　MAX
        sql = sql & "LD2TAILA, "         ' LD-2　TAIL　AVE
        sql = sql & "BMDTOPMX, "         ' BMD　TOP　MAX
        sql = sql & "BMDTOPAV, "         ' BMD　TOP　AVE
        sql = sql & "BMDTAILM, "         ' BMD　TAIL　MAX
        sql = sql & "BMDTAILA, "         ' BMD　TAIL　AVE
        sql = sql & "GD1TOP, "           ' GD1 TOP
        sql = sql & "GD1TAIL, "          ' GD1 TAIL
        sql = sql & "GD2TOP, "           ' GD2 TOP
        sql = sql & "GD2TAIL, "          ' GD2 TAIL
        sql = sql & "DIA1TOP, "          ' DIA1 TOP
        sql = sql & "DIA1TAIL, "         ' DIA1 TAIL
        sql = sql & "DIA2TOP, "          ' DIA2 TOP
        sql = sql & "DIA2TAIL, "         ' DIA2 TAIL
        sql = sql & "LTFTOP, "           ' LIFETIME from TOP
        sql = sql & "LTFTAIL, "          ' LIFETIME from TAIL
        sql = sql & "EPD, "              ' EPD
        sql = sql & "HCNO, "             ' 発注No
        sql = sql & "TSTAFFID, "         ' 登録社員ID
        sql = sql & "REGDATE, "          ' 登録日付
        sql = sql & "KSTAFFID, "         ' 更新社員ID
        sql = sql & "UPDDATE, "          ' 更新日付
        sql = sql & "SENDFLAG, "         ' 送信フラグ
        sql = sql & "SENDDATE ) "         ' 送信日付
        With record
            sql = sql & " select "
            sql = sql & " '" & .blkID & "', "             ' 結晶番号
            sql = sql & "nvl(max(TRANCNT),0)+1, "                   ' 処理回数
            sql = sql & " '" & .KRPROCCD & "', "              ' 管理工程コード
            sql = sql & " '" & .PROCCODE & "', "              ' 工程コード
            sql = sql & " '" & fullHinban.hinban & "', "            ' 品番
            sql = sql & fullHinban.mnorevno & ", "
            sql = sql & " '" & fullHinban.factory & "', "
            sql = sql & " '" & fullHinban.opecond & "', "
            sql = sql & " '" & .DELETE & "', "                '受入取消区分
            sql = sql & " '" & .RBATCHNO & "', "
            sql = sql & .DMTOP(0) & ", "
            sql = sql & .DMTOP(1) & ", "
            sql = sql & .DMTAIL(0) & ", "
            sql = sql & .DMTAIL(1) & ", "
            sql = sql & "'" & .NCHPOS & "', "
            sql = sql & .NCHWIDTH(0) & ", "
            sql = sql & "-1, "
            sql = sql & "'" & .SEEDDEG & "', "
            sql = sql & .NCHDPTH(0) & ", "
            sql = sql & "-1, "
            sql = sql & .UPLENGTH & ", "
            sql = sql & .SXLPOS & ", "
            sql = sql & .BlkLen & ", "
            sql = sql & .BLKWGHT & ", "
            sql = sql & .Spec(0).Res(0) & ", "
            sql = sql & .Spec(0).Res(1) & ", "
            sql = sql & .Spec(0).Res(2) & ", "
            sql = sql & .Spec(0).Res(3) & ", "
            sql = sql & .Spec(0).Res(4) & ", "
            sql = sql & .Spec(0).RRG & ", "
            sql = sql & .Spec(1).Res(0) & ", "
            sql = sql & .Spec(1).Res(1) & ", "
            sql = sql & .Spec(1).Res(2) & ", "
            sql = sql & .Spec(1).Res(3) & ", "
            sql = sql & .Spec(1).Res(4) & ", "
            sql = sql & .Spec(1).RRG & ", "
            sql = sql & .Spec(0).Oi(0) & ", "
            sql = sql & .Spec(0).Oi(1) & ", "
            sql = sql & .Spec(0).Oi(2) & ", "
            sql = sql & .Spec(0).Oi(3) & ", "
            sql = sql & .Spec(0).Oi(4) & ", "
            sql = sql & .Spec(0).ORG & ", "
            sql = sql & .Spec(1).Oi(0) & ", "
            sql = sql & .Spec(1).Oi(1) & ", "
            sql = sql & .Spec(1).Oi(2) & ", "
            sql = sql & .Spec(1).Oi(3) & ", "
            sql = sql & .Spec(1).Oi(4) & ", "
            sql = sql & .Spec(1).ORG & ", "
            sql = sql & .Spec(0).Cs & ", "
            sql = sql & .Spec(1).Cs & ", "
            sql = sql & .Spec(0).LD1(0) & ", "
            sql = sql & .Spec(0).LD1(1) & ", "
            sql = sql & .Spec(1).LD1(0) & ", "
            sql = sql & .Spec(1).LD1(1) & ", "
            sql = sql & .Spec(0).LD2(0) & ", "
            sql = sql & .Spec(0).LD2(1) & ", "
            sql = sql & .Spec(1).LD2(0) & ", "
            sql = sql & .Spec(1).LD2(1) & ", "
            sql = sql & .Spec(0).BMD(0) & ", "
            sql = sql & .Spec(0).BMD(1) & ", "
            sql = sql & .Spec(1).BMD(0) & ", "
            sql = sql & .Spec(1).BMD(1) & ", "
            sql = sql & .Spec(0).GD(0) & ", "
            sql = sql & .Spec(1).GD(0) & ", "
            sql = sql & .Spec(0).GD(1) & ", "
            sql = sql & .Spec(1).GD(1) & ", "
            sql = sql & .Spec(0).GD(2) & ", "
            sql = sql & .Spec(1).GD(2) & ", "
            sql = sql & .Spec(0).GD(3) & ", "
            sql = sql & .Spec(1).GD(3) & ", "
            sql = sql & .Spec(0).Lt & ", "
            sql = sql & .Spec(1).Lt & ", "
            sql = sql & .Spec(1).EPD & ", "
            sql = sql & " '" & .HCNO & "', "
            sql = sql & " '" & .TSTAFFID & "', "
            sql = sql & " sysdate, "
            sql = sql & " '" & .TSTAFFID & "', "
            sql = sql & " sysdate , "
            sql = sql & " '0', "
            sql = sql & " sysdate  "
            sql = sql & " from TBCMG002 "
            sql = sql & " where CRYNUM='" & .blkID & "' "
        End With
    'Else
        '更新時
        
    '    With record
    '        sql = "UPDATE TBCMG002 SET "
    '        sql = sql & "KRPROCCD='" & .KRPROCCD & "',"
    '        sql = sql & "PROCCODE='" & .PROCCODE & "',"
    '        sql = sql & "HINBAN='" & fullHinban.HINBAN & "',"
    '        sql = sql & "MNOREVNO=" & fullHinban.mnorevno & ","
    '        sql = sql & "FACTORY='" & fullHinban.factory & "',"
    '        sql = sql & "OPECOND='" & fullHinban.opecond & "',"
    '        sql = sql & "REPCCL='" & .DELETE & "',"
    '        sql = sql & "RBATCHNO='" & .RBATCHNO & "',"
    '        sql = sql & "DMTOP1=" & .DMTOP(0) & ","
    '        sql = sql & "DMTOP2=" & .DMTOP(1) & ","
    '        sql = sql & "DMTAIL1=" & .DMTAIL(0) & ","
    '        sql = sql & "DMTAIL2=" & .DMTAIL(1) & ","
    '        sql = sql & "NCHPOS='" & .NCHPOS & "',"
    '        sql = sql & "NCHDPTH1=" & .NCHWIDTH(0) & ","
    '        sql = sql & "NCHDPTH2=" & .NCHWIDTH(1) & ","
    '        sql = sql & "NCHWID1='" & .NCHDPTH(0) & "',"
    '        sql = sql & "NCHWID2=" & .NCHDPTH(1) & ","
    '        sql = sql & "SEEDDEG=" & .SEEDDEG & ","
    '        sql = sql & "UPLENGTH=" & .UPLENGTH & ","
    '        sql = sql & "SXLPOS=" & .SXLPOS & ","
    '        sql = sql & "BLKLEN=" & .BlkLen & ","
    '        sql = sql & "BLKWGHT=" & .BLKWGHT & ","
    '        sql = sql & "CMPTOP1=" & .Spec(0).Res(0) & ","
    '        sql = sql & "CMPTOP2=" & .Spec(0).Res(1) & ","
    '        sql = sql & "CMPTOP3=" & .Spec(0).Res(2) & ","
    '        sql = sql & "CMPTOP4=" & .Spec(0).Res(3) & ","
    '        sql = sql & "CMPTOP5=" & .Spec(0).Res(4) & ","
    '        sql = sql & "CMPTOPR=" & .Spec(0).RRG & ","
    '        sql = sql & "CMPTAIL1=" & .Spec(1).Res(0) & ","
    '        sql = sql & "CMPTAIL2=" & .Spec(1).Res(1) & ","
    '        sql = sql & "CMPTAIL3=" & .Spec(1).Res(2) & ","
    '        sql = sql & "CMPTAIL4=" & .Spec(1).Res(3) & ","
    '        sql = sql & "CMPTAIL5=" & .Spec(1).Res(4) & ","
    '        sql = sql & "CMPTAILR=" & .Spec(1).RRG & ","
    '        sql = sql & "OITOP1=" & .Spec(0).Oi(0) & ","
    '        sql = sql & "OITOP2=" & .Spec(0).Oi(1) & ","
    '        sql = sql & "OITOP3=" & .Spec(0).Oi(2) & ","
    '        sql = sql & "OITOP4=" & .Spec(0).Oi(3) & ","
    '        sql = sql & "OITOP5=" & .Spec(0).Oi(4) & ","
    '        sql = sql & "OITOPR=" & .Spec(0).ORG & ","
    '        sql = sql & "OITAIL1=" & .Spec(1).Oi(0) & ","
    '        sql = sql & "OITAIL2=" & .Spec(1).Oi(1) & ","
    '        sql = sql & "OITAIL3=" & .Spec(1).Oi(2) & ","
    '        sql = sql & "OITAIL4=" & .Spec(1).Oi(3) & ","
    '        sql = sql & "OITAIL5=" & .Spec(1).Oi(4) & ","
    '        sql = sql & "OITAILR=" & .Spec(1).ORG & ","
    '        sql = sql & "CSTOP=" & .Spec(0).Cs & ","
    '        sql = sql & "CSTAIL=" & .Spec(1).Cs & ","
    '        sql = sql & "LD1TOPMX=" & .Spec(0).LD1(0) & ","
    '        sql = sql & "LD1TOPAV=" & .Spec(0).LD1(1) & ","
    '        sql = sql & "LD1TAILM=" & .Spec(1).LD1(0) & ","
    '        sql = sql & "LD1TAILA=" & .Spec(1).LD1(1) & ","
    '        sql = sql & "LD2TOPMM=" & .Spec(0).LD2(0) & ","
    '        sql = sql & "LD2TOPAV=" & .Spec(0).LD2(1) & ","
    '        sql = sql & "LD2TAILM=" & .Spec(1).LD2(0) & ","
    '        sql = sql & "LD2TAILA=" & .Spec(1).LD2(1) & ","
    '        sql = sql & "BMDTOPMX=" & .Spec(0).BMD(0) & ","
    '        sql = sql & "BMDTOPAV=" & .Spec(0).BMD(1) & ","
    '        sql = sql & "BMDTAILM=" & .Spec(1).BMD(0) & ","
    '        sql = sql & "BMDTAILA=" & .Spec(1).BMD(1) & ","
    '        sql = sql & "GD1TOP=" & .Spec(0).GD(0) & ","
    '        sql = sql & "GD1TAIL=" & .Spec(1).GD(0) & ","
    '        sql = sql & "GD2TOP=" & .Spec(0).GD(1) & ","
    '        sql = sql & "GD2TAIL=" & .Spec(1).GD(1) & ","
    '        sql = sql & "DIA1TOP=" & .Spec(0).GD(2) & ","
    '        sql = sql & "DIA1TAIL=" & .Spec(1).GD(2) & ","
    '        sql = sql & "DIA2TOP=" & .Spec(0).GD(3) & ","
    '        sql = sql & "DIA2TAIL=" & .Spec(1).GD(3) & ","
    '        sql = sql & "LTFTOP=" & .Spec(0).Lt & ","
    '        sql = sql & "LTFTAIL=" & .Spec(1).Lt & ","
    '        sql = sql & "EPD=" & .Spec(1).EPD & ","
    '        sql = sql & "HCNO='" & .HCNO & "',"
    '        sql = sql & "KSTAFFID='" & .TSTAFFID & "',"
    '        sql = sql & "UPDDATE=sysdate,"
    '        sql = sql & "SENDFLAG='0',"
    '        sql = sql & "SENDDATE=sysdate "
    '        sql = sql & "WHERE CRYNUM='" & .blkID & "'"

    '    End With
    'End If
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_KCryTbl_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
'概要      :ブロック管理の更新
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:BlockMng　　　,I  ,typ_TBCME040   　,ブロック管理
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'-------使用しないほうがよい（蔵本）---------
Public Function DBDRV_BlockMng_Upd_SS(BlockMng As typ_TBCME040) As FUNCTION_RETURN

    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_BlockMng_Upd"

    '' ブロック管理テーブルの更新
    With BlockMng
        sql = "update TBCME040 set "
        sql = sql & "INGOTPOS=" & .INGOTPOS & ", "
        sql = sql & "LENGTH=" & .LENGTH & ", "              ' 長さ
        sql = sql & "REALLEN=" & .REALLEN & ", "            ' 実長さ
        sql = sql & "BLOCKID='" & .BLOCKID & "', "          ' ブロックID
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "        ' 現在管理工程
        sql = sql & "NOWPROC='" & .NOWPROC & "', "          ' 現在工程
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "    ' 最終通過管理工程
        sql = sql & "LASTPASS='" & .LASTPASS & "', "        ' 最終通過工程
        sql = sql & "DELCLS='" & .DELCLS & "',"             ' 削除区分
        sql = sql & "LSTATCLS='" & .LSTATCLS & "', "        ' 最終状態区分
        sql = sql & "RSTATCLS='" & .RSTATCLS & "', "        ' 流動状態区分
        sql = sql & "HOLDCLS='" & .HOLDCLS & "', "          ' ホールド区分
        sql = sql & "BDCAUS='" & .BDCAUS & "', "            ' 不良理由
        sql = sql & "UPDDATE=sysdate, "                     ' 更新日付
        sql = sql & "SENDFLAG='0' "                        ' 送信フラグ
        sql = sql & "where CRYNUM='" & .CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & Key.POSITION
    End With
    Debug.Print sql
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_FAILURE
    Else
        DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_SUCCESS
    End If

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BlockMng_Upd_SS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function



'概要      :結晶加工払出用 結晶番号入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sCryNum 　　　,I  ,String         　,結晶番号
'      　　:pCryInf 　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'      　　:pPupEnd 　　　,O  ,typ_TBCMH004   　,引上げ終了実績
'      　　:pHinSpec　　　,O  ,typ_HinSpec1   　,製品仕様
'      　　:pCutInd 　　　,O  ,typ_CutInd     　,切断指示
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001b_Disp(sCrynum As String, _
                                           pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, _
                                           pHinSpec() As typ_HinSpec1, _
                                           pCutInd() As typ_CutInd, _
                                           pCryOld() As typ_XSDCS, _
                                           sErrMsg As String, _
                                           fullHinban As tFullHinban) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim ctcen As Double
    Dim cycen As Double
    Dim iLp2 As Integer
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Disp"
    sErrMsg = ""
    
        '品番は1件しかないので固定する
    ReDim pHinDsn(1)
    pHinDsn(1).CRYNUM = sCrynum
    pHinDsn(1).INGOTPOS = 0 'Debug すること 0でいいかと思うけど
    pHinDsn(1).hinban = fullHinban.hinban
    pHinDsn(1).REVNUM = fullHinban.mnorevno
    pHinDsn(1).FACT = fullHinban.factory
    pHinDsn(1).OPCOND = fullHinban.opecond
    'pHinDsn(1).LENGTH =   'LoadDataが終わった時点で入れる
    'pHinDsn(1).USECLASS
    'pHinDsn(1).REGDATE
    'pHinDsn(1).Update
    'pHinDsn(1).SENDFLAG
    'pHinDsn(1).SENDDATE

    
    'ダミーで1件っていうことにする
    recCnt = 1
    
    '' 製品仕様の取得
' 払出規制項目追加対応 yakimura 2002.12.01 start
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).hinban)
        If sHin <> "G" And sHin <> "Z" Then
            
            For iLp2 = 1 To j
                If (sHin = pHinSpec(iLp2).HIN.hinban) And _
                   (pHinDsn(i).OPCOND = pHinSpec(iLp2).HIN.opecond) And _
                   (pHinDsn(i).REVNUM = pHinSpec(iLp2).HIN.mnorevno) And _
                   (pHinDsn(i).FACT = pHinSpec(iLp2).HIN.factory) Then
                    Exit For
                End If
            Next iLp2
            
            If (iLp2 > j) Then
                sql = "select "
                sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
                sql = sql & " ,NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
                sql = sql & " from TBCME018 E018,TBCME036 E036"
                sql = sql & " where E018.HINBAN='" & pHinDsn(i).hinban & "'"
                sql = sql & " and E018.MNOREVNO=" & pHinDsn(i).REVNUM
                sql = sql & " and E018.FACTORY='" & pHinDsn(i).FACT & "'"
                sql = sql & " and E018.OPECOND='" & pHinDsn(i).OPCOND & "'"
                sql = sql & " and E036.HINBAN='" & pHinDsn(i).hinban & "'"
                sql = sql & " and E036.MNOREVNO=" & pHinDsn(i).REVNUM
                sql = sql & " and E036.FACTORY='" & pHinDsn(i).FACT & "'"
                sql = sql & " and E036.OPECOND='" & pHinDsn(i).OPCOND & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    rs.Close
                    sErrMsg = GetMsgStr("EGET2", sDbName)
                    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                j = j + 1
                With pHinSpec(j)
                    .HIN.hinban = pHinDsn(i).hinban
                    .HIN.mnorevno = pHinDsn(i).REVNUM
                    .HIN.factory = pHinDsn(i).FACT
                    .HIN.opecond = pHinDsn(i).OPCOND
                    .HSXTYPE = rs("HSXTYPE")    ' タイプ
                    .HSXCDIR = rs("HSXCDIR")    ' 方位
                    .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' 直径 →NULL対応
                    .HSXDOP = rs("HSXDOP")      ' 結晶ドープ
                    .HSXDPDIR = rs("HSXDPDIR")          ' 品ＳＸ溝位置方位
                    .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' 品ＳＸ溝深下限 →NULL対応
                    .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' 品ＳＸ溝深上限 →NULL対応
                    ctcen = Abs(CDbl(rs("HSXCTCEN")))
                    cycen = Abs(CDbl(rs("HSXCYCEN")))
                    .TOPREG = fncNullCheck(rs("TOPREG"))              ' TOP規制
                    .TAILREG = fncNullCheck(rs("TAILREG"))            ' TAIL規制
                    .BTMSPRT = fncNullCheck(rs("BTMSPRT"))            ' ボトム析出規制
                    If ((ctcen = 2.83) And (cycen = 2.83)) _
                    Or ((ctcen = 4) And (cycen = 0)) _
                    Or ((ctcen = 0) And (cycen = 4)) Then
                        .HSXSDSLP = 4
                    Else
                        .HSXSDSLP = 0
                    End If
                End With
                rs.Close
            End If
        End If
    Next i
    ReDim Preserve pHinSpec(j)
    
    ReDim pCutInd(1) '1件しかないので固定でよい
    
    'For i = 1 To recCnt
    '    With pCutInd(i)
    '        .INGOTPOS = rs("INGOTPOS")      ' カット位置
    '        .LENGTH = rs("LENGTH")          ' 長さ
    '    End With
    '    rs.MoveNext
    'Next i
    'rs.Close
    
    ' LoadDataが終わった時点で代入する
    'pCutInd(1).INGOTPOS = 'カット位置(全長だな)
    'pCutInd(1).LENGTH = '長さ(全長だな)

    
    ' 結晶サンプル管理の取得
    sDbName = "E043"
    sql = " where CRYNUMCS='" & sCrynum & "'"
    If DBDRV_GetTBCME043(pCryOld(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

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
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
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

'概要      :テーブル「XSDCS」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCS    ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME043_SQL.basより移動)
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
    ''SQLを組み立てる
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS," & _
              " CRYSMPLIDOICS, CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2CS, CRYINDB2CS, " & _
              " CRYRESB2CS, CRYSMPLIDB3CS, CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS, " & _
              " CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, " & _
              " CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, " & _
              " CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, " & _
              " KDAYCS, SNDKCS, SNDDAYCS ,LIVKCS "
    sqlBase = sqlBase & "From XSDCS"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    Debug.Print sql
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
            .CRYNUMCS = rs("CRYNUMCS")          ' 結晶番号
            .SMPKBNCS = rs("SMPKBNCS")           ' サンプル区分
            .REPSMPLIDCS = rs("REPSMPLIDCS")           ' サンプルNo
            .XTALCS = rs("CRYNUMCS")           ' 結晶番号
            .INPOSCS = rs("INPOSCS")       ' 結晶内位置
            .HINBCS = rs("HINBCS")           ' 品番
            .REVNUMCS = rs("REVNUMCS")           ' 製品番号改訂番号
            .FACTORYCS = rs("FACTORYCS")         ' 工場
            .OPECS = rs("OPECS")         ' 操業条件
            .KTKBNCS = rs("KTKBNCS")             ' 確定区分
            .CRYINDRSCS = rs("CRYINDRSCS")       ' 状態FLG（Rs)
            .CRYRESRS1CS = rs("CRYRESRS1CS")       ' 実績FLG1（Rs)
            .CRYINDOICS = rs("CRYINDOICS")       ' 状態FLG（Oi)
            .CRYRESOICS = rs("CRYRESOICS")       ' 実績FLG（Oi)
            .CRYINDB1CS = rs("CRYINDB1CS")       ' 状態FLG（B1)
            .CRYRESB1CS = rs("CRYRESB1CS")       ' 実績FLG（B1)
            .CRYINDB2CS = rs("CRYINDB2CS")       ' 状態FLG（B2）
            .CRYRESB2CS = rs("CRYRESB2CS")       ' 実績FLG（B2）
            .CRYINDB3CS = rs("CRYINDB3CS")       ' 状態FLG（B3)
            .CRYRESB3CS = rs("CRYRESB3CS")       ' 実績FLG（B3)
            .CRYINDL1CS = rs("CRYINDL1CS")       ' 状態FLG（L1)
            .CRYRESL1CS = rs("CRYRESL1CS")       ' 実績FLG（L1)
            .CRYINDL2CS = rs("CRYINDL2CS")       ' 状態FLG（L2)
            .CRYRESL2CS = rs("CRYRESL2CS")       ' 実績FLG（L2)
            .CRYINDL3CS = rs("CRYINDL3CS")       ' 状態FLG（L3)
            .CRYRESL3CS = rs("CRYRESL3CS")       ' 実績FLG（L3)
            .CRYINDL4CS = rs("CRYINDL4CS")       ' 状態FLG（L4)
            .CRYRESL4CS = rs("CRYRESL4CS")       ' 実績FLG（L4)
            .CRYINDCSCS = rs("CRYINDCSCS")       ' 状態FLG（Cs)
            .CRYRESCSCS = rs("CRYRESCSCS")       ' 実績FLG（Cs)
            .CRYINDGDCS = rs("CRYINDGDCS")       ' 状態FLG（GD)
            .CRYRESGDCS = rs("CRYRESGDCS")       ' 実績FLG（GD)
            .CRYINDTCS = rs("CRYINDTCS")         ' 状態FLG（T)
            .CRYRESTCS = rs("CRYRESTCS")         ' 実績FLG（T)
            .CRYINDEPCS = rs("CRYINDEPCS")       ' 状態FLG（EPD)
            .CRYRESEPCS = rs("CRYRESEPCS")       ' 実績FLG（EPD)
            .SMPLNUMCS = rs("SMPLNUMCS")         ' サンプル枚数
            .SMPLPATCS = rs("SMPLPATCS")         ' サンプルパターン
            .TDAYCS = rs("TDAYCS")         ' 登録日付
            .KDAYCS = rs("KDAYCS")         ' 更新日付
            .SNDKCS = rs("SNDKCS")       ' 送信フラグ
            .SNDDAYCS = rs("SNDDAYCS")       ' 送信日付
            
            .BLKKTFLAGCS = rs("BLKKTFLAGCS")
            .CRYRESRS2CS = rs("CRYRESRS2CS")
            .LIVKCS = rs("LIVKCS")
            .TBKBNCS = rs("TBKBNCS")
            .KSTAFFCS = rs("KSTAFFCS")
            .TSTAFFCS = rs("TSTAFFCS")
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
End Function

'保留 本来は、s_cmzcDBdriverCOMに登録する。管理が森さんなので待ち
'概要      :結晶サンプル管理の削除
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型                  ,説明
'      　　:CrySmpOld　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（旧）
'      　　:CrySmpNew　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（新）
'      　　:戻り値         ,O  ,FUNCTION_RETURN　   ,書き込みの成否
'説明      :受け入れ取り消しの場合
'履歴      :2003/09/25  作成 二渡
Public Function DBDRV_CrySmp_Del(CrySmpOld() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_Del"

    DBDRV_CrySmp_Del = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(CrySmpOld)
        With CrySmpOld(i)
            sql = "Delete XSDCS where "
            sql = sql & "CRYNUMCS = '" & .CRYNUMCS & "' and "
            sql = sql & "TBKBNCS = '" & .TBKBNCS & "'"
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_CrySmp_Del = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_Del = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

Public Function DBDRV_TBCMI002_Exec(record As PURCHASE_CRYSTAL, UpDateFlag As Boolean) As FUNCTION_RETURN

    Dim sql As String
    Dim fullHinban As tFullHinban
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_KCryTbl_Exec"

    DBDRV_TBCMI002_Exec = FUNCTION_RETURN_SUCCESS

    '12桁品番を求める
    If GetLastHinban(record.hinban, fullHinban) = FUNCTION_RETURN_FAILURE Then
        DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
'    If Not UpDateFlag Then
        
        'TBCMI002 へ追加
        
        sql = "insert into TBCMI002 ( "
        sql = sql & "CRYNUM,"
        sql = sql & "INGOTPOS,"
        sql = sql & "LENGTH,"
        sql = sql & "TRANCNT,"
        sql = sql & "KRPROCCD,"
        sql = sql & "PROCCODE,"
        sql = sql & "DMTOP1,"
        sql = sql & "DMTOP2,"
        sql = sql & "DMTAIL1,"
        sql = sql & "DMTAIL2,"
        sql = sql & "NCHPOS,"
        sql = sql & "NCHDPTH,"
        sql = sql & "NCHWIDTH,"
        sql = sql & "BDLNTOP,"
        sql = sql & "BDCDTOP,"
        sql = sql & "BDLNTAIL,"
        sql = sql & "BDCDTAIL,"
        sql = sql & "TSTAFFID,"
        sql = sql & "REGDATE,"
        sql = sql & "KSTAFFID,"
        sql = sql & "UPDDATE,"
        sql = sql & "SENDFLAG,"
        sql = sql & "SENDDATE) "
        With record
            sql = sql & " select "
            sql = sql & "'" & Left(.blkID, 9) & "000" & "',"
            sql = sql & .SXLPOS & ","
            sql = sql & .BlkLen & ","
            sql = sql & "nvl(max(TRANCNT), 0) + 1 ,"
            sql = sql & "'" & .KRPROCCD & "',"
            sql = sql & "'" & .PROCCODE & "',"
            sql = sql & .DMTOP(0) & ","
            sql = sql & .DMTOP(1) & ","
            sql = sql & .DMTAIL(0) & ","
            sql = sql & .DMTAIL(1) & ","
            sql = sql & "'" & .NCHPOS & "',"
            sql = sql & .NCHDPTH(0) & ","
            sql = sql & .NCHWIDTH(0) & ","
            sql = sql & "0,"
            sql = sql & "' ',"
            sql = sql & "0,"
            sql = sql & "' ',"
            sql = sql & "'" & .TSTAFFID & "',"
            sql = sql & "sysdate ,"
            sql = sql & "'" & .TSTAFFID & "',"
            sql = sql & "sysdate,"
            sql = sql & "'0',"
            sql = sql & "sysdate  "
            sql = sql & " from TBCMI002 "
            sql = sql & " where CRYNUM='" & .blkID & "' "
        End With
'    Else
'        '更新時
'
'        With record
'            sql = "UPDATE TBCMI002 SET "
'            sql = sql & "INGOTPOS= " & .SXLPOS & ","
'            sql = sql & "LENGTH= " & .BlkLen & ","
'            sql = sql & "KRPROCCD= '" & .KRPROCCD & "',"
'            sql = sql & "PROCCODE= '" & .PROCCODE & "',"
'            sql = sql & "DMTOP1= " & .DMTOP(0) & ","
'            sql = sql & "DMTOP2= " & .DMTOP(1) & ","
'            sql = sql & "DMTAIL1= " & .DMTAIL(0) & ","
'            sql = sql & "DMTAIL2= " & .DMTAIL(1) & ","
'            sql = sql & "NCHPOS= '" & .NCHPOS & "',"
'            sql = sql & "NCHDPTH= " & .NCHDPTH(0) & ","
'            sql = sql & "NCHWIDTH= " & .NCHWIDTH(0) & ","
'            sql = sql & "TSTAFFID= '" & .TSTAFFID & "',"
'            sql = sql & "REGDATE= sysdate ,"
'            sql = sql & "KSTAFFID= '" & .TSTAFFID & "',"
'            sql = sql & "UPDDATE= sysdate,"
'            sql = sql & "SENDFLAG= '0',"
'            sql = sql & "SENDDATE= sysdate  "
'            sql = sql & "WHERE CRYNUM='" & Left(.blkID, 9) & "000" & "'"
'        End With
'    End If
    If 0 >= OraDB.ExecuteSQL(sql) Then
        DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_TBCMI002_Exec = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'概要      :結晶サンプル管理の挿入／更新
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型                  ,説明
'      　　:CrySmpOld　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（旧）
'      　　:CrySmpNew　　　,I  ,typ_XSDCS   　      ,新サンプル管理（ブロック）（新）
'      　　:戻り値         ,O  ,FUNCTION_RETURN　   ,書き込みの成否
'説明      :古いレコードをみて更新か挿入かを判別する
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_CrySmp_UpdIns037Only(CrySmpOld() As typ_XSDCS, CrySmpNew() As typ_XSDCS) As FUNCTION_RETURN

    Dim sql As String
    Dim lFlg As Boolean
    Dim i As Long
    Dim j As Long
    Dim result As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_CrySmp_UpdIns"

    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_SUCCESS
    
    For i = 1 To UBound(CrySmpNew)
        With CrySmpNew(i)
            lFlg = False
''            For j = 1 To UBound(CrySmpOld)
''                If CrySmpOld(j).XTALCS = .XTALCS And _
''                   CrySmpOld(j).SMPKBNCS = .SMPKBNCS Then
''                    sql = "update XSDCS set "
''                    sql = sql & "HINBCS='" & .HINBCS & "', "                ' 品番
''                    sql = sql & "REVNUMCS=" & .REVNUMCS & ", "              ' 製品番号改訂番号
''                    sql = sql & "FACTORYCS='" & .FACTORYCS & "', "          ' 工場
''                    sql = sql & "OPECS='" & .OPECS & "', "                  ' 操業条件
''                    sql = sql & "KTKBNCS='" & .KTKBNCS & "', "              ' 確定区分
''                    sql = sql & "REPSMPLIDCS='" & Abs(.REPSMPLIDCS) & "', " ' サンプルＮｏ
''                    If .CRYINDRSCS = "2" Then
''                        .CRYINDRSCS = "1"
''                    End If
''                    sql = sql & "CRYINDRSCS='" & .CRYINDRSCS & "', "        ' 状態FLG（Rs)
''                    If .CRYINDOICS = "2" Then
''                        .CRYINDOICS = "1"
''                    End If
''                    sql = sql & "CRYINDOICS='" & .CRYINDOICS & "', "        ' 状態FLG（Oi)
''                    If .CRYINDB1CS = "2" Then
''                        .CRYINDB1CS = "1"
''                    End If
''                    sql = sql & "CRYINDB1CS='" & .CRYINDB1CS & "', "        ' 状態FLG（B1)
''                    If .CRYINDB2CS = "2" Then
''                        .CRYINDB2CS = "1"
''                    End If
''                    sql = sql & "CRYINDB2CS='" & .CRYINDB2CS & "', "        ' 状態FLG（B2)
''                    If .CRYINDB3CS = "2" Then
''                        .CRYINDB3CS = "1"
''                    End If
''                    sql = sql & "CRYINDB3CS='" & .CRYINDB3CS & "', "        ' 状態FLG（B3)
''                    If .CRYINDL1CS = "2" Then
''                        .CRYINDL1CS = "1"
''                    End If
''                    sql = sql & "CRYINDL1CS='" & .CRYINDL1CS & "', "        ' 状態FLG（L1)
''                    If .CRYINDL2CS = "2" Then
''                        .CRYINDL2CS = "1"
''                    End If
''                    sql = sql & "CRYINDL2CS='" & .CRYINDL2CS & "', "        ' 状態FLG（L2)
''                    If .CRYINDL3CS = "2" Then
''                        .CRYINDL3CS = "1"
''                    End If
''                    sql = sql & "CRYINDL3CS='" & .CRYINDL3CS & "', "        ' 状態FLG（L3)
''                    If .CRYINDL4CS = "2" Then
''                        .CRYINDL4CS = "1"
''                    End If
''                    sql = sql & "CRYINDL4CS='" & .CRYINDL4CS & "', "        ' 状態FLG（L4)
''                    If .CRYINDCSCS = "2" Then
''                        .CRYINDCSCS = "1"
''                    End If
''                    sql = sql & "CRYINDCSCS='" & .CRYINDCSCS & "', "        ' 状態FLG（Cs)
''                    If .CRYINDGDCS = "2" Then
''                        .CRYINDGDCS = "1"
''                    End If
''                    sql = sql & "CRYINDGDCS='" & .CRYINDGDCS & "', "        ' 状態FLG（GD)
''                    If .CRYINDTCS = "2" Then
''                        .CRYINDTCS = "1"
''                    End If
''                    sql = sql & "CRYINDTCS='" & .CRYINDTCS & "', "          ' 状態FLG（T)
''                    If .CRYINDEPCS = "2" Then
''                        .CRYINDEPCS = "1"
''                    End If
''                    sql = sql & "CRYINDEPCS='" & .CRYINDEPCS & "', "        ' 状態FLG（EPD)
''
''                    sql = sql & "CRYRESRS1CS='" & .CRYRESRS1CS & "', "      ' 実績FLG1（Rs)
''                    sql = sql & "CRYRESOICS='" & .CRYRESOICS & "', "        ' 実績FLG（Oi)
''                    sql = sql & "CRYRESB1CS='" & .CRYRESB1CS & "', "        ' 実績FLG（B1)
''                    sql = sql & "CRYRESB2CS='" & .CRYRESB2CS & "', "        ' 実績FLG（B2)
''                    sql = sql & "CRYRESB3CS='" & .CRYRESB3CS & "', "        ' 実績FLG（B3)
''                    sql = sql & "CRYRESL1CS='" & .CRYRESL1CS & "', "        ' 実績FLG（L1)
''                    sql = sql & "CRYRESL2CS='" & .CRYRESL2CS & "', "        ' 実績FLG（L2)
''                    sql = sql & "CRYRESL3CS='" & .CRYRESL3CS & "', "        ' 実績FLG（L3)
''                    sql = sql & "CRYRESL4CS='" & .CRYRESL4CS & "', "        ' 実績FLG（L4)
''                    sql = sql & "CRYRESCSCS='" & .CRYRESCSCS & "', "        ' 実績FLG（Cs)
''                    sql = sql & "CRYRESGDCS='" & .CRYRESGDCS & "', "        ' 実績FLG（GD)
''                    sql = sql & "CRYRESTCS='" & .CRYRESTCS & "', "          ' 実績FLG（T)
''                    sql = sql & "CRYRESEPCS='" & .CRYRESEPCS & "', "        ' 実績FLG（EPD)
''                    sql = sql & "SMPLNUMCS=" & .SMPLNUMCS & ", "            ' サンプル枚数
''                    sql = sql & "SMPLPATCS='" & .SMPLPATCS & "', "          ' サンプルパターン
''                    sql = sql & "KDAYCS=sysdate, "                          ' 更新日付
''                    sql = sql & "SNDKCS='0' "                               ' 送信フラグ
''                    sql = sql & " where XTALCS='" & .XTALCS & "'"
''                    sql = sql & " and TBKBNCS='" & .TBKBNCS & "'"
''
''                    WriteDBLog sql
''                    Debug.Print sql
''                    If OraDB.ExecuteSQL(sql) <= 0 Then
''                        DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
''                        GoTo proc_exit
''                    End If
''                    lFlg = True
''                    Exit For
''                End If
''            Next j

            If lFlg <> True Then
                sql = "insert into XSDCS ("
                sql = sql & "CRYNUMCS,"         'ブロックID
                sql = sql & "SMPKBNCS,"         'サンプル区分
                sql = sql & "TBKBNCS,"          'T/B区分
                sql = sql & "REPSMPLIDCS,"      '代表サンプルID
                sql = sql & "XTALCS,"           '結晶番号
                sql = sql & "INPOSCS,"          '結晶内位置
                sql = sql & "HINBCS,"           '品番
                sql = sql & "REVNUMCS,"         '製品番号改訂番号
                sql = sql & "FACTORYCS,"        '工場
                sql = sql & "OPECS,"            '操業番号
                sql = sql & "KTKBNCS,"          '確定区分
                sql = sql & "BLKKTFLAGCS,"      'ブロック確定フラグ
                sql = sql & "CRYSMPLIDRSCS,"    'サンプルID(Rs)
                sql = sql & "CRYSMPLIDRS1CS,"   '推定サンプルID1（Rs）
                sql = sql & "CRYSMPLIDRS2CS,"   '推定サンプルID2（Rs）
                sql = sql & "CRYINDRSCS,"       '状態FLG(Rs)
                sql = sql & "CRYRESRS1CS,"      '実績FLG1(Rs)
                sql = sql & "CRYRESRS2CS,"      '実績FLG2(Rs)
                sql = sql & "CRYSMPLIDOICS,"    'サンプルID（Oi）
                sql = sql & "CRYINDOICS,"       '状態FLG（Oi）
                sql = sql & "CRYRESOICS,"       '実績FLG（Oi）
                sql = sql & "CRYSMPLIDB1CS,"    'サンプルID（B1）
                sql = sql & "CRYINDB1CS,"       '状態FLG（B1）
                sql = sql & "CRYRESB1CS,"       '実績FLG（B1）
                sql = sql & "CRYSMPLIDB2CS,"    'サンプルID（B2）
                sql = sql & "CRYINDB2CS,"       '状態FLG（B2）
                sql = sql & "CRYRESB2CS,"       '実績FLG（B2）
                sql = sql & "CRYSMPLIDB3CS,"    'サンプルID（B3）
                sql = sql & "CRYINDB3CS,"       '状態FLG（B3）
                sql = sql & "CRYRESB3CS,"       '実績FLG（B3）
                sql = sql & "CRYSMPLIDL1CS,"    'サンプルID（L1）
                sql = sql & "CRYINDL1CS,"       '状態FLG（L1）
                sql = sql & "CRYRESL1CS,"       '実績FLG（L1）
                sql = sql & "CRYSMPLIDL2CS,"    'サンプルID（L2）
                sql = sql & "CRYINDL2CS,"       '状態FLG（L2）
                sql = sql & "CRYRESL2CS,"       '実績FLG（L2）
                sql = sql & "CRYSMPLIDL3CS,"    'サンプルID（L3）
                sql = sql & "CRYINDL3CS,"       '状態FLG（L3）
                sql = sql & "CRYRESL3CS,"       '実績FLG（L3）
                sql = sql & "CRYSMPLIDL4CS,"    'サンプルID（L4）
                sql = sql & "CRYINDL4CS,"       '状態FLG（L4）
                sql = sql & "CRYRESL4CS,"       '実績FLG（L4）
                sql = sql & "CRYSMPLIDCSCS,"    'サンプルID（CS）
                sql = sql & "CRYINDCSCS,"       '状態FLG（CS）
                sql = sql & "CRYRESCSCS,"       '実績FLG（CS）
                sql = sql & "CRYSMPLIDGDCS,"    'サンプルID（GD）
                sql = sql & "CRYINDGDCS,"       '状態FLG（GD）
                sql = sql & "CRYRESGDCS,"       '実績FLG（GD）
                sql = sql & "CRYSMPLIDTCS,"     'サンプルID（T）
                sql = sql & "CRYINDTCS,"        '状態FLG（T）
                sql = sql & "CRYRESTCS,"        '実績FLG（T）
                sql = sql & "CRYSMPLIDEPCS,"    'サンプルID（EPD）
                sql = sql & "CRYINDEPCS,"       '状態FLG（EPD）
                sql = sql & "CRYRESEPCS,"       '実績FLG（EPD）
                sql = sql & "SMPLNUMCS,"        'サンプル枚数
                sql = sql & "SMPLPATCS,"        'サンプルパターン
                sql = sql & "TSTAFFCS,"         '登録社員ID
                sql = sql & "TDAYCS,"           '登録日付
                sql = sql & "KSTAFFCS,"         '更新社員ID
                sql = sql & "KDAYCS,"           '更新日付
                sql = sql & "SNDKCS,"           '送信フラグ
                sql = sql & "SNDDAYCS,"         '送信日付
                sql = sql & "LIVKCS)"           '生死区分
                sql = sql & " values ('"
                sql = sql & .CRYNUMCS & "', '"          'ブロックID
                sql = sql & .SMPKBNCS & "', '"          'サンプル区分
                sql = sql & .TBKBNCS & "', "            'T/B区分
                sql = sql & .REPSMPLIDCS & ", '"        '代表サンプルID
                sql = sql & .XTALCS & "', "             '結晶番号
                sql = sql & .INPOSCS & ", '"            '結晶内位置
                sql = sql & .HINBCS & "', "             '品番
                sql = sql & .REVNUMCS & ", '"           '製品番号改訂番号
                sql = sql & .FACTORYCS & "', '"         '工場
                sql = sql & .OPECS & "', '"             '操業条件
                sql = sql & .KTKBNCS & "', '"           '確定区分
                sql = sql & .BLKKTFLAGCS & "', "        'ブロック確定フラグ
                
                If .CRYINDRSCS = "2" Then
                    .CRYINDRSCS = "1"
                End If
                If .CRYINDRSCS = "1" Then
                    .CRYSMPLIDRSCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDRSCS = 0
                End If
                sql = sql & .CRYSMPLIDRSCS & ", "       'サンプルID（Rs）
                sql = sql & .CRYSMPLIDRS1CS & ", "      '推定サンプルID1（Rs）
                sql = sql & .CRYSMPLIDRS2CS & ", '"     '推定サンプルID2（Rs）
                sql = sql & .CRYINDRSCS & "', '"        '状態FLG（Rs）
                sql = sql & .CRYRESRS1CS & "', '"       '実績FLG1（Rs）
                sql = sql & .CRYRESRS2CS & "', "        '実績FLG2（Rs）
                
                If .CRYINDOICS = "2" Then
                    .CRYINDOICS = "1"
                End If
                If .CRYINDOICS = "1" Then
                    .CRYSMPLIDOICS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDOICS = 0
                End If
                sql = sql & .CRYSMPLIDOICS & ", '"      'サンプルID（Oi）
                sql = sql & .CRYINDOICS & "', '"        '状態FLG（Oi）
                sql = sql & .CRYRESOICS & "', "         '実績FLG（Oi）
                
                If .CRYINDB1CS = "2" Then
                    .CRYINDB1CS = "1"
                End If
                If .CRYINDB1CS = "1" Then
                    .CRYSMPLIDB1CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB1CS = 0
                End If
                sql = sql & .CRYSMPLIDB1CS & ", '"      'サンプルID（B1）
                sql = sql & .CRYINDB1CS & "', '"        '状態FLG（B1）
                sql = sql & .CRYRESB1CS & "', "         '実績FLG（B1）
                
                
                If .CRYINDB2CS = "2" Then
                    .CRYINDB2CS = "1"
                End If
                If .CRYINDB2CS = "1" Then
                    .CRYSMPLIDB2CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB2CS = 0
                End If
                sql = sql & .CRYSMPLIDB2CS & ", '"      'サンプルID（B2）
                sql = sql & .CRYINDB2CS & "', '"        '状態FLG（B2）
                sql = sql & .CRYRESB2CS & "', "         '実績FLG（B2）
                
                If .CRYINDB3CS = "2" Then
                    .CRYINDB3CS = "1"
                End If
                If .CRYINDB3CS = "1" Then
                    .CRYSMPLIDB3CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDB3CS = 0
                End If
                sql = sql & .CRYSMPLIDB3CS & ", '"      'サンプルID（B3）
                sql = sql & .CRYINDB3CS & "', '"        '状態FLG（B3）
                sql = sql & .CRYRESB3CS & "', "         '実績FLG（B3）
                
                If .CRYINDL1CS = "2" Then
                    .CRYINDL1CS = "1"
                End If
                If .CRYINDL1CS = "1" Then
                    .CRYSMPLIDL1CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL1CS = 0
                End If
                sql = sql & .CRYSMPLIDL1CS & ", '"      'サンプルID（L1）
                sql = sql & .CRYINDL1CS & "', '"        '状態FLG（L1）
                sql = sql & .CRYRESL1CS & "', "         '実績FLG（L1）
                
                If .CRYINDL2CS = "2" Then
                    .CRYINDL2CS = "1"
                End If
                If .CRYINDL2CS = "1" Then
                    .CRYSMPLIDL2CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL2CS = 0
                End If
                sql = sql & .CRYSMPLIDL2CS & ", '"      'サンプルID（L2）
                sql = sql & .CRYINDL2CS & "', '"        '状態FLG（L2）
                sql = sql & .CRYRESL2CS & "', "         '実績FLG（L2）
                
                If .CRYINDL3CS = "2" Then
                    .CRYINDL3CS = "1"
                End If
                If .CRYINDL3CS = "1" Then
                    .CRYSMPLIDL3CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL3CS = 0
                End If
                sql = sql & .CRYSMPLIDL3CS & ", '"      'サンプルID（L3）
                sql = sql & .CRYINDL3CS & "', '"        '状態FLG（L3）
                sql = sql & .CRYRESL3CS & "', "         '実績FLG（L3）
                
                If .CRYINDL4CS = "2" Then
                    .CRYINDL4CS = "1"
                End If
                If .CRYINDL4CS = "1" Then
                    .CRYSMPLIDL4CS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDL4CS = 0
                End If
                sql = sql & .CRYSMPLIDL4CS & ", '"      'サンプルID（L4）
                sql = sql & .CRYINDL4CS & "', '"        '状態FLG（L4）
                sql = sql & .CRYRESL4CS & "', "         '実績FLG（L4）
                
                If .CRYINDCSCS = "2" Then
                    .CRYINDCSCS = "1"
                End If
                If .CRYINDCSCS = "1" Then
                    .CRYSMPLIDCSCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDCSCS = 0
                End If
                sql = sql & .CRYSMPLIDCSCS & ", '"      'サンプルID（CS）
                sql = sql & .CRYINDCSCS & "', '"        '状態FLG（CS）
                sql = sql & .CRYRESCSCS & "', "         '実績FLG（CS）
                
                If .CRYINDGDCS = "2" Then
                    .CRYINDGDCS = "1"
                End If
                If .CRYINDGDCS = "1" Then
                    .CRYSMPLIDGDCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDGDCS = 0
                End If
                sql = sql & .CRYSMPLIDGDCS & ", '"      'サンプルID（GD）
                sql = sql & .CRYINDGDCS & "', '"        '状態FLG（GD）
                sql = sql & .CRYRESGDCS & "', "         '実績FLG（GD）
                
                If .CRYINDTCS = "2" Then
                    .CRYINDTCS = "1"
                End If
                If .CRYINDTCS = "1" Then
                    .CRYSMPLIDTCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDTCS = 0
                End If
                sql = sql & .CRYSMPLIDTCS & ", '"       'サンプルID（T）
                sql = sql & .CRYINDTCS & "', '"         '状態FLG（T）
                sql = sql & .CRYRESTCS & "', "          '実績FLG（T）
                
                If .CRYINDEPCS = "2" Then
                    .CRYINDEPCS = "1"
                End If
                If .CRYINDEPCS = "1" Then
                    .CRYSMPLIDEPCS = .REPSMPLIDCS
                Else
                    .CRYSMPLIDEPCS = 0
                End If
                sql = sql & .CRYSMPLIDEPCS & ", '"      'サンプルID（EPD）
                sql = sql & .CRYINDEPCS & "', '"        '状態FLG（EPD）
                sql = sql & .CRYRESEPCS & "', "         '実績FLG（EPD）
                sql = sql & .SMPLNUMCS & ", "           'サンプル枚数
                sql = sql & "' ', '"                    'サンプルパターン
                sql = sql & .TSTAFFCS & "', "           '登録社員ID
                sql = sql & "sysdate, '"                '登録日付
                sql = sql & .KSTAFFCS & "', "           '更新社員ID
                sql = sql & "sysdate, "                 '更新日付
                sql = sql & "'0', "                     '送信フラグ
                sql = sql & "sysdate,"                  '送信日付
                sql = sql & "'0')"                      '生死区分
                
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_CrySmp_UpdIns037Only = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


'概要      :テーブルの削除処理
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型                  ,説明
'      　　:sql　　      　,I  ,String      　      ,削除SQL文
'          :sTable        ,I  ,String              ,削除テーブル
'      　　:戻り値         ,O  ,FUNCTION_RETURN　   ,書き込みの成否
'説明      :購入単結晶受入/取消時、既に存在するデータを削除する
'履歴      :2003/10/31 ooba

Public Function DBDRV_DeleteTable(sql As String, sTable As String) As FUNCTION_RETURN


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc037_SQL.bas -- Function DBDRV_DeleteTable"

    DBDRV_DeleteTable = FUNCTION_RETURN_FAILURE
    
Debug.Print sql
    
    If OraDB.ExecuteSQL(sql) < 1 Then
        Debug.Print "<" & sTable & "> 削除データ無し"
    Else
        '' WriteDBLog sql
    End If
    
    DBDRV_DeleteTable = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
    
End Function

