Attribute VB_Name = "s_cmzcXSDC_WF_Basic"
'===============================================================================
'   構造体定義
'===============================================================================
'●基本情報構造体
Public Type type_KIHON      '基本情報構造体
    STAFFID     As String   '担当者コード
    NEWPROC     As String   '次工程
    NOWPROC     As String   '現工程
    DIAMETER    As Long     '直径
    ALLSCRAP    As String   '全数スクラップ（'Y'：あり、'N'：なし）
    FURYOUMU    As String   '不良有無（'Y'：あり、'N'：なし）
    CNTHINOLD   As Integer  '分割結晶（品番：前工程）件数
    CNTHINNOW   As Integer  '分割結晶（品番：良品）件数
End Type

Public Kihon    As type_KIHON          '基本情報
Public BlkOld As typ_XSDC2_Update      '分割結晶(ブロック)：前工程
Public BlkNow As typ_XSDC2_Update      '分割結晶(ブロック)：良品
Public HinOld() As typ_XSDCA_Update    '分割結晶(品番)：前工程
Public HinNow() As typ_XSDCA_Update    '分割結晶(品番)：良品
Public Furyou   As typ_XSDC4_Update    '不良内訳

Private blkInfo As typ_cmkc001f_Block
Private bMapErrFlg As Boolean           'WFﾏｯﾌﾟ位置ﾁｪｯｸﾌﾗｸﾞ

Private HSXCTCEN As Double      ' 品ＳＸ結晶面傾縦中心
Private HSXCYCEN As Double      ' 品ＳＸ結晶面傾横中心

'WF枚数計算用のパラメータ
Private SEEDDEG As Integer      ' SEED傾き
Private Loss0 As Integer        ' 傾き差0度のときの傾きロス
Private Loss4 As Integer        ' 傾き差4度のときの傾きロス
Private Mlt4 As Double          ' 傾き差4度の時の係数
Private Pitch As Double         ' ワイヤソーメインローラピッチ

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
    HSXTMMAX As Double                ' 品ＳＸ転位密度上限
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

'在庫減情報登録用
Public Type typ_stock_info
    hinban As String * 8        ' 品番
    GENZAL As Long              ' 現在長さ
    HARAIL As Long              ' 払い出し長さ
    FURYOL As Long              ' 不良長さ
    GENZAW As Long              ' 現在重量
    HARAIW As Long              '払い出し重量
    FuryoW As Long              ' 不良重量
    GENZAM As Long              ' 現在枚数
    HARAIM As Long              ' 払い出し枚数
    FURYOM As Long              ' 不良枚数
    KCKNT  As Integer           '工程連番
    REVNUM As Integer           ' 製品改訂番号
    factory As String           ' 工場
    OPE As String               ' 操業条件
End Type

'不良情報
Public Type typ_bad_info
    pos As Double              ' 品番
    LEN As Double
End Type

'品番振替情報
Public Type typ_trans_info
    hinban As String * 8        ' 品番
    LEN As Long                 ' 長さ
    WAT As Long                 ' 重量
    MAI As Long                 ' 枚数
    KCKNT  As Integer           ' 工程連番
    REVNUM As Integer           ' 製品改訂番号
    factory As String           ' 工場
    OPE As String               ' 操業条件
End Type

Public STOCKINFO() As typ_stock_info    'XSDC3Proc2()とXSDC3Proc3()で使用
Public giInpos  As Integer
Public strSxlData   As String

'*******************************************************************************
'*    関数名        : KihonProc
'*
'*    処理概要      : 1.画面からの基本処理を行う
'*                      （DBへの登録【XSDC2,XSDC3,XSDC4,XSDCA,XSDCB】）
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function KihonProc() As FUNCTION_RETURN
'   内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          '復帰情報
    Dim sSQL            As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sErrMsg         As String

    'エラーハンドラの設定
    On Error GoTo proc_err

    KihonProc = FUNCTION_RETURN_FAILURE
    'XSDCAProc、XSDC2Procの順番入れ替え
'    ≪分割結晶（品番）登録≫
    intRtn = XSDCAProc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCAProc()：XSDCA登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ≪分割結晶（ブロック）登録≫
    intRtn = XSDC2Proc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDC2Proc()：XSDC2登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ≪不良内訳登録≫
    '不良有無がある時
    If Kihon.FURYOUMU = "Y" Then
                                                ' 登録日付
        Furyou.TDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                ' 更新日付
        Furyou.KDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")

        '不良長さ・重量・枚数を再取得-------------
        Furyou.PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2)
        Furyou.PUCUTWC4 = CLng(BlkOld.GNWC2) - CLng(BlkNow.GNWC2)
        Furyou.PUCUTMC4 = CLng(BlkOld.GNMC2) - CLng(BlkNow.GNMC2)

        '不良内訳追加
        intRtn = CreateXSDC4(Furyou, sErrMsg)

        '不良内訳追加エラー
        If intRtn = FUNCTION_RETURN_FAILURE Then
            MsgBox sErrMsg
            Debug.Print "CreateXSDC4()：XSDC4登録エラー"
            KihonProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    Debug.Print HinNow(0).SXLIDCA

    ' ≪工程実績登録≫
    intRtn = XSDC3Proc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        Debug.Print "XSDC3Proc()：XSDC3登録エラー"
        KihonProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    Select Case Kihon.NOWPROC
        Case "CW740", "CW760"
            ' ≪在庫減情報登録≫
            intRtn = XSDC3Proc2()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3在庫減情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            ' ≪振替情報登録≫
            intRtn = XSDC3Proc3()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3振替情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Case "CC730"
            ' ≪在庫減情報登録≫
            intRtn = XSDC3Proc4()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3在庫減情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            ' ≪振替情報登録≫
            intRtn = XSDC3Proc5()
            If intRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3振替情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    End Select

    Debug.Print HinNow(0).SXLIDCA

    ' ≪分割結晶（ＳＸＬ）登録≫
    intRtn = XSDCBProc()
    If intRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCBProc()：XSDCB登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    KihonProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    KihonProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        :
'*
'*    処理概要      : 1.分割結晶（ブロック）登録処理を行う(XSDC2)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC2Proc()
    ' 内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sSqlWhere       As String           ' WHERE句
    Dim sErrMsg         As String
    Dim intSyoriKaisu   As Integer          ' 現在処理回数
    Dim intHantei       As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err

    XSDC2Proc = FUNCTION_RETURN_FAILURE

    '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
    Select Case Kihon.NOWPROC
        Case "CC730"
            intHantei = CInt(BlkNow.GNLC2)
        Case Else
            intHantei = CInt(BlkNow.GNMC2)
    End Select

    If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
        BlkNow.LSTATBC2 = "H"                   ' 最終状態区分（廃棄）
        BlkNow.LDFRBC2 = "2"                    ' 格下区分（ハイキ）
        BlkNow.LIVKC2 = "1"                     ' 生死区分（死ロット）
        BlkNow.GNWKNTC2 = " "                   ' 現在工程
        BlkNow.GNMACOC2 = "0"                   ' 現在処理回数
    Else
        ' 処理回数取得ロジック変更
        intSyoriKaisu = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.GNWKNTC2)
        If BlkNow.GNWKNTC2 = BlkNow.NEWKNTC2 Then
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        BlkNow.GNMACOC2 = intSyoriKaisu                                 ' 現在処理回数
        BlkNow.NEMACOC2 = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.NEWKNTC2)  ' 最終通過処理回数
    End If


    BlkNow.KDAYC2 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
    BlkNow.PLANTCATC2 = sCmbMukesaki
    sSqlWhere = "WHERE CRYNUMC2 = '" & BlkNow.CRYNUMC2 & "' "

    intRtn = UpdateXSDC2(BlkNow, sSqlWhere)
    '分割結晶（ブロック）更新エラー
    If intRtn = FUNCTION_RETURN_FAILURE Then
        MsgBox "XSDCB UPDATET ERROR"
        Exit Function
    End If

    XSDC2Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC2Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDCAProc
'*
'*    処理概要      : 1.分割結晶（品番）登録処理を行う(XSDCA)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDCAProc()
    ' 内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sSqlWhere       As String           ' WHERE句
    Dim sErrMsg         As String
    Dim intLivFlg       As Integer          ' 存在フラグ
    Dim udtHinban()     As typ_XSDCA        ' 分割結晶（品番）ワーク領域
    Dim udtHinban_UP()  As typ_XSDCA_Update ' 分割結晶（品番）ワーク領域
    Dim intDataCnt      As Integer          ' 該当データ件数
    Dim lngSumGNWCA     As Long
    Dim lngSumGNMCA     As Long
    Dim blChgFlg        As Boolean
    Dim intSyoriKaisu   As Integer          ' 現在処理回数
    Dim intHantei       As Integer
    Dim lngGetLength    As Long             ' TBCME040より、ブロック長さを取得する

    ' エラーハンドラの設定
    On Error GoTo proc_err

    XSDCAProc = FUNCTION_RETURN_FAILURE
    lngSumGNWCA = 0
    lngSumGNMCA = 0
    blChgFlg = False

    ' 品番の重量・枚数計算
    If Kihon.CNTHINNOW = Kihon.CNTHINOLD Then   ' 前工程と現在工程の件数が同じで、各長さも同じ場合は計算処理をしない
        For i = 0 To Kihon.CNTHINNOW - 1
            If CLng(HinNow(i).GNLCA) <> CLng(HinOld(i).GNLCA) Then
                blChgFlg = True
            End If
        Next
    Else            ' 前工程と現在工程の件数が違う場合は、計算処理を行う
        blChgFlg = True
    End If
    ' 重量・枚数計算処理

    If blChgFlg = True Then
        ' CW740,CW760用追加
        If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
            For i = 0 To Kihon.CNTHINNOW - 1
                With HinNow(i)  ' BLKOLD基準に変更
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkOld.GNWC2
                        HinNow(i).GNLCA = BlkOld.GNLC2
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkOld.GNWC2) - lngSumGNWCA
                        HinNow(i).GNLCA = CLng(BlkOld.GNLC2) - lngSumGNLCA
                        'Add Start 2010/10/14 Y.Hitomi
                        If HinNow(i).GNLCA <= 0 And HinNow(i).GNMCA = 1 Then
                            HinNow(i).GNLCA = 1
                        End If
                        'Add End 2010/10/14 Y.Hitomi
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkOld.GNWC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        HinNow(i).GNLCA = Round(CLng(BlkOld.GNLC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        'Add Start 2010/10/14 Y.Hitomi
                        If HinNow(i).GNLCA <= 0 And HinNow(i).GNMCA = 1 Then
                            HinNow(i).GNLCA = 1
                        End If
                        'Add End 2010/10/14 Y.Hitomi
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNLCA = lngSumGNLCA + CLng(HinNow(i).GNLCA)
                        .SUMITLCA = .GNLCA
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    End If
                End With
            Next
        Else
            If BlkOld.GNLC2 = BlkNow.GNLC2 Then ' 長さが同じ場合はBLKOLD基準。長さが異なる場合はBLKNOW基準にする
                BlkNow.GNWC2 = BlkOld.GNWC2
                BlkNow.GNMC2 = BlkOld.GNMC2
            Else
                BlkNow.GNWC2 = Round(CLng(BlkOld.GNWC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
                BlkNow.GNMC2 = Round(CLng(BlkOld.GNMC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
            End If
            For i = 0 To Kihon.CNTHINNOW - 1    ' BLKOLD基準に変更
                With HinNow(i)
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkNow.GNWC2
                        HinNow(i).GNMCA = BlkNow.GNMC2
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkNow.GNWC2) - lngSumGNWCA
                        HinNow(i).GNMCA = CLng(BlkNow.GNMC2) - lngSumGNMCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkNow.GNWC2) * (CLng(HinNow(i).GNLCA) / CLng(BlkNow.GNLC2)))
                        HinNow(i).GNMCA = Round(CLng(BlkNow.GNMC2) * (CLng(HinNow(i).GNLCA) / CLng(BlkNow.GNLC2)))
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNMCA = lngSumGNMCA + CLng(HinNow(i).GNMCA)
                    End If
                End With
            Next
        End If
    End If

    ' XSDC2の重量、長さをXSDCAの合計にするためにBlkNow再計算
    BlkNow.GNLC2 = 0
    BlkNow.GNWC2 = 0
    For i = 0 To Kihon.CNTHINNOW - 1
        BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + CLng(HinNow(i).GNLCA)   '2003/05/24 clng追加
        BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + CLng(HinNow(i).GNWCA)
    Next i

    ' 前工程の分割結晶（品番）と良品情報の品番・位置を比較する
    For i = 0 To Kihon.CNTHINOLD - 1

        intLivFlg = 0
        For j = 0 To Kihon.CNTHINNOW - 1
            If (HinOld(i).HINBCA = HinNow(j).HINBCA) And (HinOld(i).INPOSCA = HinNow(j).INPOSCA) Then
                intLivFlg = 1
            End If
        Next j

        ' 前工程の分割結晶（品番）にあって良品情報にないものは死ロットとする
        If intLivFlg = 0 Then
            sSqlWhere = "WHERE CRYNUMCA = '" & HinOld(i).CRYNUMCA & "' "
            sSqlWhere = sSqlWhere & "AND HINBCA = '" & HinOld(i).HINBCA & "' "
            sSqlWhere = sSqlWhere & "AND INPOSCA = '" & HinOld(i).INPOSCA & "' "
            ReDim udtHinban(0) As typ_XSDCA

            ' データの件数を取得
            intRtn = SelCntXSDCA(sSqlWhere, intDataCnt)
            If intRtn = FUNCTION_RETURN_FAILURE Then    ' エラー
                MsgBox "XSDCA SELECT ERROR"
                Exit Function
            Else                                        ' 正常
                If intDataCnt = 0 Then
                    ' 前工程の情報は必ずあるはずなので、0件はエラー
                    Exit Function
                ElseIf intDataCnt > 0 Then
                    ' 前工程
                    intRtn = DBDRV_GetXSDCA(udtHinban(), sSqlWhere)

                    ' 存在しない時エラー
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA SELECT ERROR"
                        Exit Function
                    End If
                    ReDim udtHinban_UP(0) As typ_XSDCA_Update

                    udtHinban_UP(0).CRYNUMCA = HinOld(i).CRYNUMCA
                    udtHinban_UP(0).INPOSCA = HinOld(i).INPOSCA
                    udtHinban_UP(0).HINBCA = HinOld(i).HINBCA

                    ' 生死区分に死ロットをセット
                    udtHinban_UP(0).LIVKCA = "1"              ' 生死区分
                    udtHinban_UP(0).LSTATBCA = "H"            ' 最終状態区分
                    udtHinban_UP(0).LDFRBCA = "2"             ' 格下区分
                    udtHinban_UP(0).KANKCA = "0"              ' 完了区分
                    udtHinban_UP(0).SUMITBCA = "0"
                    udtHinban_UP(0).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                    udtHinban_UP(0).PLANTCATCA = sCmbMukesaki

                    '分割結晶（品番）を更新
                    intRtn = UpdateXSDCA(udtHinban_UP(0), sSqlWhere)

                    '分割結晶（品番）更新エラー
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA UPDATE ERROR"
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    ' 分割結晶（品番）分繰り返し
    For i = 0 To Kihon.CNTHINNOW - 1
        ' 結晶番号、品番、位置で検索
        sSqlWhere = "WHERE CRYNUMCA = '" & HinNow(i).CRYNUMCA & "' "
        sSqlWhere = sSqlWhere & "AND HINBCA = '" & HinNow(i).HINBCA & "' "
        sSqlWhere = sSqlWhere & "AND INPOSCA = '" & HinNow(i).INPOSCA & "' "

        ' データの件数を取得
        intRtn = SelCntXSDCA(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then    ' エラー
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        Else                                        ' 正常
            ' データがある場合は更新処理
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCA(udtHinban, sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ' 分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(HinNow(i).GNMCA) ' 0枚を廃棄にする処理を、ブロック単位ではなく品番単位に変更
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    HinNow(i).LIVKCA = "1"              ' 生死区分
                    HinNow(i).LSTATBCA = "H"            ' 最終状態区分
                    HinNow(i).LDFRBCA = "2"             ' 格下区分
                    HinNow(i).GNWKNTCA = " "            ' 現在工程
                    HinNow(i).GNMACOCA = "0"            ' 現在処理回数
                Else
                    HinNow(i).LIVKCA = "0"              ' 生死区分（生ロット）
                    HinNow(i).LSTATBCA = "T"            ' 最終状態区分（通常）
                    HinNow(i).LDFRBCA = "0"             ' 格下区分（通常）

                    ' 処理回数取得ロジック変更
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu    '現在処理回数
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '最終通過処理回数
                End If
                ' 完了区分フラグ変更
                HinNow(i).KANKCA = "0"              ' 完了区分
                HinNow(i).SUMITBCA = "0"
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                HinNow(i).PLANTCATCA = sCmbMukesaki

                ' 良品情報で置き換え
                intRtn = UpdateXSDCA(HinNow(i), sSqlWhere)

                ' 分割結晶（品番）更新エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA UPDATE ERROR"
                    Exit Function
                End If
            ' 存在しない時追加
            ElseIf intDataCnt = 0 Then
                '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(HinNow(i).GNMCA) ' 0枚を廃棄にする処理を、ブロック単位ではなく品番単位に変更
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    HinNow(i).LIVKCA = "1"              ' 生死区分
                    HinNow(i).LSTATBCA = "H"            ' 最終状態区分
                    HinNow(i).LDFRBCA = "2"             ' 格下区分
                    HinNow(i).GNWKNTCA = " "            ' 現在工程
                    HinNow(i).GNMACOCA = "0"            ' 現在処理回数
                Else
                    HinNow(i).LIVKCA = "0"              ' 生死区分（生ロット）
                    HinNow(i).LSTATBCA = "T"            ' 最終状態区分（通常）
                    HinNow(i).LDFRBCA = "0"             ' 格下区分（通常）

                    ' 処理回数取得ロジック変更
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu  ' 現在処理回数
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '最終通過処理回数
                End If

                ' 完了区分フラグ変更
                HinNow(i).KANKCA = "0"                  ' 完了区分
                                                        ' 登録日付
                HinNow(i).TDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                HinNow(i).SUMITBCA = "0"
                HinNow(i).PLANTCATCA = sCmbMukesaki

                intRtn = CreateXSDCA(HinNow(i), sErrMsg)

                '分割結晶（品番）更新エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDCAProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDCAProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDC3Proc
'*
'*    処理概要      : 1.工程実績登録処理を行う(XSDC3)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc()
    ' 内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sSqlWhere       As String           ' WHERE句
    Dim lngFULC3        As Long             ' 分割結晶（品番）の不良長さ
    Dim lngFUWC3        As Long             ' 分割結晶（品番）の不良重量
    Dim lngFUMC3        As Long             ' 分割結晶（品番）の不良枚数
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update ' 工程実績
    Dim rsKCNTC         As OraDynaset       ' レコードセット
    Dim intNextCnt      As Integer
    Dim intOldCnt       As Integer
    Dim blNewRec        As Boolean          ' 前工程の無いレコードがあった場合
    Dim sSUMITLC3       As String           ' SUMIT長さ
    Dim sSUMITWC3       As String           ' SUMIT重量
    Dim sSUMITMC3       As String           ' SUMIT枚数
    Dim dtmSumcoTime    As Date             ' SUMCO時間
    Dim vChoseiTime     As Variant          ' 調整時間
    Dim intLoopCnt      As Integer
    Dim intHantei       As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err

    blNewRec = False                        ' フラグ初期化
    XSDC3Proc = FUNCTION_RETURN_FAILURE

    ' 工程実績からブロックＩＤ、品番が一致する工程連番の最大を取得
    sSQL = "SELECT MAX(KCNTC3) as wKCNTC3 "
    sSQL = sSQL & " FROM XSDC3 "
    sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、エラー
    If rs Is Nothing Then
        MsgBox "XSDC3 MAX KCNT SELECT ERROR"
        Exit Function
    End If

    If rs.EOF = False Then
        If IsNull(rs.Fields("wKCNTC3")) = True Then
            intNextCnt = 1
        Else
            intNextCnt = CInt(rs.Fields("wKCNTC3")) + 1
        End If
    Else
        intNextCnt = 1
    End If
    rs.Close

    ' 前工程実績の無いレコードがあるかチェックし、あったらフラグをたてる
    For i = 0 To Kihon.CNTHINNOW - 1
        ' 工程実績から前工程のデータを読み込む
        sSQL = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3,WKKTC3,MACOC3 "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sSQL = sSQL & " AND KCNTC3 = " & intNextCnt - 1 & ""  ' intOldCntは使えないので、intNextCnt - 1を変わりに使用

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、エラー
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            blNewRec = True  ' 前工程がない場合はフラグをたてる
        End If
        rs.Close
    Next

    If Kihon.NOWPROC = "CC730" Then
        blNewRec = True
    End If

    For i = 0 To Kihon.CNTHINNOW - 1
        ' 不良内訳からブロックＩＤ、品番、開始位置が一致する不良長さを取得する
        intOldCnt = 0
        lngFULC3 = 0
        lngFUWC3 = 0
        lngFUMC3 = 0

        ' 工程実績からブロックＩＤ、品番が一致する工程連番の最大を取得
        sSQL = "SELECT MAX(KCNTC3) as wKCNTC3 "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、エラー
        If rs Is Nothing Then
            MsgBox "XSDC3 MAX KCNT SELECT ERROR"
            Exit Function
        End If

        If rs.EOF = False Then
            If IsNull(rs.Fields("wKCNTC3")) = True Then
                intOldCnt = 0
            Else
                intOldCnt = CInt(rs.Fields("wKCNTC3"))
            End If
        Else
            intOldCnt = 0
        End If
        rs.Close

        ' 工程実績から前工程ののデータを読み込む
        sSQL = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3, "
        sSQL = sSQL & " SUMITLC3, SUMITWC3, SUMITMC3"
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sSQL = sSQL & " AND KCNTC3 = " & intOldCnt & ""

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、エラー
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            lngFULC3 = 0
            lngFUWC3 = 0
            lngFUMC3 = 0
            sSUMITLC3 = "0" 'SUMIT長さ
            sSUMITWC3 = "0" 'SUMIT重量
            sSUMITMC3 = "0" 'SUMIT枚数
        Else
            If IsNull(rs.Fields("STATIMEC3")) = True Then
                ' 何も入れない
            Else
                udtKoutei.STATIMEC3 = rs.Fields("STATIMEC3")
            End If
                                                    ' 処理時間終了
            If IsNull(rs.Fields("STOTIMEC3")) = True Then
                ' 何も入れない
            Else
                udtKoutei.STOTIMEC3 = rs.Fields("STOTIMEC3")
            End If

            If IsNull(rs.Fields("TOLC3")) = True Then   ' 不良長さ
                lngFULC3 = 0
                udtKoutei.FRLC3 = "0"
            Else
                udtKoutei.FRLC3 = CLng(rs.Fields("TOLC3"))
                lngFULC3 = CLng(rs.Fields("TOLC3"))
            End If
            If IsNull(rs.Fields("TOWC3")) = True Then   ' 不良重量
                lngFUWC3 = 0
                udtKoutei.FRWC3 = "0"
            Else
                udtKoutei.FRWC3 = CLng((rs.Fields("TOWC3")))
                lngFUWC3 = CLng((rs.Fields("TOWC3")))
            End If
            If IsNull(rs.Fields("TOMC3")) = True Then   ' 不良枚数
                lngFUMC3 = 0
                udtKoutei.FRMC3 = "0"
            Else
                udtKoutei.FRMC3 = CLng((rs.Fields("TOMC3")))
                lngFUMC3 = CLng((rs.Fields("TOMC3")))
            End If
            If IsNull(rs.Fields("SUMITLC3")) = True Then   ' SUMIT長さ
                sSUMITLC3 = "0"
            Else
                sSUMITLC3 = CLng((rs.Fields("SUMITLC3")))
            End If
            If IsNull(rs.Fields("SUMITWC3")) = True Then   ' SUMIT重量
                sSUMITWC3 = "0"
            Else
                sSUMITWC3 = CLng((rs.Fields("SUMITWC3")))
            End If
            If IsNull(rs.Fields("SUMITMC3")) = True Then   ' SUMIT枚数
                sSUMITMC3 = "0"
            Else
                sSUMITMC3 = CLng((rs.Fields("SUMITMC3")))
            End If
        End If

        ' 処理回数取得ロジック変更
        If IsNull(HinOld(0).NEWKNTCA) = True Then          ' (受入）工程
            udtKoutei.FRWKKTC3 = "0"
        Else
            udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
        End If

        If IsNull(HinOld(0).NEMACOCA) = True Then          '（受入）処理回数
            udtKoutei.FRMACOC3 = "0"
        Else
            udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
        End If

        ' 分割結晶（品番）から工程実績を追加
        udtKoutei.CRYNUMC3 = HinNow(i).CRYNUMCA            ' ﾌﾞﾛｯｸID･結晶番号
        udtKoutei.INPOSC3 = HinNow(i).INPOSCA              ' 結晶内開始位置
        udtKoutei.KCNTC3 = intNextCnt                      ' 工程連番
        udtKoutei.HINBC3 = HinNow(i).HINBCA                ' 品番
        udtKoutei.REVNUMC3 = HinNow(i).REVNUMCA            ' 製品番号改訂番号
        udtKoutei.FACTORYC3 = HinNow(i).FACTORYCA          ' 工場
        udtKoutei.OPEC3 = HinNow(i).OPECA                  ' 操業条件
        udtKoutei.PLANTCATC3 = HinNow(i).PLANTCATCA        ' 向先

        '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.LENC3 = 0                            ' 長さ
        Else
            udtKoutei.LENC3 = HinNow(i).GNLCA              ' 長さ
        End If

        udtKoutei.XTALC3 = HinNow(i).XTALCA                ' 結晶番号
        udtKoutei.SXLIDC3 = HinNow(i).SXLIDCA              ' SXLID

        Select Case Kihon.NOWPROC                          ' CW740，CW760工程で、管理工程に現在工程＋３を書き込む
            Case "CW740", "CW760", "CC730"
                udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                      CStr(CInt(Right(Kihon.NOWPROC, 1)) + 3)   ' 管理工程(現在工程+3)
            Case Else
                udtKoutei.KNKTC3 = HinNow(i).GNKKNTCA           ' 管理工程
        End Select
        udtKoutei.WKKTC3 = Kihon.NOWPROC                   ' 工程
        udtKoutei.WKKBC3 = HinNow(i).GNWKKBCA              ' 作業区分
        udtKoutei.MACOC3 = HinNow(i).NEMACOCA              ' 処理回数
        udtKoutei.MODKBC3 = ""                             ' 赤黒区分
        udtKoutei.SUMKBC3 = "0"                            ' 集計区分
        udtKoutei.FRKNKTC3 = " "                           ' (受入)管理工程
        udtKoutei.FRWKKBC3 = " "                           ' (受入)作業区分
        udtKoutei.TOWNKTC3 = " "                           ' (払出)管理工程

        ' 分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.TOWKKTC3 = " "                       ' (払出)工程
            udtKoutei.TOMACOC3 = "0"                       ' (払出)処理回数
        Else
            udtKoutei.TOWKKTC3 = HinNow(i).GNWKNTCA        ' (払出)工程
        End If

        udtKoutei.TOMACOC3 = HinNow(i).GNMACOCA            ' (払出)処理回
        udtKoutei.LOSWC3 = ""                              ' ロス長さ
        udtKoutei.LOSLC3 = ""                              ' ロス重量
        udtKoutei.LOSMC3 = ""                              ' ロス枚数

        If blNewRec = True Then                            ' 前工程が無いデータが存在している場合は、払出し数量を受入数量に入れる
            udtKoutei.FRLC3 = HinNow(i).GNLCA              ' 受入長さ<=払出長さ
            udtKoutei.FRWC3 = HinNow(i).GNWCA              ' 受入重量<=払出重量
            udtKoutei.FRMC3 = HinNow(i).GNMCA              ' 受入枚数<=払出枚数
            udtKoutei.TOLC3 = HinNow(i).GNLCA              ' 払出長さ
            udtKoutei.TOWC3 = HinNow(i).GNWCA              ' 払出重量（関数）
            udtKoutei.TOMC3 = HinNow(i).GNMCA              ' 払出枚数（関数）
            udtKoutei.FULC3 = 0                            ' 不良長さ
            udtKoutei.FUWC3 = 0                            ' 不良重量
            udtKoutei.FUMC3 = 0                            ' 不良枚数
        Else
            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOLC3 = 0                             ' 払出長さ
                udtKoutei.TOWC3 = 0                             ' 払出重量（関数）
                udtKoutei.TOMC3 = 0                             ' 払出枚数（関数）
            Else
                udtKoutei.TOLC3 = HinNow(i).GNLCA               ' 払出長さ
                udtKoutei.TOWC3 = HinNow(i).GNWCA               ' 払出重量（関数）
                udtKoutei.TOMC3 = HinNow(i).GNMCA               ' 払出枚数（関数）
            End If
            udtKoutei.FULC3 = lngFULC3 - CLng(udtKoutei.TOLC3)  ' 不良長さ
            udtKoutei.FUWC3 = lngFUWC3 - CLng(udtKoutei.TOWC3)  ' 不良重量
            udtKoutei.FUMC3 = lngFUMC3 - CLng(udtKoutei.TOMC3)  ' 不良枚数
        End If
        If udtKoutei.TOLC3 = "" Then
            udtKoutei.TOLC3 = "0"
        End If
        If udtKoutei.TOWC3 = "" Then
            udtKoutei.TOWC3 = "0"
        End If
        If udtKoutei.TOMC3 = "" Then
            udtKoutei.TOMC3 = "0"
        End If
        ' SUMIT長さに工程別に値をセットする--------------------
        udtKoutei.SUMITLC3 = 0                                  ' SUMIT長さ
        udtKoutei.SUMITWC3 = 0                                  ' SUMIT重量
        udtKoutei.SUMITMC3 = 0                                  ' SUMIT枚数

        For intLoopCnt = 0 To Kihon.CNTHINOLD - 1
            If (udtKoutei.CRYNUMC3 = HinOld(intLoopCnt).CRYNUMCA) _
                And (udtKoutei.INPOSC3 = HinOld(intLoopCnt).INPOSCA) Then
                    udtKoutei.SUMITLC3 = HinOld(intLoopCnt).SUMITLCA     ' SUMIT長さ=前工程SUMIT長さ
                    udtKoutei.SUMITWC3 = HinOld(intLoopCnt).SUMITWCA     ' SUMIT重量=前工程SUMIT重量
                    udtKoutei.SUMITMC3 = HinOld(intLoopCnt).SUMITMCA     ' SUMIT枚数=前工程SUMIT枚数
                    Exit For
            End If
        Next
        udtKoutei.MOTHINC3 = " "                                ' 振替品番(元)
        udtKoutei.XTWORKC3 = "42"                               ' 製造工場
        udtKoutei.WFWORKC3 = " "                                ' ｳｪｰﾊ製造
        udtKoutei.HOLDCC3 = " "                                 ' ホールドコード
        udtKoutei.HOLDBC3 = "0"                                 ' ホールド区分
        udtKoutei.LDFRCC3 = " "                                 ' 格下コード

        '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
        Select Case Kihon.NOWPROC
            Case "CC730"
                intHantei = CInt(BlkNow.GNLC2)
            Case Else
                intHantei = CInt(BlkNow.GNMC2)
        End Select

        If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
            udtKoutei.LDFRBC3 = "2"                             ' 格下区分（ハイキ）
        Else
            udtKoutei.LDFRBC3 = "0"                             ' 格下区分
        End If
        udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' 登録社員ID

        udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
        udtKoutei.KSTAFFC3 = Kihon.STAFFID                      ' 更新社員ID
        udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
        udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)
        udtKoutei.SUMITBC3 = "0"                                ' SUMIT送信フラグ
        udtKoutei.SNDKC3 = "0"                                  ' 送信フラグ
        udtKoutei.MODMACOC3 = "00"                              ' 赤黒の処理回数
        udtKoutei.KAKUCC3 = " "                                 ' 確定コード

        ' 画面で使用しているデータのみ更新を行う
        udtKoutei.PLANTCATC3 = sCmbMukesaki

        Select Case Kihon.NOWPROC
            Case "CW750"                                        ' 総合判定
                If udtKoutei.SXLIDC3 = Trim(f_cmbc039_2.txtSxlID.text) Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)
                    ' 工程実績追加エラー
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case "CW760"    ' 再抜試
                If (SIngotP <= udtKoutei.INPOSC3) And (udtKoutei.INPOSC3 < EIngotP) Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)
                    ' 工程実績追加エラー
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case "CW800"    ' シングル確定
                If udtKoutei.SXLIDC3 = strSxlData Then
                    intRtn = CreateXSDC3(udtKoutei, sErrMsg)

                    ' 工程実績追加エラー
                    If intRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox sErrMsg
                        Exit Function
                    End If
                End If
            Case Else
                intRtn = CreateXSDC3(udtKoutei, sErrMsg)

                ' 工程実績追加エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
        End Select
    Next i

    XSDC3Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDCBProc
'*
'*    処理概要      : 1.分割結晶（udtSXL）登録処理を行う(XSDCB)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDCBProc()
    ' 内部変数
    Dim i               As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sSqlWhere       As String           ' WHERE句
    Dim lngGNLCA        As Long             ' 分割結晶（品番）の合計長さ
    Dim lngGNMCA        As Long             ' 分割結晶（品番）の合計枚数
    Dim sErrMsg         As String
    Dim udtSXL()        As typ_XSDCB_Update ' 分割結晶(udtSXL)
    Dim udtWSXL()       As typ_XSDCB        ' 分割結晶(udtSXL)
    Dim intDataCnt      As Integer          ' 該当データ件数
    Dim sBlockId        As String
    Dim lngRYOMAI       As Long             ' 工程毎の良品枚数
    Dim lngFRYMAI       As Long             ' 工程毎の不良品枚数
    Dim lngLen          As Long             ' 長さ
    Dim lngMAI          As Long             ' 枚数
    Dim lngMAI800       As Long             ' CW800枚数
    Dim lngFUR          As Long             ' 不良枚数
    Dim lngFURKEI       As Long             ' 不良枚数合計
    Dim lngSAM          As Long             ' サンプル枚数
    Dim lngSIJ          As Long             ' サンプル抜指示枚数
    Dim lngSAMFUR       As Long             ' サンプル抜指示不良枚数
    Dim intLoopBkHinGet As Integer          ' 元品番
    Dim m               As Integer          ' カウンタ

    'エラーハンドラの設定
    On Error GoTo proc_err

    XSDCBProc = FUNCTION_RETURN_FAILURE

    For i = 0 To Kihon.CNTHINNOW - 1
        ' 工程実績から同じＳＸＬＩＤの長さ、枚数、不良枚数の合計を取得
        intRtn = XSDCBSum(Kihon.NOWPROC, HinNow(i).SXLIDCA, lngLen, lngMAI, lngMAI800, lngFUR, lngFURKEI, lngSAM, wSAMSIJ, lngSAMFUR)

        ' 分割結晶（品番）：良品のＳＸＬＩＤで分割結晶（udtSXL）を検索
        sSqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
        ReDim udtWSXL(0) As typ_XSDCB

        ' データの件数を取得
        intRtn = SelCntXSDCB(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then            ' エラー
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                                ' 正常
            ' データが存在する場合はUPDATE
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCB(udtWSXL(), sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then    ' エラー
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ReDim udtSXL(0) As typ_XSDCB_Update

                ' 分割結晶（udtSXL）を更新
                udtSXL(0).LENCB = lngLen

                ' 更新時も品番変更する
                udtSXL(0).HINBCB = HinNow(i).HINBCA         ' 品番
                udtSXL(0).MAICB = lngMAI
                udtSXL(0).KCNTCB = BlkNow.KCNTC2            ' 工程連番

                ' シングル確定時、最終状態区分='S'にする
                udtSXL(0).LIVKCB = "0"
                udtSXL(0).KANKCB = "0"                      ' 完了区分
                udtSXL(0).LSTCCB = "T"                      ' 最終状態区分
                udtSXL(0).LDFRBCB = "0"                     ' 格下区分
                udtSXL(0).LENCB = lngLen                    ' 長さ

                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                    udtSXL(0).LSTCCB = "S"
                End If

                ' 工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                     Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR       ' 廃棄WF枚数
                        udtSXL(0).NEWKNTCB = "CW740"       ' 最終通過工程
                        udtSXL(0).GNWKNTCB = "CW750"       ' 現在工程
                     Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ         ' サンプル抜指示枚数
                        udtSXL(0).SNMAICB = lngSAMFUR      ' サンプル抜指示不良枚数
                        udtSXL(0).STMAICB = lngSAM         ' サンプル枚数

                        If SelectSxlID039 = HinNow(i).SXLIDCA Then
                           udtSXL(0).NEWKNTCB = "CW750"    ' 最終通過工程
                           udtSXL(0).GNWKNTCB = "CW800"    ' 現在工程
                        End If

                        ' 処理SXL以外は､最終状態区分を変更しない
                        If Trim(HinNow(i).SXLIDCA) <> SelectSxlID039 Then
                            udtSXL(0).LSTCCB = udtWSXL(1).LSTCCB  ' 最終状態区分
                        End If
                     Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR       ' 廃棄WF枚数

                        ' Z品番の場合は廃棄とする。
                        For m = 1 To UBound(tblWfSxlMng())
                            If HinNow(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                            udtSXL(0).NEWKNTCB = "CW760"       ' 最終通過工程
                            udtSXL(0).GNWKNTCB = "CW750"       ' 現在工程
                            If Trim(udtWSXL(1).FACTORYCB) = "" Then udtSXL(0).FACTORYCB = HinNow(i).FACTORYCA   ' 工場
                            If Trim(udtWSXL(1).OPECB) = "" Then udtSXL(0).OPECB = HinNow(i).OPECA               ' 操業条件
                            If Trim(udtWSXL(1).MOTHINCB) = "" Then
                                udtSXL(0).MOTHINCB = vbNullString ' 初期化
                                udtSXL(0).INPOSCB = udtWSXL(1).INPOSCB
                                If udtSXL(0).INPOSCB = "" Then udtSXL(0).INPOSCB = HinNow(i).INPOSCA
                                For intLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                                    If (CInt(HinOld(intLoopBkHinGet).INPOSCA) <= CInt(udtSXL(0).INPOSCB)) And (CInt(udtSXL(0).INPOSCB) <= CInt(HinOld(intLoopBkHinGet).INPOSCA) + CInt(HinOld(intLoopBkHinGet).GNLCA)) Then
                                        udtSXL(0).MOTHINCB = HinOld(intLoopBkHinGet).HINBCA
                                        Exit For
                                    End If
                                Next

                                ' もし該当HINOLDが無かったら自分の品番を元品番とする
                                If udtSXL(0).MOTHINCB = vbNullString Then
                                    udtSXL(0).MOTHINCB = udtSXL(0).HINBCB
                                End If
                            End If
                               ' 廃棄の場合
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"         ' 最終通過工程
                                   udtSXL(0).GNWKNTCB = "TX860"         ' 現在工程
                                   udtSXL(0).LSTCCB = "H"               ' 最終状態区分
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI800                 ' SXL指示（良品）
                        udtSXL(0).WFCNMAICB = lngFURKEI                 ' WFC内欠落枚数
                End Select
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba

                intRtn = UpdateXSDCB(udtSXL(0), sSqlWhere)

                ' 分割結晶（udtSXL）更新エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             ' 存在しない時、追加
             ElseIf intDataCnt = 0 Then
                ReDim udtSXL(0) As typ_XSDCB_Update
                udtSXL(0).SXLIDCB = HinNow(i).SXLIDCA                   ' SXLID
                udtSXL(0).KCNTCB = BlkNow.KCNTC2                        ' 工程連番
                udtSXL(0).XTALCB = HinNow(i).XTALCA                     ' 結晶番号
                udtSXL(0).INPOSCB = HinNow(i).INPOSCA                   ' 結晶内開始位置
                udtSXL(0).LENCB = lngLen                                ' 長さ
                udtSXL(0).HINBCB = HinNow(i).HINBCA                     ' 品番
                udtSXL(0).REVNUMCB = HinNow(i).REVNUMCA                 ' 電話番号改訂番号
                udtSXL(0).FACTORYCB = HinNow(i).FACTORYCA               ' 工場
                udtSXL(0).OPECB = HinNow(i).OPECA                       ' 操業条件
                udtSXL(0).MAICB = lngMAI                                ' 実枚数
                udtSXL(0).WSRMAICB = 0                                  ' WS洗後枚数
                udtSXL(0).WSNMAICB = 0                                  ' WS洗浄欠落枚数
                udtSXL(0).WFCMAICB = 0                                  ' 受入枚数
                udtSXL(0).SXLRMAICB = 0                                 ' SXL指示(良品)
                udtSXL(0).SXLEMAICB = 0                                 ' SXL確定枚数
                udtSXL(0).FURIMAICB = ""                                ' 振替枚数
                udtSXL(0).XTWORKCB = "42"                               ' 製造工場
                udtSXL(0).WFWORKCB = " "                                ' ウェーハ製造
                udtSXL(0).FURYCCB = " "                                 ' 不良理由
                udtSXL(0).LSTCCB = "T"                                  ' 採取状態区分

                ' シングル確定時、最終状態区分='S'にする
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).LUFRCCB = " "                                 ' 格上コード
                udtSXL(0).LUFRBCB = " "                                 ' 格上区分
                udtSXL(0).LDERCCB = " "                                 ' 格下コード
                udtSXL(0).LDFRBCB = "0"                                 ' 格下区分
                udtSXL(0).HOLDCCB = " "                                 ' ホールドコード
                udtSXL(0).HOLDBCB = " "                                 ' ホールド区分
                udtSXL(0).EXKUBCB = " "                                 ' 例外区分
                udtSXL(0).HENPKCB = " "                                 ' 返品区分
                udtSXL(0).KANKCB = "0"                                  ' 完了区分
                udtSXL(0).LIVKCB = "0"                                  ' 生死区分
                udtSXL(0).NFCB = "0"                                    ' 入庫区分
                udtSXL(0).SAKJCB = "0"                                  ' 削除区分
                udtSXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                udtSXL(0).SUMITCB = "0"                                 ' SUMIT送信フラグ
                udtSXL(0).SNDKCB = "0"                                  ' 返品区分
                udtSXL(0).SNDAYCB = ""                                  ' 送信日付
                udtSXL(0).PLANTCATCB = HinNow(i).PLANTCATCA             ' 向先
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba


                ' 工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR                    ' 廃棄WF枚数
                        udtSXL(0).NEWKNTCB = "CW740"                    ' 最終通過工程
                        udtSXL(0).GNWKNTCB = "CW750"                    ' 現在工程
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ                      ' サンプル抜指示枚数
                        udtSXL(0).SNMAICB = lngSAMFUR                   ' サンプル抜指示不良枚数
                        udtSXL(0).STMAICB = lngSAM                      ' サンプル枚数

                        ' ブロック単位で変更されてしまうためコメント化
                        If SelectSxlID039 = HinNow(i).SXLIDCA Then
                           udtSXL(0).NEWKNTCB = "CW750"                 ' 最終通過工程
                           udtSXL(0).GNWKNTCB = "CW800"                 ' 現在工程
                        End If
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR                    ' 廃棄WF枚数
                        udtSXL(0).NEWKNTCB = "CW760"                    ' 最終通過工程
                        udtSXL(0).GNWKNTCB = "CW750"                    ' 現在工程

                        ' Z品番の場合は廃棄とする。
                        For m = 1 To UBound(tblWfSxlMng())
                           If HinNow(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"         ' 最終通過工程
                                   udtSXL(0).GNWKNTCB = "TX860 "        ' 現在工程
                                   udtSXL(0).LSTCCB = "H"               ' 最終状態区分

                                   ' 登録の場合のみ設定
                                   udtSXL(0).FURYCCB = tblWfSxlMng(m).BDCAUS        ' 不良理由
                                   udtSXL(0).RLENCB = tblWfSxlMng(m).LENGTH         ' 理論長さ
                                   udtSXL(0).SHOLDCLSCB = tblWfSxlMng(m).HOLDCLS    ' ホールド区分
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI                    ' SXL指示（良品）
                        udtSXL(0).WFCNMAICB = lngFURKEI                 ' WFC内欠落枚数
                End Select

                ' 前ﾃﾞｰﾀがなく、元品番が取得できないので、HINOLDから該当位置の品番を取得し、それを元品番とする　---------------
                udtSXL(0).MOTHINCB = vbNullString '初期化
                ' CW740,CW760のみに変更
                If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
                    For intLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                        If (CInt(HinOld(intLoopBkHinGet).INPOSCA) <= CInt(udtSXL(0).INPOSCB)) And (CInt(udtSXL(0).INPOSCB) <= CInt(HinOld(intLoopBkHinGet).INPOSCA) + CInt(HinOld(intLoopBkHinGet).GNLCA)) Then
                            udtSXL(0).MOTHINCB = HinOld(intLoopBkHinGet).HINBCA
                            Exit For
                        End If
                    Next
                   If udtSXL(0).MOTHINCB = vbNullString Then ' もし該当HINOLDが無かったら自分の品番を元品番とする
                       udtSXL(0).MOTHINCB = udtSXL(0).HINBCB
                   End If
                End If
                intRtn = CreateXSDCB(udtSXL(0), sErrMsg)

                ' 分割結晶（udtSXL）追加エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
             End If
        End If
    Next i

''''' sta：分割結晶（品番）の合計長さが0になった時、
''''' 分割結晶（udtSXL）を死ロットにする
    For i = 0 To Kihon.CNTHINOLD - 1
        ' 工程実績から同じＳＸＬＩＤの長さ、枚数、不良枚数の合計を取得
        intRtn = XSDCBSum(Kihon.NOWPROC, HinOld(i).SXLIDCA, lngLen, lngMAI, pMAI800, lngFUR, lngFURKEI, lngSAM, wSAMSIJ, lngSAMFUR)

        ' 分割結晶（品番）：ＳＸＬＩＤで分割結晶（udtSXL）を検索
        sSqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
        ReDim udtWSXL(0) As typ_XSDCB

        ' データの件数を取得
        intRtn = SelCntXSDCB(sSqlWhere, intDataCnt)
        If intRtn = FUNCTION_RETURN_FAILURE Then    ' エラー
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                        ' 正常
            ' データが存在する場合はUPDATE
            If intDataCnt > 0 Then
                intRtn = DBDRV_GetXSDCB(udtWSXL(), sSqlWhere)
                If intRtn = FUNCTION_RETURN_FAILURE Then  'エラー
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                ReDim udtSXL(0) As typ_XSDCB_Update

                ' 分割結晶（udtSXL）を更新
                udtSXL(0).LENCB = lngLen
                udtSXL(0).MAICB = lngMAI
                udtSXL(0).KCNTCB = BlkNow.KCNTC2

                '長さが0の時、死ロットとする
                If (lngMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (lngLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     udtSXL(0).LIVKCB = "1"                 ' 生死区分
                     udtSXL(0).KANKCB = "2"                 ' 完了区分
                     udtSXL(0).LSTCCB = "H"                 ' 最終状態区分
                     udtSXL(0).LDFRBCB = "2"                ' 格下区分
                 Else
                     udtSXL(0).LIVKCB = "0"
                     udtSXL(0).KANKCB = "0"                 ' 完了区分
                     udtSXL(0).LSTCCB = "T"                 ' 最終状態区分
                     udtSXL(0).LDFRBCB = "0"                ' 格下区分
                End If
                udtSXL(0).KANKCB = "0"                      ' 完了区分

                ' 工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR                ' 廃棄WF枚数
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ                  ' サンプル抜指示枚数
                        udtSXL(0).SNMAICB = lngSAMFUR               ' サンプル抜指示不良枚数
                        udtSXL(0).STMAICB = lngSAM                  ' サンプル枚数

                        ' 処理SXL以外は､最終状態区分を変更しない
                        If Trim(HinOld(i).SXLIDCA) <> SelectSxlID039 Then
                            udtSXL(0).LSTCCB = udtWSXL(1).LSTCCB    ' 最終状態区分
                        End If
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR               ' 廃棄WF枚数

                        ' Z品番の場合は廃棄とする。
                        For m = 1 To UBound(tblWfSxlMng())
                            If HinOld(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                                If Trim(udtWSXL(1).FACTORYCB) = "" Then udtSXL(0).FACTORYCB = HinOld(i).FACTORYCA   ' 工場
                                If Trim(udtWSXL(1).OPECB) = "" Then udtSXL(0).OPECB = HinOld(i).OPECA               ' 操業条件
                                ' 廃棄の場合
                                If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                    udtSXL(0).NEWKNTCB = "CW760"    ' 最終通過工程
                                    udtSXL(0).GNWKNTCB = "TX860"    ' 現在工程
                                    udtSXL(0).LSTCCB = "H"          ' 最終状態区分
                                End If
                                Exit For
                            End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI         ' SXL指示（良品）
                        udtSXL(0).WFCNMAICB = lngFURKEI      ' WFC内欠落枚数
                End Select

                ' シングル確定時、最終状態区分='S'にする
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba

                intRtn = UpdateXSDCB(udtSXL(0), sSqlWhere)

                ' 分割結晶（udtSXL）更新エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             ' 存在しない時、追加
             ElseIf intDataCnt = 0 Then
                ReDim udtSXL(0) As typ_XSDCB_Update
                udtSXL(0).SXLIDCB = HinOld(i).SXLIDCA      ' SXLID
                udtSXL(0).KCNTCB = BlkNow.KCNTC2           ' 工程連番
                udtSXL(0).XTALCB = HinOld(i).XTALCA        ' 結晶番号
                udtSXL(0).INPOSCB = HinOld(i).INPOSCA      ' 結晶内開始位置
                udtSXL(0).LENCB = lngLen                   ' 長さ
                udtSXL(0).HINBCB = HinOld(i).HINBCA        ' 品番
                udtSXL(0).REVNUMCB = HinOld(i).REVNUMCA    ' 電話番号改訂番号
                udtSXL(0).FACTORYCB = HinOld(i).FACTORYCA  ' 工場
                udtSXL(0).OPECB = HinOld(i).OPECA          ' 操業条件
                udtSXL(0).MAICB = lngMAI                   ' 実枚数
                udtSXL(0).WSRMAICB = 0                     ' WS洗後枚数
                udtSXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
                udtSXL(0).WFCMAICB = 0                     ' 受入枚数
                udtSXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
                udtSXL(0).WFCMAICB = 0                     ' 受入枚数
                udtSXL(0).SXLEMAICB = 0                    ' SXL確定枚数

                udtSXL(0).FURIMAICB = ""                   ' 振替枚数
                udtSXL(0).XTWORKCB = "42"                  ' 製造工場
                udtSXL(0).WFWORKCB = " "                   ' ウェーハ製造
                udtSXL(0).FURYCCB = " "                    ' 不良理由
                udtSXL(0).LSTCCB = "T"                     ' 採取状態区分
                udtSXL(0).LUFRCCB = " "                    ' 格上コード
                udtSXL(0).LUFRBCB = " "                    ' 格上区分
                udtSXL(0).LDERCCB = " "                    ' 格下コード

                ' 長さが0の時、廃棄とする
                If wLENCB = 0 Then
                    udtSXL(0).LDFRBCB = "2"                ' 格下区分
                Else
                    udtSXL(0).LDFRBCB = "0"
                End If

                udtSXL(0).HOLDCCB = " "                    ' ホールドコード
                udtSXL(0).HOLDBCB = " "                    ' ホールド区分
                udtSXL(0).EXKUBCB = " "                    ' 例外区分
                udtSXL(0).HENPKCB = " "                    ' 返品区分

                ' 長さが0の時、死ロットとする
                If (lngMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (lngLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                    udtSXL(0).LIVKCB = "1"                 ' 生死区分
                    udtSXL(0).KANKCB = "2"                 ' 完了区分
                    udtSXL(0).LSTCCB = "H"                 ' 最終状態区分
                    udtSXL(0).LDFRBCB = "2"                ' 格下区分
                Else
                    udtSXL(0).LIVKCB = "0"
                    udtSXL(0).KANKCB = "0"                 ' 完了区分
                    udtSXL(0).LSTCCB = "T"                 ' 最終状態区分
                    udtSXL(0).LDFRBCB = "0"                ' 格下区分
                End If

                ' 工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                    Case "CW740"
                        udtSXL(0).SXLNMAICB = lngFUR         ' 廃棄WF枚数
                    Case "CW750"
                        udtSXL(0).SRMAICB = lngSIJ           ' サンプル抜指示枚数
                        udtSXL(0).SNMAICB = lngSAMFUR        ' サンプル抜指示不良枚数
                        udtSXL(0).STMAICB = lngSAM           ' サンプル枚数
                    Case "CW760"
                        udtSXL(0).SXLNMAICB = lngFUR         ' 廃棄WF枚数

                        ''Z品番の場合は廃棄とする。
                        For m = 1 To UBound(tblWfSxlMng())
                           If HinOld(i).SXLIDCA = tblWfSxlMng(m).SXLID Then
                               If Trim(tblWfSxlMng(m).hinban) = "Z" Then
                                   udtSXL(0).NEWKNTCB = "CW760"               ' 最終通過工程
                                   udtSXL(0).GNWKNTCB = "TX860"               ' 現在工程
                                   udtSXL(0).LSTCCB = "H"                     ' 最終状態区分

                                   ' 登録の場合のみ設定
                                   udtSXL(0).FURYCCB = tblWfSxlMng(m).BDCAUS  ' 不良理由
                                   udtSXL(0).RLENCB = tblWfSxlMng(m).LENGTH   ' 理論長さ
                                   udtSXL(0).SHOLDCLSCB = tblWfSxlMng(m).HOLDCLS  ''ホールド区分
                               End If
                               Exit For
                           End If
                        Next m
                    Case "CW800"
                        udtSXL(0).SXLRMAICB = lngMAI         ' SXL指示（良品）
                        udtSXL(0).WFCNMAICB = lngFURKEI      ' WFC内欠落枚数
                End Select

                ' 完了区分フラグ変更
                udtSXL(0).KANKCB = "0"                 ' 完了区分

                ' シングル確定時、最終状態区分='S'にする
                If Kihon.NOWPROC = PROCD_udtSXL_KAKUTEI Then
                   udtSXL(0).LSTCCB = "S"
                End If

                udtSXL(0).NFCB = "0"                                    ' 入庫区分
                udtSXL(0).SAKJCB = "0"                                  ' 削除区分
                udtSXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
                udtSXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                udtSXL(0).SUMITCB = "0"                                 ' SUMIT送信フラグ
                udtSXL(0).SNDKCB = "0"                                  ' 返品区分
                udtSXL(0).SNDAYCB = ""                                  ' 送信日付
                udtSXL(0).PLANTCATCB = sCmbMukesaki
                If sKanrenFlg = "1" Then udtSXL(0).KBLKFLGCB = sKanrenFlg      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　08/01/31 ooba

                intRtn = CreateXSDCB(udtSXL(0), sErrMsg)

                ' 分割結晶（udtSXL）追加エラー
                If intRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox sErrMsg
                    Exit Function
                End If
             End If
        End If
    Next i

    XSDCBProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDCBProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'************************************************************************************************
'*    関数名        : XSDCBSum
'*
'*    処理概要      : 1.工程実績から指定された工程、ＳＸＬＩＤ（ブロックＩＤ、位置、品番）の長さ、
'*                      枚数、不良枚数を集計する受信用のテーブルから指定されたＳＸＬＩＤ）の
'*                      サンプル枚数、サンプル抜指示枚数、サンプル抜指示不良枚数を集計する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    pKKTC          I   string   工程
'*                    pSXLID         I   string   ＳＸＬＩＤ
'*                    pLEN           O   NUMBER   長さ
'*                    pMAI           O   NUMBER   枚数
'*                    pMAI800        O   NUMBER   CW800枚数
'*                    pFUR           O   NUMBER   不良枚数
'*                    pFURKEI        O   NUMBER   不良枚数合計
'*                    pSAM           O   NUMBER   サンプル枚数
'*                    pSAMNUK        O   NUMBER   サンプル抜指示枚数
'*                    pSAMFUR        O   NUMBER   サンプル抜指示不良枚数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'************************************************************************************************
Public Function XSDCBSum(ByVal pKKTC, ByVal pSXLID, ByRef pLEN, ByRef pMAI, ByRef pMAI800, ByRef pFUR, ByRef pFURKEI, ByRef pSAM, ByRef pSAMSIJ, ByRef pSAMFUR)

    ' 内部変数
    Dim i               As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sCRYNUMCA       As String           ' ブロックＩＤ
    Dim lngINPOSCA      As Long             ' 開始位置
    Dim sHINBCA         As String           ' 品番
    Dim lngLen          As Long             ' 長さ
    Dim lngMAI          As Long             ' 枚数
    Dim lngMAI800       As Long             ' CW800を通過した枚数
    Dim lngFUR          As Long             ' 不良枚数
    Dim lngFURKEI       As Long             ' 不良枚数合計
    Dim sKCNTC3         As String           ' 工程連番最大
    Dim sSAMFUR         As String           ' サンプル抜試指示不良枚数
    Dim rsXsdca         As OraDynaset
    Dim rsMain          As OraDynaset

    ' エラーハンドラの設定
    On Error GoTo proc_err

    ' パラメータ初期化
    pLEN = 0
    pMAI = 0
    pMAI800 = 0
    pFUR = 0
    pFURKE = 0
    pSAM = 0
    pSAMSIJ = 0
    pSAMFUR = 0

    ' 分割結晶（品番）からパラメータのＳＸＬＩＤの長さ、枚数を取得
    sSQL = "SELECT SUM(GNLCA) AS wLEN, SUM(GNMCA) AS lngMAI "
    sSQL = sSQL & " FROM XSDCA "
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "' "
    sSQL = sSQL & " AND LIVKCA = '0' "

    Set rsXsdca = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、次へ
    If rsXsdca.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsXsdca.Close
        GoTo CW800_CAL
    End If

    ' 抽出結果を格納する
    If IsNull(rsXsdca.Fields("wLEN")) = True Then
        pLEN = 0
    Else
        pLEN = rsXsdca.Fields("wLEN")                   ' 長さ
    End If
    If IsNull(rsXsdca.Fields("lngMAI")) = True Then
        pMAI = 0
    Else
        pMAI = rsXsdca.Fields("lngMAI")                 ' 枚数
    End If

    rsXsdca.Close

CW800_CAL:
    ' 分割結晶（品番）から同じＳＸＬＩＤのブロックＩＤ、開始位置、品番を取得
    sSQL = "SELECT CRYNUMCA, INPOSCA, HINBCA "
    sSQL = sSQL & " FROM XSDCA "
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "' "
    sSQL = sSQL & " AND LIVKCA = '0' "
    sSQL = sSQL & " ORDER BY CRYNUMCA,INPOSCA"

    Set rsMain = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、次へ
    If rsMain.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsMain.Close
        GoTo SAMPLE_CAL
    End If

    Do Until rsMain.EOF
        ' 抽出結果を格納する
        sCRYNUMCA = rsMain.Fields("CRYNUMCA")
        lngINPOSCA = rsMain.Fields("INPOSCA")
        sHINBCA = rsMain.Fields("HINBCA")

        ' 取得したブロックＩＤ､開始位置､品番で工程実績の該当工程で、工程連番の最大を取得する
        sSQL = "SELECT TOMC3 AS lngMAI800,FUMC3 AS lngFUR "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = " & lngINPOSCA & ""
        sSQL = sSQL & " AND KCNTC3  = (SELECT MAX(KCNTC3)"
        sSQL = sSQL & "                  FROM XSDC3"
        sSQL = sSQL & "                 WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & "                   AND HINBC3 = '" & sHINBCA & "'"
        sSQL = sSQL & "                   AND INPOSC3 = '" & lngINPOSCA & "'"
        sSQL = sSQL & "                   AND WKKTC3 = '" & pKKTC & "' "
        sSQL = sSQL & "                   AND (SUMKBC3 = '0' "
        sSQL = sSQL & "                    OR SUMKBC3 = ' ' "
        sSQL = sSQL & "                    OR SUMKBC3 is null)) "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If

        ' 抽出結果を格納する
        If IsNull(rs.Fields("lngMAI800")) = True Then
            pMAI800 = pMAI800 + 0
        Else
            pMAI800 = pMAI800 + CInt(rs.Fields("lngMAI800"))     '枚数
        End If
        If IsNull(rs.Fields("lngFUR")) = True Then
            pFUR = pFUR + 0
        Else
            pFUR = pFUR + CInt(rs.Fields("lngFUR"))              '不良長さ
        End If

        ' 取得したブロックＩＤ､開始位置､品番で工程実績の不良合計を取得する
        sSQL = "SELECT SUM(FUMC3) AS lngFURKEI "
        sSQL = sSQL & " FROM XSDC3 "
        sSQL = sSQL & " WHERE CRYNUMC3 = '" & sCRYNUMCA & "' "
        sSQL = sSQL & " AND INPOSC3 = " & lngINPOSCA & ""
        sSQL = sSQL & " AND HINBC3 = '" & sHINBCA & "' "
        sSQL = sSQL & " AND SUMKBC3 = '1' "
        sSQL = sSQL & " AND MODKBC3 = '0' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If

        ' 抽出結果を格納する
        If IsNull(rs.Fields("lngFURKEI")) = True Then
            pFURKEI = pFURKEI + 0 '不良長さ
        Else
            pFURKEI = pFURKEI + CInt(rs.Fields("lngFURKEI")) '不良長さ
        End If
        rsMain.MoveNext
    Loop

    rs.Close
    rsMain.Close

SAMPLE_CAL:
    ' 評価結果受信レコード数よりサンプル枚数を取得する
    ' ■サンプル枚数　-　評価結果受信レコード数（Y013)
    ' 取得条件変更(ｽﾋﾟｰﾄﾞ化) 09/05/25 ooba
    sSQL = "SELECT COUNT(SAMPLEID) AS wSAM "
    sSQL = sSQL & " FROM TBCMY013 Y013"
    sSQL = sSQL & " WHERE  SAMPLEID in ( "
    sSQL = sSQL & " SELECT REPSMPLIDCW FROM XSDCW"
    sSQL = sSQL & " WHERE SXLIDCW = '" & pSXLID & "'"
    sSQL = sSQL & " )"
''    sSql = sSql & " SELECT E044.REPSMPLIDCW "
''    sSql = sSql & " FROM XSDCW E044 "
''    sSql = sSql & "  ,("
''    sSql = sSql & "    SELECT"
''    sSql = sSql & "      XTALCB as CRYNUM"
''    sSql = sSql & "     ,INPOSCB as INGOTPOS"
''    sSql = sSql & "     ,RLENCB as LENGTH"
''    sSql = sSql & "    FROM"
''    sSql = sSql & "      XSDCB"
''    sSql = sSql & "    WHERE SXLIDCB = '" & pSXLID & "'"
''    sSql = sSql & "   ) E042"
''    sSql = sSql & " WHERE (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND  E044.INPOSCW = E042.INGOTPOS "
''    sSql = sSql & " AND E044.SMPKBNCW = 'T' ) "
''    sSql = sSql & " OR (E044.XTALCW = E042.CRYNUM"
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
''    sSql = sSql & " AND E044.SMPKBNCW = 'B' ))"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    ' 抽出結果を格納する
    pSAM = rs.Fields("wSAM") ' サンプル枚数

    rs.Close

    ' 抜試指示レコード数よりサンプル指示枚数を取得する
    ' ■サンプル抜指示枚数（良品）　-　抜試指示レコード数（Y003)
    ' 取得条件変更(ｽﾋﾟｰﾄﾞ化) 09/05/25 ooba
    sSQL = "SELECT COUNT(SAMPLEID) AS wSIJ"
    sSQL = sSQL & " FROM TBCMY003 "
    sSQL = sSQL & " WHERE SAMPLEID in ( "
    sSQL = sSQL & " SELECT REPSMPLIDCW FROM XSDCW"
    sSQL = sSQL & " WHERE SXLIDCW = '" & pSXLID & "'"
    sSQL = sSQL & " )"
''    sSql = sSql & " SELECT E044.REPSMPLIDCW "
''    sSql = sSql & " FROM XSDCW E044 "
''    sSql = sSql & "  ,("
''    sSql = sSql & "    SELECT"
''    sSql = sSql & "      XTALCB as CRYNUM"
''    sSql = sSql & "     ,INPOSCB as INGOTPOS"
''    sSql = sSql & "     ,RLENCB as LENGTH"
''    sSql = sSql & "    FROM"
''    sSql = sSql & "      XSDCB"
''    sSql = sSql & "    WHERE SXLIDCB = '" & pSXLID & "'"
''    sSql = sSql & "   ) E042"
''    sSql = sSql & " WHERE (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS "
''    sSql = sSql & " AND E044.SMPKBNCW = 'T' ) "
''    sSql = sSql & " OR (E044.XTALCW = E042.CRYNUM "
''    sSql = sSql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
''    sSql = sSql & " AND  E044.SMPKBNCW = 'B' )) "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    ' 抽出結果を格納する
'    pSIJ = rs.Fields("wSIJ") ' サンプル抜指示枚数
    pSAMSIJ = rs.Fields("wSIJ") ' サンプル抜指示枚数 09/05/25 ooba

    rs.Close

    ' C欠落枚数よりサンプル抜指示不良枚数を取得する
    ' ■サンプル抜試指示不良枚数　-　C欠落枚数　-（Y012）

    ' 対象のブロックID取得
    sSQL = "SELECT DISTINCT(CRYNUMCA) "
    sSQL = sSQL & " FROM XSDCA"
    sSQL = sSQL & " WHERE CRYNUMCA like '" & left(pSXLID, 9) & "%' "    'ｲﾝﾃﾞｯｸｽ項目追加 09/05/25 ooba
    sSQL = sSQL & " AND SXLIDCA = '" & pSXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
    End If

    Do Until rs.EOF
        ' 欠落情報COUNT(ブロックIDごとループしSUMする）
        ' 抽出結果を格納する
        sCRYNUMCA = rs.Fields("CRYNUMCA") ' ブロックID

        sSQL = "SELECT COUNT(Y012.LOTID) AS sSAMFUR "
        sSQL = sSQL & " FROM TBCMY012 Y012 "
        sSQL = sSQL & "  ,("
        sSQL = sSQL & "    SELECT"
        sSQL = sSQL & "      XTALCB as CRYNUM"
        sSQL = sSQL & "     ,INPOSCB as INGOTPOS"
        sSQL = sSQL & "     ,RLENCB as LENGTH"
        sSQL = sSQL & "    FROM"
        sSQL = sSQL & "      XSDCB"
        sSQL = sSQL & "    WHERE SXLIDCB = '" & pSXLID & "'"
        sSQL = sSQL & "   ) E042"
        sSQL = sSQL & " ,(SELECT CRYNUM, INGOTPOS, LENGTH, BLOCKID "
        sSQL = sSQL & " FROM TBCME040 "
        sSQL = sSQL & " WHERE BLOCKID =  '" & sCRYNUMCA & "' ) E040 "
        sSQL = sSQL & " WHERE Y012.LOTID = E040.BLOCKID "
        sSQL = sSQL & " AND E042.INGOTPOS <= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sSQL = sSQL & " AND E042.INGOTPOS + E042.LENGTH  >= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sSQL = sSQL & " AND REJCAT = 'C' "

        Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

        ' 存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
        End If

        ' 抽出結果を格納する
        pSAMFUR = rs.Fields("sSAMFUR") ' サンプル抜試指示不良枚数
        rs.MoveNext
    Loop

    rs.Close

    XSDCBSum = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    XSDCBSum = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : WfCount
'*
'*    処理概要      : 1.WF枚数を計算する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                   sSelectBlkID    ,I  ,Integer  ,ブロックID
'*                   intBlkLen      ,I  ,Integer  ,ブロック長さ
'*                   intWfCnt       ,O  ,Integer  ,枚数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function WfCount(ByVal sSelectBlkID As String, ByVal intBlkLen As Integer, ByRef intWFcnt As Integer) As FUNCTION_RETURN
    Dim udtRec()    As typ_cmkc001f_Disp
    Dim RET         As FUNCTION_RETURN
    Dim lngRecCnt   As Long
    Dim i           As Long
    Dim j           As Integer
    Dim s           As String
    Dim intWfNum    As Integer '枚数

    ' 枚数計算関数用パラメータ（HSXCTCEN & HSXCYCEN）
    ' 仕様・実績を読み込む
    RET = DBDRV_fcmkc001f_Disp(Trim(sSelectBlkID), blkInfo, udtRec)   ' SelectBlkId=ブロックID,blkInfo=ブロック管理構造体
    If RET = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    ElseIf UBound(udtRec) Then
        HSXCTCEN = udtRec(1).HSXCTCEN
        HSXCYCEN = udtRec(1).HSXCYCEN
    End If

    ' WF枚数計算用の基本値を取得する
    Loss0 = val(GetCodeField("LG", "01", "LOSS0", "INFO1"))
    Loss4 = val(GetCodeField("LG", "01", "LOSS4", "INFO1"))
    Mlt4 = val(GetCodeField("LG", "01", "MLT4", "INFO1"))
    Pitch = val(GetCodeField("LG", "01", "PITCH", "INFO1"))

    ' 枚数計算関数用パラメータ（SEEDDEG）取得
    ' 引き上げ結晶の場合
    s = GetCodeField("SC", "28", left$(blkInfo.SEED, 1), "INFO3")
    If left$(s, 1) = "4" Then
        SEEDDEG = 4
    Else
        SEEDDEG = 0
    End If

    ' 枚数取得関数
    intWFcnt = GetWfCount(val(intBlkLen), SEEDDEG, HSXCTCEN, HSXCYCEN)  'intWfCount=枚数
    WfCount = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    Exit Function

End Function

'*******************************************************************************
'*    関数名        : GetWfCount
'*
'*    処理概要      : 1.WF枚数を計算する
'*
'*    パラメータ    : 変数名       ,IO ,型        ,説明
'*                   blkLen        ,I  ,Integer   ,ブロック長さ
'*                   seedDeg       ,I  ,Integer   ,結晶のSEED傾き
'*                   dblHinDegT       ,I  ,Double    ,品番傾き（縦）
'*                   dblHinDegY       ,I  ,Double    ,品番傾き（横）
'*
'*    戻り値        : WF枚数
'*
'*******************************************************************************
Public Function GetWfCount(ByVal BlkLen%, ByVal SEEDDEG%, ByVal dblHinDegT As Double, ByVal dblHinDegY As Double) As Integer
    Dim intHinDeg   As Integer
    Dim s           As String
    Dim intWFcnt    As Integer

    If Pitch = 0# Then
        GetWfCount = 0
        Exit Function
    End If

    ' 品番傾きを得る
    ' 結晶最終払い出し、品番傾きの求め方変更
    If (Abs(dblHinDegT) = 2.83) And (Abs(dblHinDegY) = 2.83) Then
        intHinDeg = 4
    ElseIf (Abs(dblHinDegT) = 4) And (dblHinDegY = 0) Then
        intHinDeg = 4
    ElseIf (dblHinDegT = 0) And (Abs(dblHinDegY) = 4) Then
        intHinDeg = 4
    Else
        intHinDeg = 0
    End If

    ' WF枚数を計算する
    If SEEDDEG = intHinDeg Then
        ' 通常品の場合
        intWFcnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        intWFcnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If

    GetWfCount = intWFcnt
End Function

'概要      :結晶最終払出入力 表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                   ,説明
'      　　:sBlockID_in　 ,I  ,String               ,ブロックID
'      　　:udtBlkInfo　　　,O  ,typ_cmkc001f_Block   ,ブロック情報
'      　　:udtRecords　　　,O  ,typ_cmkc001f_Disp    ,製品仕様取得用
'      　　:戻り値       ,O  ,FUNCTION_RETURN      ,読み込みの成否
'*******************************************************************************
'*    関数名        : DBDRV_fcmkc001f_Disp
'*
'*    処理概要      : 1.結晶最終払出入力 表示用ＤＢドライバ
'*
'*    パラメータ    : 変数名       ,IO ,型                   ,説明
'*               　　sBlockID_in　  ,I  ,String               ,ブロックID
'*      　　     　　udtBlkInfo　　　 ,O  ,typ_cmkc001f_Block   ,ブロック情報
'*      　　     　　udtRecords　　　 ,O  ,typ_cmkc001f_Disp    ,製品仕様取得用
'*
'*    戻り値        : WF枚数
'*
'*******************************************************************************
Public Function DBDRV_fcmkc001f_Disp(sBlockID_in As String, udtBlkInfo As typ_cmkc001f_Block, udtRecords() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sSQL        As String
    Dim rs          As OraDynaset
    Dim intRecCnt   As Integer
    Dim i           As Long
    Dim n           As Integer

    ' エラーハンドラの設定
    On Error GoTo proc_err

    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS

    ' ブロック情報を得る
    sSQL = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & sBlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With udtBlkInfo
        .INGOTPOS = rs("INGOTPOS")      ' 結晶内開始位置
        .LENGTH = rs("LENGTH")          ' 長さ
        .REALLEN = rs("REALLEN")        ' 実長さ
        .KRPROCCD = rs("KRPROCCD")      ' 現在管理工程
        .NOWPROC = rs("NOWPROC")        ' 現在工程
        .LPKRPROCCD = rs("LPKRPROCCD")  ' 最終通過管理工程
        .LASTPASS = rs("LASTPASS")      ' 最終通過工程
        .DELCLS = rs("DELCLS")          ' 削除区分
        .RSTATCLS = rs("RSTATCLS")      ' 流動状態区分
        .LSTATCLS = rs("LSTATCLS")      ' 最終状態区分
        .SEED = rs("SEED")              ' SEED
    End With
    rs.Close

    ' 製品仕様を得る
    sSQL = "select "
    sSQL = sSQL & "BH.E041HINBAN, "           ' 品番
    sSQL = sSQL & "BH.E041INGOTPOS, "         ' 結晶内開始位置
    sSQL = sSQL & "BH.E041REVNUM, "           ' 製品番号改訂番号
    sSQL = sSQL & "BH.E041FACTORY, "          ' 工場
    sSQL = sSQL & "BH.E041OPECOND, "          ' 操業条件
    sSQL = sSQL & "BH.E041LENGTH, "           ' 長さ

    ' 製品仕様SXLデータ
    sSQL = sSQL & "S.E018HSXD1CEN, "          ' 品ＳＸ直径１中心
    sSQL = sSQL & "S.E018HSXRMIN, "           ' 品ＳＸ比抵抗下限
    sSQL = sSQL & "S.E018HSXRMAX, "           ' 品ＳＸ比抵抗上限
    sSQL = sSQL & "S.E018HSXRMBNP, "          ' 品ＳＸ比抵抗面内分布
    sSQL = sSQL & "S.E018HSXRHWYS, "          ' 品ＳＸ比抵抗保証方法＿処
    sSQL = sSQL & "S.E019HSXONMIN, "          ' 品ＳＸ酸素濃度下限
    sSQL = sSQL & "S.E019HSXONMAX, "          ' 品ＳＸ酸素濃度上限
    sSQL = sSQL & "S.E019HSXONMBP, "          ' 品ＳＸ酸素濃度面内分布
    sSQL = sSQL & "S.E019HSXONHWS, "          ' 品ＳＸ酸素濃度保証方法＿処
    sSQL = sSQL & "S.E019HSXCNMIN, "          ' 品ＳＸ炭素濃度下限
    sSQL = sSQL & "S.E019HSXCNMAX, "          ' 品ＳＸ炭素濃度上限
    sSQL = sSQL & "S.E019HSXCNHWS, "          ' 品ＳＸ炭素濃度保証方法＿処
    sSQL = sSQL & "S.E019HSXTMMAXN, "         ' 品ＳＸ転位密度上限             項目追加，修正対応 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXBM1AN, "          ' 品ＳＸＢＭＤ１平均下限
    sSQL = sSQL & "S.E020HSXBM1AX, "          ' 品ＳＸＢＭＤ１平均上限
    sSQL = sSQL & "S.E020HSXBM1HS, "          ' 品ＳＸＢＭＤ１保証方法＿処
    sSQL = sSQL & "S.E020HSXBM2AN, "          ' 品ＳＸＢＭＤ２平均下限
    sSQL = sSQL & "S.E020HSXBM2AX, "          ' 品ＳＸＢＭＤ２平均上限
    sSQL = sSQL & "S.E020HSXBM2HS, "          ' 品ＳＸＢＭＤ２保証方法＿処
    sSQL = sSQL & "S.E020HSXBM3AN, "          ' 品ＳＸＢＭＤ３平均下限
    sSQL = sSQL & "S.E020HSXBM3AX, "          ' 品ＳＸＢＭＤ３平均上限
    sSQL = sSQL & "S.E020HSXBM3HS, "          ' 品ＳＸＢＭＤ３保証方法＿処
    sSQL = sSQL & "S.E020HSXOF1AX, "          ' 品ＳＸＯＳＦ１平均上限
    sSQL = sSQL & "S.E020HSXOF1MX, "          ' 品ＳＸＯＳＦ１上限
    sSQL = sSQL & "S.E020HSXOF1HS, "          ' 品ＳＸＯＳＦ１ 保証方法＿処
    sSQL = sSQL & "S.E020HSXOF2AX, "          ' 品ＳＸＯＳＦ２平均上限
    sSQL = sSQL & "S.E020HSXOF2MX, "          ' 品ＳＸＯＳＦ２上限
    sSQL = sSQL & "S.E020HSXOF2HS, "          ' 品ＳＸＯＳＦ２ 保証方法＿処
    sSQL = sSQL & "S.E020HSXOF3AX, "          ' 品ＳＸＯＳＦ３平均上限
    sSQL = sSQL & "S.E020HSXOF3MX, "          ' 品ＳＸＯＳＦ３上限
    sSQL = sSQL & "S.E020HSXOF3HS, "          ' 品ＳＸＯＳＦ３ 保証方法＿処
    sSQL = sSQL & "S.E020HSXOF4AX, "          ' 品ＳＸＯＳＦ４平均上限
    sSQL = sSQL & "S.E020HSXOF4MX, "          ' 品ＳＸＯＳＦ４上限
    sSQL = sSQL & "S.E020HSXOF4HS, "          ' 品ＳＸＯＳＦ４ 保証方法＿処
    sSQL = sSQL & "S.E020HSXDENMX, "          ' 品ＳＸＤｅｎ上限
    sSQL = sSQL & "S.E020HSXDENMN, "          ' 品ＳＸＤｅｎ下限
    sSQL = sSQL & "S.E020HSXDENHS, "          ' 品ＳＸＤｅｎ保証方法＿処
    sSQL = sSQL & "S.E020HSXDVDMXN, "         ' 品ＳＸＤＶＤ２上限           項目追加，修正対応 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXDVDMNN, "         ' 品ＳＸＤＶＤ２下限           項目追加，修正対応 2003.05.20 yakimura
    sSQL = sSQL & "S.E020HSXDVDHS, "          ' 品ＳＸＤＶＤ２保証方法＿処
    sSQL = sSQL & "S.E020HSXLDLMX, "          ' 品ＳＸＬ／ＤＬ上限
    sSQL = sSQL & "S.E020HSXLDLMN, "          ' 品ＳＸＬ／ＤＬ下限
    sSQL = sSQL & "S.E020HSXLDLHS, "          ' 品ＳＸＬ／ＤＬ保証方法＿処
    sSQL = sSQL & "S.E019HSXLTMIN, "          ' 品ＳＸＬタイム下限
    sSQL = sSQL & "S.E019HSXLTMAX, "          ' 品ＳＸＬタイム上限
    sSQL = sSQL & "S.E019HSXLTHWS, "          ' 品ＳＸＬタイム保証方法＿処
    sSQL = sSQL & "S.E018HSXDPDIR, "          ' 品ＳＸ溝位置方位
    sSQL = sSQL & "S.E018HSXDPDRC, "          ' 品ＳＸ溝位置方向
    sSQL = sSQL & "S.E018HSXDWMIN, "          ' 品ＳＸ溝巾下限
    sSQL = sSQL & "S.E018HSXDWMAX, "          ' 品ＳＸ溝巾上限
    sSQL = sSQL & "S.E018HSXDDMIN, "          ' 品ＳＸ溝深下限
    sSQL = sSQL & "S.E018HSXDDMAX, "          ' 品ＳＸ溝深上限
    sSQL = sSQL & "S.E018HSXD1MIN, "          ' 品ＳＸ直径１下限
    sSQL = sSQL & "S.E018HSXD1MAX, "          ' 品ＳＸ直径１上限
    sSQL = sSQL & "S.E018HSXCTCEN, "          ' 品ＳＸ結晶面傾縦中心
    sSQL = sSQL & "S.E018HSXCYCEN, "          ' 品ＳＸ結晶面傾横中心
    sSQL = sSQL & "U.EPDUP "                  ' 結晶内側管理 EPD　上限
    sSQL = sSQL & " from VECME009 BH, VECME001 S, TBCME036 U "
    sSQL = sSQL & " where BH.E040BLOCKID='" & sBlockID_in & "' "
    sSQL = sSQL & " and S.E018HINBAN=BH.E041HINBAN "
    sSQL = sSQL & " and S.E018MNOREVNO=BH.E041REVNUM "
    sSQL = sSQL & " and S.E018FACTORY=BH.E041FACTORY "
    sSQL = sSQL & " and S.E018OPECOND=BH.E041OPECOND "
    sSQL = sSQL & " and U.HINBAN=BH.E041HINBAN "
    sSQL = sSQL & " and U.MNOREVNO=BH.E041REVNUM "
    sSQL = sSQL & " and U.FACTORY=BH.E041FACTORY "
    sSQL = sSQL & " and U.OPECOND=BH.E041OPECOND "

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_NO_BLANKSTRIP)

    If rs.RecordCount = 0 Then
        ReDim udtRecords(0)
        rs.Close
        GoTo proc_exit
    End If

    intRecCnt = rs.RecordCount
    ReDim udtRecords(intRecCnt)
    For i = 1 To intRecCnt
        With udtRecords(i)
            ' 品番管理
            .hinban = rs("E041HINBAN")              ' 品番
            .INGOTPOS = rs("E041INGOTPOS")          ' 結晶内開始位置
            .REVNUM = rs("E041REVNUM")              ' 製品番号改訂番号
            .factory = rs("E041FACTORY")            ' 工場
            .opecond = rs("E041OPECOND")            ' 操業条件
            .LENGTH = rs("E041LENGTH")              ' 長さ

            ' 製品仕様SXLデータ
            .HSXD1CEN = rs("E018HSXD1CEN")          ' 品ＳＸ直径１中心
            .HSXRMIN = rs("E018HSXRMIN")            ' 品ＳＸ比抵抗下限
            .HSXRMAX = rs("E018HSXRMAX")            ' 品ＳＸ比抵抗上限
            .HSXRMBNP = rs("E018HSXRMBNP")          ' 品ＳＸ比抵抗面内分布
            .HSXRHWYS = rs("E018HSXRHWYS")          ' 品ＳＸ比抵抗保証方法＿処
            .HSXONMIN = rs("E019HSXONMIN")          ' 品ＳＸ酸素濃度下限
            .HSXONMAX = rs("E019HSXONMAX")          ' 品ＳＸ酸素濃度上限
            .HSXONMBP = rs("E019HSXONMBP")          ' 品ＳＸ酸素濃度面内分布
            .HSXONHWS = rs("E019HSXONHWS")          ' 品ＳＸ酸素濃度保証方法＿処
            .HSXCNMIN = rs("E019HSXCNMIN")          ' 品ＳＸ炭素濃度下限
            .HSXCNMAX = rs("E019HSXCNMAX")          ' 品ＳＸ炭素濃度上限
            .HSXCNHWS = rs("E019HSXCNHWS")          ' 品ＳＸ炭素濃度保証方法＿処
            .HSXTMMAX = rs("E019HSXTMMAXN")         ' 品ＳＸ転位密度上限           項目追加，修正対応 2003.05.20 yakimura

            For n = 1 To 3
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")  ' 品ＳＸＢＭＤn 保証方法＿処
            Next

            For n = 1 To 4
                If IsNull(rs("E020HSXOF" & n & "AX")) = False Then .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' 品ＳＸＯＳＦn 平均上限         '05/03/29 ooba NULL対応
                If IsNull(rs("E020HSXOF" & n & "MX")) = False Then .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' 品ＳＸＯＳＦn 上限             '05/03/29 ooba NULL対応
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")   ' 品ＳＸＯＳＦn 保証方法＿処
            Next

            .HSXDENMX = rs("E020HSXDENMX")          ' 品ＳＸＤｅｎ上限
            .HSXDENMN = rs("E020HSXDENMN")          ' 品ＳＸＤｅｎ下限
            .HSXDENHS = rs("E020HSXDENHS")          ' 品ＳＸＤｅｎ保証方法＿処
            .HSXDVDMX = rs("E020HSXDVDMXN")         ' 品ＳＸＤＶＤ２上限        項目追加，修正対応 2003.05.20 yakimura
            .HSXDVDMN = rs("E020HSXDVDMNN")         ' 品ＳＸＤＶＤ２下限        項目追加，修正対応 2003.05.20 yakimura
            .HSXDVDHS = rs("E020HSXDVDHS")          ' 品ＳＸＤＶＤ２保証方法＿処
            .HSXLDLMX = rs("E020HSXLDLMX")          ' 品ＳＸＬ／ＤＬ上限
            .HSXLDLMN = rs("E020HSXLDLMN")          ' 品ＳＸＬ／ＤＬ下限
            .HSXLDLHS = rs("E020HSXLDLHS")          ' 品ＳＸＬ／ＤＬ保証方法＿処
            .HSXLTMIN = rs("E019HSXLTMIN")          ' 品ＳＸＬタイム下限
            .HSXLTMAX = rs("E019HSXLTMAX")          ' 品ＳＸＬタイム上限
            .HSXLTHWS = rs("E019HSXLTHWS")          ' 品ＳＸＬタイム保証方法＿処
            .HSXDPDIR = rs("E018HSXDPDIR")          ' 品ＳＸ溝位置方位
            .HSXDPDRC = rs("E018HSXDPDRC")          ' 品ＳＸ溝位置方向
            .HSXDWMIN = rs("E018HSXDWMIN")          ' 品ＳＸ溝巾下限
            .HSXDWMAX = rs("E018HSXDWMAX")          ' 品ＳＸ溝巾上限
            .HSXDDMIN = rs("E018HSXDDMIN")          ' 品ＳＸ溝深下限
            .HSXDDMAX = rs("E018HSXDDMAX")          ' 品ＳＸ溝深上限
            .HSXD1MIN = rs("E018HSXD1MIN")          ' 品ＳＸ直径１下限
            .HSXD1MAX = rs("E018HSXD1MAX")          ' 品ＳＸ直径１上限
            If IsNull(rs("E018HSXCTCEN")) = False Then .HSXCTCEN = rs("E018HSXCTCEN")       ' 品ＳＸ結晶面傾縦中心      '05/03/29 ooba NULL対応
            If IsNull(rs("E018HSXCYCEN")) = False Then .HSXCYCEN = rs("E018HSXCYCEN")       ' 品ＳＸ結晶面傾横中心      '05/03/29 ooba NULL対応
            .EPDUP = rs("EPDUP")                    ' 結晶内側管理 EPD　上限
        End With
        rs.MoveNext
    Next
    rs.Close

proc_exit:
    '終了
'    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    Resume proc_exit
End Function

'******************************************************************************************
'*    関数名        : SelCntXSDC4
'*
'*    処理概要      : 1.分割結晶（不良内訳）から、指定した条件に該当するデータの行数を取得
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　          sStrWhere      ,I  ,String   ,SELECT条件文
'*　　　　　          intCnt        ,O  ,Integer  ,件数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'******************************************************************************************
Public Function SelCntXSDC4(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' ＳＱＬ
    Dim rs          As OraDynaset       ' レコードセット
    Dim sSqlWhere   As String           ' WHERE句

    ' エラーハンドラの設定
    On Error GoTo proc_err

    SelCntXSDC4 = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDC4 " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、エラー
    If rs Is Nothing Then
        SelCntXSDC4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDC4 = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' 終了
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*************************************************************************************
'*    関数名        : SelCntXSDCA
'*
'*    処理概要      : 1.分割結晶（品番）から、指定した条件に該当するデータの行数を取得
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　          sStrWhere      ,I  ,String   ,SELECT条件文
'*　　　　　          intCnt        ,O  ,Integer  ,件数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************
Public Function SelCntXSDCA(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' ＳＱＬ
    Dim rs          As OraDynaset       ' レコードセット
    Dim sSqlWhere   As String           ' WHERE句

    ' エラーハンドラの設定
    On Error GoTo proc_err

    SelCntXSDCA = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDCA " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、エラー
    If rs Is Nothing Then
        SelCntXSDCA = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDCA = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'*************************************************************************************
'*    関数名        : SelCntXSDCB
'*
'*    処理概要      : 1.分割結晶（SXL）から、指定した条件に該当するデータの行数を取得
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　          sStrWhere      ,I  ,String   ,SELECT条件文
'*　　　　　          intCnt        ,O  ,Integer  ,件数
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*************************************************************************************
Public Function SelCntXSDCB(ByVal sStrWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    Dim sSQL        As String           ' ＳＱＬ
    Dim rs          As OraDynaset       ' レコードセット
    Dim sSqlWhere   As String           ' WHERE句

    ' エラーハンドラの設定
    On Error GoTo proc_err

    SelCntXSDCB = FUNCTION_RETURN_FAILURE

    sSQL = "      SELECT count(*) cnt "
    sSQL = sSQL & "  FROM XSDCB " & sStrWhere

    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)

    ' 存在しない時、エラー
    If rs Is Nothing Then
        SelCntXSDCB = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    If rs.RecordCount = 0 Then
        intCnt = 0
    Else
        intCnt = CInt(rs("cnt"))
    End If
    rs.Close

    SelCntXSDCB = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSQL
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*************************************************************************************
'*    関数名        : clearType
'*
'*    処理概要      : 1.新DB構造体初期化
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*************************************************************************************
Public Sub clearType()
    On Error Resume Next

    ' 基本情報
    With Kihon
        .ALLSCRAP = ""
        .CNTHINNOW = 0
        .CNTHINOLD = 0
        .DIAMETER = 0
        .FURYOUMU = ""
        .NEWPROC = ""
        .NOWPROC = ""
        .STAFFID = ""
    End With

    ' 分割結晶(ブロック)：前工程
    With BlkOld
        .CRYNUMC2 = ""
        .KCNTC2 = ""
        .XTALC2 = ""
        .INPOSC2 = ""
        .NEKKNTC2 = ""
        .NEWKNTC2 = ""
        .NEWKKBC2 = ""
        .NEMACOC2 = ""
        .GNKKNTC2 = ""
        .GNWKNTC2 = ""
        .GNWKKBC2 = ""
        .GNMACOC2 = ""
        .GNDAYC2 = ""
        .GNLC2 = ""
        .GNWC2 = ""
        .GNMC2 = ""
        .SUMITLC2 = ""
        .SUMITWC2 = ""
        .SUMITMC2 = ""
        .CHGC2 = ""
        .KAKOUBC2 = ""
        .KEIDAYC2 = ""
        .GNTKUBC2 = ""
        .GNTNOC2 = ""
        .XTWORKC2 = ""
        .WFWORKC2 = ""
        .LSTATBC2 = ""
        .RSTATBC2 = ""
        .LUFRCC2 = ""
        .LUFRBC2 = ""
        .LDFRCC2 = ""
        .LDFRBC2 = ""
        .HOLDCC2 = ""
        .HOLDBC2 = ""
        .EXKUBC2 = ""
        .HENPKC2 = ""
        .LIVKC2 = ""
        .KANKC2 = ""
        .NFC2 = ""
        .SAKJC2 = ""
        .TDAYC2 = ""
        .KDAYC2 = ""
        .SUMITBC2 = ""
        .SNDKC2 = ""
        .SNDDAYC2 = ""
    End With
    With BlkNow
        .CRYNUMC2 = ""
        .KCNTC2 = ""
        .XTALC2 = ""
        .INPOSC2 = ""
        .NEKKNTC2 = ""
        .NEWKNTC2 = ""
        .NEWKKBC2 = ""
        .NEMACOC2 = ""
        .GNKKNTC2 = ""
        .GNWKNTC2 = ""
        .GNWKKBC2 = ""
        .GNMACOC2 = ""
        .GNDAYC2 = ""
        .GNLC2 = ""
        .GNWC2 = ""
        .GNMC2 = ""
        .SUMITLC2 = ""
        .SUMITWC2 = ""
        .SUMITMC2 = ""
        .CHGC2 = ""
        .KAKOUBC2 = ""
        .KEIDAYC2 = ""
        .GNTKUBC2 = ""
        .GNTNOC2 = ""
        .XTWORKC2 = ""
        .WFWORKC2 = ""
        .LSTATBC2 = ""
        .RSTATBC2 = ""
        .LUFRCC2 = ""
        .LUFRBC2 = ""
        .LDFRCC2 = ""
        .LDFRBC2 = ""
        .HOLDCC2 = ""
        .HOLDBC2 = ""
        .EXKUBC2 = ""
        .HENPKC2 = ""
        .LIVKC2 = ""
        .KANKC2 = ""
        .NFC2 = ""
        .SAKJC2 = ""
        .TDAYC2 = ""
        .KDAYC2 = ""
        .SUMITBC2 = ""
        .SNDKC2 = ""
        .SNDDAYC2 = ""
    End With

    ReDim HinOld(0) As typ_XSDCA_Update
    ReDim HinNow(0) As typ_XSDCA_Update

    With Furyou
        .XTALC4 = ""
        .INPOSC4 = ""
        .KCKNTC4 = ""
        .HINBC4 = ""
        .REVNUMC4 = ""
        .FACTORYC4 = ""
        .OPEC4 = ""
        .KNKTC4 = ""
        .WKKTC4 = ""
        .WKKDC4 = ""
        .MACOC4 = ""
        .SXLIDC4 = ""
        .FCODEC4 = ""
        .PUCUTLC4 = ""
        .PUCUTWC4 = ""
        .PUCUTMC4 = ""
        .FKUBC4 = ""
        .TDAYC4 = ""
        .KDAYC4 = ""
        .SUMITBC3 = ""
        .SNDKC3 = ""
        .SNDDAYC3 = ""
    End With

End Sub

'*******************************************************************************
'*    関数名        : calculateWfNum
'*
'*    処理概要      : 1.WF枚数を計算する
'*
'*    パラメータ    : 変数名        ,IO ,型        ,説明
'*                    blkLen        ,I  ,Integer   ,ブロック長さ
'*                    seedDeg       ,I  ,Integer   ,結晶のSEED傾き
'*                    dblHinDegT       ,I  ,Double    ,品番傾き（縦）
'*                    dblHinDegY       ,I  ,Double    ,品番傾き（横）
'*
'*    戻り値        : WF枚数
'*
'*******************************************************************************
Private Function calculateWfNum(ByVal BlkLen%, ByVal SEEDDEG%, ByVal dblHinDegT As Double, ByVal dblHinDegY As Double) As Integer
    Dim intHinDeg   As Integer
    Dim s           As String
    Dim intWFcnt    As Integer

    If Pitch = 0# Then
        calculateWfNum = 0
        Exit Function
    End If

    ' 品番傾きを得る
    ' 結晶最終払い出し、品番傾きの求め方変更
    If (Abs(dblHinDegT) = 2.83) And (Abs(dblHinDegY) = 2.83) Then
        intHinDeg = 4
    ElseIf (Abs(dblHinDegT) = 4) And (dblHinDegY = 0) Then
        intHinDeg = 4
    ElseIf (dblHinDegT = 0) And (Abs(dblHinDegY) = 4) Then
        intHinDeg = 4
    Else
        intHinDeg = 0
    End If

    ' WF枚数を計算する
    If SEEDDEG = intHinDeg Then
        ' 通常品の場合
        intWFcnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        intWFcnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
    If intWFcnt < 0 Then intWFcnt = 0
    calculateWfNum = intWFcnt
End Function

'*******************************************************************************
'*    関数名        : XSDC3Proc2
'*
'*    処理概要      : 1.工程実績登録処理を行う(在庫減情報：CW740,CW760用)
'*
'*    パラメータ    : 変数名        ,IO ,型        ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc2() As FUNCTION_RETURN
    ' 内部変数
    Dim i, j, k         As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim rs              As OraDynaset       ' レコードセット
    Dim sSqlWhere       As String           ' WHERE句
    Dim sErrMsg         As String           ' エラーメッセージ
    Dim udtKoutei       As typ_XSDC3_Update ' 工程実績
    Dim rsKCNTC         As OraDynaset       ' レコードセット

    Dim udtWSTOCKINFO() As typ_stock_info   ' 現在工程の情報
    Dim vGetData        As Variant          ' 画面取込用work
    Dim sOldHinban      As String           ' 旧品番
    Dim sNowHinban      As String           ' 現品番
    Dim vBlkId          As Variant          ' 画面取込用work
    Dim sOldBlkID       As String           ' 旧ブロックID
    Dim vREVNUM         As Variant          ' 画面取込用work
    Dim vFACTORY        As Variant          ' 画面取込用work
    Dim vOPE            As Variant          ' 画面取込用work
    Dim intREVNUM       As Integer          ' 製品改訂番号
    Dim sFACTORY        As String           ' 工場
    Dim sOPE            As String           ' 操業条件
    Dim sBlkId          As String           ' ブロックID

    Dim intMapSt        As Integer          ' マップ開始位置
    Dim intMapEd        As Integer          ' マップ終了位置
    Dim blHinFlg        As Boolean          ' 品番比較用フラグ
    Dim lngTMaisu       As Long             ' 合計枚数
    Dim intGetHinInpos  As Integer          ' 結晶内位置
    Dim objGamenSpd     As Object           ' 画面ID
    Dim intHantei       As Integer

    ' エラーハンドラの設定
    On Error GoTo 0

    ' 初期設定
    XSDC3Proc2 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(0)
    ReDim udtWSTOCKINFO(0)

   ' 前工程長さ合計
    For i = 0 To Kihon.CNTHINOLD - 1
        If (Kihon.NOWPROC = "CW760") _
           And ((SIngotP > CLng(HinOld(i).INPOSCA)) Or (HinOld(i).INPOSCA >= EIngotP)) Then
            ' 処理なし
        Else
            ' 前工程枚数が0の時、処理終了
            If HinOld(i).GNMCA <= 0 Then
                XSDC3Proc2 = FUNCTION_RETURN_SUCCESS
                Exit Function
            End If

            ReDim Preserve STOCKINFO(UBound(STOCKINFO) + 1)  ' 配列の追加

            ' 不良､払いの初期設定
            STOCKINFO(UBound(STOCKINFO)).hinban = HinOld(i).HINBCA
            STOCKINFO(UBound(STOCKINFO)).GENZAL = CLng(HinOld(i).GNLCA)
            STOCKINFO(UBound(STOCKINFO)).FURYOL = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIL = CLng(HinOld(i).GNLCA)
            STOCKINFO(UBound(STOCKINFO)).GENZAW = CLng(HinOld(i).GNWCA)
            STOCKINFO(UBound(STOCKINFO)).FuryoW = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIW = CLng(HinOld(i).GNWCA)
            STOCKINFO(UBound(STOCKINFO)).GENZAM = CLng(HinOld(i).GNMCA)
            STOCKINFO(UBound(STOCKINFO)).FURYOM = 0
            STOCKINFO(UBound(STOCKINFO)).HARAIM = CLng(HinOld(i).GNMCA)
            STOCKINFO(UBound(STOCKINFO)).KCKNT = CLng(HinOld(i).KCKNTCA)
        End If
    Next i

    ' 最抜試指示画面から品番の払い出しと欠落をマップ位置項目から求める
    ' STOCKINFO配列に格納するがSTOCKINFOの品番はHinOldの品番の登録順序と一致しているとは限らない
    If Kihon.NOWPROC = "CW740" Then
        Set objGamenSpd = f_cmbc036_2.sprExamine    ' 抜試変更ｽﾌﾟﾚｯﾄﾞ
    Else
        Set objGamenSpd = f_cmbc039_3.sprExamine    ' 再抜試ｽﾌﾟﾚｯﾄﾞ
    End If

    ' 品番を1列追加したことによる列の変更
    With objGamenSpd
        bMapErrFlg = False                          ' WFﾏｯﾌﾟ位置ﾁｪｯｸﾌﾗｸﾞ初期化

        ' エピ先行評価追加対応
        .GetText 39, 1, vBlkId                      ' ブロックID
        sOldBlkID = CStr(Trim(vBlkId))
        For i = 1 To .MaxRows Step 2                ' ｽﾌﾟﾚｯﾄﾞからﾃﾞｰﾀを入力(2行ずつ確認)
            ' エピ先行評価追加対応
            .GetText 40, i, vGetData                ' 古い品番取得
            sOldHinban = Trim(CStr(vGetData))
            .GetText 2, i, vGetData                 ' 新しい品番取得

            ' 品番が"Z"の時は新品番=旧品番
            If Trim(CStr(vGetData)) = "Z" Then
                sNowHinban = sOldHinban
            Else
                sNowHinban = Trim(CStr(vGetData))
            End If
            .GetText 5, i, vGetData                 ' 結晶位置
            intGetHinInpos = val(vGetData)
            .GetText 6, i, vGetData                 ' マップ開始位置
            intMapSt = val(vGetData)
            .GetText 6, i + 1, vGetData             ' マップ終了位置
            intMapEd = val(vGetData)

            ' エピ先行評価追加対応
            .GetText 39, i, vBlkId                  ' ブロックID
            If vBlkId = "" Then                     ' ブロックがNULLだったら、前回のブロックを使用
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & sOldBlkID
            Else
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & vBlkId
            End If

            ' 良品枚数はマップ位置ではなくテーブルから取得する
            sBlkId = vBlkId
            intREVNUM = gtSprWfMap(i).REVNUM        ' 製品改訂番号
            sFACTORY = gtSprWfMap(i).factory        ' 工場
            sOPE = gtSprWfMap(i).opecond            ' 操業条件

            If ((Kihon.NOWPROC = "CW760") Or (Kihon.NOWPROC = "CW740")) And (vBlkId <> BlkNow.CRYNUMC2) Then
                ' 処理なし
            Else
                ' SXL範囲外はｴﾗｰとする
'                If (Kihon.NOWPROC = "CW760") And ((SIngotP > intGetHinInpos) Or (intGetHinInpos >= EIngotP)) Then
'2010/05/10 Change Y.Hitomi 最終マップ位置緩和（999.9→1000）対応
                If (Kihon.NOWPROC = "CW760") And ((SIngotP > intGetHinInpos) Or (intGetHinInpos > EIngotP)) Then
                    bMapErrFlg = True
                End If

                ' 良品枚数はマップ位置ではなくテーブルから取得する
                sSQL = "SELECT COUNT(*) AS SXLCNT"
                sSQL = sSQL & " FROM TBCMY011 "
                sSQL = sSQL & " WHERE LOTID = '" & sBlkId & "'"
                sSQL = sSQL & " AND (WFSTA ='0' OR WFSTA = '1') "
                sSQL = sSQL & " AND BLOCKSEQ >= " & intMapSt & ""
                sSQL = sSQL & " AND BLOCKSEQ <= " & intMapEd & ""

                Debug.Print sSQL

                Set rs = OraDB.CreateDynaset(sSQL, ORADYN_DEFAULT)

                ' みつからなかったらエラー
                If rs.RecordCount = 0 Then
                    SXLCnt = 0
                Else ' 見つかったら、良品枚数を取得する
                    SXLCnt = val(rs("SXLCNT"))
                End If
                Debug.Print SXLCnt

                blHinFlg = False ' 既存の配列に同じ品番が登録されているかのﾌﾗｸﾞ

                ' udtWSTOCKINFO()は1から開始
                For j = 1 To UBound(udtWSTOCKINFO)
                    If (udtWSTOCKINFO(j).hinban = sOldHinban) Then  ' 既に登録してある品番
                        blHinFlg = True
                        udtWSTOCKINFO(j).HARAIM = udtWSTOCKINFO(j).HARAIM + SXLCnt
                    End If
                Next j

                If (blHinFlg = False) Then   ' udtWSTOCKINFO()の配列に品番が登録なかった時新規にwSTOCKINFO()に登録
                    ReDim Preserve udtWSTOCKINFO(UBound(udtWSTOCKINFO) + 1)     ' 配列の追加
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).hinban = sOldHinban    ' 品番
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).HARAIM = 0             ' 配列初期設定
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).FURYOM = 0             ' 配列初期設定
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).HARAIM = SXLCnt        ' 画面のﾃﾞｰﾀ
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).REVNUM = intREVNUM     ' 画面のﾃﾞｰﾀ
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).factory = sFACTORY     ' 画面のﾃﾞｰﾀ
                    udtWSTOCKINFO(UBound(udtWSTOCKINFO)).OPE = sOPE             ' 画面のﾃﾞｰﾀ
                End If
            End If
        Next i
    End With

    ' STOCKINFとwSTOCKINFの突合せをしてSTOCKINFに品番、長さ、重量、枚数の払い出しと不良を格納
    ' HinOldにないデータはなくなる
    ' STOCKINFO()は添字0から開始
    ' STOCKINFO()の品番はHinOldのﾃﾞｰﾀ、ﾃﾞｰﾀはHinNowのﾃﾞｰﾀ
    For i = 1 To UBound(STOCKINFO)
        STOCKINFO(i).HARAIM = 0
        STOCKINFO(i).FURYOM = 0
        For j = 1 To UBound(udtWSTOCKINFO)
            If (STOCKINFO(i).hinban = udtWSTOCKINFO(j).hinban) Then    ' 品番が等しい時登録する
                STOCKINFO(i).HARAIM = udtWSTOCKINFO(j).HARAIM
                STOCKINFO(i).FURYOM = udtWSTOCKINFO(j).FURYOM
                STOCKINFO(i).REVNUM = udtWSTOCKINFO(j).REVNUM
                STOCKINFO(i).factory = udtWSTOCKINFO(j).factory
                STOCKINFO(i).OPE = udtWSTOCKINFO(j).OPE
                lngTMaisu = udtWSTOCKINFO(j).HARAIM + udtWSTOCKINFO(j).FURYOM   ' 枚数の合計
                If (lngTMaisu > 0) Then   ' 枚数があるか確認
                    STOCKINFO(i).HARAIW = udtWSTOCKINFO(j).HARAIM / lngTMaisu * CLng(STOCKINFO(i).HARAIW)   ' 不良、払出は枚数の比率で算出
                    STOCKINFO(i).FuryoW = CLng(STOCKINFO(i).GENZAW) - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIL = udtWSTOCKINFO(j).HARAIM / lngTMaisu * CLng(STOCKINFO(i).HARAIL)   ' 不良、払出は枚数の比率で算出
                    STOCKINFO(i).FURYOL = CLng(STOCKINFO(i).GENZAL) - STOCKINFO(i).HARAIL
                 End If
            End If
        Next j
    Next i

    ' 不良がある場合現在庫減情報の作成
    For i = 1 To UBound(STOCKINFO)
        If STOCKINFO(i).FURYOM > 0 Then
            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' ブロックＩＤ
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                 ' 位置
            udtKoutei.KCNTC3 = STOCKINFO(i).KCKNT + 1   ' 工程連番
            udtKoutei.HINBC3 = STOCKINFO(i).hinban      ' 品番
            udtKoutei.REVNUMC3 = STOCKINFO(i).REVNUM    ' 製品改訂番号
            udtKoutei.FACTORYC3 = STOCKINFO(i).factory  ' 工場
            udtKoutei.OPEC3 = STOCKINFO(i).OPE          ' 操業条件
            udtKoutei.LENC3 = STOCKINFO(i).HARAIL       ' 長さ
            udtKoutei.XTALC3 = HinNow(0).XTALCA         ' 結晶番号
            udtKoutei.SXLIDC3 = ""                      ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1)   ' 管理工程(現在工程+1)
            udtKoutei.WKKTC3 = Kihon.NOWPROC            ' 工程
            udtKoutei.WKKBC3 = ""                       ' 作業区分
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' 処理回数
            udtKoutei.MODKBC3 = ""                      ' 赤黒区分
            udtKoutei.SUMKBC3 = ""                      ' 集計区分
            udtKoutei.FRKNKTC3 = ""                     ' (受入)管理工程

            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If

            udtKoutei.FRWKKBC3 = ""                     ' (受入)作業区分

            If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                udtKoutei.TOWKKTC3 = " "                            ' (払出)工程
                udtKoutei.TOMACOC3 = "0"                            ' (払出)処理回数
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (払出)工程
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (払出)処理回
            End If

            udtKoutei.FRLC3 = STOCKINFO(i).GENZAL                   ' 受入長さ
            udtKoutei.FRWC3 = STOCKINFO(i).GENZAW                   ' 受入重量
            udtKoutei.FRMC3 = STOCKINFO(i).GENZAM                   ' 受入枚数
            udtKoutei.FULC3 = STOCKINFO(i).FURYOL                   ' 不良長さ
            udtKoutei.FUWC3 = STOCKINFO(i).FuryoW                   ' 不良重量
            udtKoutei.FUMC3 = STOCKINFO(i).FURYOM                   ' 不良枚数
            udtKoutei.LOSWC3 = ""                                   ' ロス長さ

            udtKoutei.LOSLC3 = ""                                   ' ロス重量
            udtKoutei.LOSMC3 = ""                                   ' ロス枚数
            udtKoutei.TOLC3 = STOCKINFO(i).HARAIL                   ' 払出長さ
            udtKoutei.TOWC3 = STOCKINFO(i).HARAIW                   ' 払出重量
            udtKoutei.TOMC3 = STOCKINFO(i).HARAIM                   ' 払出枚数
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT長さ
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT重量
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT枚数
            udtKoutei.MOTHINC3 = ""                                 ' 振替品番(元)
            udtKoutei.XTWORKC3 = "42"                               ' 製造工場

            udtKoutei.WFWORKC3 = ""                                 ' ｳｪｰﾊ製造
            udtKoutei.HOLDCC3 = " "                                 ' ホールドコード
            udtKoutei.HOLDBC3 = "0"                                 ' ホールド区分
            udtKoutei.LDFRCC3 = ""                                  ' 格下コード
            udtKoutei.LDFRBC3 = "0"                                 ' 格下区分（ハイキ）
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' 登録社員ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付

            udtKoutei.KSTAFFC3 = ""                                 ' 更新社員ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT送信フラグ
            udtKoutei.SNDKC3 = ""                                   ' 送信フラグ
            udtKoutei.MODMACOC3 = ""                                ' 赤黒の処理回数
            udtKoutei.KAKUCC3 = ""                                  ' 確定コード
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO時間
            udtKoutei.PAYCLASSC3 = ""                               ' 転送先工場フラグ
            udtKoutei.PLANTCATC3 = HinNow(0).PLANTCATCA             ' 向先
            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' 工程実績に在庫減情報登録
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' 工程実績追加エラー
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDC3Proc3
'*
'*    処理概要      : 1.工程実績登録処理を行う(品番振替情報：CW740,CW760用)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc3() As FUNCTION_RETURN
    ' 内部変数
    Dim i, j            As Integer
    Dim intRtn          As Integer          ' 復帰情報
    Dim sSQL            As String           ' ＳＱＬ
    Dim sSqlWhere       As String           ' WHERE句
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update ' 工程実績

    Dim lngLen          As Long
    Dim lngCHKPOS       As Long

    Dim udtWOINF()      As typ_trans_info   ' 前品番並び替え用
    Dim udtWNINF()      As typ_trans_info   ' 後品番並び替え用
    Dim udtWWINF()      As typ_trans_info   ' 並び替え用ワーク
    Dim intBuf          As Integer
    Dim intOINFrecCnt   As Integer
    Dim intNINFrecCnt   As Integer
    Dim intOINFFLG      As Integer
    Dim intNINFFLG      As Integer
    Dim intCnt          As Integer
    Dim intWNINFMAX     As Integer
    Dim intWOINFMAX     As Integer
    Dim intHantei       As Integer

    ' エラーハンドラの設定
    On Error GoTo proc_err

    ' 初期設定
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE

    ReDim udtWOINF(UBound(STOCKINFO))      ' 品番ごとにソート用
    ReDim udtWNINF(Kihon.CNTHINNOW)        ' 品番ごとにソート用
    ReDim udtWWINF(1)                      ' 品番ごとにソート用

    ' 在庫減情報より取り込み
    For i = 1 To UBound(STOCKINFO)
        If Trim(STOCKINFO(i).hinban) <> "" Then
            udtWOINF(i).hinban = STOCKINFO(i).hinban
            udtWOINF(i).LEN = STOCKINFO(i).HARAIL       ' 前工程長さ合計
            udtWOINF(i).WAT = STOCKINFO(i).HARAIW       ' 前工程重量合計
            udtWOINF(i).MAI = STOCKINFO(i).HARAIM       ' 前工程枚数合計
        End If
    Next i

    ' 良品情報より取り込み
    For j = 0 To Kihon.CNTHINNOW - 1
        If (Kihon.NOWPROC = "CW760") _
            And ((SIngotP > HinNow(j).INPOSCA) Or (HinNow(j).INPOSCA >= EIngotP)) Then
                ' 処理なし
        Else
            udtWNINF(j).hinban = HinNow(j).HINBCA
            udtWNINF(j).REVNUM = HinNow(j).REVNUMCA
            udtWNINF(j).factory = HinNow(j).FACTORYCA
            udtWNINF(j).OPE = HinNow(j).OPECA
            udtWNINF(j).LEN = HinNow(j).GNLCA           ' 後工程長さ合計
            udtWNINF(j).WAT = HinNow(j).GNWCA           ' 後工程重量合計
            udtWNINF(j).MAI = HinNow(j).GNMCA           ' 後工程枚数合計
            udtWNINF(j).KCKNT = HinNow(j).KCKNTCA       ' 後工程連番
        End If
    Next j

    ' 同じ品番同士、長さ等を打ち消す
    For i = 1 To UBound(STOCKINFO)
        For j = 0 To Kihon.CNTHINNOW - 1
           If udtWOINF(i).hinban = udtWNINF(j).hinban Then
                ' 小さい方の数字を両方から引く
                If udtWOINF(i).MAI <= udtWNINF(j).MAI Then
                    udtWNINF(j).LEN = udtWNINF(j).LEN - udtWOINF(i).LEN
                    udtWNINF(j).WAT = udtWNINF(j).WAT - udtWOINF(i).WAT
                    udtWNINF(j).MAI = udtWNINF(j).MAI - udtWOINF(i).MAI
                    udtWOINF(i).LEN = 0
                    udtWOINF(i).WAT = 0
                    udtWOINF(i).MAI = 0
                Else
                    udtWOINF(i).LEN = udtWOINF(i).LEN - udtWNINF(j).LEN
                    udtWOINF(i).WAT = udtWOINF(i).WAT - udtWNINF(j).WAT
                    udtWOINF(i).MAI = udtWOINF(i).MAI - udtWNINF(j).MAI
                    If udtWOINF(i).MAI < 0 Then
                        udtWOINF(i).MAI = 0
                    End If
                    udtWNINF(j).LEN = 0
                    udtWNINF(j).WAT = 0
                    udtWNINF(j).MAI = 0
                End If
            End If
        Next
    Next

    For i = 0 To UBound(udtWOINF) - 2
        For j = i + 1 To UBound(udtWOINF) - 1
            If (StrComp(udtWOINF(i).hinban, udtWOINF(j).hinban, _
                vbTextCompare)) = 1 Then ' 品番の入替必要
                udtWWINF(0) = udtWOINF(j)
                udtWOINF(j) = udtWOINF(i)
                udtWOINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' wNINFの品番をソートする
    For i = 0 To UBound(udtWNINF) - 2
        For j = i + 1 To UBound(udtWNINF) - 1
            If (StrComp(udtWNINF(i).hinban, udtWNINF(j).hinban, _
                vbTextCompare)) = 1 Then ' 品番の入替必要
                udtWWINF(0) = udtWNINF(j)
                udtWNINF(j) = udtWNINF(i)
                udtWNINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' 空きの配列削除する(配列のデータを詰める)
    For i = 0 To intWOINFMAX
        If udtWOINF(i).MAI <= 0 Then
            intCnt = i
            Call HairetuOpe_Mai(udtWOINF(), intCnt, -1)
        End If
    Next i

    ' 空きの配列削除する(配列のデータを詰める)
    For i = 0 To intWNINFMAX
        If udtWNINF(i).MAI <= 0 Then
            intCnt = i
            Call HairetuOpe_Mai(udtWNINF(), intCnt, -1)
        End If
    Next i

    ' 品番入替情報を作成する
    i = 0 ' 前品番の位置
    j = 0 ' 後品番の位置
    Do
        ' 枚数を突き合わせて数量が同じでなかったら大きい値の品番を分割する
        If (udtWOINF(i).MAI = udtWNINF(j).MAI) Then         ' 品番長さが同じ時両方とも次に進む
        ElseIf (udtWOINF(i).MAI > udtWNINF(j).MAI) Then     ' 品番長さが異なる時
            intCnt = i
            Call HairetuOpe(udtWOINF(), intCnt, 1)          ' 配列の追加
            udtWOINF(i + 1).hinban = udtWOINF(i).hinban
            udtWOINF(i + 1).LEN = udtWOINF(i).LEN - udtWNINF(j).LEN
            udtWOINF(i + 1).WAT = udtWOINF(i).WAT - udtWNINF(j).WAT
            udtWOINF(i + 1).MAI = udtWOINF(i).MAI - udtWNINF(j).MAI
            udtWOINF(i).LEN = udtWNINF(j).LEN
            udtWOINF(i).WAT = udtWNINF(j).WAT
            udtWOINF(i).MAI = udtWNINF(j).MAI
        ElseIf (udtWOINF(i).MAI < udtWNINF(j).MAI) Then     ' 品番数量が異なる時
            intCnt = j
            Call HairetuOpe(udtWNINF(), intCnt, 1)
            udtWNINF(j + 1).hinban = udtWNINF(i).hinban
            udtWNINF(j + 1).LEN = udtWNINF(j).LEN - udtWOINF(i).LEN
            udtWNINF(j + 1).WAT = udtWNINF(j).WAT - udtWOINF(i).WAT
            udtWNINF(j + 1).MAI = udtWNINF(j).MAI - udtWOINF(i).MAI
            udtWNINF(j).LEN = udtWOINF(i).LEN
            udtWNINF(j).WAT = udtWOINF(i).WAT
            udtWNINF(j).MAI = udtWOINF(i).MAI
        End If
        intOINFrecCnt = UBound(udtWOINF())
        intNINFrecCnt = UBound(udtWNINF())
        i = i + 1
        j = j + 1
        If (i > intOINFrecCnt) Then
            Exit Do

        End If
        If (j > intNINFrecCnt) Then
            Exit Do
        End If
        If (udtWOINF(i).MAI) <= 0 Then
            Exit Do
        End If
        If (udtWNINF(j).MAI) <= 0 Then
            Exit Do
        End If
    Loop

    intOINFrecCnt = UBound(udtWOINF())
    For i = 0 To intOINFrecCnt
        If (StrComp(udtWNINF(i).hinban, udtWOINF(i).hinban, vbTextCompare) <> 0) Then  ' 品番が異なる時振替情報に登録する
            If Trim(udtWNINF(i).hinban) <> "" And udtWNINF(i).LEN > 0 Then
                ' SXL範囲外はｴﾗｰとする
                If bMapErrFlg Then
                    MsgBox "結晶Pチェックエラー (SXL位置 ： " & SIngotP & "－" & EIngotP & ")"
                    Exit Function
                End If

                udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' ブロックＩＤ
                giInpos = giInpos + 1
                udtKoutei.INPOSC3 = giInpos                 ' 位置
                udtKoutei.KCNTC3 = udtWNINF(i).KCKNT        ' 工程連番
                udtKoutei.HINBC3 = udtWNINF(i).hinban       ' 品番
                udtKoutei.REVNUMC3 = udtWNINF(i).REVNUM     ' 製品改訂番号
                udtKoutei.FACTORYC3 = udtWNINF(i).factory   ' 工場
                udtKoutei.OPEC3 = udtWNINF(i).OPE           ' 操業条件
                udtKoutei.LENC3 = udtWNINF(i).LEN           ' 受入長さ
                udtKoutei.XTALC3 = HinNow(0).XTALCA         ' 結晶番号
                udtKoutei.SXLIDC3 = ""                      ' SXLID

                udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                  CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2)   ' 管理工程(現在工程+2)
                udtKoutei.WKKTC3 = Kihon.NOWPROC            ' 工程
                udtKoutei.WKKBC3 = ""                       ' 作業区分
                udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' 処理回数
                udtKoutei.MODKBC3 = ""                      ' 赤黒区分
                udtKoutei.SUMKBC3 = ""                      ' 集計区分
                udtKoutei.FRKNKTC3 = ""                     ' (受入)管理工程
                If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                    udtKoutei.FRWKKTC3 = ""
                Else
                    udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                End If
                udtKoutei.FRWKKBC3 = ""                     ' (受入)作業区分
                If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                    udtKoutei.FRMACOC3 = "0"
                Else
                    udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
                End If

                Select Case Kihon.NOWPROC
                    Case "CC730"
                        intHantei = CInt(BlkNow.GNLC2)
                    Case Else
                        intHantei = CInt(BlkNow.GNMC2)
                End Select

                If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                    udtKoutei.TOWKKTC3 = " "                            ' (払出)工程
                    udtKoutei.TOMACOC3 = "0"                            '(払出)処理回数
                Else
                    udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (払出)工程
                    udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (払出)処理回数
                End If
                udtKoutei.FRLC3 = udtWNINF(i).LEN                       '受入長さ
                udtKoutei.FRWC3 = udtWNINF(i).WAT                       '受入重量
                udtKoutei.FRMC3 = udtWNINF(i).MAI                       '受入枚数
                udtKoutei.FULC3 = 0                                     '不良長さ
                udtKoutei.FUWC3 = 0                                     '不良重量
                udtKoutei.FUMC3 = 0                                     '不良枚数
                udtKoutei.LOSWC3 = ""                                   ' ロス長さ

                udtKoutei.LOSLC3 = ""                                   ' ロス重量
                udtKoutei.LOSMC3 = ""                                   ' ロス枚数
                udtKoutei.TOLC3 = udtWNINF(i).LEN                       '払出長さ
                udtKoutei.TOWC3 = udtWNINF(i).WAT                       '払出重量
                udtKoutei.TOMC3 = udtWNINF(i).MAI                       '払出枚数
                udtKoutei.SUMITLC3 = ""                                 ' SUMIT長さ
                udtKoutei.SUMITWC3 = ""                                 ' SUMIT重量
                udtKoutei.SUMITMC3 = ""                                 ' SUMIT枚数
                udtKoutei.MOTHINC3 = udtWOINF(i).hinban                 '元品番
                udtKoutei.XTWORKC3 = "42"                               ' 製造工場

                udtKoutei.WFWORKC3 = ""                                 ' ｳｪｰﾊ製造
                udtKoutei.HOLDCC3 = " "                                 ' ホールドコード
                udtKoutei.HOLDBC3 = "0"                                 ' ホールド区分
                udtKoutei.LDFRCC3 = ""                                  ' 格下コード
                udtKoutei.LDFRBC3 = "0"                                 ' 格下区分（ハイキ）
                udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' 登録社員ID
                udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付

                udtKoutei.KSTAFFC3 = ""                                 ' 更新社員ID
                udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                udtKoutei.SUMITBC3 = ""                                 ' SUMIT送信フラグ
                udtKoutei.SNDKC3 = ""                                   ' 送信フラグ
                udtKoutei.MODMACOC3 = ""                                ' 赤黒の処理回数
                udtKoutei.KAKUCC3 = ""                                  ' 確定コード
                udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO時間
                udtKoutei.PAYCLASSC3 = ""                               ' 転送先工場フラグ
                udtKoutei.PLANTCATC3 = HinNow(0).PLANTCATCA             ' 向先

                intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' 工程実績に在庫減情報登録
                If intRtn = FUNCTION_RETURN_FAILURE Then                ' 工程実績追加エラー
                    MsgBox sErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDC3Proc3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDC3Proc4
'*
'*    処理概要      : 1.工程実績登録処理を行う(在庫減情報：CC730用)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc4() As FUNCTION_RETURN
    ' 内部変数
    Dim i, j                    As Integer
    Dim intRtn                  As Integer          ' 復帰情報
    Dim sSQL                    As String           ' ＳＱＬ
    Dim rs                      As OraDynaset       ' レコードセット
    Dim sSqlWhere               As String           ' WHERE句
    Dim sErrMsg                 As String
    Dim udtKoutei               As typ_XSDC3_Update ' 工程実績
    Dim rsKCNTC                 As OraDynaset       ' レコードセット
    Dim intNextCnt              As Integer
    Dim lngLen                  As Long
    Dim lngCHKPOS               As Long
    Dim intBadcnt               As Integer
    Dim udtBADINFO()            As typ_bad_info
    Dim udtWSTOCKINFO()         As typ_stock_info
    Dim intLoopCnt              As Integer
    Dim vGetMaxPos              As Variant
    Dim vGetData                As Variant
    Dim sOldHinban, sNowHinban  As String
    Dim intSXLCnt               As Integer
    Dim intMapSt, intMapEd      As Integer
    Dim blHinFlg                As Boolean
    Dim lngTMaisu               As Long
    Dim vNukisiFlg              As Variant
    Dim intHantei               As Integer

    ' エラーハンドラの設定
    On Error GoTo 0

    ' 初期設定
    XSDC3Proc4 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(Kihon.CNTHINOLD)
    ReDim udtWSTOCKINFO(0)

    ' HinOldから前工程長さ,重量,枚数合計取得(長さは0)
    For i = 0 To Kihon.CNTHINOLD - 1
        FRLC3Sum = FRLC3Sum + CLng(HinOld(i).GNLCA)     ' 前工程長さ合計
        FRWC3Sum = FRWC3Sum + CLng(HinOld(i).GNWCA)     ' 前工程重量合計
        FRMC3Sum = FRMC3Sum + CLng(HinOld(i).GNMCA)     ' 前工程枚数合計

        ' 不良､払いの初期設定
        STOCKINFO(i).hinban = HinOld(i).HINBCA
        STOCKINFO(i).FURYOL = 0
        STOCKINFO(i).HARAIL = CLng(HinOld(i).GNLCA)
        STOCKINFO(i).FuryoW = CLng(HinOld(i).GNWCA)     ' 不良重量に払い重量を仮に代入して後で計算する
        STOCKINFO(i).HARAIW = CLng(HinOld(i).GNWCA)
        STOCKINFO(i).FURYOM = CLng(HinOld(i).GNMCA)     ' 不良枚数に払い枚数を仮に代入し後で計算する
        STOCKINFO(i).HARAIM = CLng(HinOld(i).GNMCA)
        STOCKINFO(i).KCKNT = CLng(HinOld(i).KCKNTCA)
        STOCKINFO(i).REVNUM = HinOld(i).REVNUMCA        ' 製品改訂番号
        STOCKINFO(i).factory = HinOld(i).FACTORYCA      ' 工場
        STOCKINFO(i).OPE = HinOld(i).OPECA              ' 製品改訂番号
    Next i

    ' 最抜試指示画面から品番の払い出しと欠落をマップ位置項目から求める
    ' STOCKINFO配列に格納するがSTOCKINFOの品番はHinOldの品番の登録順序と一致しているとは限らない
    intBadcnt = 0  ' 不良数初期設定

    ' 不良が先頭にないか確認
    If ((CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)) > 0) Then ' 前後開始位置を比較して差があれば不良位置登録
        intBadcnt = intBadcnt + 1
        ReDim Preserve udtBADINFO(intBadcnt)
        udtBADINFO(intBadcnt).pos = CLng(HinOld(0).INPOSCA)
        udtBADINFO(intBadcnt).LEN = CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)
    End If

    ' 不良長さが品番間にないか確認
    For i = 0 To Kihon.CNTHINNOW - 2
        If (CLng(HinNow(i + 1).INPOSCA) > (CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA))) Then ' 品番間に不良有
            intBadcnt = intBadcnt + 1 ' 不良位置の登録
            ReDim Preserve udtBADINFO(intBadcnt)
            udtBADINFO(intBadcnt).pos = CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA)
            udtBADINFO(intBadcnt).LEN = CLng(HinNow(i + 1).INPOSCA) - CLng(HinNow(i).INPOSCA) - CLng(HinNow(i).GNLCA)
        End If
    Next i

    ' 不良が最後にないか前の確認(結晶内開始位置+長さで比較)
    If ((CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA)) _
        <> (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))) Then ' 終了位置の確認
        intBadcnt = intBadcnt + 1
        ReDim Preserve udtBADINFO(intBadcnt)
        udtBADINFO(intBadcnt).pos = (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
        udtBADINFO(intBadcnt).LEN = CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA) - (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
    End If

    If (intBadcnt = 0) Then  ' 前と後で不良なし 処理終了
        XSDC3Proc4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

    ' 不良位置を振替前の結晶位置を確認して不良位置に相当する品番を登録する
    If BlkOld.GNLC2 < BlkNow.GNLC2 Then
        For i = 1 To intBadcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + udtBADINFO(i).LEN   ' 不良の長さ(HinOld(i)のﾁｪｯｸしている品番の長さ不良)
                STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - udtBADINFO(i).LEN   ' 良品長さ
            Next j
        Next i
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then           ' 不良数が存在したら不良重さを不良比率で求める
                If i = Kihon.CNTHINOLD - 1 Then
                    STOCKINFO(i).HARAIW = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIW)  'STOCKINFO(i).HARAIWは入力済み  ’2003/08/06 hitec)matsumoto ROUND追加
                    STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIM = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIM)  'STOCKINFO(i).HARAIMは入力済み  ’2003/08/06 hitec)matsumoto ROUND追加
                    STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM
                Else
                    STOCKINFO(i).HARAIW = HinOld(i).GNWCA
                    STOCKINFO(i).FuryoW = 0
                    STOCKINFO(i).HARAIM = HinOld(i).GNMCA
                    STOCKINFO(i).FURYOM = 0
                End If
            End If
        Next i
    Else
        ' STOCKINFOの払いは既に入力済み
        For i = 1 To intBadcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                If (udtBADINFO(i).pos >= CLng(HinOld(j).INPOSCA) And _
                    udtBADINFO(i).pos < CLng(HinOld(j).INPOSCA) + CLng(HinOld(j).GNLCA)) Then
                    STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + udtBADINFO(i).LEN   ' 不良の長さ(HinOld(i)のﾁｪｯｸしている品番の長さ不良)
                    STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - udtBADINFO(i).LEN   ' 良品長さ
                End If
            Next j
        Next i

        ' 重量と枚数の不良と払い出しの値を設定する
        ' STOCKINFO(i).HARAIWは入力済み
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then           ' 不良数が存在したら不良重さを不良比率で求める
                STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW
                STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW
                STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM
                STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM
            End If
        Next i
    End If

    ' 不良がある場合現在庫減情報の作成
    For i = 0 To Kihon.CNTHINOLD - 1
        If STOCKINFO(i).FURYOL <> 0 Then
            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA                 ' ブロックＩＤ
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                             ' 位置
            udtKoutei.KCNTC3 = STOCKINFO(i).KCKNT + 1               ' 工程連番
            udtKoutei.HINBC3 = HinOld(i).HINBCA                     ' 品番
            udtKoutei.REVNUMC3 = HinOld(i).REVNUMCA                 ' 製品改訂番号
            udtKoutei.FACTORYC3 = HinOld(i).FACTORYCA               ' 工場
            udtKoutei.OPEC3 = HinOld(i).OPECA                       ' 操業条件
            udtKoutei.LENC3 = STOCKINFO(i).HARAIL                   ' 長さ
            udtKoutei.XTALC3 = HinOld(i).XTALCA                     ' 結晶番号
            udtKoutei.SXLIDC3 = ""                                  ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1)               ' 管理工程(現在工程+1)

            udtKoutei.WKKTC3 = Kihon.NOWPROC                        ' 工程
            udtKoutei.WKKBC3 = ""                                   ' 作業区分
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA                   ' 処理回数
            udtKoutei.MODKBC3 = ""                                  ' 赤黒区分
            udtKoutei.SUMKBC3 = ""                                  ' 集計区分
            udtKoutei.FRKNKTC3 = ""                                 ' (受入)管理工程

            If IsNull(HinOld(0).NEWKNTCA) = True Then               '(受入）工程
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If
            udtKoutei.FRWKKBC3 = ""                                 ' (受入)作業区分
            If IsNull(HinOld(0).NEMACOCA) = True Then               '（受入）処理回数
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOWKKTC3 = " "                            ' (払出)工程
                udtKoutei.TOMACOC3 = "0"                            ' (払出)処理回数
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (払出)工程
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (払出)処理回
            End If
            udtKoutei.FRLC3 = HinOld(i).GNLCA                       ' 受入長さ
            udtKoutei.FRWC3 = HinOld(i).GNWCA                       ' 受入重量
            udtKoutei.FRMC3 = HinOld(i).GNMCA                       ' 受入枚数
            udtKoutei.FULC3 = STOCKINFO(i).FURYOL                   ' 不良長さ
            udtKoutei.FUWC3 = STOCKINFO(i).FuryoW                   ' 不良重量
            udtKoutei.FUMC3 = STOCKINFO(i).FURYOM                   ' 不良枚数
            udtKoutei.LOSWC3 = ""                                   ' ロス長さ

            udtKoutei.LOSLC3 = ""                                   ' ロス重量
            udtKoutei.LOSMC3 = ""                                   ' ロス枚数
            udtKoutei.TOLC3 = STOCKINFO(i).HARAIL                   ' 払出長さ
            udtKoutei.TOWC3 = STOCKINFO(i).HARAIW                   ' 払出重量
            udtKoutei.TOMC3 = STOCKINFO(i).HARAIM                   ' 払出枚数
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT長さ
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT重量
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT枚数
            udtKoutei.MOTHINC3 = " "                                ' 元品番
            udtKoutei.XTWORKC3 = "42"                               ' 製造工場

            udtKoutei.WFWORKC3 = ""                                 ' ｳｪｰﾊ製造
            udtKoutei.HOLDCC3 = " "                                 ' ホールドコード
            udtKoutei.HOLDBC3 = "0"                                 ' ホールド区分
            udtKoutei.LDFRCC3 = ""                                  ' 格下コード
            udtKoutei.LDFRBC3 = "0"                                 ' 格下区分（ハイキ）
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' 登録社員ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付

            udtKoutei.KSTAFFC3 = ""                                 ' 更新社員ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT送信フラグ
            udtKoutei.SNDKC3 = ""                                   ' 送信フラグ
'           udtKoutei.SNDDAYC3 = ""                                 ' 送信日付
            udtKoutei.MODMACOC3 = ""                                ' 赤黒の処理回数
            udtKoutei.KAKUCC3 = ""                                  ' 確定コード
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO時間
            udtKoutei.PAYCLASSC3 = ""                               ' 転送先工場フラグ

            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' 工程実績に在庫減情報登録
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' 工程実績追加エラー
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function

proc_err:
    ' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'*******************************************************************************
'*    関数名        : XSDC3Proc5
'*
'*    処理概要      : 1.工程実績登録処理を行う(品番振替情報：CC730用)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Public Function XSDC3Proc5() As FUNCTION_RETURN
    ' 内部変数
    Dim i, j            As Integer
    Dim intRtn          As Integer                  ' 復帰情報
    Dim sSQL            As String                   ' ＳＱＬ
    Dim sSqlWhere       As String                   ' WHERE句
    Dim sErrMsg         As String
    Dim udtKoutei       As typ_XSDC3_Update         ' 工程実績

    Dim lngLen          As Long
    Dim lngCHKPOS       As Long

    Dim udtWCHKPOS()    As typ_trans_info           ' 前品番並び替え用
    Dim udtWNINF()      As typ_trans_info           ' 後品番並び替え用
    Dim udtWWINF()      As typ_trans_info           ' 並び替え用ワーク
    Dim intBuf          As Integer
    Dim intOINFrecCnt   As Integer
    Dim intNINFrecCnt   As Integer
    Dim intOINFFLG      As Integer
    Dim intNINFFLG      As Integer
    Dim intPoint        As Integer
    Dim intNINFMAX      As Integer
    Dim intOINFMAX      As Integer
    Dim intHantei       As Integer

    ' エラーハンドラの設定
    On Error GoTo proc_err

    ' 初期設定
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE

    ReDim udtWCHKPOS(UBound(STOCKINFO))             ' 品番ごとにソート用
    ReDim udtWNINF(Kihon.CNTHINNOW)                 ' 品番ごとにソート用
    ReDim udtWWINF(1)                               ' 品番ごとにソート用

    ' 在庫減情報より取り込み
    For i = 0 To UBound(STOCKINFO) - 1
        udtWCHKPOS(i).hinban = STOCKINFO(i).hinban
        udtWCHKPOS(i).LEN = STOCKINFO(i).HARAIL     ' 前工程長さ合計
        udtWCHKPOS(i).WAT = STOCKINFO(i).HARAIW     ' 前工程重量合計
        udtWCHKPOS(i).MAI = STOCKINFO(i).HARAIM     ' 前工程枚数合計
    Next i

    ' 良品情報より取り込み
    For j = 0 To Kihon.CNTHINNOW - 1
        udtWNINF(j).hinban = HinNow(j).HINBCA
        udtWNINF(j).LEN = HinNow(j).GNLCA           ' 後工程長さ合計
        udtWNINF(j).WAT = HinNow(j).GNWCA           ' 後工程重量合計
        udtWNINF(j).MAI = HinNow(j).GNMCA           ' 後工程枚数合計
        udtWNINF(j).KCKNT = HinNow(j).KCKNTCA       ' 後工程連番
    Next j

    ' 同じ品番同士、長さ等を打ち消す
    For i = 0 To UBound(STOCKINFO) - 1
        For j = 0 To Kihon.CNTHINNOW - 1
           If udtWCHKPOS(i).hinban = udtWNINF(j).hinban Then
                ' 小さい方の数字を両方から引く
                If udtWCHKPOS(i).LEN <= udtWNINF(j).LEN Then
                    udtWNINF(j).LEN = udtWNINF(j).LEN - udtWCHKPOS(i).LEN
                    udtWNINF(j).WAT = udtWNINF(j).WAT - udtWCHKPOS(i).WAT
                    udtWNINF(j).MAI = udtWNINF(j).MAI - udtWCHKPOS(i).MAI
                    udtWCHKPOS(i).LEN = 0
                    udtWCHKPOS(i).WAT = 0
                    udtWCHKPOS(i).MAI = 0
                Else
                    udtWCHKPOS(i).LEN = udtWCHKPOS(i).LEN - udtWNINF(j).LEN
                    udtWCHKPOS(i).WAT = udtWCHKPOS(i).WAT - udtWNINF(j).WAT
                    udtWCHKPOS(i).MAI = udtWCHKPOS(i).MAI - udtWNINF(j).MAI
                    If udtWCHKPOS(i).MAI < 0 Then
                        udtWCHKPOS(i).MAI = 0
                    End If
                    udtWNINF(j).LEN = 0
                    udtWNINF(j).WAT = 0
                    udtWNINF(j).MAI = 0
                End If
            End If
        Next
    Next

    For i = 0 To UBound(udtWCHKPOS) - 2
        For j = i + 1 To UBound(udtWCHKPOS) - 1
            If (StrComp(udtWCHKPOS(i).hinban, udtWCHKPOS(j).hinban, _
                vbTextCompare)) = 1 Then ' 品番の入替必要
                udtWWINF(0) = udtWCHKPOS(j)
                udtWCHKPOS(j) = udtWCHKPOS(i)
                udtWCHKPOS(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' wNINFの品番をソートする
    For i = 0 To UBound(udtWNINF) - 2
        For j = i + 1 To UBound(udtWNINF) - 1
            If (StrComp(udtWNINF(i).hinban, udtWNINF(j).hinban, _
                vbTextCompare)) = 1 Then ' 品番の入替必要
                udtWWINF(0) = udtWNINF(j)
                udtWNINF(j) = udtWNINF(i)
                udtWNINF(i) = udtWWINF(0)
            End If
        Next j
    Next i

    ' 空きの配列削除する(配列のデータを詰める)
    For i = 0 To intOINFMAX
        If udtWCHKPOS(i).LEN <= 0 Then
            intPoint = i
            Call HairetuOpe(udtWCHKPOS(), intPoint, -1)
        End If
    Next i

    ' 空きの配列削除する(配列のデータを詰める)
    For i = 0 To intNINFMAX
        If udtWNINF(i).LEN <= 0 Then
            intPoint = i
            Call HairetuOpe(udtWNINF(), intPoint, -1)
        End If
    Next i

    ' 品番入替情報を作成する
    i = 0 ' 前品番の位置
    j = 0 ' 後品番の位置
    Do
        ' 長さを突き合わせて数量が同じでなかったら大きい値の品番を分割する
        If (udtWCHKPOS(i).LEN = udtWNINF(j).LEN And udtWCHKPOS(i).hinban = udtWNINF(j).hinban) Then   ' 品番長さが同じ時両方とも次に進む
        ElseIf (udtWCHKPOS(i).LEN >= udtWNINF(j).LEN) Then  ' 品番枚数が異なる時
            intPoint = i
            Call HairetuOpe(udtWCHKPOS(), intPoint, 1)      ' 配列の追加
            udtWCHKPOS(i + 1).hinban = udtWCHKPOS(i).hinban
            udtWCHKPOS(i + 1).LEN = udtWCHKPOS(i).LEN - udtWNINF(j).LEN
            udtWCHKPOS(i + 1).WAT = udtWCHKPOS(i).WAT - udtWNINF(j).WAT
            udtWCHKPOS(i + 1).MAI = udtWCHKPOS(i).MAI - udtWNINF(j).MAI
            If udtWCHKPOS(i + 1).MAI < 0 Then
                udtWCHKPOS(i + 1).MAI = 0
            End If
            udtWCHKPOS(i).LEN = udtWNINF(j).LEN
            udtWCHKPOS(i).WAT = udtWNINF(j).WAT
            udtWCHKPOS(i).MAI = udtWNINF(j).MAI
            Debug.Print "HINBAN=", i, udtWCHKPOS(i).hinban
            Debug.Print "LEN=", i, udtWCHKPOS(i).LEN
        ElseIf (udtWCHKPOS(i).LEN < udtWNINF(j).LEN) Then   ' 品番数量が異なる時
            intPoint = j
            Call HairetuOpe(udtWNINF(), intPoint, 1)
            udtWNINF(j + 1).hinban = udtWNINF(i).hinban
            udtWNINF(j + 1).LEN = udtWNINF(j).LEN - udtWCHKPOS(i).LEN
            udtWNINF(j + 1).WAT = udtWNINF(j).WAT - udtWCHKPOS(i).WAT
            udtWNINF(j + 1).MAI = udtWNINF(j).MAI - udtWCHKPOS(i).MAI
            udtWNINF(j).LEN = udtWCHKPOS(i).LEN
            udtWNINF(j).WAT = udtWCHKPOS(i).WAT
            udtWNINF(j).MAI = udtWCHKPOS(i).MAI
            Debug.Print "HINBAN=", i, udtWNINF(i).hinban
            Debug.Print "LEN=", i, udtWNINF(i).LEN
        End If

        intOINFrecCnt = UBound(udtWCHKPOS())
        intNINFrecCnt = UBound(udtWNINF())
        i = i + 1
        j = j + 1
        If (i > intOINFrecCnt) Then
            Exit Do

        End If
        If (j > intNINFrecCnt) Then
            Exit Do
        End If

        If (udtWCHKPOS(i).LEN) <= 0 Then
            Exit Do
        End If

        If (udtWNINF(j).LEN) <= 0 Then
            Exit Do
        End If
    Loop

    intOINFrecCnt = UBound(udtWCHKPOS())
    For i = 0 To intOINFrecCnt - 1
        If (StrComp(udtWNINF(i).hinban, udtWCHKPOS(i).hinban, vbTextCompare) <> 0 _
            And Len(Trim(udtWNINF(i).hinban) > 0)) And (udtWNINF(i).LEN > 0) Then ' 品番が異なる時振替情報に登録する   'upd 2003/05/31 hitec)matsumoto udtWNINF(i).LEN > 0追加

            udtKoutei.CRYNUMC3 = HinNow(0).CRYNUMCA     ' ブロックＩＤ
            giInpos = giInpos + 1
            udtKoutei.INPOSC3 = giInpos                 ' 位置
            udtKoutei.KCNTC3 = udtWNINF(i).KCKNT        ' 工程連番
            udtKoutei.HINBC3 = udtWNINF(i).hinban       ' 品番
            udtKoutei.REVNUMC3 = HinNow(0).REVNUMCA     ' 製品改訂番号
            udtKoutei.FACTORYC3 = HinNow(0).FACTORYCA   ' 工場
            udtKoutei.OPEC3 = HinNow(0).OPECA           ' 操業条件
            udtKoutei.LENC3 = udtWNINF(i).LEN           ' 受入長さ
            udtKoutei.XTALC3 = HinNow(0).XTALCA         ' 結晶番号
            udtKoutei.SXLIDC3 = ""                      ' SXLID

            udtKoutei.KNKTC3 = left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2)   ' 管理工程(現在工程+2)
            udtKoutei.WKKTC3 = Kihon.NOWPROC            ' 工程
            udtKoutei.WKKBC3 = ""                       ' 作業区分
            udtKoutei.MACOC3 = HinNow(0).NEMACOCA       ' 処理回数
            udtKoutei.MODKBC3 = ""                      ' 赤黒区分
            udtKoutei.SUMKBC3 = ""                      ' 集計区分
            udtKoutei.FRKNKTC3 = ""                     ' (受入)管理工程

            If IsNull(HinOld(0).NEWKNTCA) = True Then   ' (受入）工程
                udtKoutei.FRWKKTC3 = ""
            Else
                udtKoutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            End If

            udtKoutei.FRWKKBC3 = ""                     ' (受入)作業区分

            If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                udtKoutei.FRMACOC3 = "0"
            Else
                udtKoutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If

            Select Case Kihon.NOWPROC
                Case "CC730"
                    intHantei = CInt(BlkNow.GNLC2)
                Case Else
                    intHantei = CInt(BlkNow.GNMC2)
            End Select

            If intHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then
                udtKoutei.TOWKKTC3 = " "                            ' (払出)工程
                udtKoutei.TOMACOC3 = "0"                            ' (払出)処理回数
            Else
                udtKoutei.TOWKKTC3 = HinNow(0).GNWKNTCA             ' (払出)工程
                udtKoutei.TOMACOC3 = HinNow(0).GNMACOCA             ' (払出)処理回
            End If
            udtKoutei.FRLC3 = udtWNINF(i).LEN                       ' 受入長さ
            udtKoutei.FRWC3 = udtWNINF(i).WAT                       ' 受入重量
            udtKoutei.FRMC3 = udtWNINF(i).MAI                       ' 受入枚数
            udtKoutei.FULC3 = 0                                     ' 不良長さ
            udtKoutei.FUWC3 = 0                                     ' 不良重量
            udtKoutei.FUMC3 = 0                                     ' 不良枚数
            udtKoutei.LOSWC3 = ""                                   ' ロス長さ

            udtKoutei.LOSLC3 = ""                                   ' ロス重量
            udtKoutei.LOSMC3 = ""                                   ' ロス枚数
            udtKoutei.TOLC3 = udtWNINF(i).LEN                       ' 払出長さ
            udtKoutei.TOWC3 = udtWNINF(i).WAT                       ' 払出重量
            udtKoutei.TOMC3 = udtWNINF(i).MAI                       ' 払出枚数
            udtKoutei.SUMITLC3 = ""                                 ' SUMIT長さ
            udtKoutei.SUMITWC3 = ""                                 ' SUMIT重量
            udtKoutei.SUMITMC3 = ""                                 ' SUMIT枚数
            udtKoutei.MOTHINC3 = udtWCHKPOS(i).hinban               ' 元品番
            udtKoutei.XTWORKC3 = "42"                               ' 製造工場

            udtKoutei.WFWORKC3 = ""                                 ' ｳｪｰﾊ製造
            udtKoutei.HOLDCC3 = " "                                 ' ホールドコード
            udtKoutei.HOLDBC3 = "0"                                 ' ホールド区分
            udtKoutei.LDFRCC3 = ""                                  ' 格下コード
            udtKoutei.LDFRBC3 = "0"                                 ' 格下区分（ハイキ）
            udtKoutei.TSTAFFC3 = Kihon.STAFFID                      ' 登録社員ID
            udtKoutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付

            udtKoutei.KSTAFFC3 = ""                                 ' 更新社員ID
            udtKoutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            udtKoutei.SUMITBC3 = ""                                 ' SUMIT送信フラグ
            udtKoutei.SNDKC3 = ""                                   ' 送信フラグ
            udtKoutei.MODMACOC3 = ""                                ' 赤黒の処理回数
            udtKoutei.KAKUCC3 = ""                                  ' 確定コード
            udtKoutei.SUMDAYC3 = CalcSumcoTime(udtKoutei.KDAYC3)    ' SUMCO時間
            udtKoutei.PAYCLASSC3 = ""                               ' 転送先工場フラグ

            intRtn = CreateXSDC3(udtKoutei, sErrMsg)                ' 工程実績に在庫減情報登録
            If intRtn = FUNCTION_RETURN_FAILURE Then                ' 工程実績追加エラー
                MsgBox sErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc5 = FUNCTION_RETURN_SUCCESS

proc_exit:
    ' 終了
    Exit Function
proc_err:
    ' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'**************************************************************************************
'*    関数名        : HairetuOpe
'*
'*    処理概要      : 1.HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
'*                    2.HinNum番目の配列を削除する(配列データを前につめる)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*                    udtHinInf        ,I  ,typ_trans_info ,品番振替情報
'*                    intHinNum        ,I  ,Integer        ,品番振替情報
'*                    intHINFLG        ,I  ,Integer        ,品番振替情報
'*
'*    戻り値        : なし
'*
'**************************************************************************************
Public Function HairetuOpe(udtHinInf() As typ_trans_info, intHinNum As Integer, intHINFLG As Integer)
    Dim intRecCnt   As Integer
    Dim i, j        As Integer
    Dim intSflg     As Integer

    intSflg = 0
    intRecCnt = UBound(udtHinInf())

    If (intHINFLG = 1) Then    ' HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
        For i = intHinNum + 1 To intRecCnt ' 既存の配列に空き場所を探す
            If (udtHinInf(i).LEN <= 0) Then    ' i番目に空きがあったのでデータをずらす
                For j = i To intHinNum + 1 Step -1
                    udtHinInf(j).hinban = udtHinInf(j - 1).hinban
                    udtHinInf(j).LEN = udtHinInf(j - 1).LEN
                    udtHinInf(j).WAT = udtHinInf(j - 1).WAT
                    udtHinInf(j).MAI = udtHinInf(j - 1).MAI
                    udtHinInf(j).KCKNT = udtHinInf(j - 1).KCKNT
                Next j
                intSflg = 1
                Exit For
            End If
        Next i
        If (intSflg = 0) Then  ' 空き見つからず
            ReDim Preserve udtHinInf(intRecCnt + 1)
            For i = intRecCnt + 1 To intHinNum + 1 Step -1
                udtHinInf(i).hinban = udtHinInf(i - 1).hinban
                udtHinInf(i).LEN = udtHinInf(i - 1).LEN
                udtHinInf(i).WAT = udtHinInf(i - 1).WAT
                udtHinInf(i).MAI = udtHinInf(i - 1).MAI
                udtHinInf(i).KCKNT = udtHinInf(i - 1).KCKNT
            Next i
        End If

        ' intHinNum+1番目を空きにする
        udtHinInf(intHinNum + 1).hinban = ""
        udtHinInf(intHinNum + 1).LEN = 0
        udtHinInf(intHinNum + 1).MAI = 0
        udtHinInf(intHinNum + 1).WAT = 0
        udtHinInf(intHinNum + 1).KCKNT = 0
    Else    ' HinNum番目の配列を削除する(配列データを前につめる)
        i = intHinNum
        udtHinInf(intHinNum).hinban = ""
        udtHinInf(intHinNum).LEN = 0
        udtHinInf(intHinNum).MAI = 0
        udtHinInf(intHinNum).WAT = 0
        udtHinInf(intHinNum).KCKNT = 0

        For j = intHinNum + 1 To intRecCnt
            If (udtHinInf(j).LEN > 0) Then ' HinNum以降でデータが存在していた時
                udtHinInf(i).hinban = udtHinInf(j).hinban
                udtHinInf(i).LEN = udtHinInf(j).LEN
                udtHinInf(i).MAI = udtHinInf(j).MAI
                udtHinInf(i).WAT = udtHinInf(j).WAT
                udtHinInf(i).KCKNT = udtHinInf(j).KCKNT
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
                i = i + 1
            Else
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
             End If
        Next j
    End If
End Function

'**************************************************************************************
'*    関数名        : HairetuOpe_Mai
'*
'*    処理概要      : 1.HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
'*                    2.HinNum番目の配列を削除する(配列データを前につめる)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*                    udtHinInf        ,I  ,typ_trans_info ,品番振替情報
'*                    intHinNum        ,I  ,Integer        ,品番振替情報
'*                    intHINFLG        ,I  ,Integer        ,品番振替情報
'*
'*    戻り値        : なし
'*
'**************************************************************************************
Public Function HairetuOpe_Mai(udtHinInf() As typ_trans_info, intHinNum As Integer, intHINFLG As Integer)
    Dim intRecCnt   As Integer
    Dim i, j        As Integer
    Dim intSflg     As Integer

    intSflg = 0
    intRecCnt = UBound(udtHinInf())

    If (intHINFLG = 1) Then    ' HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
        For i = intHinNum + 1 To intRecCnt ' 既存の配列に空き場所を探す
            If (udtHinInf(i).MAI <= 0) Then    ' i番目に空きがあったのでデータをずらす
                For j = i To intHinNum + 1 Step -1
                    udtHinInf(j).hinban = udtHinInf(j - 1).hinban
                    udtHinInf(j).LEN = udtHinInf(j - 1).LEN
                    udtHinInf(j).WAT = udtHinInf(j - 1).WAT
                    udtHinInf(j).MAI = udtHinInf(j - 1).MAI
                    udtHinInf(j).KCKNT = udtHinInf(j - 1).KCKNT
                Next j
                intSflg = 1
                Exit For
            End If
        Next i
        If (intSflg = 0) Then  ' 空き見つからず
            ReDim Preserve udtHinInf(intRecCnt + 1)
            For i = intRecCnt + 1 To intHinNum + 1 Step -1
                udtHinInf(i).hinban = udtHinInf(i - 1).hinban
                udtHinInf(i).LEN = udtHinInf(i - 1).LEN
                udtHinInf(i).WAT = udtHinInf(i - 1).WAT
                udtHinInf(i).MAI = udtHinInf(i - 1).MAI
                udtHinInf(i).KCKNT = udtHinInf(i - 1).KCKNT
            Next i
        End If

        ' intHinNum+1番目を空きにする
        udtHinInf(intHinNum + 1).hinban = ""
        udtHinInf(intHinNum + 1).LEN = 0
        udtHinInf(intHinNum + 1).MAI = 0
        udtHinInf(intHinNum + 1).WAT = 0
        udtHinInf(intHinNum + 1).KCKNT = 0
    Else    ' HinNum番目の配列を削除する(配列データを前につめる)
        i = intHinNum
        udtHinInf(intHinNum).hinban = ""
        udtHinInf(intHinNum).LEN = 0
        udtHinInf(intHinNum).MAI = 0
        udtHinInf(intHinNum).WAT = 0
        udtHinInf(intHinNum).KCKNT = 0
        For j = intHinNum + 1 To intRecCnt
            If (udtHinInf(j).MAI > 0) Then ' HinNum以降でデータが存在していた時
                udtHinInf(i).hinban = udtHinInf(j).hinban
                udtHinInf(i).LEN = udtHinInf(j).LEN
                udtHinInf(i).MAI = udtHinInf(j).MAI
                udtHinInf(i).WAT = udtHinInf(j).WAT
                udtHinInf(i).KCKNT = udtHinInf(j).KCKNT
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
                i = i + 1
            Else
                udtHinInf(j).hinban = ""
                udtHinInf(j).LEN = 0
                udtHinInf(j).MAI = 0
                udtHinInf(j).WAT = 0
                udtHinInf(j).KCKNT = 0
             End If
        Next j
    End If
End Function
