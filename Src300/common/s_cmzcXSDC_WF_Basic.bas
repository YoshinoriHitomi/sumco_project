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
    HSXTMMAX As Double                ' 品ＳＸ転位密度上限             項目追加，修正対応 2003.05.20 yakimura
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

'add 2003/03/29 hitec)sada --------------
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
'add 2003/03/29 hitec)sada ---------
Public giInpos  As Integer  ' add 2003/04/09 hitec)matsumoto
Public strSxlData   As String   'Add. 03/05/01 後藤

'概要      :画面からの基本処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Public Function KihonProc() As FUNCTION_RETURN

'   内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim wErrMsg         As String
        
    'エラーハンドラの設定
    On Error GoTo proc_err

    KihonProc = FUNCTION_RETURN_FAILURE
    '########################################### 2003/05/23 okazaki
    'XSDCAProc、XSDC2Procの順番入れ替え
'    ≪分割結晶（品番）登録≫
    iRtn = XSDCAProc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCAProc()：XSDCA登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ≪分割結晶（ブロック）登録≫
    iRtn = XSDC2Proc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDC2Proc()：XSDC2登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    '########################################### 2003/05/23 okazaki
'    ≪不良内訳登録≫
    '不良有無がある時
    If Kihon.FURYOUMU = "Y" Then
                                                ' 登録日付
        Furyou.TDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                ' 更新日付
        Furyou.KDAYC4 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        'add start 2003/06/01 hitec)matsumoto 不良長さ・重量・枚数を再取得-------------
'''        Furyou.PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2)
'''        Furyou.PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)
'''        Furyou.PUCUTMC4 = CLng(BlkNow.GNMC2) - CLng(BlkOld.GNMC2)
        Furyou.PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2)
        Furyou.PUCUTWC4 = CLng(BlkOld.GNWC2) - CLng(BlkNow.GNWC2)
        Furyou.PUCUTMC4 = CLng(BlkOld.GNMC2) - CLng(BlkNow.GNMC2)
        'add end 2003/06/01 hitec)matsumoto -------------
        '不良内訳追加
        iRtn = CreateXSDC4(Furyou, wErrMsg)
        '不良内訳追加エラー
        If iRtn = FUNCTION_RETURN_FAILURE Then
            MsgBox wErrMsg
            Debug.Print "CreateXSDC4()：XSDC4登録エラー"
            KihonProc = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    Debug.Print HinNow(0).SXLIDCA
'    ≪工程実績登録≫
    iRtn = XSDC3Proc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        Debug.Print "XSDC3Proc()：XSDC3登録エラー"
        KihonProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    Select Case Kihon.NOWPROC
        Case "CW740", "CW760"
        '    ≪在庫減情報登録≫
            iRtn = XSDC3Proc2()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3在庫減情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        '    ≪振替情報登録≫
            iRtn = XSDC3Proc3()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3振替情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Case "CC730"
        '    ≪在庫減情報登録≫
            iRtn = XSDC3Proc4()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3在庫減情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        '    ≪振替情報登録≫
            iRtn = XSDC3Proc5()
            If iRtn = FUNCTION_RETURN_FAILURE Then
                Debug.Print "XSDC3Proc()：XSDC3振替情報登録エラー"
                KihonProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
    End Select
    Debug.Print HinNow(0).SXLIDCA
'    ≪分割結晶（ＳＸＬ）登録≫
    iRtn = XSDCBProc()
    If iRtn = FUNCTION_RETURN_FAILURE Then
        KihonProc = FUNCTION_RETURN_FAILURE
        Debug.Print "XSDCBProc()：XSDCB登録エラー"
        GoTo proc_exit
    End If
    Debug.Print HinNow(0).SXLIDCA
    KihonProc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    KihonProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'概要      :分割結晶（ブロック）登録処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function XSDC2Proc()

'   内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String
    Dim intSyoriKaisu   As Integer          '現在処理回数
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    XSDC2Proc = FUNCTION_RETURN_FAILURE
    
    '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''    If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
'''    If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
    Select Case Kihon.NOWPROC
    Case "CC730"
        iHantei = CInt(BlkNow.GNLC2)
    Case Else
        iHantei = CInt(BlkNow.GNMC2)
    End Select
    If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
        BlkNow.LSTATBC2 = "H"                   '最終状態区分（廃棄）
        BlkNow.LDFRBC2 = "2"                    '格下区分（ハイキ）
        BlkNow.LIVKC2 = "1"                     '生死区分（死ロット）
    '2002/12/13 ooba 完了区分フラグ変更
    '   BlkNow.KANKC2 = "2"                     '完了区分（終了）
        '2002/12/02 ooba
        BlkNow.GNWKNTC2 = " "                   '現在工程
        BlkNow.GNMACOC2 = "0"                   '現在処理回数
    Else
        '2002/11/24 tuku 処理回数取得ロジック変更
        intSyoriKaisu = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.GNWKNTC2)
        If BlkNow.GNWKNTC2 = BlkNow.NEWKNTC2 Then
            intSyoriKaisu = intSyoriKaisu + 1
        End If
        BlkNow.GNMACOC2 = intSyoriKaisu                             '現在処理回数
'        BlkNow.NEMACOC2 = GetNEMACOC2(BlkNow.CRYNUMC2)               '最終通過処理回数
        
        '2002/12/02 tuku
        BlkNow.NEMACOC2 = GetGNMACOC(BlkNow.CRYNUMC2, BlkNow.NEWKNTC2)   '最終通過処理回数
    End If
    
                                                ' 更新日付
    BlkNow.KDAYC2 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
    
    sqlWhere = "WHERE CRYNUMC2 = '" & BlkNow.CRYNUMC2 & "' "

    iRtn = UpdateXSDC2(BlkNow, sqlWhere)
    '分割結晶（ブロック）更新エラー
    If iRtn = FUNCTION_RETURN_FAILURE Then
        MsgBox "XSDCB UPDATET ERROR"
        Exit Function
    End If
    
    XSDC2Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC2Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'概要      :分割結晶（品番）登録処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function XSDCAProc()

'   内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String
    Dim LivFlg          As Integer          '存在フラグ
    Dim wHinban()         As typ_XSDCA        '分割結晶（品番）ワーク領域
    Dim wHinban_UP()      As typ_XSDCA_Update '分割結晶（品番）ワーク領域
    Dim intDataCnt      As Integer          '該当データ件数
    Dim lngSumGNWCA     As Long
    Dim lngSumGNMCA     As Long
    Dim bChgFlg         As Boolean
    Dim intSyoriKaisu As Integer    '現在処理回数
    Dim iHantei         As Integer  'add 2003/05/27 hitec)matsumoto
    Dim lGetLength      As Long     'add 2003/06/02 hitec)matsumoto TBCME040より、ブロック長さを取得する
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    XSDCAProc = FUNCTION_RETURN_FAILURE
    lngSumGNWCA = 0
    lngSumGNMCA = 0
    bChgFlg = False
    '品番の重量・枚数計算
    If Kihon.CNTHINNOW = Kihon.CNTHINOLD Then   '前工程と現在工程の件数が同じで、各長さも同じ場合は計算処理をしない
        For i = 0 To Kihon.CNTHINNOW - 1
            If CLng(HinNow(i).GNLCA) <> CLng(HinOld(i).GNLCA) Then
                bChgFlg = True
            End If
        Next
    Else            '前工程と現在工程の件数が違う場合は、計算処理を行う
        bChgFlg = True
    End If
    '重量・枚数計算処理

    If bChgFlg = True Then
' VVVVV 2003/04/30 ALT BY HITEC)会田：CW740,CW760用追加
        If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
            For i = 0 To Kihon.CNTHINNOW - 1
                With HinNow(i)  'upd 203/05/19 hitec)matsumoto BLKOLD基準に変更
                    If Kihon.CNTHINNOW = 1 Then
                        HinNow(i).GNWCA = BlkOld.GNWC2
                        HinNow(i).GNLCA = BlkOld.GNLC2
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    ElseIf i = Kihon.CNTHINNOW - 1 Then
                        HinNow(i).GNWCA = CLng(BlkOld.GNWC2) - lngSumGNWCA
                        HinNow(i).GNLCA = CLng(BlkOld.GNLC2) - lngSumGNLCA
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    Else
                        HinNow(i).GNWCA = Round(CLng(BlkOld.GNWC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        HinNow(i).GNLCA = Round(CLng(BlkOld.GNLC2) * (CLng(HinNow(i).GNMCA) / CLng(BlkOld.GNMC2)))
                        lngSumGNWCA = lngSumGNWCA + CLng(HinNow(i).GNWCA)
                        lngSumGNLCA = lngSumGNLCA + CLng(HinNow(i).GNLCA)
                        .SUMITLCA = .GNLCA   '' 03/05/18 matsumoto
                        .SUMITMCA = .GNMCA
                        .SUMITWCA = .GNWCA
                    End If
                End With
            Next
        Else
            If BlkOld.GNLC2 = BlkNow.GNLC2 Then 'upd 2003/06/01 hitec)matsumoto 長さが同じ場合はBLKOLD基準。長さが異なる場合はBLKNOW基準にする
                BlkNow.GNWC2 = BlkOld.GNWC2
                BlkNow.GNMC2 = BlkOld.GNMC2
            Else
                BlkNow.GNWC2 = Round(CLng(BlkOld.GNWC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
                BlkNow.GNMC2 = Round(CLng(BlkOld.GNMC2) * (CLng(BlkNow.GNLC2) / CLng(BlkOld.GNLC2)))
            End If
            For i = 0 To Kihon.CNTHINNOW - 1    'upd 203/05/19 hitec)matsumoto BLKOLD基準に変更
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
    '########################################### 2003/05/23 okazaki
    'XSDC2の重量、長さをXSDCAの合計にするためにBlkNow再計算
    BlkNow.GNLC2 = 0
    BlkNow.GNWC2 = 0
    For i = 0 To Kihon.CNTHINNOW - 1
        BlkNow.GNLC2 = CLng(BlkNow.GNLC2) + CLng(HinNow(i).GNLCA)   '2003/05/24 clng追加
        BlkNow.GNWC2 = CLng(BlkNow.GNWC2) + CLng(HinNow(i).GNWCA)
    Next i
    '########################################### 2003/05/23 end

    '前工程の分割結晶（品番）と良品情報の品番・位置を比較する
    For i = 0 To Kihon.CNTHINOLD - 1
    
        LivFlg = 0
        For j = 0 To Kihon.CNTHINNOW - 1
            If (HinOld(i).HINBCA = HinNow(j).HINBCA) And (HinOld(i).INPOSCA = HinNow(j).INPOSCA) Then
                LivFlg = 1
            End If
        Next j
        '前工程の分割結晶（品番）にあって良品情報にないものは死ロットとする
        If LivFlg = 0 Then
            sqlWhere = "WHERE CRYNUMCA = '" & HinOld(i).CRYNUMCA & "' "
            sqlWhere = sqlWhere & "AND HINBCA = '" & HinOld(i).HINBCA & "' "
            sqlWhere = sqlWhere & "AND INPOSCA = '" & HinOld(i).INPOSCA & "' "
            ReDim wHinban(0) As typ_XSDCA
            
            'データの件数を取得
            iRtn = SelCntXSDCA(sqlWhere, intDataCnt)
            If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
                MsgBox "XSDCA SELECT ERROR"
                Exit Function
            Else                                    '正常
                If intDataCnt = 0 Then
                    '前工程の情報は必ずあるはずなので、0件はエラー
''''                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                ElseIf intDataCnt > 0 Then
                    '前工程
                    iRtn = DBDRV_GetXSDCA(wHinban(), sqlWhere)
                    '存在しない時エラー
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA SELECT ERROR"
                        Exit Function
                    End If
                    ReDim wHinban_UP(0) As typ_XSDCA_Update
                    
                    wHinban_UP(0).CRYNUMCA = HinOld(i).CRYNUMCA
                    wHinban_UP(0).INPOSCA = HinOld(i).INPOSCA
                    wHinban_UP(0).HINBCA = HinOld(i).HINBCA
                    
                    '生死区分に死ロットをセット
                    wHinban_UP(0).LIVKCA = "1"              ' 生死区分
                    wHinban_UP(0).LSTATBCA = "H"            ' 最終状態区分
                    wHinban_UP(0).LDFRBCA = "2"             ' 格下区分
                    ' 2002/12/13 ooba 完了区分フラグ変更
                    wHinban_UP(0).KANKCA = "0"              ' 完了区分
'                    wHinban_UP(0).KANKCA = "2"              ' 完了区分

                    
                    wHinban_UP(0).SUMITBCA = "0"
''''                    wHinban_UP(0).SUMITLCA = "0"    'del 2003/05/18 hitec)matsumoto
''''                    wHinban_UP(0).SUMITMCA = "0"
''''                    wHinban_UP(0).SUMITWCA = "0"
                                                        ' 更新日付
                    wHinban_UP(0).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                    
                    '分割結晶（品番）を更新
                    iRtn = UpdateXSDCA(wHinban_UP(0), sqlWhere)
                    '分割結晶（品番）更新エラー
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox "XSDCA UPDATE ERROR"
                        Exit Function
                    End If
                End If
            End If
        
        End If
    
    Next i
    
    '分割結晶（品番）分繰り返し
    For i = 0 To Kihon.CNTHINNOW - 1
        '結晶番号、品番、位置で検索
        sqlWhere = "WHERE CRYNUMCA = '" & HinNow(i).CRYNUMCA & "' "
        sqlWhere = sqlWhere & "AND HINBCA = '" & HinNow(i).HINBCA & "' "
        sqlWhere = sqlWhere & "AND INPOSCA = '" & HinNow(i).INPOSCA & "' "

        'データの件数を取得
        iRtn = SelCntXSDCA(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        Else                                    '正常
            'データがある場合は更新処理
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCA(wHinban, sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If

                '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
''''                    iHantei = CInt(BlkNow.GNMC2)
                    iHantei = CInt(HinNow(i).GNMCA) 'upd 2003/06/05 hitec)matsumoto 0枚を廃棄にする処理を、ブロック単位ではなく品番単位に変更
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    HinNow(i).LIVKCA = "1"              ' 生死区分
                    HinNow(i).LSTATBCA = "H"            ' 最終状態区分
                    HinNow(i).LDFRBCA = "2"             ' 格下区分
'                    HinNow(i).KANKCA = "2"              ' 完了区分
                    '2002/12/02 ooba
                    HinNow(i).GNWKNTCA = " "            '現在工程
                    HinNow(i).GNMACOCA = "0"            '現在処理回数
                Else
                    HinNow(i).LIVKCA = "0"               ' 生死区分（生ロット）
                    HinNow(i).LSTATBCA = "T"            ' 最終状態区分（通常）
                    HinNow(i).LDFRBCA = "0"             ' 格下区分（通常）
 '                    HinNow(i).KANKCA = "0"              ' 完了区分
                    '2002/11/24 tuku 処理回数取得ロジック変更
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu    '現在処理回数
'                    HinNow(i).NEMACOCA = GetNEMACOC2(HinNow(i).CRYNUMCA)               '最終通過処理回数

                    '2002/12/02 ooba
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '最終通過処理回数
                End If
                ' 2002/12/13 ooba 完了区分フラグ変更
                HinNow(i).KANKCA = "0"              ' 完了区分
                
                HinNow(i).SUMITBCA = "0"
''''                HinNow(i).SUMITLCA = "0"    'del 2003/05/18 hitec)matsumoto
''''                HinNow(i).SUMITMCA = "0"
''''                HinNow(i).SUMITWCA = "0"

                                                    ' 更新日付
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
''''                'シングル確定時、完了区分='2'・生死区分='1'にする
''''                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    HinNow(i).LIVKCA = "1"
''''                    HinNow(i).KANKCA = "2"
''''                End If
                
                '良品情報で置き換え
                iRtn = UpdateXSDCA(HinNow(i), sqlWhere)
                '分割結晶（品番）更新エラー
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCA UPDATE ERROR"
                    Exit Function
                End If
            '存在しない時追加
            ElseIf intDataCnt = 0 Then
                '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''                If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
''''                    iHantei = CInt(BlkNow.GNMC2)
                    iHantei = CInt(HinNow(i).GNMCA) 'upd 2003/06/05 hitec)matsumoto 0枚を廃棄にする処理を、ブロック単位ではなく品番単位に変更
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    HinNow(i).LIVKCA = "1"              ' 生死区分
                    HinNow(i).LSTATBCA = "H"            ' 最終状態区分
                    HinNow(i).LDFRBCA = "2"             ' 格下区分
'                    HinNow(i).KANKCA = "2"              ' 完了区分
                    '2002/12/02 ooba
                    HinNow(i).GNWKNTCA = " "            '現在工程
                    HinNow(i).GNMACOCA = "0"            '現在処理回数
                Else
                    HinNow(i).LIVKCA = "0"               ' 生死区分（生ロット）
                    HinNow(i).LSTATBCA = "T"            ' 最終状態区分（通常）
                    HinNow(i).LDFRBCA = "0"             ' 格下区分（通常）
'                    HinNow(i).KANKCA = "0"              ' 完了区分
                    '2002/11/24 tuku 処理回数取得ロジック変更
                    intSyoriKaisu = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).GNWKNTCA)
                    If HinNow(i).GNWKNTCA = HinNow(i).NEWKNTCA Then
                          intSyoriKaisu = intSyoriKaisu + 1
                    End If
                    HinNow(i).GNMACOCA = intSyoriKaisu                      '現在処理回数
'                    HinNow(i).NEMACOCA = GetNEMACOC2(HinNow(i).CRYNUMCA)    '最終通過処理回数

                    '2002/12/02 ooba
                    HinNow(i).NEMACOCA = GetGNMACOC(HinNow(i).CRYNUMCA, HinNow(i).NEWKNTCA)   '最終通過処理回数
                End If
                ' 2002/12/13 ooba 完了区分フラグ変更
                HinNow(i).KANKCA = "0"              ' 完了区分
                                                    ' 登録日付
                HinNow(i).TDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                    ' 更新日付
                HinNow(i).KDAYCA = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
''''                'シングル確定時、完了区分='2'・生死区分='1'にする
''''                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    HinNow(i).LIVKCA = "1"
''''                    HinNow(i).KANKCA = "2"
''''                End If
                HinNow(i).SUMITBCA = "0"
'''                HinNow(i).SUMITLCA = "0"     'del 2003/05/18 hitec)matsumoto
'''                HinNow(i).SUMITMCA = "0"
'''                HinNow(i).SUMITWCA = "0"
                
                iRtn = CreateXSDCA(HinNow(i), wErrMsg)
                '分割結晶（品番）更新エラー
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox wErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i
    
    XSDCAProc = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDCAProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function


'概要      :工程実績登録処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function XSDC3Proc()

'   内部変数
    Dim i               As Integer
    Dim j               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wFULC3          As Long             '分割結晶（品番）の不良長さ
    Dim wFUWC3          As Long             '分割結晶（品番）の不良重量
    Dim wFUMC3          As Long             '分割結晶（品番）の不良枚数
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '工程実績
    Dim rsKCNTC         As OraDynaset       'レコードセット
    Dim intNextCnt      As Integer
    Dim intOldCnt       As Integer
    'add start 2003/03/28 hitec)matsumoto ------------------
    Dim bNewRec         As Boolean          '前工程の無いレコードがあった場合
    Dim sSUMITLC3       As String           'SUMIT長さ
    Dim sSUMITWC3       As String           'SUMIT重量
    Dim sSUMITMC3       As String           'SUMIT枚数
    Dim dSumcoTime      As Date             'SUMCO時間
    Dim vChoseiTime     As Variant          '調整時間
    'add end   2003/03/28 hitec)matsumoto ------------------
    'add 03/05/17 後藤 -------------------------------------
    Dim iLoopCnt        As Integer
    '--------------------------------------end 03/05/17 ---
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    bNewRec = False 'add 2003/03/27 hitec)matsumoto フラグ初期化
    XSDC3Proc = FUNCTION_RETURN_FAILURE
    
''''    'SUMCO時間作成の為、調整時間取得  add   2003/04/01 hitec)matsumoto ----
''''    sql = "SELECT KCODE01A9"
''''    sql = sql & " FROM koda9 "
''''    sql = sql & " WHERE SYSCA9 = 'X'"
''''    sql = sql & "   AND SHUCA9 = '80'"
''''    sql = sql & "   AND CODEA9 = '1'"
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''    '存在しない時、エラー
''''    If rs Is Nothing Then
''''        MsgBox "koda9 KCODE01A9 SELECT ERROR"
''''        Exit Function
''''    End If
''''    If Not rs.EOF Then
''''        If IsNull(rs.Fields("KCODE01A9")) = True Then
''''            MsgBox "koda9 KCODE01A9 SELECT ERROR"
''''            Exit Function
''''        Else
''''            vChoseiTime = CDate(rs.Fields("KCODE01A9"))
''''        End If
''''    End If
''''    rs.Close
''''    '----------------------------------------------------------------------
    
    '工程実績からブロックＩＤ、品番が一致する工程連番の最大を取得
    sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
    sql = sql & " FROM XSDC3 "
    sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
'        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "

    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、エラー
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
    
    'add 2003/03/27 hitec)matsumoto 前工程実績の無いレコードがあるかチェックし、あったらフラグをたてる-----------------------
    For i = 0 To Kihon.CNTHINNOW - 1
        '工程実績から前工程のデータを読み込む
        sql = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3,WKKTC3,MACOC3 "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND KCNTC3 = " & intOldCnt & ""
        sql = sql & " AND KCNTC3 = " & intNextCnt - 1 & ""  'upd 2003/05/31 hitec)matsumoto intOldCntは使えないので、intNextCnt - 1を変わりに使用
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "


        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、エラー
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            bNewRec = True  '前工程がない場合はフラグをたてる
        End If
        rs.Close
    Next
    'add end 2003/03/27 hitec)matsumoto ---------------------------
    If Kihon.NOWPROC = "CC730" Then
        bNewRec = True
    End If
    For i = 0 To Kihon.CNTHINNOW - 1
        '不良内訳からブロックＩＤ、品番、開始位置が一致する不良長さを取得する
        intOldCnt = 0
        wFULC3 = 0
        wFUWC3 = 0
        wFUMC3 = 0
'            Koutei.FRWKKTC3 = " "
'            Koutei.FRMACOC3 = 0
''''        For j = 0 To Kihon.CNTHINNOW - 1
''''            If HinNow(i).CRYNUMCA = Furyou.XTALC4 And HinNow(i).HINBCA = Furyou.HINBC4 And HinNow(i).INPOSCA = Furyou.INPOSC4 Then
''''                wFULC3 = Furyou.PUCUTLC4
''''                wFUWC3 = Furyou.PUCUTWC4
''''                wFUMC3 = Furyou.PUCUTMC4
''''            End If
''''        Next j
''''
''''        If Furyou.HINBC4 = "Z" Then '廃棄
''''            wFULC3 = Furyou.PUCUTLC4
''''            wFUWC3 = Furyou.PUCUTWC4
''''            wFUMC3 = Furyou.PUCUTMC4
''''        End If
    
''''' 02/09/20 Add By 会田@HITEC  sta
''''        '工程実績からブロックＩＤ、品番が一致する工程連番の最大を取得
''''        sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
''''        sql = sql & " FROM XSDC3 "
''''        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
'''''        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、エラー
''''        If rs Is Nothing Then
''''            MsgBox "XSDC3 MAX KCNT SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        If rs.EOF = False Then
''''            If IsNull(rs.Fields("wKCNTC3")) = True Then
''''                intNextCnt = 1
''''                intOldCnt = 0
''''            Else
''''                intNextCnt = CInt(rs.Fields("wKCNTC3")) + 1
''''                intOldCnt = CInt(rs.Fields("wKCNTC3"))
''''            End If
''''        Else
''''            intNextCnt = 1
''''            intOldCnt = 0
''''        End If
''''        rs.Close

        '工程実績からブロックＩＤ、品番が一致する工程連番の最大を取得
        sql = "SELECT MAX(KCNTC3) as wKCNTC3 "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "
    
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、エラー
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
    
        '工程実績から前工程ののデータを読み込む
        sql = "SELECT STATIMEC3, STOTIMEC3 , TOLC3,TOWC3,TOMC3, "
        'add start 2003/03/28 hitec)matsumoto -------
        sql = sql & " SUMITLC3, SUMITWC3, SUMITMC3"
        'add end   2003/03/28 hitec)matsumoto -------
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = '" & HinNow(i).INPOSCA & "' "
        sql = sql & " AND KCNTC3 = " & intOldCnt & ""
''''        sql = sql & " AND ((SUMKBC3 ='0') OR (SUMKBC3 = ' ') OR (SUMKBC3 IS NULL)) "


        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、エラー
        If rs Is Nothing Then
            MsgBox "XSDCA SELECT ERROR"
            Exit Function
        End If
        If rs.RecordCount = 0 Then
            'upd end 2003/03/27 hitec)matsumoto 前工程無し時は、払出し長さを受入長さに入れる------------
''''            Koutei.FRLC3 = ""                       ' 受入長さクリア
''''            Koutei.FRWC3 = ""                       ' 受入重量クリア
''''            Koutei.FRMC3 = ""                       ' 受入枚数クリア
            'upd end 2003/03/27 hitec)matsumoto -------------
            wFULC3 = 0
            wFUWC3 = 0
            wFUMC3 = 0
            'add start 2003/03/28 hitec)matsumoto --------
            sSUMITLC3 = "0" 'SUMIT長さ
            sSUMITWC3 = "0" 'SUMIT重量
            sSUMITMC3 = "0" 'SUMIT枚数
            'add end 2003/03/28 hietc)matsumoto -----------
        Else
            If IsNull(rs.Fields("STATIMEC3")) = True Then
                '何も入れない
            Else
                Koutei.STATIMEC3 = rs.Fields("STATIMEC3")
            End If
                                                    ' 処理時間終了
            If IsNull(rs.Fields("STOTIMEC3")) = True Then
                '何も入れない
            Else
                Koutei.STOTIMEC3 = rs.Fields("STOTIMEC3")
            End If
            
            If IsNull(rs.Fields("TOLC3")) = True Then   '不良長さ
                wFULC3 = 0
                Koutei.FRLC3 = "0"
            Else
                Koutei.FRLC3 = CLng(rs.Fields("TOLC3"))
                wFULC3 = CLng(rs.Fields("TOLC3"))
            End If
            If IsNull(rs.Fields("TOWC3")) = True Then   '不良重量
                wFUWC3 = 0
                Koutei.FRWC3 = "0"
            Else
                Koutei.FRWC3 = CLng((rs.Fields("TOWC3")))
                wFUWC3 = CLng((rs.Fields("TOWC3")))
            End If
            If IsNull(rs.Fields("TOMC3")) = True Then   '不良枚数
                wFUMC3 = 0
                Koutei.FRMC3 = "0"
            Else
                Koutei.FRMC3 = CLng((rs.Fields("TOMC3")))
                wFUMC3 = CLng((rs.Fields("TOMC3")))
            End If

            'add start 2003/03/28 hitec)matsumoto --------
            If IsNull(rs.Fields("SUMITLC3")) = True Then   'SUMIT長さ
                sSUMITLC3 = "0"
            Else
                sSUMITLC3 = CLng((rs.Fields("SUMITLC3")))
            End If
            If IsNull(rs.Fields("SUMITWC3")) = True Then   'SUMIT重量
                sSUMITWC3 = "0"
            Else
                sSUMITWC3 = CLng((rs.Fields("SUMITWC3")))
            End If
            If IsNull(rs.Fields("SUMITMC3")) = True Then   'SUMIT枚数
                sSUMITMC3 = "0"
            Else
                sSUMITMC3 = CLng((rs.Fields("SUMITMC3")))
            End If
'            '2002/11/24 tuku 処理回数取得ロジック変更
'            If IsNull(rs.Fields("WKKTC3")) = True Then   '(受入）工程
'                Koutei.FRWKKTC3 = "0"
'            Else
'                Koutei.FRWKKTC3 = CStr((rs.Fields("WKKTC3")))
'            End If
'
'            If IsNull(rs.Fields("MACOC3")) = True Then   '（受入）処理回数
'                Koutei.FRMACOC3 = "0"
'            Else
'                Koutei.FRMACOC3 = CLng((rs.Fields("MACOC3")))
'            End If
        End If
        
        '2002/11/29 ooba 処理回数取得ロジック変更
        If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
            Koutei.FRWKKTC3 = "0"
        Else
            Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
            'add end 2003/03/28 hitec)matsumoto --------
        End If
        
        If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
            Koutei.FRMACOC3 = "0"
        Else
            Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
        End If
        '分割結晶（品番）から工程実績を追加
        Koutei.CRYNUMC3 = HinNow(i).CRYNUMCA    ' ﾌﾞﾛｯｸID･結晶番号
        Koutei.INPOSC3 = HinNow(i).INPOSCA      ' 結晶内開始位置
            
        
        
'''        '工程連番のMAXを取得
'''        sql = ""
'''        sql = "SELECT MAX(KCNTC3) KCNTC3 FROM XSDC3"
'''        sql = sql & "  WHERE CRYNUMC3 = '" & HinNow(i).CRYNUMCA & "'"
'''        sql = sql & "    AND INPOSC3 = " & HinNow(i).INPOSCA
'''        sql = sql & "  group by CRYNUMC3,INPOSC3 "
'''
'''        Set rsKCNTC = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''
'''        If rsKCNTC Is Nothing Then
'''            XSDC3Proc = FUNCTION_RETURN_FAILURE
'''            Exit Function
'''        End If
'''        '工程連番のMAX+1の値を使用
'''        If rsKCNTC.EOF = False Then
'''            If IsNull(rsKCNTC("KCNTC3")) = True Then
'''                lngProcCnt = 0
'''            Else
'''                lngProcCnt = rsKCNTC("KCNTC3") + 1
'''            End If
'''        Else
'''            lngProcCnt = 0
'''        End If
        
'''        rsKCNTC.Close
        Koutei.KCNTC3 = intNextCnt       ' 工程連番
        Koutei.HINBC3 = HinNow(i).HINBCA        ' 品番
        Koutei.REVNUMC3 = HinNow(i).REVNUMCA    ' 製品番号改訂番号
        Koutei.FACTORYC3 = HinNow(i).FACTORYCA  ' 工場
        Koutei.OPEC3 = HinNow(i).OPECA          ' 操業条件
        '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.LENC3 = 0                    ' 長さ
        Else
            Koutei.LENC3 = HinNow(i).GNLCA      ' 長さ
        End If
        Koutei.XTALC3 = HinNow(i).XTALCA        ' 結晶番号
        Koutei.SXLIDC3 = HinNow(i).SXLIDCA      ' SXLID
        Select Case Kihon.NOWPROC   'upd 2003/04/05 hitec)matsumoto  CW740，CW760工程で、管理工程に現在工程＋３を書き込む
            Case "CW740", "CW760", "CC730"
                Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                      CStr(CInt(Right(Kihon.NOWPROC, 1)) + 3) ' 管理工程(現在工程+3)
            Case Else
                Koutei.KNKTC3 = HinNow(i).GNKKNTCA      ' 管理工程
        End Select
        Koutei.WKKTC3 = Kihon.NOWPROC           ' 工程
        Koutei.WKKBC3 = HinNow(i).GNWKKBCA      ' 作業区分
        Koutei.MACOC3 = HinNow(i).NEMACOCA      ' 処理回数
        Koutei.MODKBC3 = ""                     ' 赤黒区分
        Koutei.SUMKBC3 = "0"                    ' 集計区分
        Koutei.FRKNKTC3 = " "                   ' (受入)管理工程
'        Koutei.FRWKKTC3 = HinOld(0).NEWKNTCA    ' (受入)工程
        Koutei.FRWKKBC3 = " "                   ' (受入)作業区分
'        Koutei.FRWKKTC3 = HinOld(0).NEWKNTCA    ' (受入)工程
        Koutei.TOWNKTC3 = " "                   ' (払出)管理工程
        '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.TOWKKTC3 = " "               ' (払出)工程
            '2002/12/02 ooba
            Koutei.TOMACOC3 = "0"               '(払出)処理回数
        Else
            Koutei.TOWKKTC3 = HinNow(i).GNWKNTCA     ' (払出)工程
        End If
        Koutei.TOMACOC3 = HinNow(i).GNMACOCA    ' (払出)処理回
        
''''        Koutei.FRLC3 = ""                       ' 受入長さクリア
''''        Koutei.FRWC3 = ""                       ' 受入重量クリア
''''        Koutei.FRMC3 = ""                       ' 受入枚数クリア
''''        For j = 0 To Kihon.CNTHINOLD - 1
''''            If (HinNow(i).CRYNUMCA = HinOld(j).CRYNUMCA) And (HinNow(i).INPOSCA = HinOld(j).INPOSCA) And (HinNow(i).KCKNTCA = HinOld(j).KCKNTCA) Then
''''                Koutei.FRLC3 = HinOld(i).GNLCA  ' 受入長さ
''''                Koutei.FRWC3 = HinOld(i).GNWCA  ' 受入重量
''''                Koutei.FRMC3 = HinOld(i).GNMCA  ' 受入枚数
''''                Exit For
''''            End If
''''        Next j

        Koutei.LOSWC3 = ""                      ' ロス長さ
        Koutei.LOSLC3 = ""                      ' ロス重量
        Koutei.LOSMC3 = ""                      ' ロス枚数
        If bNewRec = True Then  'add 2003/03/27 hitec)matsumoto 前工程が無いデータが存在している場合は、払出し数量を受入数量に入れる
            Koutei.FRLC3 = HinNow(i).GNLCA      ' 受入長さ<=払出長さ
            Koutei.FRWC3 = HinNow(i).GNWCA      ' 受入重量<=払出重量
            Koutei.FRMC3 = HinNow(i).GNMCA      ' 受入枚数<=払出枚数
            Koutei.TOLC3 = HinNow(i).GNLCA      ' 払出長さ
            Koutei.TOWC3 = HinNow(i).GNWCA      ' 払出重量（関数）
            Koutei.TOMC3 = HinNow(i).GNMCA      ' 払出枚数（関数）
            Koutei.FULC3 = 0                    ' 不良長さ
            Koutei.FUWC3 = 0                    ' 不良重量
            Koutei.FUMC3 = 0                    ' 不良枚数
        Else
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOLC3 = 0                    ' 払出長さ
                Koutei.TOWC3 = 0                    ' 払出重量（関数）
                Koutei.TOMC3 = 0                    ' 払出枚数（関数）
            Else
                Koutei.TOLC3 = HinNow(i).GNLCA      ' 払出長さ
                Koutei.TOWC3 = HinNow(i).GNWCA      ' 払出重量（関数）
                Koutei.TOMC3 = HinNow(i).GNMCA      ' 払出枚数（関数）
            End If
            Koutei.FULC3 = wFULC3 - CLng(Koutei.TOLC3)                  ' 不良長さ
            Koutei.FUWC3 = wFUWC3 - CLng(Koutei.TOWC3)                  ' 不良重量
            Koutei.FUMC3 = wFUMC3 - CLng(Koutei.TOMC3)                  ' 不良枚数
        End If
        If Koutei.TOLC3 = "" Then
            Koutei.TOLC3 = "0"
        End If
        If Koutei.TOWC3 = "" Then
            Koutei.TOWC3 = "0"
        End If
        If Koutei.TOMC3 = "" Then
            Koutei.TOMC3 = "0"
        End If
        'upd start 2003/03/28 hitec)matsumoto SUMIT長さに工程別に値をセットする--------------------
        Koutei.SUMITLC3 = 0                     ' SUMIT長さ
        Koutei.SUMITWC3 = 0                     ' SUMIT重量
        Koutei.SUMITMC3 = 0                     ' SUMIT枚数
'----------------------------------------------------- 03/05/13 後藤 ----------------------------
'''        Select Case Kihon.NOWPROC
'''            Case "CW740", "CW760"
'''                Koutei.SUMITLC3 = HinOld(i).SUMITLCA    ' SUMIT長さ=払出長さ
'''                Koutei.SUMITWC3 = HinOld(i).SUMITWCA    ' SUMIT重量=払出重量
'''                Koutei.SUMITMC3 = HinOld(i).SUMITMCA    ' SUMIT枚数=払出枚数
'''            Case "CW750", "CW800"
'''                Koutei.SUMITLC3 = HinNow(i).SUMITLCA    ' SUMIT長さ=払出長さ
'''                Koutei.SUMITWC3 = HinNow(i).SUMITWCA    ' SUMIT重量=払出重量
'''                Koutei.SUMITMC3 = HinNow(i).SUMITMCA    ' SUMIT枚数=払出枚数
'''        End Select
        For iLoopCnt = 0 To Kihon.CNTHINOLD - 1     '' 03/05/17 後藤
            If (Koutei.CRYNUMC3 = HinOld(iLoopCnt).CRYNUMCA) _
                And (Koutei.INPOSC3 = HinOld(iLoopCnt).INPOSCA) Then
                    Koutei.SUMITLC3 = HinOld(iLoopCnt).SUMITLCA     ' SUMIT長さ=前工程SUMIT長さ
                    Koutei.SUMITWC3 = HinOld(iLoopCnt).SUMITWCA     ' SUMIT重量=前工程SUMIT重量
                    Koutei.SUMITMC3 = HinOld(iLoopCnt).SUMITMCA     ' SUMIT枚数=前工程SUMIT枚数
                    Exit For
            End If
        Next
'----------------------------------------------------------------------------- END 03/05/13 ------
        'upd end 2003/03/28 hitec)matsumoto ---------------
        Koutei.MOTHINC3 = " "                   ' 振替品番(元)
        Koutei.XTWORKC3 = "42"                  ' 製造工場
        Koutei.WFWORKC3 = " "                   ' ｳｪｰﾊ製造
                                                ' 処理時間開始
''''        Koutei.ETIMEC3 = ""                     ' 実績時間は入れない
        Koutei.HOLDCC3 = " "                    ' ホールドコード
        Koutei.HOLDBC3 = "0"                    ' ホールド区分
        Koutei.LDFRCC3 = " "                    ' 格下コード
        '分割結晶（ブロック）の良品長さ<=0 or 全数スクラップの時
''''        If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''        If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
        Select Case Kihon.NOWPROC
        Case "CC730"
            iHantei = CInt(BlkNow.GNLC2)
        Case Else
            iHantei = CInt(BlkNow.GNMC2)
        End Select
        If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
            Koutei.LDFRBC3 = "2"                ' 格下区分（ハイキ）
        Else
            Koutei.LDFRBC3 = "0"                ' 格下区分
        End If
        Koutei.TSTAFFC3 = Kihon.STAFFID         ' 登録社員ID
                                                ' 登録日付
        Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        Koutei.KSTAFFC3 = Kihon.STAFFID         ' 更新社員ID
                                                ' 更新日付
        Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS")
        'SUMCO時間=工程実績.更新日付-KODA9.調整時間 add 2003/04/01 hitec)matsumoto --
''''        dSumcoTime = CDate(Koutei.KDAYC3) - CDate(vChoseiTime)
''''        Koutei.SUMDAYC3 = Format(dSumcoTime, "YYYY/MM/DD")
        Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3)
        '-------------------------------------------------------------------------------
        Koutei.SUMITBC3 = "0"                   ' SUMIT送信フラグ
        Koutei.SNDKC3 = "0"                     ' 送信フラグ
'        Koutei.SNDDAYC3 = ""                   ' 送信日付
        Koutei.MODMACOC3 = "00"                 ' 赤黒の処理回数
        Koutei.KAKUCC3 = " "                    ' 確定コード
        'upd start 2003/03/25 hitec)matsumoto 画面で使用しているデータのみ更新を行う。 -------------
        Select Case Kihon.NOWPROC
            Case "CW750"    '総合判定
                If Koutei.SXLIDC3 = Trim(f_cmbc039_2.txtSxlID.Text) Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '工程実績追加エラー
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case "CW760"    '再抜試
                If (SIngotP <= Koutei.INPOSC3) And (Koutei.INPOSC3 < EIngotP) Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '工程実績追加エラー
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case "CW800"    'シングル確定   03/05/01 Add.後藤
                If Koutei.SXLIDC3 = strSxlData Then
                    iRtn = CreateXSDC3(Koutei, wErrMsg)
                    '工程実績追加エラー
                    If iRtn = FUNCTION_RETURN_FAILURE Then
                        MsgBox wErrMsg
                        Exit Function
                    End If
                End If
            Case Else
                iRtn = CreateXSDC3(Koutei, wErrMsg)
                '工程実績追加エラー
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox wErrMsg
                    Exit Function
                End If
        End Select
        'upd end  2003/03/25 hitec)matsumoto ------------------------------------------
    Next i

    XSDC3Proc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function


'''''概要      :分割結晶（ＳＸＬ）登録処理を行う
'''''ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'''''      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'''''説明      :
''''
''''Public Function XSDCBProc()
''''
'''''   内部変数
''''    Dim i               As Integer
''''    Dim iRtn            As Integer          '復帰情報
''''    Dim sql             As String           'ＳＱＬ
''''    Dim rs              As OraDynaset       'レコードセット
''''    Dim sqlWhere        As String           'WHERE句
''''    Dim wGNLCA          As Long             '分割結晶（品番）の合計長さ
''''    Dim wGNMCA          As Long             '分割結晶（品番）の合計枚数
''''    Dim wLENCB          As Long             '合計長さ
''''    Dim wMAICB          As Long             '合計枚数
''''    Dim wPUCUTMC4       As Long             '不良内訳の合計不良枚数
''''    Dim wPUCUTMCB       As Long             '合計不良枚数
''''    Dim wErrMsg         As String
''''    Dim SXL()           As typ_XSDCB_Update '分割結晶(ＳＸＬ)
''''    Dim wSXL()          As typ_XSDCB        '分割結晶(ＳＸＬ)
''''    Dim intDataCnt      As Integer          '該当データ件数
''''    Dim strBlockID      As String
''''
''''    'エラーハンドラの設定
''''    On Error GoTo proc_err
''''
''''    XSDCBProc = FUNCTION_RETURN_FAILURE
''''
''''    For i = 0 To Kihon.CNTHINNOW - 1
''''        '分割結晶（品番）から同じＳＸＬＩＤの長さの合計を取得
''''        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
''''        sql = sql & " SUM(GNMCA) AS wGNMCA "
''''        sql = sql & " FROM XSDCA "
''''        sql = sql & " WHERE SXLIDCA = '" & HinNow(i).SXLIDCA & "' "
''''        sql = sql & " AND LIVKCA = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、エラー
''''        If rs Is Nothing Then
''''            MsgBox "XSDCA SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '抽出結果を格納する
''''        If IsNull(rs.Fields("wGNLCA")) = False Then
''''            wLENCB = rs.Fields("wGNLCA")
''''        Else
''''            wLENCB = 0
''''        End If
''''        If IsNull(rs.Fields("wGNMCA")) = False Then
''''            wMAICB = rs.Fields("wGNMCA")
''''        Else
''''            wMAICB = 0
''''        End If
''''
''''        '不良内訳から同じＳＸＬＩＤの不良枚数の合計を取得
''''        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
''''        sql = sql & " FROM XSDC4 "
''''        sql = sql & " WHERE SXLIDC4 = '" & HinNow(i).SXLIDCA & "' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、エラー
''''        If rs Is Nothing Then
''''            MsgBox "XSDC4 SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '抽出結果を格納する
''''        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
''''            wPUCUTMCB = rs.Fields("wPUCUTMC4")
''''        Else
''''            wPUCUTMCB = 0
''''        End If
''''
''''        '分割結晶（品番）：良品のＳＸＬＩＤで分割結晶（ＳＸＬ）を検索
''''        sqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
''''        ReDim wSXL(0) As typ_XSDCB
''''
''''        'データの件数を取得
''''        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
''''        If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
''''            MsgBox "XSDCB SELECT ERROR"
''''            Exit Function
''''        Else                                    '正常
''''            'データが存在する場合はUPDATE
''''            If intDataCnt > 0 Then
''''                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
''''                If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
''''                    MsgBox "XSDCA SELECT ERROR"
''''                    Exit Function
''''                End If
''''
''''                ReDim SXL(0) As typ_XSDCB_Update
''''
''''                '分割結晶（ＳＸＬ）を更新
''''                SXL(0).LENCB = wLENCB
''''                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''
''''                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
''''                '分割結晶（ＳＸＬ）更新エラー
''''                If iRtn = FUNCTION_RETURN_FAILURE Then
''''                    MsgBox "XSDCB UPDATET ERROR"
''''                    Exit Function
''''                End If
''''             '存在しない時、追加
''''             ElseIf intDataCnt = 0 Then
''''                 ReDim SXL(0) As typ_XSDCB_Update
''''                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
''''                 SXL(0).KCNTCB = HinNow(i).KCKNTCA       ' 工程連番
''''                 SXL(0).XTALCB = HinNow(i).XTALCA        ' 結晶番号
''''                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' 結晶内開始位置
''''                 SXL(0).LENCB = wLENCB                   ' 長さ
''''                 SXL(0).HINBCB = HinNow(i).HINBCA        ' 品番
''''                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' 電話番号改訂番号
''''                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' 工場
''''                 SXL(0).OPECB = HinNow(i).OPECA          ' 操業条件
''''                 SXL(0).MAICB = wMAICB                   ' 実枚数
''''                 SXL(0).WSRMAICB = 0                     ' WS洗後枚数
''''                 SXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
''''                 SXL(0).WFCMAICB = 0                     ' 受入枚数
''''
''''                 SXL(0).SXLRMAICB = 0                    ' SXL指示(良品)
''''                 SXL(0).SXLNMAICB = 0                    ' SXL指示(不良)
''''                 SXL(0).WFCNMAICB = 0                    ' WFC内欠落枚数
''''                 SXL(0).SXLEMAICB = 0                    ' SXL確定枚数
''''                 SXL(0).SRMAICB = 0                      ' サンプル抜指示(良品)
''''                 SXL(0).SNMAICB = 0                      ' サンプル抜指示(不良)
''''                 SXL(0).STMAICB = 0                      ' サンプル枚数
''''                 '工程により振分（とりあえず画面分）
''''                 Select Case Kihon.NOWPROC
''''                     Case "CW740"
''''                         SXL(0).SRMAICB = wMAICB         ' サンプル抜指示(良品)
''''                         SXL(0).SNMAICB = wPUCUTMCB      ' サンプル抜指示(不良)
''''                     Case "CW800"
''''                         SXL(0).SXLEMAICB = wMAICB       ' SXL確定枚数
''''                 End Select
''''                 SXL(0).FURIMAICB = ""                 ' 振替枚数
''''                 SXL(0).XTWORKCB = "42"                  ' 製造工場
''''                 SXL(0).WFWORKCB = " "                   ' ウェーハ製造
''''                 SXL(0).FURYCCB = " "                     ' 不良理由
''''                 SXL(0).LSTCCB = "T"                     ' 採取状態区分
''''                 SXL(0).LUFRCCB = " "                    ' 格上コード
''''                 SXL(0).LUFRBCB = " "                    ' 格上区分
''''                 SXL(0).LDERCCB = " "                    ' 格下コード
''''                 SXL(0).LDFRBCB = "0"                    ' 格下区分
''''                 SXL(0).HOLDCCB = " "                    ' ホールドコード
''''                 SXL(0).HOLDBCB = " "                    ' ホールド区分
''''                 SXL(0).EXKUBCB = " "                    ' 例外区分
''''                 SXL(0).HENPKCB = " "                    ' 返品区分
''''''''                 'シングル確定時、完了区分='2'・生死区分='1'にする
''''''''                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''''''                    SXL(0).KANKCB = "2"                ' 完了区分
''''''''                    SXL(0).LIVKCB = "1"                 ' 生死区分
''''''''                 Else
''''                    SXL(0).KANKCB = "0"                ' 完了区分
''''                    SXL(0).LIVKCB = "0"                 ' 生死区分
''''''''                 End If
''''                 SXL(0).NFCB = "0"                       ' 入庫区分
''''                 SXL(0).SAKJCB = "0"                     ' 削除区分
''''                                                      ' 登録日付
''''                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                                                      ' 更新日付
''''                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                 SXL(0).SUMITCB = "0"                    ' SUMIT送信フラグ
''''                 SXL(0).SNDKCB = "0"                    ' 返品区分
''''                 SXL(0).SNDAYCB = ""                   ' 送信日付
''''
''''                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
''''                 '分割結晶（ＳＸＬ）追加エラー
''''                 If iRtn = FUNCTION_RETURN_FAILURE Then
''''                     MsgBox wErrMsg
''''                     Exit Function
''''                 End If
''''             End If
''''        End If
''''    Next i
''''
''''''''' 02/09/20 Add By 会田@HITEC  sta：分割結晶（品番）の合計長さが0になった時、
'''''''''                                 分割結晶（ＳＸＬ）を死ロットにする
''''    For i = 0 To Kihon.CNTHINOLD - 1
''''        '分割結晶（品番）前工程から死ロットの同じＳＸＬＩＤの長さの合計を取得
''''        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
''''        sql = sql & " SUM(GNMCA) AS wGNMCA, "
''''        sql = sql & " max(CRYNUMCA) AS CRYNUMCA"
''''        sql = sql & " FROM XSDCA "
''''        sql = sql & " WHERE SXLIDCA = '" & HinOld(i).SXLIDCA & "' "
''''        sql = sql & " AND LIVKCA = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、エラー
''''        If rs Is Nothing Then
''''            MsgBox "XSDCA SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '抽出結果を格納する
''''        If IsNull(rs.Fields("wGNLCA")) = False Then
''''            wLENCB = rs.Fields("wGNLCA")
''''        Else
''''            wLENCB = 0
''''        End If
''''        If IsNull(rs.Fields("wGNMCA")) = False Then
''''            wMAICB = rs.Fields("wGNMCA")
''''        Else
''''            wMAICB = 0
''''        End If
''''
''''        '不良内訳から同じＳＸＬＩＤの不良枚数の合計を取得
''''        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
''''        sql = sql & " FROM XSDC4 "
''''        sql = sql & " WHERE SXLIDC4 = '" & HinOld(i).SXLIDCA & "' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、エラー
''''        If rs Is Nothing Then
''''            MsgBox "XSDC4 SELECT ERROR"
''''            Exit Function
''''        End If
''''
''''        '抽出結果を格納する
''''        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
''''            wPUCUTMCB = rs.Fields("wPUCUTMC4")
''''        Else
''''            wPUCUTMCB = 0
''''        End If
''''
''''        '分割結晶（品番）：ＳＸＬＩＤで分割結晶（ＳＸＬ）を検索
''''        sqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
''''        ReDim wSXL(0) As typ_XSDCB
''''
''''        'データの件数を取得
''''        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
''''        If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
''''            MsgBox "XSDCB SELECT ERROR"
''''            Exit Function
''''        Else                                    '正常
''''            'データが存在する場合はUPDATE
''''            If intDataCnt > 0 Then
''''                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
''''                If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
''''                    MsgBox "XSDCA SELECT ERROR"
''''                    Exit Function
''''                End If
''''
''''                ReDim sxl(0) As typ_XSDCB_Update
''''
''''                '分割結晶（ＳＸＬ）を更新
''''                SXL(0).LENCB = wLENCB
''''                '長さが0の時、死ロットとする
''''                If wLENCB = 0 Then
''''                    SXL(0).LDFRBCB = "2"
''''                    SXL(0).LIVKCB = "1"                 ' 生死区分
''''                    SXL(0).LDFRBCB = "H"
''''                    SXL(0).KANKCB = "2"
''''                End If
''''
''''                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''
''''                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
''''                '分割結晶（ＳＸＬ）更新エラー
''''                If iRtn = FUNCTION_RETURN_FAILURE Then
''''                    MsgBox "XSDCB UPDATET ERROR"
''''                    Exit Function
''''                End If
''''             '存在しない時、追加
''''             ElseIf intDataCnt = 0 Then
''''                 ReDim SXL(0) As typ_XSDCB_Update
''''                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
''''                 SXL(0).KCNTCB = HinNow(i).KCKNTCA       ' 工程連番
''''                 SXL(0).XTALCB = HinNow(i).XTALCA        ' 結晶番号
''''                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' 結晶内開始位置
''''                 SXL(0).LENCB = wLENCB                   ' 長さ
''''                 SXL(0).HINBCB = HinNow(i).HINBCA        ' 品番
''''                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' 電話番号改訂番号
''''                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' 工場
''''                 SXL(0).OPECB = HinNow(i).OPECA          ' 操業条件
''''                 SXL(0).MAICB = wMAICB                   ' 実枚数
''''                 SXL(0).WSRMAICB = 0                     ' WS洗後枚数
''''                 SXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
''''                 SXL(0).WFCMAICB = 0                     ' 受入枚数
''''                 SXL(0).SXLRMAICB = 0                    ' SXL指示(良品)
''''                 SXL(0).SXLNMAICB = 0                    ' SXL指示(不良)
''''                 SXL(0).WFCNMAICB = 0                    ' WFC内欠落枚数
''''                 SXL(0).SXLEMAICB = 0                    ' SXL確定枚数
''''                 SXL(0).SRMAICB = 0                      ' サンプル抜指示(良品)
''''                 SXL(0).SNMAICB = 0                      ' サンプル抜指示(不良)
''''                 SXL(0).STMAICB = 0                      ' サンプル枚数
''''                 '工程により振分（とりあえず画面分）
''''                 Select Case Kihon.NOWPROC
''''                     Case "CW740"
''''                         SXL(0).SRMAICB = wMAICB         ' サンプル抜指示(良品)
''''                         SXL(0).SNMAICB = wPUCUTMCB      ' サンプル抜指示(不良)
''''                     Case "CW800"
''''                         SXL(0).SXLEMAICB = wMAICB       ' SXL確定枚数
''''                 End Select
''''                 SXL(0).FURIMAICB = ""                 ' 振替枚数
''''                 SXL(0).XTWORKCB = "42"                  ' 製造工場
''''                 SXL(0).WFWORKCB = " "                   ' ウェーハ製造
''''                 SXL(0).FURYCCB = " "                     ' 不良理由
''''                 SXL(0).LSTCCB = "T"                     ' 採取状態区分
''''                 SXL(0).LUFRCCB = " "                    ' 格上コード
''''                 SXL(0).LUFRBCB = " "                    ' 格上区分
''''                 SXL(0).LDERCCB = " "                    ' 格下コード
''''                '長さが0の時、廃棄とする
''''                 If wLENCB = 0 Then
''''                     SXL(0).LDFRBCB = "2"                ' 格下区分
''''                 Else
''''                     SXL(0).LDFRBCB = "0"
''''                 End If
''''                 SXL(0).HOLDCCB = " "                    ' ホールドコード
''''                 SXL(0).HOLDBCB = " "                    ' ホールド区分
''''                 SXL(0).EXKUBCB = " "                    ' 例外区分
''''                 SXL(0).HENPKCB = " "                    ' 返品区分
''''                 '長さが0の時、死ロットとする
''''                 If wLENCB = 0 Then
''''                     SXL(0).LIVKCB = "1"                 ' 生死区分
''''                     SXL(0).KANKCB = "2"                    ' 完了区分
''''                 Else
''''                     SXL(0).LIVKCB = "0"
''''                     SXL(0).KANKCB = "0"                    ' 完了区分
''''                 End If
''''                 SXL(0).KANKCB = "0"                    ' 完了区分
''''                 SXL(0).NFCB = "0"                       ' 入庫区分
''''                 SXL(0).SAKJCB = "0"                     ' 削除区分
''''                                                      ' 登録日付
''''                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                                                      ' 更新日付
''''                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
''''                 SXL(0).SUMITCB = "0"                    ' SUMIT送信フラグ
''''                 SXL(0).SNDKCB = "0"                    ' 返品区分
''''                 SXL(0).SNDAYCB = ""                   ' 送信日付
''''
''''                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
''''                 '分割結晶（ＳＸＬ）追加エラー
''''                 If iRtn = FUNCTION_RETURN_FAILURE Then
''''                     MsgBox wErrMsg
''''                     Exit Function
''''                 End If
''''             End If
''''        End If
''''    Next i
''''
''''''''' 02/09/20 Add                end
''''
''''    XSDCBProc = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '' 終了
'''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '' エラーハンドラ
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    XSDCBProc = FUNCTION_RETURN_FAILURE
'''''    gErr.HandleError
'''''    sErrMsg = GetMsgStr("EXXXX", sDBName)
''''    Resume proc_exit
''''
''''End Function
''''




'概要      :分割結晶（ＳＸＬ）登録処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function XSDCBProc()
    
'   内部変数
    Dim i               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wGNLCA          As Long             '分割結晶（品番）の合計長さ
    Dim wGNMCA          As Long             '分割結晶（品番）の合計枚数
'    Dim wLENCB          As Long             '合計長さ
'    Dim wMAICB          As Long             '合計枚数
'    Dim wPUCUTMC4       As Long             '不良内訳の合計不良枚数
'    Dim wPUCUTMCB       As Long             '合計不良枚数
    Dim wErrMsg         As String
    Dim SXL()           As typ_XSDCB_Update '分割結晶(ＳＸＬ)
    Dim wSXL()          As typ_XSDCB        '分割結晶(ＳＸＬ)
    Dim intDataCnt      As Integer          '該当データ件数
    Dim strBlockID      As String
''''' 02/09/21 Add bY 会田@hitec  sta
    Dim wRYOMAI         As Long             '工程毎の良品枚数
    Dim wFRYMAI         As Long             '工程毎の不良品枚数
    Dim wLen            As Long             '長さ
    Dim wMAI            As Long             '枚数
    Dim wMAI800         As Long             'CW800枚数
    Dim wFUR            As Long             '不良枚数
    Dim wFURKEI         As Long             '不良枚数合計
    Dim wSAM            As Long             'サンプル枚数
    Dim wSIJ            As Long             'サンプル抜指示枚数
    Dim wSAMFUR         As Long             'サンプル抜指示不良枚数
''''' 02/09/21 Add                end
    Dim iLoopBkHinGet   As Integer          '元品番 'add 2003/04/03 hitec)matsumoto
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    XSDCBProc = FUNCTION_RETURN_FAILURE
    
    For i = 0 To Kihon.CNTHINNOW - 1
''''' 02/09/21 Dlt bY 会田@hitec  sta
'        '分割結晶（品番）から同じＳＸＬＩＤの長さの合計を取得
'        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
'        sql = sql & " SUM(GNMCA) AS wGNMCA "
'        sql = sql & " FROM XSDCA "
'        sql = sql & " WHERE SXLIDCA = '" & HinNow(i).SXLIDCA & "' "
'        sql = sql & " AND LIVKCA = '0' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '存在しない時、エラー
'        If rs Is Nothing Then
'            MsgBox "XSDCA SELECT ERROR"
'            Exit Function
'        End If
'
'        '抽出結果を格納する
'        If IsNull(rs.Fields("wGNLCA")) = False Then
'            wLENCB = rs.Fields("wGNLCA")
'        Else
'            wLENCB = 0
'        End If
'        If IsNull(rs.Fields("wGNMCA")) = False Then
'            wMAICB = rs.Fields("wGNMCA")
'        Else
'            wMAICB = 0
'        End If
'
'        '不良内訳から同じＳＸＬＩＤの不良枚数の合計を取得
'        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
'        sql = sql & " FROM XSDC4 "
'        sql = sql & " WHERE SXLIDC4 = '" & HinNow(i).SXLIDCA & "' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '存在しない時、エラー
'        If rs Is Nothing Then
'            MsgBox "XSDC4 SELECT ERROR"
'            Exit Function
'        End If
'
'        '抽出結果を格納する
'        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
'            wPUCUTMCB = rs.Fields("wPUCUTMC4")
'        Else
'            wPUCUTMCB = 0
'        End If
'
''''' 02/09/21 Dlt                end

''''' 02/09/21 Add bY 会田@hitec  sta
'        '工程実績から同じＳＸＬＩＤの長さ、枚数、不良枚数の合計を取得
        iRtn = XSDCBSum(Kihon.NOWPROC, HinNow(i).SXLIDCA, wLen, wMAI, wMAI800, wFUR, wFURKEI, wSAM, wSAMSIJ, wSAMFUR)
''''' 02/09/21 Add                end

        '分割結晶（品番）：良品のＳＸＬＩＤで分割結晶（ＳＸＬ）を検索
        sqlWhere = "WHERE SXLIDCB = '" & HinNow(i).SXLIDCA & "' "
        ReDim wSXL(0) As typ_XSDCB

        'データの件数を取得
        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                    '正常
            'データが存在する場合はUPDATE
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If
                
                ReDim SXL(0) As typ_XSDCB_Update
            
                '分割結晶（ＳＸＬ）を更新
'''''                SXL(0).LENCB = wLENCB
                SXL(0).LENCB = wLen
' VVVVV 2003/05/02 ADD BY HITEC)会田：更新時も品番変更する
                SXL(0).HINBCB = HinNow(i).HINBCA        ' 品番
' ^^^^^ 2003/05/02 ADD BY HITEC)会田  END
                SXL(0).MAICB = wMAI
                SXL(0).KCNTCB = BlkNow.KCNTC2           ' 工程連番
                'シングル確定時、最終状態区分='S'にする
                SXL(0).LIVKCB = "0"
                SXL(0).KANKCB = "0"                 ' 完了区分
                SXL(0).LSTCCB = "T"                 ' 最終状態区分
                SXL(0).LDFRBCB = "0"                ' 格下区分
                'add start  2003/06/09 hitec)matsumoto ---------------
                SXL(0).INPOSCB = HinNow(i).INPOSCA      ' 結晶内開始位置
                SXL(0).LENCB = wLen                     ' 長さ
                'add end    2003/06/09 hitec)matsumoto ---------------
                If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                End If
                 '工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                         'add start 2003/03/25 hitec)matsumoto
                         SXL(0).NEWKNTCB = "CW740"       ' 最終通過工程
                         SXL(0).GNWKNTCB = "CW750"       ' 現在工程
                         'add end 2003/03/25 hitec)matsumoto
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' サンプル抜指示枚数
                         SXL(0).SNMAICB = wSAMFUR        ' サンプル抜指示不良枚数
                         SXL(0).STMAICB = wSAM           ' サンプル枚数
                         'add start 2003/10/22 tuku → 03/11/05 ブロック単位で変更されてしまうためコメント化
'                         SXL(0).NEWKNTCB = "CW750"       ' 最終通過工程
'                         SXL(0).GNWKNTCB = "CW800"       ' 現在工程
                         'add end 2003/10/22 tuku
                         
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                         'add start 2003/10/22
                         SXL(0).NEWKNTCB = "CW760"       ' 最終通過工程
                         SXL(0).GNWKNTCB = "CW750"       ' 現在工程
                         'add end 2003/10/22 tuku
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI800      ' SXL指示（良品）
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC内欠落枚数
                End Select
                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                
                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
                '分割結晶（ＳＸＬ）更新エラー
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             '存在しない時、追加
             ElseIf intDataCnt = 0 Then
                 ReDim SXL(0) As typ_XSDCB_Update
                 SXL(0).SXLIDCB = HinNow(i).SXLIDCA      ' SXLID
                 SXL(0).KCNTCB = BlkNow.KCNTC2           ' 工程連番
                 SXL(0).XTALCB = HinNow(i).XTALCA        ' 結晶番号
                 SXL(0).INPOSCB = HinNow(i).INPOSCA      ' 結晶内開始位置
                 SXL(0).LENCB = wLen                     ' 長さ
                 SXL(0).HINBCB = HinNow(i).HINBCA        ' 品番
                 SXL(0).REVNUMCB = HinNow(i).REVNUMCA    ' 電話番号改訂番号
                 SXL(0).FACTORYCB = HinNow(i).FACTORYCA  ' 工場
                 SXL(0).OPECB = HinNow(i).OPECA          ' 操業条件
                 SXL(0).MAICB = wMAI                     ' 実枚数
                 SXL(0).WSRMAICB = 0                     ' WS洗後枚数
                 SXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
                 SXL(0).WFCMAICB = 0                     ' 受入枚数
                 SXL(0).SXLRMAICB = 0                    ' SXL指示(良品)
                 SXL(0).SXLEMAICB = 0                    ' SXL確定枚数
                 '工程により振分（とりあえず画面分）
                 Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                         'add start 2003/03/25 hitec)matsumoto
                         SXL(0).NEWKNTCB = "CW740"       ' 最終通過工程
                         SXL(0).GNWKNTCB = "CW750"       ' 現在工程
                         'add end 2003/03/25 hitec)matsumoto
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' サンプル抜指示枚数
                         SXL(0).SNMAICB = wSAMFUR        ' サンプル抜指示不良枚数
                         SXL(0).STMAICB = wSAM           ' サンプル枚数
                         'add start 2003/10/22 tuku　→ 03/11/05 ブロック単位で変更されてしまうためコメント化
'                         SXL(0).NEWKNTCB = "CW750"       ' 最終通過工程
'                         SXL(0).GNWKNTCB = "CW800"       ' 現在工程
                         'add end 2003/10/22 tuku
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                         'add start 2003/10/22
                         SXL(0).NEWKNTCB = "CW760"       ' 最終通過工程
                         SXL(0).GNWKNTCB = "CW750"       ' 現在工程
                         'add end 2003/10/22 tuku
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL指示（良品）
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC内欠落枚数
                 End Select
                 SXL(0).FURIMAICB = ""                   ' 振替枚数
                 SXL(0).XTWORKCB = "42"                  ' 製造工場
                 SXL(0).WFWORKCB = " "                   ' ウェーハ製造
                 SXL(0).FURYCCB = " "                     ' 不良理由
                 SXL(0).LSTCCB = "T"                     ' 採取状態区分
                 'シングル確定時、最終状態区分='S'にする
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
                 SXL(0).LUFRCCB = " "                    ' 格上コード
                 SXL(0).LUFRBCB = " "                    ' 格上区分
                 SXL(0).LDERCCB = " "                    ' 格下コード
                 SXL(0).LDFRBCB = "0"                    ' 格下区分
                 SXL(0).HOLDCCB = " "                    ' ホールドコード
                 SXL(0).HOLDBCB = " "                    ' ホールド区分
                 SXL(0).EXKUBCB = " "                    ' 例外区分
                 SXL(0).HENPKCB = " "                    ' 返品区分
''''                 'シングル確定時、完了区分='2'・生死区分='1'にする
''''                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
''''                    SXL(0).KANKCB = "2"                ' 完了区分
''''                    SXL(0).LIVKCB = "1"                ' 生死区分
''''                 Else
                    SXL(0).KANKCB = "0"                  ' 完了区分
                    SXL(0).LIVKCB = "0"                  ' 生死区分
''''                 End If
                 SXL(0).NFCB = "0"                       ' 入庫区分
                 SXL(0).SAKJCB = "0"                     ' 削除区分
                                                         ' 登録日付
                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                      ' 更新日付
                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                 SXL(0).SUMITCB = "0"                    ' SUMIT送信フラグ
                 SXL(0).SNDKCB = "0"                     ' 返品区分
                 SXL(0).SNDAYCB = ""                     ' 送信日付
            
                 'add start 2003/04/03 hitec)matsumoto 前ﾃﾞｰﾀがなく、元品番が取得できないので、HINOLDから該当位置の品番を取得し、それを元品番とする　---------------
                 SXL(0).MOTHINCB = vbNullString '初期化
' VVVVV 2003/04/30 ALT BY HITEC)会田：CW740,CW760のみに変更
                 If Kihon.NOWPROC = "CW740" Or Kihon.NOWPROC = "CW760" Then
                    For iLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                        If (CInt(HinOld(iLoopBkHinGet).INPOSCA) <= CInt(SXL(0).INPOSCB)) And (CInt(SXL(0).INPOSCB) <= CInt(HinOld(iLoopBkHinGet).INPOSCA) + CInt(HinOld(iLoopBkHinGet).GNLCA)) Then
                             SXL(0).MOTHINCB = HinOld(iLoopBkHinGet).HINBCA
                             Exit For
                        End If
                    Next
                    If SXL(0).MOTHINCB = vbNullString Then 'もし該当HINOLDが無かったら自分の品番を元品番とする
                        SXL(0).MOTHINCB = SXL(0).HINBCB
                    End If
                 End If
' ^^^^^^ 2003/04/30 ALT BY HITEC)会田  END
                 'add end   2003/04/03 hitec)matsumoto ---------------
            
                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
                 '分割結晶（ＳＸＬ）追加エラー
                 If iRtn = FUNCTION_RETURN_FAILURE Then
                     MsgBox wErrMsg
                     Exit Function
                 End If
             End If
        End If
    Next i
    
''''' 02/09/20 Add By 会田@HITEC  sta：分割結晶（品番）の合計長さが0になった時、
'''''                                 分割結晶（ＳＸＬ）を死ロットにする
    For i = 0 To Kihon.CNTHINOLD - 1
''''' 02/09/21 Dlt bY 会田@hitec  sta
'        '分割結晶（品番）前工程から死ロットの同じＳＸＬＩＤの長さの合計を取得
'        sql = "SELECT SUM(GNLCA) AS wGNLCA, "
'        sql = sql & " SUM(GNMCA) AS wGNMCA, "
'        sql = sql & " max(CRYNUMCA) AS CRYNUMCA"
'        sql = sql & " FROM XSDCA "
'        sql = sql & " WHERE SXLIDCA = '" & HinOld(i).SXLIDCA & "' "
'        sql = sql & " AND LIVKCA = '0' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '存在しない時、エラー
'        If rs Is Nothing Then
'            MsgBox "XSDCA SELECT ERROR"
'            Exit Function
'        End If
'
'        '抽出結果を格納する
'        If IsNull(rs.Fields("wGNLCA")) = False Then
'            wLENCB = rs.Fields("wGNLCA")
'        Else
'            wLENCB = 0
'        End If
'        If IsNull(rs.Fields("wGNMCA")) = False Then
'            wMAICB = rs.Fields("wGNMCA")
'        Else
'            wMAICB = 0
'        End If
'
'        '不良内訳から同じＳＸＬＩＤの不良枚数の合計を取得
'        sql = "SELECT SUM(PUCUTMC4) AS wPUCUTMC4 "
'        sql = sql & " FROM XSDC4 "
'        sql = sql & " WHERE SXLIDC4 = '" & HinOld(i).SXLIDCA & "' "
'
'        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'        '存在しない時、エラー
'        If rs Is Nothing Then
'            MsgBox "XSDC4 SELECT ERROR"
'            Exit Function
'        End If
'
'        '抽出結果を格納する
'        If IsNull(rs.Fields("wPUCUTMC4")) = False Then
'            wPUCUTMCB = rs.Fields("wPUCUTMC4")
'        Else
'            wPUCUTMCB = 0
'        End If
    
''''' 02/09/21 Dlt                end

''''' 02/09/21 Add bY 会田@hitec  sta
'        '工程実績から同じＳＸＬＩＤの長さ、枚数、不良枚数の合計を取得
        iRtn = XSDCBSum(Kihon.NOWPROC, HinOld(i).SXLIDCA, wLen, wMAI, pMAI800, wFUR, wFURKEI, wSAM, wSAMSIJ, wSAMFUR)
''''' 02/09/21 Add                end
        
        '分割結晶（品番）：ＳＸＬＩＤで分割結晶（ＳＸＬ）を検索
        sqlWhere = "WHERE SXLIDCB = '" & HinOld(i).SXLIDCA & "' "
        ReDim wSXL(0) As typ_XSDCB
        
        'データの件数を取得
        iRtn = SelCntXSDCB(sqlWhere, intDataCnt)
        If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
            MsgBox "XSDCB SELECT ERROR"
            Exit Function
        Else                                    '正常
            'データが存在する場合はUPDATE
            If intDataCnt > 0 Then
                iRtn = DBDRV_GetXSDCB(wSXL(), sqlWhere)
                If iRtn = FUNCTION_RETURN_FAILURE Then  'エラー
                    MsgBox "XSDCA SELECT ERROR"
                    Exit Function
                End If
                
                ReDim SXL(0) As typ_XSDCB_Update
            
                '分割結晶（ＳＸＬ）を更新
'''                SXL(0).LENCB = wLENCB
                SXL(0).LENCB = wLen
''''' VVVVV 2003/05/02 ADD BY HITEC)会田：更新時も品番変更する
''''                SXL(0).HINBCB = HinNow(i).HINBCA        ' 品番
''''' ^^^^^ 2003/05/02 ADD BY HITEC)会田  END
                SXL(0).MAICB = wMAI
                SXL(0).KCNTCB = BlkNow.KCNTC2
                '工程により振分（とりあえず画面分）
                Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' サンプル抜指示枚数
                         SXL(0).SNMAICB = wSAMFUR        ' サンプル抜指示不良枚数
                         SXL(0).STMAICB = wSAM           ' サンプル枚数
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL指示（良品）
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC内欠落枚数
                End Select
                '長さが0の時、死ロットとする
''''                If wLen = 0 Then
'                If wMAI = 0 Then    'upd 2003/06/05 hitec)matsumoto
                If (wMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (wLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     SXL(0).LIVKCB = "1"                 ' 生死区分
                     SXL(0).KANKCB = "2"                 ' 完了区分
                     SXL(0).LSTCCB = "H"                 ' 最終状態区分
                     SXL(0).LDFRBCB = "2"                ' 格下区分
                 Else
                     SXL(0).LIVKCB = "0"
                     SXL(0).KANKCB = "0"                 ' 完了区分
                     SXL(0).LSTCCB = "T"                 ' 最終状態区分
                     SXL(0).LDFRBCB = "0"                ' 格下区分
                End If
                ' 2002/12/13 ooba 完了区分フラグ変更
                SXL(0).KANKCB = "0"                 ' 完了区分

                 'シングル確定時、最終状態区分='S'にする
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
                SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
            
                iRtn = UpdateXSDCB(SXL(0), sqlWhere)
                '分割結晶（ＳＸＬ）更新エラー
                If iRtn = FUNCTION_RETURN_FAILURE Then
                    MsgBox "XSDCB UPDATET ERROR"
                    Exit Function
                End If
             '存在しない時、追加
             ElseIf intDataCnt = 0 Then
                 ReDim SXL(0) As typ_XSDCB_Update
                 SXL(0).SXLIDCB = HinOld(i).SXLIDCA      ' SXLID
                 SXL(0).KCNTCB = BlkNow.KCNTC2       ' 工程連番
                 SXL(0).XTALCB = HinOld(i).XTALCA        ' 結晶番号
                 SXL(0).INPOSCB = HinOld(i).INPOSCA      ' 結晶内開始位置
                 SXL(0).LENCB = wLen                     ' 長さ
                 SXL(0).HINBCB = HinOld(i).HINBCA        ' 品番
                 SXL(0).REVNUMCB = HinOld(i).REVNUMCA    ' 電話番号改訂番号
                 SXL(0).FACTORYCB = HinOld(i).FACTORYCA  ' 工場
                 SXL(0).OPECB = HinOld(i).OPECA          ' 操業条件
                 SXL(0).MAICB = wMAI                     ' 実枚数
                 SXL(0).WSRMAICB = 0                     ' WS洗後枚数
                 SXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
                 SXL(0).WFCMAICB = 0                     ' 受入枚数
                 SXL(0).WSNMAICB = 0                     ' WS洗浄欠落枚数
                 SXL(0).WFCMAICB = 0                     ' 受入枚数
                 SXL(0).SXLEMAICB = 0                    ' SXL確定枚数
                 '工程により振分（とりあえず画面分）
                 Select Case Kihon.NOWPROC
                     Case "CW740"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                     Case "CW750"
                         SXL(0).SRMAICB = wSIJ           ' サンプル抜指示枚数
                         SXL(0).SNMAICB = wSAMFUR        ' サンプル抜指示不良枚数
                         SXL(0).STMAICB = wSAM           ' サンプル枚数
                     Case "CW760"
                         SXL(0).SXLNMAICB = wFUR         ' 廃棄WF枚数
                     Case "CW800"
                         SXL(0).SXLRMAICB = wMAI         ' SXL指示（良品）
                         SXL(0).WFCNMAICB = wFURKEI      ' WFC内欠落枚数
                 End Select
                 SXL(0).FURIMAICB = ""                   ' 振替枚数
                 SXL(0).XTWORKCB = "42"                  ' 製造工場
                 SXL(0).WFWORKCB = " "                   ' ウェーハ製造
                 SXL(0).FURYCCB = " "                    ' 不良理由
                 SXL(0).LSTCCB = "T"                     ' 採取状態区分
                 SXL(0).LUFRCCB = " "                    ' 格上コード
                 SXL(0).LUFRBCB = " "                    ' 格上区分
                 SXL(0).LDERCCB = " "                    ' 格下コード
                '長さが0の時、廃棄とする
                 If wLENCB = 0 Then
                     SXL(0).LDFRBCB = "2"                ' 格下区分
                 Else
                     SXL(0).LDFRBCB = "0"
                 End If
                 SXL(0).HOLDCCB = " "                    ' ホールドコード
                 SXL(0).HOLDBCB = " "                    ' ホールド区分
                 SXL(0).EXKUBCB = " "                    ' 例外区分
                 SXL(0).HENPKCB = " "                    ' 返品区分
                 '長さが0の時、死ロットとする
''''                 If wLENCB = 0 Then
'                 If wMAI = 0 Then    'upd 2003/06/05 hitec)matsumoto
                If (wMAI = 0 And Kihon.NOWPROC <> PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Or _
                        (wLen = 0 And Kihon.NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU) Then    '05/03/29 ooba
                     SXL(0).LIVKCB = "1"                 ' 生死区分
                     SXL(0).KANKCB = "2"                 ' 完了区分
                     SXL(0).LSTCCB = "H"                 ' 最終状態区分
                     SXL(0).LDFRBCB = "2"                ' 格下区分
                 Else
                     SXL(0).LIVKCB = "0"
                     SXL(0).KANKCB = "0"                 ' 完了区分
                     SXL(0).LSTCCB = "T"                 ' 最終状態区分
                     SXL(0).LDFRBCB = "0"                ' 格下区分
                 End If
                 ' 2002/12/13 ooba 完了区分フラグ変更
                 SXL(0).KANKCB = "0"                 ' 完了区分
                 'シングル確定時、最終状態区分='S'にする
                 If Kihon.NOWPROC = PROCD_SXL_KAKUTEI Then
                    SXL(0).LSTCCB = "S"
                 End If
''''                 SXL(0).KANKCB = "0"                     ' 完了区分
                 SXL(0).NFCB = "0"                       ' 入庫区分
                 SXL(0).SAKJCB = "0"                     ' 削除区分
                                                         ' 登録日付
                 SXL(0).TDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                                                      ' 更新日付
                 SXL(0).KDAYCB = Format(Now(), "YYYY/MM/DD HH:NN:SS")
                 SXL(0).SUMITCB = "0"                    ' SUMIT送信フラグ
                 SXL(0).SNDKCB = "0"                     ' 返品区分
                 SXL(0).SNDAYCB = ""                    ' 送信日付
            
                 iRtn = CreateXSDCB(SXL(0), wErrMsg)
                 '分割結晶（ＳＸＬ）追加エラー
                 If iRtn = FUNCTION_RETURN_FAILURE Then
                     MsgBox wErrMsg
                     Exit Function
                 End If
             End If
        End If
    Next i

''''' 02/09/20 Add                end
    
    XSDCBProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDCBProc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'概要      :工程実績から指定された工程、ＳＸＬＩＤ（ブロックＩＤ、位置、品番）の長さ、枚数、不良枚数を集計する
'          :受信用のテーブルから指定されたＳＸＬＩＤ）のサンプル枚数、サンプル抜指示枚数、サンプル抜指示不良枚数
'           を集計する
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'           pKKTC          I   string            工程
'           pSXLID         I   string            ＳＸＬＩＤ
'           pLEN           O   NUMBER            長さ
'           pMAI           O   NUMBER            枚数
'           pMAI800        O   NUMBER            CW800枚数
'           pFUR           O   NUMBER            不良枚数
'           pFURKEI        O   NUMBER            不良枚数合計
'           pSAM           O   NUMBER            サンプル枚数
'           pSAMNUK        O   NUMBER            サンプル抜指示枚数
'           pSAMFUR        O   NUMBER            サンプル抜指示不良枚数
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function XSDCBSum(ByVal pKKTC, ByVal pSXLID, ByRef pLEN, ByRef pMAI, ByRef pMAI800, ByRef pFUR, ByRef pFURKEI, ByRef pSAM, ByRef pSAMSIJ, ByRef pSAMFUR)
    
'   内部変数
    Dim i               As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim wCRYNUMCA       As String           'ブロックＩＤ
    Dim wINPOSCA        As Long             '開始位置
    Dim wHINBCA         As String           '品番
    Dim wLen            As Long             '長さ
    Dim wMAI            As Long             '枚数
    Dim wMAI800         As Long             'CW800を通過した枚数
    Dim wFUR            As Long             '不良枚数
    Dim wFURKEI         As Long             '不良枚数合計
    Dim wKCNTC3         As String           '工程連番最大
    Dim wSAMFUR         As String           'サンプル抜試指示不良枚数
    Dim rsXsdca         As OraDynaset
    Dim rsMain          As OraDynaset

    'エラーハンドラの設定
    On Error GoTo proc_err
        
    'パラメータノ初期化
    pLEN = 0
    pMAI = 0
    pMAI800 = 0
    pFUR = 0
    pFURKE = 0
    pSAM = 0
    pSAMSIJ = 0
    pSAMFUR = 0

    '分割結晶（品番）からパラメータのＳＸＬＩＤの長さ、枚数を取得
    sql = "SELECT SUM(GNLCA) AS wLEN, SUM(GNMCA) AS wMAI "
    sql = sql & " FROM XSDCA "
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "' "
    sql = sql & " AND LIVKCA = '0' "

    Set rsXsdca = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、次へ
    If rsXsdca.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsXsdca.Close
        GoTo CW800_CAL
    End If
    
    '抽出結果を格納する
    If IsNull(rsXsdca.Fields("wLEN")) = True Then
        pLEN = 0
    Else
        pLEN = rsXsdca.Fields("wLEN")                '長さ
    End If
    If IsNull(rsXsdca.Fields("wMAI")) = True Then
        pMAI = 0
    Else
        pMAI = rsXsdca.Fields("wMAI")                '枚数
    End If
    
    rsXsdca.Close

CW800_CAL:
    
    '分割結晶（品番）から同じＳＸＬＩＤのブロックＩＤ、開始位置、品番を取得
    sql = "SELECT CRYNUMCA, INPOSCA, HINBCA "
    sql = sql & " FROM XSDCA "
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "' "
    sql = sql & " AND LIVKCA = '0' "
    sql = sql & " ORDER BY CRYNUMCA,INPOSCA"

    Set rsMain = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、次へ
    If rsMain.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rsMain.Close
        GoTo SAMPLE_CAL
    End If

    Do Until rsMain.EOF
        '抽出結果を格納する
        wCRYNUMCA = rsMain.Fields("CRYNUMCA")
        wINPOSCA = rsMain.Fields("INPOSCA")
        wHINBCA = rsMain.Fields("HINBCA")

        '取得したブロックＩＤ､開始位置､品番で工程実績の該当工程で、工程連番の最大を取得する
''''        sql = "SELECT MAX(KCNTC3) AS wKCNTC3 "
''''        sql = sql & " FROM XSDC3 "
''''        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
''''        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
''''        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
''''        sql = sql & " AND WKKTC3 = '" & pKKTC & "' "
''''''''        sql = sql & " AND LIVKC3 = '0' "
''''        sql = sql & " AND ((SUMKBC3 = '0') "
''''        sql = sql & "  OR (SUMKBC3 = ' ') "
''''        sql = sql & "  OR (SUMKBC3 is null)) "
''''''''        sql = sql & " AND KKCNTC3  = '0' "
''''
''''        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
''''        '存在しない時、続行
''''        If rs.RecordCount = 0 Then
''''            XSDCBSum = FUNCTION_RETURN_FAILURE
''''            rs.Close
''''            GoTo SAMPLE_CAL
''''        End If
''''
''''        '抽出結果を格納する
''''        If IsNull(rs.Fields("wKCNTC3")) = True Then
''''            XSDCBSum = FUNCTION_RETURN_FAILURE
''''            wKCNTC3 = 0
''''            rs.Close
''''            GoTo SAMPLE_CAL
''''        Else
''''            wKCNTC3 = rs.Fields("wKCNTC3")
''''        End If
        
        '取得したブロックＩＤ､開始位置､品番、工程連番で工程実績から長さ、枚数、不良枚数を取得する
''''        sql = "SELECT SUM(LENC3) AS wLEN, SUM(TOMC3) AS wMAI,SUM(FUMC3) AS wFUR "
        sql = "SELECT TOMC3 AS wMAI800,FUMC3 AS wFUR "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
''''        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
''''        sql = sql & " AND WKKTC3 = '" & pKKTC & "' "
''''        sql = sql & " AND LIVKC3 = '0' "
''''        sql = sql & " AND (SUMKBC3 = '0' "
''''        sql = sql & "  OR SUMKBC3 = ' ' "
''''        sql = sql & "  OR SUMKBC3 is null) "
        sql = sql & " AND KCNTC3  = (SELECT MAX(KCNTC3)"
        sql = sql & "                  FROM XSDC3"
        sql = sql & "                 WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & "                   AND HINBC3 = '" & wHINBCA & "'"
        sql = sql & "                   AND INPOSC3 = '" & wINPOSCA & "'"
        sql = sql & "                   AND WKKTC3 = '" & pKKTC & "' "
        sql = sql & "                   AND (SUMKBC3 = '0' "
        sql = sql & "                    OR SUMKBC3 = ' ' "
        sql = sql & "                    OR SUMKBC3 is null)) "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If
        
        '抽出結果を格納する
'''''        If IsNull(rs.Fields("wLEN")) = True Then
'''''            pLEN = pLEN + 0
'''''        Else
'''''            pLEN = pLEN + rs.Fields("wLEN")         '長さ
'''''        End If
        If IsNull(rs.Fields("wMAI800")) = True Then
            pMAI800 = pMAI800 + 0
        Else
            pMAI800 = pMAI800 + CInt(rs.Fields("wMAI800"))     '枚数
        End If
        If IsNull(rs.Fields("wFUR")) = True Then
            pFUR = pFUR + 0
        Else
            pFUR = pFUR + CInt(rs.Fields("wFUR"))              '不良長さ
        End If
        
        '取得したブロックＩＤ､開始位置､品番で工程実績の不良合計を取得する
        sql = "SELECT SUM(FUMC3) AS wFURKEI "
        sql = sql & " FROM XSDC3 "
        sql = sql & " WHERE CRYNUMC3 = '" & wCRYNUMCA & "' "
        sql = sql & " AND INPOSC3 = " & wINPOSCA & ""
        sql = sql & " AND HINBC3 = '" & wHINBCA & "' "
        sql = sql & " AND SUMKBC3 = '1' "
        sql = sql & " AND MODKBC3 = '0' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
            GoTo SAMPLE_CAL
        End If
        
        '抽出結果を格納する
        If IsNull(rs.Fields("wFURKEI")) = True Then
            pFURKEI = pFURKEI + 0 '不良長さ
        Else
            pFURKEI = pFURKEI + CInt(rs.Fields("wFURKEI")) '不良長さ    'upd 2003/05/20
        End If
        rsMain.MoveNext

    Loop

    rs.Close
    rsMain.Close

SAMPLE_CAL:
    '評価結果受信レコード数よりサンプル枚数を取得する
    '■サンプル枚数　-　評価結果受信レコード数（Y013)
'    sql = "SELECT COUNT(SAMPLEID) AS wSAM "
'    sql = sql & " FROM TBCMY013 "
'    sql = sql & " WHERE  SAMPLEID in ( "
'    sql = sql & " SELECT E044.SMPLID "
'    sql = sql & " FROM TBCME044 E044 "
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & "  FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
'    sql = sql & " WHERE (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND  E044.INGOTPOS = E042.INGOTPOS "
'    sql = sql & " AND SMPKBN = 'T' ) "
'    sql = sql & " OR (E044.CRYNUM = E042.CRYNUM"
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS + E042.LENGTH "
'    sql = sql & " AND SMPKBN = 'B' ))"

    sql = "SELECT COUNT(SAMPLEID) AS wSAM "
    sql = sql & " FROM TBCMY013 Y013"
    sql = sql & " WHERE  SAMPLEID in ( "
    sql = sql & " SELECT E044.REPSMPLIDCW "
    sql = sql & " FROM XSDCW E044 "
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      XTALCB as CRYNUM"
    sql = sql & "     ,INPOSCB as INGOTPOS"
    sql = sql & "     ,RLENCB as LENGTH"
    sql = sql & "    FROM"
    sql = sql & "      XSDCB"
    sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & "  FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = sql & " WHERE (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND  E044.INPOSCW = E042.INGOTPOS "
    sql = sql & " AND E044.SMPKBNCW = 'T' ) "
    sql = sql & " OR (E044.XTALCW = E042.CRYNUM"
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
    sql = sql & " AND E044.SMPKBNCW = 'B' ))"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    '抽出結果を格納する
    pSAM = rs.Fields("wSAM") 'サンプル枚数

    rs.Close

    '抜試指示レコード数よりサンプル指示枚数を取得する
    '■サンプル抜指示枚数（良品）　-　抜試指示レコード数（Y003)
'    sql = "SELECT COUNT(SAMPLEID) AS wSIJ"
'    sql = sql & " FROM TBCMY003 "
'    sql = sql & " WHERE SAMPLEID in ( "
'    sql = sql & " SELECT E044.SMPLID "
'    sql = sql & " FROM TBCME044 E044 "
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & " FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
'    sql = sql & " WHERE (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS "
'    sql = sql & " AND SMPKBN = 'T' ) "
'    sql = sql & " OR (E044.CRYNUM = E042.CRYNUM "
'    sql = sql & " AND E044.INGOTPOS = E042.INGOTPOS + E042.LENGTH "
'    sql = sql & " AND  SMPKBN = 'B' )) "

    sql = "SELECT COUNT(SAMPLEID) AS wSIJ"
    sql = sql & " FROM TBCMY003 "
    sql = sql & " WHERE SAMPLEID in ( "
    sql = sql & " SELECT E044.REPSMPLIDCW "
    sql = sql & " FROM XSDCW E044 "
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = sql & "  ,("
    sql = sql & "    SELECT"
    sql = sql & "      XTALCB as CRYNUM"
    sql = sql & "     ,INPOSCB as INGOTPOS"
    sql = sql & "     ,RLENCB as LENGTH"
    sql = sql & "    FROM"
    sql = sql & "      XSDCB"
    sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'    sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'    sql = sql & " FROM TBCME042 "
'    sql = sql & " WHERE SXLID = '" & pSXLID & "') E042 "
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
    sql = sql & " WHERE (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS "
    sql = sql & " AND E044.SMPKBNCW = 'T' ) "
    sql = sql & " OR (E044.XTALCW = E042.CRYNUM "
    sql = sql & " AND E044.INPOSCW = E042.INGOTPOS + E042.LENGTH "
    sql = sql & " AND  E044.SMPKBNCW = 'B' )) "

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    '抽出結果を格納する
    pSIJ = rs.Fields("wSIJ") 'サンプル抜指示枚数

    rs.Close

    'C欠落枚数よりサンプル抜指示不良枚数を取得する
    '■サンプル抜試指示不良枚数　-　C欠落枚数　-（Y012）

    '対象のブロックID取得
    sql = "SELECT DISTINCT(CRYNUMCA) "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE SXLIDCA = '" & pSXLID & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、続行
    If rs.RecordCount = 0 Then
        XSDCBSum = FUNCTION_RETURN_FAILURE
        rs.Close
'        GoTo PROC_EXIT
    End If

    Do Until rs.EOF
        '欠落情報COUNT(ブロックIDごとループしSUMする）

        '抽出結果を格納する
        wCRYNUMCA = rs.Fields("CRYNUMCA") 'ブロックID

        sql = "SELECT COUNT(Y012.LOTID) AS wSAMFUR "
        sql = sql & " FROM TBCMY012 Y012 "
''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
        sql = sql & "  ,("
        sql = sql & "    SELECT"
        sql = sql & "      XTALCB as CRYNUM"
        sql = sql & "     ,INPOSCB as INGOTPOS"
        sql = sql & "     ,RLENCB as LENGTH"
        sql = sql & "    FROM"
        sql = sql & "      XSDCB"
        sql = sql & "    WHERE SXLIDCB = '" & pSXLID & "'"
    sql = sql & "   ) E042"
''↑追加END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
'        sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH "
'        sql = sql & " FROM TBCME042 "
'        sql = sql & " WHERE SXLID = '" & pSXLID & "' ) E042 "
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/21 SMP岡本
        sql = sql & " ,(SELECT CRYNUM, INGOTPOS, LENGTH, BLOCKID "
        sql = sql & " FROM TBCME040 "
        sql = sql & " WHERE BLOCKID =  '" & wCRYNUMCA & "' ) E040 "
        sql = sql & " WHERE Y012.LOTID = E040.BLOCKID "
        sql = sql & " AND E042.INGOTPOS <= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sql = sql & " AND E042.INGOTPOS + E042.LENGTH  >= Y012.TOP_POS / 10 + E040.INGOTPOS "
        sql = sql & " AND REJCAT = 'C' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        '存在しない時、続行
        If rs.RecordCount = 0 Then
            XSDCBSum = FUNCTION_RETURN_FAILURE
            rs.Close
'            GoTo PROC_EXIT
        End If

        '抽出結果を格納する
        pSAMFUR = rs.Fields("wSAMFUR") 'サンプル抜試指示不良枚数

        rs.MoveNext

    Loop

    rs.Close

    XSDCBSum = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    XSDCBSum = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function









'###  枚数取得関数  ###########################

'概要      :WF枚数を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型        ,説明
'          :SelectBlkID     ,I  ,Integer   ,ブロックID
'          :intBlkLen       ,I  ,Integer   ,ブロック長さ
'          :intWfCnt        ,O  ,Integer   ,枚数
'          :戻り値          ,O  ,Integer   ,WF枚数
'説明      :
'履歴      :2002/09/12 ADD hitec)N.MATSUMOTO
Public Function WfCount(ByVal SelectBlkID As String, ByVal intBlkLen As Integer, ByRef intWfCnt As Integer) As FUNCTION_RETURN


Dim rec() As typ_cmkc001f_Disp
Dim ret As FUNCTION_RETURN
Dim recCnt As Long
Dim i As Long
Dim j As Integer
Dim s As String
Dim intWfNum    As Integer '枚数

    '###　枚数計算関数用パラメータ（HSXCTCEN & HSXCYCEN）取得 ###
    
    ''仕様・実績を読み込む
    ret = DBDRV_fcmkc001f_Disp(Trim(SelectBlkID), blkInfo, rec)   'SelectBlkId=ブロックID,blkInfo=ブロック管理構造体
    If ret = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    ElseIf UBound(rec) Then
        HSXCTCEN = rec(1).HSXCTCEN
        HSXCYCEN = rec(1).HSXCYCEN
    End If
    
    '########################################################
    
    
    '###  WF枚数計算用の基本値を取得する  ###
    Loss0 = val(GetCodeField("LG", "01", "LOSS0", "INFO1"))
    Loss4 = val(GetCodeField("LG", "01", "LOSS4", "INFO1"))
    Mlt4 = val(GetCodeField("LG", "01", "MLT4", "INFO1"))
    Pitch = val(GetCodeField("LG", "01", "PITCH", "INFO1"))
    '######################################
    
    
    '###　枚数計算関数用パラメータ（SEEDDEG）取得 ###
    
'頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota
'    If Left(Trim(SelectBlkID), 1) = "8" Then
'        '購入単結晶の場合
'        If DBDRV_getSEEDDEG(Trim(SelectBlkID), SEEDDEG) = FUNCTION_RETURN_FAILURE Then
'            GoTo proc_exit
'        End If
'    Else
        '引き上げ結晶の場合
        s = GetCodeField("SC", "28", Left$(blkInfo.SEED, 1), "INFO3")
        If Left$(s, 1) = "4" Then
            SEEDDEG = 4
        Else
            SEEDDEG = 0
        End If
'    End If
    
    '#############################################


    '###  枚数取得関数  ###########################
    intWfCnt = GetWfCount(val(intBlkLen), SEEDDEG, HSXCTCEN, HSXCYCEN)  'intWfCount=枚数
    WfCount = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

End Function
'2002/09/12 ADD hitec)N.MATSUMOTO

'2002/09/12 ADD hitec)N.MATSUMOTO End

'###  枚数取得関数  ###########################

'概要      :WF枚数を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :blkLen        ,I  ,Integer   ,ブロック長さ
'          :seedDeg       ,I  ,Integer   ,結晶のSEED傾き
'          :hinDegT       ,I  ,Double    ,品番傾き（縦）
'          :hinDegY       ,I  ,Double    ,品番傾き（横）
'          :戻り値        ,O  ,Integer   ,WF枚数
'説明      :
'履歴      :2001/8/30 作成  野村
Public Function GetWfCount(ByVal BlkLen%, ByVal SEEDDEG%, ByVal hinDegT As Double, ByVal hinDegY As Double) As Integer
Dim hinDeg As Integer
Dim s As String
Dim WfCnt As Integer

    If Pitch = 0# Then
        GetWfCount = 0
        Exit Function
    End If

    ''品番傾きを得る
    '結晶最終払い出し、品番傾きの求め方変更
    If (Abs(hinDegT) = 2.83) And (Abs(hinDegY) = 2.83) Then
        hinDeg = 4
    ElseIf (Abs(hinDegT) = 4) And (hinDegY = 0) Then
        hinDeg = 4
    ElseIf (hinDegT = 0) And (Abs(hinDegY) = 4) Then
        hinDeg = 4
    Else
        hinDeg = 0
    End If
    
    ''WF枚数を計算する
    If SEEDDEG = hinDeg Then
        '通常品の場合
        WfCnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        WfCnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
''''    If WfCnt < 0 Then WfCnt = 0
    GetWfCount = WfCnt
End Function
'##########################################################


'2002/09/12 ADD hitec)N.MATSUMOTO Start
'###　枚数計算関数用パラメータ（HSXCTCEN & HSXCYCEN）取得 ###

'概要      :結晶最終払出入力 表示用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                   ,説明
'      　　:BlockID_in　 ,I  ,String               ,ブロックID
'      　　:blkInfo　　　,O  ,typ_cmkc001f_Block   ,ブロック情報
'      　　:records　　　,O  ,typ_cmkc001f_Disp    ,製品仕様取得用
'      　　:戻り値       ,O  ,FUNCTION_RETURN      ,読み込みの成否
Public Function DBDRV_fcmkc001f_Disp(BlockID_in As String, blkInfo As typ_cmkc001f_Block, records() As typ_cmkc001f_Disp) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Integer
    Dim i As Long
    Dim n As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_SUCCESS
    
    ''ブロック情報を得る
    sql = "Select BLK.INGOTPOS, BLK.LENGTH, BLK.REALLEN, BLK.KRPROCCD, BLK.NOWPROC, BLK.LPKRPROCCD, " & _
          "BLK.LASTPASS, BLK.DELCLS, BLK.RSTATCLS, BLK.LSTATCLS, CRY.SEED " & _
          "From TBCME040 BLK, TBCME037 CRY " & _
          "Where (BLOCKID='" & BlockID_in & "') and (BLK.CRYNUM=CRY.CRYNUM)"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If
    With blkInfo
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
    
    
    
    ''製品仕様を得る
    sql = "select "
    sql = sql & "BH.E041HINBAN, "           ' 品番
    sql = sql & "BH.E041INGOTPOS, "         ' 結晶内開始位置
    sql = sql & "BH.E041REVNUM, "           ' 製品番号改訂番号
    sql = sql & "BH.E041FACTORY, "          ' 工場
    sql = sql & "BH.E041OPECOND, "          ' 操業条件
    sql = sql & "BH.E041LENGTH, "           ' 長さ
    '製品仕様SXLデータ
    sql = sql & "S.E018HSXD1CEN, "          ' 品ＳＸ直径１中心
    sql = sql & "S.E018HSXRMIN, "           ' 品ＳＸ比抵抗下限
    sql = sql & "S.E018HSXRMAX, "           ' 品ＳＸ比抵抗上限
    sql = sql & "S.E018HSXRMBNP, "          ' 品ＳＸ比抵抗面内分布
    sql = sql & "S.E018HSXRHWYS, "          ' 品ＳＸ比抵抗保証方法＿処
    sql = sql & "S.E019HSXONMIN, "          ' 品ＳＸ酸素濃度下限
    sql = sql & "S.E019HSXONMAX, "          ' 品ＳＸ酸素濃度上限
    sql = sql & "S.E019HSXONMBP, "          ' 品ＳＸ酸素濃度面内分布
    sql = sql & "S.E019HSXONHWS, "          ' 品ＳＸ酸素濃度保証方法＿処
    sql = sql & "S.E019HSXCNMIN, "          ' 品ＳＸ炭素濃度下限
    sql = sql & "S.E019HSXCNMAX, "          ' 品ＳＸ炭素濃度上限
    sql = sql & "S.E019HSXCNHWS, "          ' 品ＳＸ炭素濃度保証方法＿処
    sql = sql & "S.E019HSXTMMAXN, "          ' 品ＳＸ転位密度上限             項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXBM1AN, "          ' 品ＳＸＢＭＤ１平均下限
    sql = sql & "S.E020HSXBM1AX, "          ' 品ＳＸＢＭＤ１平均上限
    sql = sql & "S.E020HSXBM1HS, "          ' 品ＳＸＢＭＤ１保証方法＿処
    sql = sql & "S.E020HSXBM2AN, "          ' 品ＳＸＢＭＤ２平均下限
    sql = sql & "S.E020HSXBM2AX, "          ' 品ＳＸＢＭＤ２平均上限
    sql = sql & "S.E020HSXBM2HS, "          ' 品ＳＸＢＭＤ２保証方法＿処
    sql = sql & "S.E020HSXBM3AN, "          ' 品ＳＸＢＭＤ３平均下限
    sql = sql & "S.E020HSXBM3AX, "          ' 品ＳＸＢＭＤ３平均上限
    sql = sql & "S.E020HSXBM3HS, "          ' 品ＳＸＢＭＤ３保証方法＿処
    sql = sql & "S.E020HSXOF1AX, "          ' 品ＳＸＯＳＦ１平均上限
    sql = sql & "S.E020HSXOF1MX, "          ' 品ＳＸＯＳＦ１上限
    sql = sql & "S.E020HSXOF1HS, "          ' 品ＳＸＯＳＦ１ 保証方法＿処
    sql = sql & "S.E020HSXOF2AX, "          ' 品ＳＸＯＳＦ２平均上限
    sql = sql & "S.E020HSXOF2MX, "          ' 品ＳＸＯＳＦ２上限
    sql = sql & "S.E020HSXOF2HS, "          ' 品ＳＸＯＳＦ２ 保証方法＿処
    sql = sql & "S.E020HSXOF3AX, "          ' 品ＳＸＯＳＦ３平均上限
    sql = sql & "S.E020HSXOF3MX, "          ' 品ＳＸＯＳＦ３上限
    sql = sql & "S.E020HSXOF3HS, "          ' 品ＳＸＯＳＦ３ 保証方法＿処
    sql = sql & "S.E020HSXOF4AX, "          ' 品ＳＸＯＳＦ４平均上限
    sql = sql & "S.E020HSXOF4MX, "          ' 品ＳＸＯＳＦ４上限
    sql = sql & "S.E020HSXOF4HS, "          ' 品ＳＸＯＳＦ４ 保証方法＿処
    sql = sql & "S.E020HSXDENMX, "          ' 品ＳＸＤｅｎ上限
    sql = sql & "S.E020HSXDENMN, "          ' 品ＳＸＤｅｎ下限
    sql = sql & "S.E020HSXDENHS, "          ' 品ＳＸＤｅｎ保証方法＿処
    sql = sql & "S.E020HSXDVDMXN, "          ' 品ＳＸＤＶＤ２上限           項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDMNN, "          ' 品ＳＸＤＶＤ２下限           項目追加，修正対応 2003.05.20 yakimura
    sql = sql & "S.E020HSXDVDHS, "          ' 品ＳＸＤＶＤ２保証方法＿処
    sql = sql & "S.E020HSXLDLMX, "          ' 品ＳＸＬ／ＤＬ上限
    sql = sql & "S.E020HSXLDLMN, "          ' 品ＳＸＬ／ＤＬ下限
    sql = sql & "S.E020HSXLDLHS, "          ' 品ＳＸＬ／ＤＬ保証方法＿処
    sql = sql & "S.E019HSXLTMIN, "          ' 品ＳＸＬタイム下限
    sql = sql & "S.E019HSXLTMAX, "          ' 品ＳＸＬタイム上限
    sql = sql & "S.E019HSXLTHWS, "          ' 品ＳＸＬタイム保証方法＿処
    sql = sql & "S.E018HSXDPDIR, "          ' 品ＳＸ溝位置方位
    sql = sql & "S.E018HSXDPDRC, "          ' 品ＳＸ溝位置方向
    sql = sql & "S.E018HSXDWMIN, "          ' 品ＳＸ溝巾下限
    sql = sql & "S.E018HSXDWMAX, "          ' 品ＳＸ溝巾上限
    sql = sql & "S.E018HSXDDMIN, "          ' 品ＳＸ溝深下限
    sql = sql & "S.E018HSXDDMAX, "          ' 品ＳＸ溝深上限
    sql = sql & "S.E018HSXD1MIN, "          ' 品ＳＸ直径１下限
    sql = sql & "S.E018HSXD1MAX, "          ' 品ＳＸ直径１上限
    sql = sql & "S.E018HSXCTCEN, "          ' 品ＳＸ結晶面傾縦中心
    sql = sql & "S.E018HSXCYCEN, "          ' 品ＳＸ結晶面傾横中心
    sql = sql & "U.EPDUP "                  ' 結晶内側管理 EPD　上限
    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
    sql = sql & " where BH.E040BLOCKID='" & BlockID_in & "' "
    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
    sql = sql & " and U.HINBAN=BH.E041HINBAN "
    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
    sql = sql & " and U.FACTORY=BH.E041FACTORY "
    sql = sql & " and U.OPECOND=BH.E041OPECOND "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs.RecordCount = 0 Then
        ReDim records(0)
        rs.Close
        GoTo proc_exit
    End If
    
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            '品番管理
            .hinban = rs("E041HINBAN")              ' 品番
            .INGOTPOS = rs("E041INGOTPOS")          ' 結晶内開始位置
            .REVNUM = rs("E041REVNUM")              ' 製品番号改訂番号
            .factory = rs("E041FACTORY")            ' 工場
            .opecond = rs("E041OPECOND")            ' 操業条件
            .LENGTH = rs("E041LENGTH")              ' 長さ
            '製品仕様SXLデータ
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
            .HSXTMMAX = rs("E019HSXTMMAXN")           ' 品ＳＸ転位密度上限           項目追加，修正対応 2003.05.20 yakimura
            For n = 1 To 3
'''                .HSXBMnAN(n) = rs("E020HSXBM" & n & "AN") * 10 ' 品ＳＸＢＭＤn 平均下限
'''                .HSXBMnAX(n) = rs("E020HSXBM" & n & "AX") * 10 ' 品ＳＸＢＭＤn 平均上限
                .HSXBMnHS(n) = rs("E020HSXBM" & n & "HS")  ' 品ＳＸＢＭＤn 保証方法＿処
            Next
            For n = 1 To 4
'''                .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' 品ＳＸＯＳＦn 平均上限
'''                .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' 品ＳＸＯＳＦn 上限
                If IsNull(rs("E020HSXOF" & n & "AX")) = False Then .HSXOFnAX(n) = rs("E020HSXOF" & n & "AX")   ' 品ＳＸＯＳＦn 平均上限         '05/03/29 ooba NULL対応
                If IsNull(rs("E020HSXOF" & n & "MX")) = False Then .HSXOFnMX(n) = rs("E020HSXOF" & n & "MX")   ' 品ＳＸＯＳＦn 上限             '05/03/29 ooba NULL対応
                .HSXOFnHS(n) = rs("E020HSXOF" & n & "HS")   ' 品ＳＸＯＳＦn 保証方法＿処
            Next
            .HSXDENMX = rs("E020HSXDENMX")          ' 品ＳＸＤｅｎ上限
            .HSXDENMN = rs("E020HSXDENMN")          ' 品ＳＸＤｅｎ下限
            .HSXDENHS = rs("E020HSXDENHS")          ' 品ＳＸＤｅｎ保証方法＿処
            .HSXDVDMX = rs("E020HSXDVDMXN")          ' 品ＳＸＤＶＤ２上限        項目追加，修正対応 2003.05.20 yakimura
            .HSXDVDMN = rs("E020HSXDVDMNN")          ' 品ＳＸＤＶＤ２下限        項目追加，修正対応 2003.05.20 yakimura
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
'''            .HSXCTCEN = rs("E018HSXCTCEN")          ' 品ＳＸ結晶面傾縦中心
'''            .HSXCYCEN = rs("E018HSXCYCEN")          ' 品ＳＸ結晶面傾横中心
            If IsNull(rs("E018HSXCTCEN")) = False Then .HSXCTCEN = rs("E018HSXCTCEN")       ' 品ＳＸ結晶面傾縦中心      '05/03/29 ooba NULL対応
            If IsNull(rs("E018HSXCYCEN")) = False Then .HSXCYCEN = rs("E018HSXCYCEN")       ' 品ＳＸ結晶面傾横中心      '05/03/29 ooba NULL対応
            .EPDUP = rs("EPDUP")                    ' 結晶内側管理 EPD　上限
        End With
        rs.MoveNext
    Next
    rs.Close


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_fcmkc001f_Disp = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'2002/09/13 Add hitec)N.MATSUMOTO  Start
'概要      :分割結晶（不良内訳）から、指定した条件に該当するデータの行数を取得
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'　　　　　：strWhere     ,I  ,String           ,SELECT条件文
'　　　　　：intCnt       ,O  ,Integer          ,件数
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function SelCntXSDC4(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDC4 " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、エラー
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
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End


'2002/09/13 Add hitec)N.MATSUMOTO  Start
'概要      :分割結晶（品番）から、指定した条件に該当するデータの行数を取得
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'　　　　　：strWhere     ,I  ,String           ,SELECT条件文
'　　　　　：intCnt       ,O  ,Integer          ,件数
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function SelCntXSDCA(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDCA " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、エラー
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
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End

'2002/09/13 Add hitec)N.MATSUMOTO  Start
'概要      :分割結晶（SXL）から、指定した条件に該当するデータの行数を取得
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'　　　　　：strWhere     ,I  ,String           ,SELECT条件文
'　　　　　：intCnt       ,O  ,Integer          ,件数
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :

Public Function SelCntXSDCB(ByVal strWhere As String, ByRef intCnt As Integer) As FUNCTION_RETURN
    
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    
    sql = "      SELECT count(*) cnt "
    sql = sql & "  FROM XSDCB " & strWhere

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    '存在しない時、エラー
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
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    SelCntXSDCB = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/13 Add hitec)N.MATSUMOTO  End


'**********************************************
'　新DB構造体初期化
'　ADD hitec)N.MATSUMOTO
'**********************************************
Public Sub clearType()

    On Error Resume Next

    With Kihon  '基本情報
        .ALLSCRAP = ""
        .CNTHINNOW = 0
        .CNTHINOLD = 0
        .DIAMETER = 0
        .FURYOUMU = ""
        .NEWPROC = ""
        .NOWPROC = ""
        .STAFFID = ""
    End With
    
    With BlkOld      '分割結晶(ブロック)：前工程
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


'''''概要      :枚数計算関数
'''''ﾊﾟﾗﾒｰﾀ    :変数名          ,IO ,型        ,説明
'''''          :strBlockId      ,I  ,Integer   ,ブロックID
'''''          :intLen          ,I  ,Integer   ,長さ
'''''          :戻り値          ,O  ,Integer   ,WF枚数
'''''説明      :
'''''履歴      :2002/09/11 ADD hitec)N.MATSUMOTO
''''Public Function GetWfNum(ByVal strBlockID As String, ByVal intLen As Integer, ByRef intWfNum As Integer) As FUNCTION_RETURN
''''
''''    Dim rs      As OraDynaset
''''    Dim sql     As String
''''    Dim intRtn  As Integer
''''
''''    '' エラーハンドラの設定
''''    On Error GoTo proc_err
''''
''''    '製品仕様を得る
''''    sql = "SELECT S.E018HSXCTCEN,S.E018HSXCYCEN "
''''    sql = sql & " from VECME009 BH, VECME001 S, TBCME036 U "
''''    sql = sql & " where BH.E040BLOCKID='" & strBlockID & "' "
''''    sql = sql & " and S.E018HINBAN=BH.E041HINBAN "
''''    sql = sql & " and S.E018MNOREVNO=BH.E041REVNUM "
''''    sql = sql & " and S.E018FACTORY=BH.E041FACTORY "
''''    sql = sql & " and S.E018OPECOND=BH.E041OPECOND "
''''    sql = sql & " and U.HINBAN=BH.E041HINBAN "
''''    sql = sql & " and U.MNOREVNO=BH.E041REVNUM "
''''    sql = sql & " and U.FACTORY=BH.E041FACTORY "
''''    sql = sql & " and U.OPECOND=BH.E041OPECOND "
''''
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs.RecordCount = 0 Then
''''        rs.Close
''''        GoTo proc_exit
''''    End If
''''
''''    ' 品ＳＸ結晶面傾縦中心
''''    If IsNull(rs("E018HSXCTCEN")) = False Then
''''        HSXCTCEN = 0
''''    Else
''''        HSXCTCEN = rs("E018HSXCTCEN")
''''    End If
''''
''''    ' 品ＳＸ結晶面傾横中心
''''    If IsNull(rs("E018HSXCYCEN")) = False Then
''''        HSXCYCEN = 0
''''    Else
''''        HSXCYCEN = rs("E018HSXCYCEN")
''''    End If
''''
''''    If Left(Trim(strBlockID), 1) = "8" Then
''''        '購入単結晶の場合
''''        If DBDRV_getSEEDDEG(Trim(strBlockID), SEEDDEG) = FUNCTION_RETURN_FAILURE Then
''''            rs.Close
''''            GoTo proc_exit
''''        End If
''''    Else
''''''        'ブロック管理TBCME037から取得
''''''
''''''        '引き上げ結晶の場合
''''''        s = GetCodeField("SC", "28", Left$(blkInfo.SEED, 1), "INFO3")
''''''        If Left$(s, 1) = "4" Then
''''''            SEEDDEG = 4
''''''        Else
''''''            SEEDDEG = 0
''''''        End If
''''    End If
''''
''''    'WF枚数を計算し、値を返す
''''    intWfNum = calculateWfNum(intLen, SEEDDEG, HSXCTCEN, HSXCYCEN)
''''
''''    GetWfNum = FUNCTION_RETURN_SUCCESS
''''
''''proc_exit:
''''    '' 終了
''''    gErr.Pop
''''    Exit Function
''''
''''proc_err:
''''    '' エラーハンドラ
''''    Debug.Print "====== Error SQL ======"
''''    Debug.Print sql
''''    gErr.HandleError
''''    GetWfNum = FUNCTION_RETURN_FAILURE
''''    Resume proc_exit
''''
''''End Function

'概要      :WF枚数を計算する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :blkLen        ,I  ,Integer   ,ブロック長さ
'          :seedDeg       ,I  ,Integer   ,結晶のSEED傾き
'          :hinDegT       ,I  ,Double    ,品番傾き（縦）
'          :hinDegY       ,I  ,Double    ,品番傾き（横）
'          :戻り値        ,O  ,Integer   ,WF枚数
'説明      :
'履歴      :2001/8/30 作成  野村
Private Function calculateWfNum(ByVal BlkLen%, ByVal SEEDDEG%, ByVal hinDegT As Double, ByVal hinDegY As Double) As Integer
Dim hinDeg As Integer
Dim s As String
Dim WfCnt As Integer

    If Pitch = 0# Then
        calculateWfNum = 0
        Exit Function
    End If

    ''品番傾きを得る
    '結晶最終払い出し、品番傾きの求め方変更
    If (Abs(hinDegT) = 2.83) And (Abs(hinDegY) = 2.83) Then
        hinDeg = 4
    ElseIf (Abs(hinDegT) = 4) And (hinDegY = 0) Then
        hinDeg = 4
    ElseIf (hinDegT = 0) And (Abs(hinDegY) = 4) Then
        hinDeg = 4
    Else
        hinDeg = 0
    End If
    
    ''WF枚数を計算する
    If SEEDDEG = hinDeg Then
        '通常品の場合
        WfCnt = Format(((BlkLen - Loss0) / Pitch) + 0.4, "0")
    Else
        WfCnt = Format(((BlkLen * Mlt4 - Loss4) / Pitch) + 0.4, "0")
    End If
    If WfCnt < 0 Then WfCnt = 0
    calculateWfNum = WfCnt
End Function
'2002/09/11 ADD hitec)N.MATSUMOTO End


'概要      :工程実績登録処理を行う(在庫減情報：CW740,CW760用)
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :工程実績(XSDC3)に在庫減情報の登録処理を行う
'履歴      :2003/04/27  HITEC)会田：ＷＦ枚数はマップ位置ではなく画面から直接取得する
'                                  品番="Z"は不良にしない

Public Function XSDC3Proc2() As FUNCTION_RETURN

'   内部変数
    Dim i, j, k         As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String           'エラーメッセージ
    Dim Koutei          As typ_XSDC3_Update '工程実績
    Dim rsKCNTC         As OraDynaset       'レコードセット
                                                
    Dim wSTOCKINFO()    As typ_stock_info   '現在工程の情報
    Dim vGetData        As Variant          '画面取込用work
    Dim sOldHinban      As String           '旧品番
    Dim sNowHinban      As String           '現品番
    Dim vBlkId          As Variant          '画面取込用work
    Dim sOldBlkID       As String           '旧ブロックID
    Dim vREVNUM         As Variant          '画面取込用work
    Dim vFACTORY        As Variant          '画面取込用work
    Dim vOPE            As Variant          '画面取込用work
    Dim iREVNUM         As Integer          '製品改訂番号
    Dim sFACTORY        As String           '工場
    Dim sOPE            As String           '操業条件
    Dim sBlkId          As String           'ブロックID
    
    Dim iMapSt          As Integer          'マップ開始位置
    Dim iMapEd          As Integer          'マップ終了位置
    Dim bHinFlg         As Boolean          '品番比較用フラグ
    Dim lTMaisu         As Long             '合計枚数
    Dim iGetHinInpos    As Integer          '結晶内位置
    Dim oGamenSpd       As Object           '画面ID
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    'エラーハンドラの設定
    'On Error GoTo PROC_ERR
    On Error GoTo 0
    
    '初期設定
    XSDC3Proc2 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(0)
    ReDim wSTOCKINFO(0)
        
   '前工程長さ合計
    For i = 0 To Kihon.CNTHINOLD - 1
        If (Kihon.NOWPROC = "CW760") _
           And ((SIngotP > CLng(HinOld(i).INPOSCA)) Or (HinOld(i).INPOSCA >= EIngotP)) Then
                '処理なし
        Else
            ' 前工程枚数が0の時、処理終了
            If HinOld(i).GNMCA <= 0 Then
                XSDC3Proc2 = FUNCTION_RETURN_SUCCESS
                Exit Function
            End If
                
            ReDim Preserve STOCKINFO(UBound(STOCKINFO) + 1)  '配列の追加
            '不良､払いの初期設定
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
    
        
'最抜試指示画面から品番の払い出しと欠落をマップ位置項目から求める
'STOCKINFO配列に格納するがSTOCKINFOの品番はHinOldの品番の登録順序と一致しているとは限らない

    If Kihon.NOWPROC = "CW740" Then
        Set oGamenSpd = f_cmbc036_2.sprExamine    '抜試変更ｽﾌﾟﾚｯﾄﾞ
    Else
        Set oGamenSpd = f_cmbc039_3.sprExamine    '再抜試ｽﾌﾟﾚｯﾄﾞ
    End If
    '品番を1列追加したことによる列の変更-------start iida 2003/09/06
    With oGamenSpd
'        .GetText 31, 1, vBlkId          'ブロックID
        ''残存酸素検査項目追加による変更　04/01/09 ooba
'        .GetText 32, 1, vBlkId          'ブロックID
        'GD追加による変更　05/01/31 ooba
'        .GetText 33, 1, vBlkId          'ブロックID
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        .GetText 39, 1, vBlkId          'ブロックID
        sOldBlkID = CStr(Trim(vBlkId))
        For i = 1 To .MaxRows Step 2    'ｽﾌﾟﾚｯﾄﾞからﾃﾞｰﾀを入力(2行ずつ確認)
''''''            .GetText 28, i, vNukisiFlg
''''''            If (vNukisiFlg = "1") Then
            ' 元品番の登録
'            .GetText 32, i, vGetData    '古い品番取得
            ''残存酸素検査項目追加による変更　04/01/09 ooba
'            .GetText 33, i, vGetData    '古い品番取得
            'GD追加による変更　05/01/31 ooba
'            .GetText 34, i, vGetData    '古い品番取得
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
            .GetText 40, i, vGetData    '古い品番取得
            sOldHinban = Trim(CStr(vGetData))
            .GetText 2, i, vGetData     '新しい品番取得
            ' 品番が"Z"の時は新品番=旧品番
            If Trim(CStr(vGetData)) = "Z" Then
                sNowHinban = sOldHinban
            Else
                sNowHinban = Trim(CStr(vGetData))
            End If
            .GetText 5, i, vGetData     '結晶位置
            iGetHinInpos = val(vGetData)
            .GetText 6, i, vGetData     'マップ開始位置
            iMapSt = val(vGetData)
            .GetText 6, i + 1, vGetData 'マップ終了位置
            iMapEd = val(vGetData)
'            .GetText 31, i, vBlkId      'ブロックID
            ''残存酸素検査項目追加による変更　04/01/09 ooba
'            .GetText 32, i, vBlkId      'ブロックID
            'GD追加による変更　05/01/31 ooba
'            .GetText 33, i, vBlkId      'ブロックID
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
            .GetText 39, i, vBlkId      'ブロックID
            If vBlkId = "" Then         'ブロックがNULLだったら、前回のブロックを使用
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & sOldBlkID
            Else
                vBlkId = Mid(BlkNow.CRYNUMC2, 1, 9) & vBlkId
            End If
' VVVVV 2003/04/27 ALT BY HITEC)会田：良品枚数はマップ位置ではなくテーブルから取得する
            sBlkId = vBlkId
'''''            SXLCnt = iMapEd - iMapSt + 1        'マップ枚数
' ^^^^^ 2003/04/27 ALT BY HITEC)会田  END
            iREVNUM = gtSprWfMap(i).REVNUM      '製品改訂番号
            sFACTORY = gtSprWfMap(i).factory    '工場
            sOPE = gtSprWfMap(i).opecond        '操業条件
            
'''            For k = 0 To Kihon.CNTHINNOW - 1    '品番を比較し、該当ﾃﾞｰﾀの結晶内開始位置を取得
'''                If sNowHinban = HinNow(k).HINBCA Then
'''                    iGetHinInpos = HinNow(k).INPOSCA
'''                    Exit For
'''                End If
'''            Next
                              
            If (((Kihon.NOWPROC = "CW760") Or (Kihon.NOWPROC = "CW740")) And (vBlkId <> BlkNow.CRYNUMC2)) Or ((Kihon.NOWPROC = "CW760") And ((SIngotP > iGetHinInpos) Or (iGetHinInpos >= EIngotP))) Then
                '処理なし
            Else
' VVVVV 2003/04/27 ALT BY HITEC)会田：良品枚数はマップ位置ではなくテーブルから取得する
                sql = "SELECT COUNT(*) AS SXLCNT"
                sql = sql & " FROM TBCMY011 "
                sql = sql & " WHERE LOTID = '" & sBlkId & "'"
                sql = sql & " AND (WFSTA ='0' OR WFSTA = '1') "
                sql = sql & " AND BLOCKSEQ >= " & iMapSt & ""
                sql = sql & " AND BLOCKSEQ <= " & iMapEd & ""
                
                Debug.Print sql
                
                Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
                ''みつからなかったらエラー
                If rs.RecordCount = 0 Then
                    SXLCnt = 0
                Else ''見つかったら、良品枚数を取得する
                    SXLCnt = val(rs("SXLCNT"))
                End If
                Debug.Print SXLCnt
                
' ^^^^^ 2003/04/27 ALT BY HITEC)会田  END
                bHinFlg = False '既存の配列に同じ品番が登録されているかのﾌﾗｸﾞ
                'wSTOCKINFO()は1から開始
                For j = 1 To UBound(wSTOCKINFO)
                    If (wSTOCKINFO(j).hinban = sOldHinban) Then  '既に登録してある品番
                        bHinFlg = True
' VVVVV 2003/04/27 ALT BY HITEC)会田：品番="Z"は不良にしない
'''''                        If (sNowHinban = "Z") Then   'Z登録の時
'''''                            wSTOCKINFO(j).FURYOM = wSTOCKINFO(j).FURYOM + SXLCnt
'''''                        Else
                        wSTOCKINFO(j).HARAIM = wSTOCKINFO(j).HARAIM + SXLCnt
'''''                        End If
' ^^^^^ 2003/04/27 ALT BY HITEC)会田  END
                    End If
                Next j
    
                If (bHinFlg = False) Then   'wSTOCKINFO()の配列に品番が登録なかった時新規にwSTOCKINFO()に登録
                    ReDim Preserve wSTOCKINFO(UBound(wSTOCKINFO) + 1)  '配列の追加
                    wSTOCKINFO(UBound(wSTOCKINFO)).hinban = sOldHinban  '品番
                    wSTOCKINFO(UBound(wSTOCKINFO)).HARAIM = 0           '配列初期設定
                    wSTOCKINFO(UBound(wSTOCKINFO)).FURYOM = 0           '配列初期設定
' VVVVV 2003/04/27 ALT BY HITEC)会田：品番="Z"は不良にしない
'''''                    If (sNowHinban = "Z") Then   'Z登録の時
'''''                        wSTOCKINFO(UBound(wSTOCKINFO)).FURYOM = SXLCnt  '画面のﾃﾞｰﾀ
'''''                    Else
                    wSTOCKINFO(UBound(wSTOCKINFO)).HARAIM = SXLCnt  '画面のﾃﾞｰﾀ
'''''                    End If
' ^^^^^ 2003/04/27 ALT BY HITEC)会田  END
                    wSTOCKINFO(UBound(wSTOCKINFO)).REVNUM = iREVNUM     '画面のﾃﾞｰﾀ
                    wSTOCKINFO(UBound(wSTOCKINFO)).factory = sFACTORY   '画面のﾃﾞｰﾀ
                    wSTOCKINFO(UBound(wSTOCKINFO)).OPE = sOPE           '画面のﾃﾞｰﾀ
                End If
            End If
        Next i
    End With
    '品番を1列追加したことによる列の変更-------end iida 2003/09/06
    'STOCKINFとwSTOCKINFの突合せをしてSTOCKINFに品番、長さ、重量、枚数の払い出しと不良を格納
    'HinOldにないデータはなくなる
    'STOCKINFO()は添字0から開始
    'STOCKINFO()の品番はHinOldのﾃﾞｰﾀ、ﾃﾞｰﾀはHinNowのﾃﾞｰﾀ
    For i = 1 To UBound(STOCKINFO)
        STOCKINFO(i).HARAIM = 0
        STOCKINFO(i).FURYOM = 0
        For j = 1 To UBound(wSTOCKINFO)
            If (STOCKINFO(i).hinban = wSTOCKINFO(j).hinban) Then    '品番が等しい時登録する
                '''STOCKINFO(i).hinban = HinOld(i).HINBCA
                STOCKINFO(i).HARAIM = wSTOCKINFO(j).HARAIM
                STOCKINFO(i).FURYOM = wSTOCKINFO(j).FURYOM
'''''                STOCKINFO(i).KCKNT = HinOld(i).KCKNTCA  '連番はHinOldから取得
                STOCKINFO(i).REVNUM = wSTOCKINFO(j).REVNUM
                STOCKINFO(i).factory = wSTOCKINFO(j).factory
                STOCKINFO(i).OPE = wSTOCKINFO(j).OPE
                lTMaisu = wSTOCKINFO(j).HARAIM + wSTOCKINFO(j).FURYOM   '枚数の合計
                If (lTMaisu > 0) Then   '枚数があるか確認
                    STOCKINFO(i).HARAIW = wSTOCKINFO(j).HARAIM / lTMaisu * CLng(STOCKINFO(i).HARAIW)        '不良、払出は枚数の比率で算出
                    STOCKINFO(i).FuryoW = CLng(STOCKINFO(i).GENZAW) - STOCKINFO(i).HARAIW
                    STOCKINFO(i).HARAIL = wSTOCKINFO(j).HARAIM / lTMaisu * CLng(STOCKINFO(i).HARAIL)       '不良、払出は枚数の比率で算出
                    STOCKINFO(i).FURYOL = CLng(STOCKINFO(i).GENZAL) - STOCKINFO(i).HARAIL
                 End If
            End If
        Next j
    Next i
   
    '不良がある場合現在庫減情報の作成
    For i = 1 To UBound(STOCKINFO)
        If STOCKINFO(i).FURYOM > 0 Then
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    'ブロックＩＤ
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '位置
            Koutei.KCNTC3 = STOCKINFO(i).KCKNT + 1  '工程連番
            Koutei.HINBC3 = STOCKINFO(i).hinban     '品番
            Koutei.REVNUMC3 = STOCKINFO(i).REVNUM   '製品改訂番号
            Koutei.FACTORYC3 = STOCKINFO(i).factory '工場
            Koutei.OPEC3 = STOCKINFO(i).OPE         '操業条件
            Koutei.LENC3 = STOCKINFO(i).HARAIL      '長さ
            Koutei.XTALC3 = HinNow(0).XTALCA        '結晶番号
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1) ' 管理工程(現在工程+1)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' 工程
            Koutei.WKKBC3 = ""                      ' 作業区分
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' 処理回数
            Koutei.MODKBC3 = ""                     ' 赤黒区分
            Koutei.SUMKBC3 = ""                     ' 集計区分
            Koutei.FRKNKTC3 = ""                    ' (受入)管理工程
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (受入)作業区分
            If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
            
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (払出)工程
                Koutei.TOMACOC3 = "0"               '(払出)処理回数
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (払出)工程
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (払出)処理回
            End If
            Koutei.FRLC3 = STOCKINFO(i).GENZAL      ' 受入長さ
            Koutei.FRWC3 = STOCKINFO(i).GENZAW      '受入重量
            Koutei.FRMC3 = STOCKINFO(i).GENZAM      '受入枚数
            Koutei.FULC3 = STOCKINFO(i).FURYOL      '不良長さ
            Koutei.FUWC3 = STOCKINFO(i).FuryoW      '不良重量
            Koutei.FUMC3 = STOCKINFO(i).FURYOM      '不良枚数
            Koutei.LOSWC3 = ""                      ' ロス長さ
            
            Koutei.LOSLC3 = ""                      ' ロス重量
            Koutei.LOSMC3 = ""                      ' ロス枚数
            Koutei.TOLC3 = STOCKINFO(i).HARAIL      '払出長さ
            Koutei.TOWC3 = STOCKINFO(i).HARAIW      '払出重量
            Koutei.TOMC3 = STOCKINFO(i).HARAIM      '払出枚数
            Koutei.SUMITLC3 = ""                    ' SUMIT長さ
            Koutei.SUMITWC3 = ""                    ' SUMIT重量
            Koutei.SUMITMC3 = ""                    ' SUMIT枚数
            Koutei.MOTHINC3 = ""                    ' 振替品番(元)
            Koutei.XTWORKC3 = "42"                  ' 製造工場
            
            Koutei.WFWORKC3 = ""                    ' ｳｪｰﾊ製造
'           Koutei.STATIMEC3 = Null                 ' 処理開始終了
'           Koutei.STOTIMEC3 = Null                 ' 処理時間終了
'           Koutei.ETIMEC3 = ""                     ' 実績時間は入れない
            Koutei.HOLDCC3 = " "                    ' ホールドコード
            Koutei.HOLDBC3 = "0"                    ' ホールド区分
            Koutei.LDFRCC3 = ""                     ' 格下コード
            Koutei.LDFRBC3 = "0"                    ' 格下区分（ハイキ）
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' 登録社員ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
            
            Koutei.KSTAFFC3 = ""                    ' 更新社員ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            Koutei.SUMITBC3 = ""                    ' SUMIT送信フラグ
            Koutei.SNDKC3 = ""                      ' 送信フラグ
'           Koutei.SNDDAYC3 = ""                    ' 送信日付
            Koutei.MODMACOC3 = ""                   ' 赤黒の処理回数
            Koutei.KAKUCC3 = ""                     ' 確定コード
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO時間
            Koutei.PAYCLASSC3 = ""                  '　転送先工場フラグ
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT送信日付
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '工程実績に在庫減情報登録
            If iRtn = FUNCTION_RETURN_FAILURE Then  '工程実績追加エラー
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

''''
'HinInf():操作する配列
'HinNum:操作する配列位置
'HinFlg:-1なら配列削除 1なら配列追加

'概要      :工程実績登録処理を行う(品番振替情報：CW740,CW760用)
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :工程実績(XSDC3)に品番振替情報の登録処理を行う

Public Function XSDC3Proc3() As FUNCTION_RETURN

'   内部変数
    Dim i, j            As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '工程実績
    
    Dim wLen            As Long
    Dim wCHKPOS         As Long
        
    Dim wOINF()         As typ_trans_info   '前品番並び替え用
    Dim wNINF()         As typ_trans_info   '後品番並び替え用
    Dim wWINF()         As typ_trans_info   '並び替え用ワーク  ' 2003/04/17 add by t.t
    Dim ibuf            As Integer
    Dim wOINFrecCnt     As Integer
    Dim wNINFrecCnt     As Integer
    Dim wOINFFLG        As Integer
    Dim wNINFFLG        As Integer
    Dim iCnt             As Integer
    Dim wNINFMAX        As Integer           ' 2003/04/17 add by t.t
    Dim wOINFMAX        As Integer           ' 2003/04/17 add by t.t
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    '初期設定
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE

'' 2003/04/17 add by t.t   start
    ReDim wOINF(UBound(STOCKINFO))      '品番ごとにソート用
    ReDim wNINF(Kihon.CNTHINNOW)        '品番ごとにソート用
    ReDim wWINF(1)                      '品番ごとにソート用
        
    ' 在庫減情報より取り込み
    For i = 1 To UBound(STOCKINFO)
        If Trim(STOCKINFO(i).hinban) <> "" Then
            wOINF(i).hinban = STOCKINFO(i).hinban
            wOINF(i).LEN = STOCKINFO(i).HARAIL       ' 前工程長さ合計
            wOINF(i).WAT = STOCKINFO(i).HARAIW       ' 前工程重量合計
            wOINF(i).MAI = STOCKINFO(i).HARAIM       ' 前工程枚数合計
        End If
    Next i
    
    ' 良品情報より取り込み
    For j = 0 To Kihon.CNTHINNOW - 1
        If (Kihon.NOWPROC = "CW760") _
            And ((SIngotP > HinNow(j).INPOSCA) Or (HinNow(j).INPOSCA >= EIngotP)) Then
                '処理なし
        Else
            wNINF(j).hinban = HinNow(j).HINBCA
            '04/01/13 ooba 追加 START ===================>
            wNINF(j).REVNUM = HinNow(j).REVNUMCA
            wNINF(j).factory = HinNow(j).FACTORYCA
            wNINF(j).OPE = HinNow(j).OPECA
            '04/01/13 ooba 追加 END =====================>
            wNINF(j).LEN = HinNow(j).GNLCA           ' 後工程長さ合計
            wNINF(j).WAT = HinNow(j).GNWCA           ' 後工程重量合計
            wNINF(j).MAI = HinNow(j).GNMCA           ' 後工程枚数合計
            wNINF(j).KCKNT = HinNow(j).KCKNTCA       ' 後工程連番
        End If
    Next j
        
    '同じ品番同士、長さ等を打ち消す
    For i = 1 To UBound(STOCKINFO)
        For j = 0 To Kihon.CNTHINNOW - 1
           If wOINF(i).hinban = wNINF(j).hinban Then
                '小さい方の数字を両方から引く
                If wOINF(i).MAI <= wNINF(j).MAI Then
                    wNINF(j).LEN = wNINF(j).LEN - wOINF(i).LEN
                    wNINF(j).WAT = wNINF(j).WAT - wOINF(i).WAT
                    wNINF(j).MAI = wNINF(j).MAI - wOINF(i).MAI
                    wOINF(i).LEN = 0
                    wOINF(i).WAT = 0
                    wOINF(i).MAI = 0
                Else
                    wOINF(i).LEN = wOINF(i).LEN - wNINF(j).LEN
                    wOINF(i).WAT = wOINF(i).WAT - wNINF(j).WAT
                    wOINF(i).MAI = wOINF(i).MAI - wNINF(j).MAI
                    If wOINF(i).MAI < 0 Then
                        wOINF(i).MAI = 0
                    End If
                    wNINF(j).LEN = 0
                    wNINF(j).WAT = 0
                    wNINF(j).MAI = 0
                End If
            End If
        Next
    Next
        
    For i = 0 To UBound(wOINF) - 2
        For j = i + 1 To UBound(wOINF) - 1
            If (StrComp(wOINF(i).hinban, wOINF(j).hinban, _
                vbTextCompare)) = 1 Then '品番の入替必要
                wWINF(0) = wOINF(j)
                wOINF(j) = wOINF(i)
                wOINF(i) = wWINF(0)
            End If
        Next j
    Next i
'' 2003/04/17 add by t.t   end

'' 2003/04/17 add by t.t   start
    'wNINFの品番をソートする
    For i = 0 To UBound(wNINF) - 2
        For j = i + 1 To UBound(wNINF) - 1
            If (StrComp(wNINF(i).hinban, wNINF(j).hinban, _
                vbTextCompare)) = 1 Then '品番の入替必要
                wWINF(0) = wNINF(j)
                wNINF(j) = wNINF(i)
                wNINF(i) = wWINF(0)
            End If
        Next j
    Next i
'' 2003/04/17 add by t.t   end

    '空きの配列削除する(配列のデータを詰める)
    For i = 0 To wOINFMAX
        If wOINF(i).MAI <= 0 Then
            iCnt = i
            Call HairetuOpe_Mai(wOINF(), iCnt, -1)
        End If
    Next i

    '空きの配列削除する(配列のデータを詰める)
    For i = 0 To wNINFMAX
        If wNINF(i).MAI <= 0 Then
            iCnt = i
            Call HairetuOpe_Mai(wNINF(), iCnt, -1)
        End If
    Next i
    
    '品番入替情報を作成する
    i = 0 '前品番の位置
    j = 0 '後品番の位置
    Do
        '枚数を突き合わせて数量が同じでなかったら大きい値の品番を分割する
        If (wOINF(i).MAI = wNINF(j).MAI) Then   '品番長さが同じ時両方とも次に進む
        ElseIf (wOINF(i).MAI > wNINF(j).MAI) Then   '品番長さが異なる時
            iCnt = i
            Call HairetuOpe(wOINF(), iCnt, 1)    '配列の追加
            wOINF(i + 1).hinban = wOINF(i).hinban
            wOINF(i + 1).LEN = wOINF(i).LEN - wNINF(j).LEN
            wOINF(i + 1).WAT = wOINF(i).WAT - wNINF(j).WAT
            wOINF(i + 1).MAI = wOINF(i).MAI - wNINF(j).MAI
            wOINF(i).LEN = wNINF(j).LEN
            wOINF(i).WAT = wNINF(j).WAT
            wOINF(i).MAI = wNINF(j).MAI
'''''        ElseIf (wOINF(i).LEN < wNINF(j).LEN) Then   '品番数量が異なる時   '2003/04/17 rep by tt
        ElseIf (wOINF(i).MAI < wNINF(j).MAI) Then   '品番数量が異なる時
            iCnt = j
            Call HairetuOpe(wNINF(), iCnt, 1)
            wNINF(j + 1).hinban = wNINF(i).hinban
            wNINF(j + 1).LEN = wNINF(j).LEN - wOINF(i).LEN
            wNINF(j + 1).WAT = wNINF(j).WAT - wOINF(i).WAT
            wNINF(j + 1).MAI = wNINF(j).MAI - wOINF(i).MAI
            wNINF(j).LEN = wOINF(i).LEN
            wNINF(j).WAT = wOINF(i).WAT
            wNINF(j).MAI = wOINF(i).MAI
        End If
        wOINFrecCnt = UBound(wOINF())
        wNINFrecCnt = UBound(wNINF())
        i = i + 1
        j = j + 1
        If (i > wOINFrecCnt) Then
            Exit Do
        
        End If
        If (j > wNINFrecCnt) Then
            Exit Do
        End If
        
'''''        If (wOINF(i).LEN) <= 0 Then   '2003/04/17 rep by tt
        If (wOINF(i).MAI) <= 0 Then
            Exit Do
        End If
        
'''''        If (wNINF(j).LEN) <= 0 Then   '2003/04/17 rep by tt
        If (wNINF(j).MAI) <= 0 Then
            Exit Do
        End If
    Loop
    
    wOINFrecCnt = UBound(wOINF())
'    For i = 0 To wOINFrecCnt - 1 'Step 1
    For i = 0 To wOINFrecCnt     '04/01/13 ooba
        If (StrComp(wNINF(i).hinban, wOINF(i).hinban, vbTextCompare) <> 0) Then  '品番が異なる時振替情報に登録する
            If Trim(wNINF(i).hinban) <> "" And wNINF(i).LEN > 0 Then
                Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    'ブロックＩＤ
                giInpos = giInpos + 1
                Koutei.INPOSC3 = giInpos                '位置
                Koutei.KCNTC3 = wNINF(i).KCKNT          ' 工程連番
                Koutei.HINBC3 = wNINF(i).hinban         '品番
'''                Koutei.REVNUMC3 = HinNow(i).REVNUMCA    '製品改訂番号
'''                Koutei.FACTORYC3 = HinNow(i).FACTORYCA  '工場
'''                Koutei.OPEC3 = HinNow(i).OPECA          '操業条件
                Koutei.REVNUMC3 = wNINF(i).REVNUM       '製品改訂番号       ''04/01/13 ooba
                Koutei.FACTORYC3 = wNINF(i).factory     '工場               ''04/01/13 ooba
                Koutei.OPEC3 = wNINF(i).OPE             '操業条件           ''04/01/13 ooba
                Koutei.LENC3 = wNINF(i).LEN             '受入長さ
                Koutei.XTALC3 = HinNow(0).XTALCA        '結晶番号
                Koutei.SXLIDC3 = ""                     ' SXLID
                
                Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
                  CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2) ' 管理工程(現在工程+2)
                Koutei.WKKTC3 = Kihon.NOWPROC           ' 工程
                Koutei.WKKBC3 = ""                      ' 作業区分
                Koutei.MACOC3 = HinNow(0).NEMACOCA      ' 処理回数
                Koutei.MODKBC3 = ""                     ' 赤黒区分
                Koutei.SUMKBC3 = ""                     ' 集計区分
                Koutei.FRKNKTC3 = ""                    ' (受入)管理工程
                If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                    Koutei.FRWKKTC3 = ""
                Else
                    Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                    'add end 2003/03/28 hitec)matsumoto --------
                End If
                Koutei.FRWKKBC3 = ""                    ' (受入)作業区分
                If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                    Koutei.FRMACOC3 = "0"
                Else
                    Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
                End If
                
''''                If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''                If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
                Select Case Kihon.NOWPROC
                Case "CC730"
                    iHantei = CInt(BlkNow.GNLC2)
                Case Else
                    iHantei = CInt(BlkNow.GNMC2)
                End Select
                If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                    Koutei.TOWKKTC3 = " "               ' (払出)工程
                    Koutei.TOMACOC3 = "0"               '(払出)処理回数
                Else
'''                    Koutei.TOWKKTC3 = HinNow(i).GNWKNTCA    ' (払出)工程
'''                    Koutei.TOMACOC3 = HinNow(i).GNMACOCA    ' (払出)処理回
                    Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (払出)工程        ''04/01/13 ooba
                    Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (払出)処理回数    ''04/01/13 ooba
                End If
                Koutei.FRLC3 = wNINF(i).LEN             '受入長さ
                Koutei.FRWC3 = wNINF(i).WAT             '受入重量
                Koutei.FRMC3 = wNINF(i).MAI             '受入枚数
                Koutei.FULC3 = 0                        '不良長さ
                Koutei.FUWC3 = 0                        '不良重量
                Koutei.FUMC3 = 0                        '不良枚数
                Koutei.LOSWC3 = ""                      ' ロス長さ
                
                Koutei.LOSLC3 = ""                      ' ロス重量
                Koutei.LOSMC3 = ""                      ' ロス枚数
                Koutei.TOLC3 = wNINF(i).LEN             '払出長さ
'''                Koutei.TOWC3 = wNINF(0).WAT             '払出重量
'''                Koutei.TOMC3 = wNINF(0).MAI             '払出枚数
                Koutei.TOWC3 = wNINF(i).WAT             '払出重量           ''04/01/13 ooba
                Koutei.TOMC3 = wNINF(i).MAI             '払出枚数           ''04/01/13 ooba
                Koutei.SUMITLC3 = ""                    ' SUMIT長さ
                Koutei.SUMITWC3 = ""                    ' SUMIT重量
                Koutei.SUMITMC3 = ""                    ' SUMIT枚数
                Koutei.MOTHINC3 = wOINF(i).hinban       '元品番
                Koutei.XTWORKC3 = "42"                  ' 製造工場
                
                Koutei.WFWORKC3 = ""                    ' ｳｪｰﾊ製造
    '           Koutei.STATIMEC3 = Null                 ' 処理開始終了
    '           Koutei.STOTIMEC3 = Null                 ' 処理時間終了
    '           Koutei.ETIMEC3 = ""                     ' 実績時間は入れない
                Koutei.HOLDCC3 = " "                    ' ホールドコード
                Koutei.HOLDBC3 = "0"                    ' ホールド区分
                Koutei.LDFRCC3 = ""                     ' 格下コード
                Koutei.LDFRBC3 = "0"                    ' 格下区分（ハイキ）
                Koutei.TSTAFFC3 = Kihon.STAFFID         ' 登録社員ID
                Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
                
                Koutei.KSTAFFC3 = ""                    ' 更新社員ID
                Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
                Koutei.SUMITBC3 = ""                    ' SUMIT送信フラグ
                Koutei.SNDKC3 = ""                      ' 送信フラグ
    '           Koutei.SNDDAYC3 = ""                    ' 送信日付
                Koutei.MODMACOC3 = ""                   ' 赤黒の処理回数
                Koutei.KAKUCC3 = ""                     ' 確定コード
                Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO時間
                Koutei.PAYCLASSC3 = ""                  '　転送先工場フラグ
    '            Koutei.SUMITSNDC3 = ""                  ' SUMIT送信日付
                
    '            Koutei.SSENDNOC3 = ""
               
                iRtn = CreateXSDC3(Koutei, wErrMsg)     '工程実績に在庫減情報登録
                If iRtn = FUNCTION_RETURN_FAILURE Then  '工程実績追加エラー
                    MsgBox wErrMsg
                    Exit Function
                End If
            End If
        End If
    Next i

    XSDC3Proc3 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc3 = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'概要      :工程実績登録処理を行う(在庫減情報：CC730用)
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :工程実績(XSDC3)に在庫減情報の登録処理を行う

Public Function XSDC3Proc4() As FUNCTION_RETURN

'   内部変数
    Dim i, j            As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '工程実績
    Dim rsKCNTC         As OraDynaset       'レコードセット
    Dim intNextCnt      As Integer
    
    Dim wLen As Long
    Dim wCHKPOS As Long
    Dim badcnt As Integer
    Dim BADINFO() As typ_bad_info
    Dim wSTOCKINFO() As typ_stock_info
    Dim iLoopCnt As Integer
    Dim vGetMaxPos  As Variant
    Dim vGetData  As Variant
    Dim sOldHinban, sNowHinban As String
    Dim iSXLcnt As Integer
    Dim iMapSt, iMapEd As Integer
    Dim bHinFlg As Boolean
    Dim lTMaisu As Long
    Dim vNukisiFlg  As Variant
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    'エラーハンドラの設定
    'On Error GoTo PROC_ERR
    On Error GoTo 0
    
    '初期設定
    XSDC3Proc4 = FUNCTION_RETURN_FAILURE

    ReDim STOCKINFO(Kihon.CNTHINOLD)
    ReDim wSTOCKINFO(0)
    
   'HinOldから前工程長さ,重量,枚数合計取得(長さは0)
    For i = 0 To Kihon.CNTHINOLD - 1
        FRLC3Sum = FRLC3Sum + CLng(HinOld(i).GNLCA)    ' 前工程長さ合計
        FRWC3Sum = FRWC3Sum + CLng(HinOld(i).GNWCA)    ' 前工程重量合計
        FRMC3Sum = FRMC3Sum + CLng(HinOld(i).GNMCA)    ' 前工程枚数合計
        '不良､払いの初期設定
        STOCKINFO(i).hinban = HinOld(i).HINBCA
        STOCKINFO(i).FURYOL = 0
        STOCKINFO(i).HARAIL = CLng(HinOld(i).GNLCA)
        STOCKINFO(i).FuryoW = CLng(HinOld(i).GNWCA) '不良重量に払い重量を仮に代入して後で計算する
        STOCKINFO(i).HARAIW = CLng(HinOld(i).GNWCA)
        STOCKINFO(i).FURYOM = CLng(HinOld(i).GNMCA) '不良枚数に払い枚数を仮に代入し後で計算する
        STOCKINFO(i).HARAIM = CLng(HinOld(i).GNMCA)
        STOCKINFO(i).KCKNT = CLng(HinOld(i).KCKNTCA)
        STOCKINFO(i).REVNUM = HinOld(i).REVNUMCA        ' 製品改訂番号
        STOCKINFO(i).factory = HinOld(i).FACTORYCA      ' 工場
        STOCKINFO(i).OPE = HinOld(i).OPECA              ' 製品改訂番号
    Next i
        
'最抜試指示画面から品番の払い出しと欠落をマップ位置項目から求める
'STOCKINFO配列に格納するがSTOCKINFOの品番はHinOldの品番の登録順序と一致しているとは限らない
    
    badcnt = 0  '不良数初期設定
'    '不良が先頭にないか確認
    If ((CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)) > 0) Then '前後開始位置を比較して差があれば不良位置登録
        badcnt = badcnt + 1
        ReDim Preserve BADINFO(badcnt)
        BADINFO(badcnt).pos = CLng(HinOld(0).INPOSCA)
        BADINFO(badcnt).LEN = CLng(HinNow(0).INPOSCA) - CLng(HinOld(0).INPOSCA)
    End If
'
    '不良長さが品番間にないか確認
    For i = 0 To Kihon.CNTHINNOW - 2
        If (CLng(HinNow(i + 1).INPOSCA) > (CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA))) Then '品番間に不良有
            badcnt = badcnt + 1 '不良位置の登録
            ReDim Preserve BADINFO(badcnt)
            BADINFO(badcnt).pos = CLng(HinNow(i).INPOSCA) + CLng(HinNow(i).GNLCA)
            BADINFO(badcnt).LEN = CLng(HinNow(i + 1).INPOSCA) - CLng(HinNow(i).INPOSCA) - CLng(HinNow(i).GNLCA)
        End If
    Next i
'
    '不良が最後にないか前の確認(結晶内開始位置+長さで比較)
    If ((CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA)) _
        <> (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))) Then '終了位置の確認  'upd 2003/05/31 hitec)matsumoto 「＞」を「＜＞」に変更
        badcnt = badcnt + 1
        ReDim Preserve BADINFO(badcnt)
        BADINFO(badcnt).pos = (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
        BADINFO(badcnt).LEN = CLng(HinOld(Kihon.CNTHINOLD - 1).INPOSCA) + CLng(HinOld(Kihon.CNTHINOLD - 1).GNLCA) - (CLng(HinNow(Kihon.CNTHINNOW - 1).INPOSCA) + CLng(HinNow(Kihon.CNTHINNOW - 1).GNLCA))
    End If
'
    If (badcnt = 0) Then  '前と後で不良なし 処理終了
        XSDC3Proc4 = FUNCTION_RETURN_SUCCESS
        Exit Function
    End If

'    '不良位置を振替前の結晶位置を確認して不良位置に相当する品番を登録する
            
    'add start 2003/05/31 hitec)matsumoto --------------
    If BlkOld.GNLC2 < BlkNow.GNLC2 Then
        For i = 1 To badcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + BADINFO(i).LEN   '不良の長さ(HinOld(i)のﾁｪｯｸしている品番の長さ不良)
                STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - BADINFO(i).LEN  '良品長さ
            Next j
        Next i
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then  '不良数が存在したら不良重さを不良比率で求める
                If i = Kihon.CNTHINOLD - 1 Then
                    'STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW 'STOCKINFO(i).HARAIWは入力済み
                    STOCKINFO(i).HARAIW = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIW)  'STOCKINFO(i).HARAIWは入力済み  ’2003/08/06 hitec)matsumoto ROUND追加
                    STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW 'STOCKINFO(i).FURYOWは入力済み
                    'STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM 'STOCKINFO(i).HARAIMは入力済み
                    STOCKINFO(i).HARAIM = Round((STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL)) * STOCKINFO(i).HARAIM)  'STOCKINFO(i).HARAIMは入力済み  ’2003/08/06 hitec)matsumoto ROUND追加
                    STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM 'STOCKINFO(i).FURYOMは入力済み
                Else
                    STOCKINFO(i).HARAIW = HinOld(i).GNWCA
                    STOCKINFO(i).FuryoW = 0
                    STOCKINFO(i).HARAIM = HinOld(i).GNMCA
                    STOCKINFO(i).FURYOM = 0
                End If
            End If
        Next i
    Else
        'STOCKINFOの払いは既に入力済み
        For i = 1 To badcnt
            For j = 0 To Kihon.CNTHINOLD - 1
                If (BADINFO(i).pos >= CLng(HinOld(j).INPOSCA) And _
                    BADINFO(i).pos < CLng(HinOld(j).INPOSCA) + CLng(HinOld(j).GNLCA)) Then
                    STOCKINFO(j).FURYOL = STOCKINFO(j).FURYOL + BADINFO(i).LEN   '不良の長さ(HinOld(i)のﾁｪｯｸしている品番の長さ不良)
                    STOCKINFO(j).HARAIL = STOCKINFO(j).HARAIL - BADINFO(i).LEN  '良品長さ
                End If
            Next j
        Next i
    
        '重量と枚数の不良と払い出しの値を設定する
        For i = 0 To Kihon.CNTHINOLD - 1
            If ((STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) > 0) Then  '不良数が存在したら不良重さを不良比率で求める
                STOCKINFO(i).HARAIW = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIW 'STOCKINFO(i).HARAIWは入力済み
                STOCKINFO(i).FuryoW = STOCKINFO(i).FuryoW - STOCKINFO(i).HARAIW 'STOCKINFO(i).FURYOWは入力済み
                STOCKINFO(i).HARAIM = STOCKINFO(i).HARAIL / (STOCKINFO(i).HARAIL + STOCKINFO(i).FURYOL) * STOCKINFO(i).HARAIM 'STOCKINFO(i).HARAIMは入力済み
                STOCKINFO(i).FURYOM = STOCKINFO(i).FURYOM - STOCKINFO(i).HARAIM 'STOCKINFO(i).FURYOMは入力済み
            End If
        Next i
    End If
    'add end 2003/05/31 hitec)matsumoto --------------
    
    '不良がある場合現在庫減情報の作成
    For i = 0 To Kihon.CNTHINOLD - 1
        If STOCKINFO(i).FURYOL <> 0 Then
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    'ブロックＩＤ
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '位置
            Koutei.KCNTC3 = STOCKINFO(i).KCKNT + 1  '工程連番
            Koutei.HINBC3 = HinOld(i).HINBCA        '品番
            Koutei.REVNUMC3 = HinOld(i).REVNUMCA    '製品改訂番号
            Koutei.FACTORYC3 = HinOld(i).FACTORYCA  '工場
            Koutei.OPEC3 = HinOld(i).OPECA          '操業条件
            Koutei.LENC3 = STOCKINFO(i).HARAIL      '長さ
            Koutei.XTALC3 = HinOld(i).XTALCA        '結晶番号
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 1) ' 管理工程(現在工程+1)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' 工程
            Koutei.WKKBC3 = ""                      ' 作業区分
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' 処理回数
            Koutei.MODKBC3 = ""                     ' 赤黒区分
            Koutei.SUMKBC3 = ""                     ' 集計区分
            Koutei.FRKNKTC3 = ""                    ' (受入)管理工程
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (受入)作業区分
            If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (払出)工程
                Koutei.TOMACOC3 = "0"               '(払出)処理回数
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (払出)工程
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (払出)処理回
            End If
            Koutei.FRLC3 = HinOld(i).GNLCA          '受入長さ
            Koutei.FRWC3 = HinOld(i).GNWCA          '受入重量
            Koutei.FRMC3 = HinOld(i).GNMCA          '受入枚数
            Koutei.FULC3 = STOCKINFO(i).FURYOL       '不良長さ
            Koutei.FUWC3 = STOCKINFO(i).FuryoW       '不良重量
            Koutei.FUMC3 = STOCKINFO(i).FURYOM       '不良枚数
            Koutei.LOSWC3 = ""                      ' ロス長さ
            
            Koutei.LOSLC3 = ""                      ' ロス重量
            Koutei.LOSMC3 = ""                      ' ロス枚数
            Koutei.TOLC3 = STOCKINFO(i).HARAIL       '払出長さ
            Koutei.TOWC3 = STOCKINFO(i).HARAIW       '払出重量
            Koutei.TOMC3 = STOCKINFO(i).HARAIM       '払出枚数
            Koutei.SUMITLC3 = ""                    ' SUMIT長さ
            Koutei.SUMITWC3 = ""                    ' SUMIT重量
            Koutei.SUMITMC3 = ""                    ' SUMIT枚数
            Koutei.MOTHINC3 = " "       '元品番
            Koutei.XTWORKC3 = "42"                  ' 製造工場
            
            Koutei.WFWORKC3 = ""                    ' ｳｪｰﾊ製造
'           Koutei.STATIMEC3 = Null                 ' 処理開始終了
'           Koutei.STOTIMEC3 = Null                 ' 処理時間終了
'           Koutei.ETIMEC3 = ""                     ' 実績時間は入れない
            Koutei.HOLDCC3 = " "                    ' ホールドコード
            Koutei.HOLDBC3 = "0"                    ' ホールド区分
            Koutei.LDFRCC3 = ""                     ' 格下コード
            Koutei.LDFRBC3 = "0"                    ' 格下区分（ハイキ）
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' 登録社員ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
            
            Koutei.KSTAFFC3 = ""                    ' 更新社員ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            Koutei.SUMITBC3 = ""                    ' SUMIT送信フラグ
            Koutei.SNDKC3 = ""                      ' 送信フラグ
'           Koutei.SNDDAYC3 = ""                    ' 送信日付
            Koutei.MODMACOC3 = ""                   ' 赤黒の処理回数
            Koutei.KAKUCC3 = ""                     ' 確定コード
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO時間
            Koutei.PAYCLASSC3 = ""                  '　転送先工場フラグ
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT送信日付
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '工程実績に在庫減情報登録
            If iRtn = FUNCTION_RETURN_FAILURE Then  '工程実績追加エラー
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

'概要      :工程実績登録処理を行う(品番振替情報：CC730用)
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :工程実績(XSDC3)に品番振替情報の登録処理を行う

Public Function XSDC3Proc5() As FUNCTION_RETURN

'   内部変数
    Dim i, j            As Integer
    Dim iRtn            As Integer          '復帰情報
    Dim sql             As String           'ＳＱＬ
    Dim sqlWhere        As String           'WHERE句
    Dim wErrMsg         As String
    Dim Koutei          As typ_XSDC3_Update    '工程実績
    
    Dim wLen            As Long
    Dim wCHKPOS         As Long
        
    Dim wOINF()         As typ_trans_info   '前品番並び替え用
    Dim wNINF()         As typ_trans_info   '後品番並び替え用
    Dim wWINF()         As typ_trans_info   '並び替え用ワーク
    Dim ibuf            As Integer
    Dim wOINFrecCnt     As Integer
    Dim wNINFrecCnt     As Integer
    Dim wOINFFLG        As Integer
    Dim wNINFFLG        As Integer
    Dim iPoint          As Integer
    Dim wNINFMAX        As Integer
    Dim wOINFMAX        As Integer
    Dim iHantei         As Integer          'add 2003/05/27 hitec)matsumoto

    'エラーハンドラの設定
    On Error GoTo proc_err
    
    '初期設定
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
    
    ReDim wOINF(UBound(STOCKINFO))      '品番ごとにソート用
    ReDim wNINF(Kihon.CNTHINNOW)        '品番ごとにソート用
    ReDim wWINF(1)                      '品番ごとにソート用
        
    ' 在庫減情報より取り込み
    For i = 0 To UBound(STOCKINFO) - 1
        wOINF(i).hinban = STOCKINFO(i).hinban
        wOINF(i).LEN = STOCKINFO(i).HARAIL       ' 前工程長さ合計
        wOINF(i).WAT = STOCKINFO(i).HARAIW       ' 前工程重量合計
        wOINF(i).MAI = STOCKINFO(i).HARAIM       ' 前工程枚数合計
    Next i
    
    ' 良品情報より取り込み
    For j = 0 To Kihon.CNTHINNOW - 1
        wNINF(j).hinban = HinNow(j).HINBCA
        wNINF(j).LEN = HinNow(j).GNLCA           ' 後工程長さ合計
        wNINF(j).WAT = HinNow(j).GNWCA           ' 後工程重量合計
        wNINF(j).MAI = HinNow(j).GNMCA           ' 後工程枚数合計
        wNINF(j).KCKNT = HinNow(j).KCKNTCA       ' 後工程連番
    Next j
        
    '同じ品番同士、長さ等を打ち消す
    For i = 0 To UBound(STOCKINFO) - 1
        For j = 0 To Kihon.CNTHINNOW - 1
           If wOINF(i).hinban = wNINF(j).hinban Then
                '小さい方の数字を両方から引く
                If wOINF(i).LEN <= wNINF(j).LEN Then
                    wNINF(j).LEN = wNINF(j).LEN - wOINF(i).LEN
                    wNINF(j).WAT = wNINF(j).WAT - wOINF(i).WAT
                    wNINF(j).MAI = wNINF(j).MAI - wOINF(i).MAI
                    wOINF(i).LEN = 0
                    wOINF(i).WAT = 0
                    wOINF(i).MAI = 0
                Else
                    wOINF(i).LEN = wOINF(i).LEN - wNINF(j).LEN
                    wOINF(i).WAT = wOINF(i).WAT - wNINF(j).WAT
                    wOINF(i).MAI = wOINF(i).MAI - wNINF(j).MAI
                    If wOINF(i).MAI < 0 Then
                        wOINF(i).MAI = 0
                    End If
                    wNINF(j).LEN = 0
                    wNINF(j).WAT = 0
                    wNINF(j).MAI = 0
                End If
            End If
        Next
    Next
        
    For i = 0 To UBound(wOINF) - 2
        For j = i + 1 To UBound(wOINF) - 1
            If (StrComp(wOINF(i).hinban, wOINF(j).hinban, _
                vbTextCompare)) = 1 Then '品番の入替必要
                wWINF(0) = wOINF(j)
                wOINF(j) = wOINF(i)
                wOINF(i) = wWINF(0)
            End If
        Next j
    Next i
    
    'wNINFの品番をソートする
    For i = 0 To UBound(wNINF) - 2
        For j = i + 1 To UBound(wNINF) - 1
            If (StrComp(wNINF(i).hinban, wNINF(j).hinban, _
                vbTextCompare)) = 1 Then '品番の入替必要
                wWINF(0) = wNINF(j)
                wNINF(j) = wNINF(i)
                wNINF(i) = wWINF(0)
            End If
        Next j
    Next i
        
        
    '空きの配列削除する(配列のデータを詰める)
    For i = 0 To wOINFMAX
        If wOINF(i).LEN <= 0 Then
            iPoint = i
            Call HairetuOpe(wOINF(), iPoint, -1)
        End If
    Next i
        
    '空きの配列削除する(配列のデータを詰める)
    For i = 0 To wNINFMAX
        If wNINF(i).LEN <= 0 Then
            iPoint = i
            Call HairetuOpe(wNINF(), iPoint, -1)
        End If
    Next i
    
    '品番入替情報を作成する
    i = 0 '前品番の位置
    j = 0 '後品番の位置
    Do
        '長さを突き合わせて数量が同じでなかったら大きい値の品番を分割する
        If (wOINF(i).LEN = wNINF(j).LEN And wOINF(i).hinban = wNINF(j).hinban) Then   '品番長さが同じ時両方とも次に進む
        ElseIf (wOINF(i).LEN >= wNINF(j).LEN) Then   '品番枚数が異なる時
            iPoint = i
            Call HairetuOpe(wOINF(), iPoint, 1)    '配列の追加
            wOINF(i + 1).hinban = wOINF(i).hinban
            wOINF(i + 1).LEN = wOINF(i).LEN - wNINF(j).LEN
            wOINF(i + 1).WAT = wOINF(i).WAT - wNINF(j).WAT
            wOINF(i + 1).MAI = wOINF(i).MAI - wNINF(j).MAI
            If wOINF(i + 1).MAI < 0 Then
                wOINF(i + 1).MAI = 0
            End If
            wOINF(i).LEN = wNINF(j).LEN
            wOINF(i).WAT = wNINF(j).WAT
            wOINF(i).MAI = wNINF(j).MAI
            Debug.Print "HINBAN=", i, wOINF(i).hinban
            Debug.Print "LEN=", i, wOINF(i).LEN
        ElseIf (wOINF(i).LEN < wNINF(j).LEN) Then   '品番数量が異なる時
            iPoint = j
            Call HairetuOpe(wNINF(), iPoint, 1)
            wNINF(j + 1).hinban = wNINF(i).hinban
            wNINF(j + 1).LEN = wNINF(j).LEN - wOINF(i).LEN
            wNINF(j + 1).WAT = wNINF(j).WAT - wOINF(i).WAT
            wNINF(j + 1).MAI = wNINF(j).MAI - wOINF(i).MAI
            wNINF(j).LEN = wOINF(i).LEN
            wNINF(j).WAT = wOINF(i).WAT
            wNINF(j).MAI = wOINF(i).MAI
            Debug.Print "HINBAN=", i, wNINF(i).hinban
            Debug.Print "LEN=", i, wNINF(i).LEN
        End If
        wOINFrecCnt = UBound(wOINF())
        wNINFrecCnt = UBound(wNINF())
        i = i + 1
        j = j + 1
        If (i > wOINFrecCnt) Then
            Exit Do
        
        End If
        If (j > wNINFrecCnt) Then
            Exit Do
        End If
        
        If (wOINF(i).LEN) <= 0 Then
            Exit Do
        End If

        If (wNINF(j).LEN) <= 0 Then
            Exit Do
        End If
    Loop

    wOINFrecCnt = UBound(wOINF())
    For i = 0 To wOINFrecCnt - 1
''''        wNINF(i).HINBAN
        If (StrComp(wNINF(i).hinban, wOINF(i).hinban, vbTextCompare) <> 0 _
            And Len(Trim(wNINF(i).hinban) > 0)) And (wNINF(i).LEN > 0) Then '品番が異なる時振替情報に登録する   'upd 2003/05/31 hitec)matsumoto wNINF(i).LEN > 0追加
            
            Koutei.CRYNUMC3 = HinNow(0).CRYNUMCA    'ブロックＩＤ
            giInpos = giInpos + 1
            Koutei.INPOSC3 = giInpos                '位置
            Koutei.KCNTC3 = wNINF(i).KCKNT          ' 工程連番
            Koutei.HINBC3 = wNINF(i).hinban         '品番
            Koutei.REVNUMC3 = HinNow(0).REVNUMCA    '製品改訂番号
            Koutei.FACTORYC3 = HinNow(0).FACTORYCA  '工場
            Koutei.OPEC3 = HinNow(0).OPECA          '操業条件
            Koutei.LENC3 = wNINF(i).LEN             '受入長さ
            Koutei.XTALC3 = HinNow(0).XTALCA        '結晶番号
            Koutei.SXLIDC3 = ""                     ' SXLID
            
            Koutei.KNKTC3 = Left(Kihon.NOWPROC, Len(Kihon.NOWPROC) - 1) & _
              CStr(CInt(Right(Kihon.NOWPROC, 1)) + 2) ' 管理工程(現在工程+2)
            Koutei.WKKTC3 = Kihon.NOWPROC           ' 工程
            Koutei.WKKBC3 = ""                      ' 作業区分
            Koutei.MACOC3 = HinNow(0).NEMACOCA      ' 処理回数
            Koutei.MODKBC3 = ""                     ' 赤黒区分
            Koutei.SUMKBC3 = ""                     ' 集計区分
            Koutei.FRKNKTC3 = ""                    ' (受入)管理工程
            If IsNull(HinOld(0).NEWKNTCA) = True Then   '(受入）工程
                Koutei.FRWKKTC3 = ""
            Else
                Koutei.FRWKKTC3 = CStr(HinOld(0).NEWKNTCA)
                'add end 2003/03/28 hitec)matsumoto --------
            End If
            Koutei.FRWKKBC3 = ""                    ' (受入)作業区分
            If IsNull(HinOld(0).NEMACOCA) = True Then   '（受入）処理回数
                Koutei.FRMACOC3 = "0"
            Else
                Koutei.FRMACOC3 = CLng(HinOld(0).NEMACOCA)
            End If
            
''''            If BlkNow.GNLC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then
''''            If BlkNow.GNMC2 <= 0 Or Kihon.ALLSCRAP = "Y" Then   '2003/05/05 hitec)matsumoto
            Select Case Kihon.NOWPROC
            Case "CC730"
                iHantei = CInt(BlkNow.GNLC2)
            Case Else
                iHantei = CInt(BlkNow.GNMC2)
            End Select
            If iHantei <= 0 Or Kihon.ALLSCRAP = "Y" Then   'upd 2003/05/27 hitec)matsumoto
                Koutei.TOWKKTC3 = " "               ' (払出)工程
                Koutei.TOMACOC3 = "0"               '(払出)処理回数
            Else
                Koutei.TOWKKTC3 = HinNow(0).GNWKNTCA    ' (払出)工程
                Koutei.TOMACOC3 = HinNow(0).GNMACOCA    ' (払出)処理回
            End If
            Koutei.FRLC3 = wNINF(i).LEN             '受入長さ
            Koutei.FRWC3 = wNINF(i).WAT             '受入重量
            Koutei.FRMC3 = wNINF(i).MAI             '受入枚数
            Koutei.FULC3 = 0                        '不良長さ
            Koutei.FUWC3 = 0                        '不良重量
            Koutei.FUMC3 = 0                        '不良枚数
            Koutei.LOSWC3 = ""                      ' ロス長さ
            
            Koutei.LOSLC3 = ""                      ' ロス重量
            Koutei.LOSMC3 = ""                      ' ロス枚数
            Koutei.TOLC3 = wNINF(i).LEN             '払出長さ
            Koutei.TOWC3 = wNINF(i).WAT             '払出重量
            Koutei.TOMC3 = wNINF(i).MAI             '払出枚数
            Koutei.SUMITLC3 = ""                    ' SUMIT長さ
            Koutei.SUMITWC3 = ""                    ' SUMIT重量
            Koutei.SUMITMC3 = ""                    ' SUMIT枚数
            Koutei.MOTHINC3 = wOINF(i).hinban       '元品番
            Koutei.XTWORKC3 = "42"                  ' 製造工場
            
            Koutei.WFWORKC3 = ""                    ' ｳｪｰﾊ製造
'           Koutei.STATIMEC3 = Null                 ' 処理開始終了
'           Koutei.STOTIMEC3 = Null                 ' 処理時間終了
'           Koutei.ETIMEC3 = ""                     ' 実績時間は入れない
            Koutei.HOLDCC3 = " "                    ' ホールドコード
            Koutei.HOLDBC3 = "0"                    ' ホールド区分
            Koutei.LDFRCC3 = ""                     ' 格下コード
            Koutei.LDFRBC3 = "0"                    ' 格下区分（ハイキ）
            Koutei.TSTAFFC3 = Kihon.STAFFID         ' 登録社員ID
            Koutei.TDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 登録日付
            
            Koutei.KSTAFFC3 = ""                    ' 更新社員ID
            Koutei.KDAYC3 = Format(Now(), "YYYY/MM/DD HH:NN:SS") ' 更新日付
            Koutei.SUMITBC3 = ""                    ' SUMIT送信フラグ
            Koutei.SNDKC3 = ""                      ' 送信フラグ
'           Koutei.SNDDAYC3 = ""                    ' 送信日付
            Koutei.MODMACOC3 = ""                   ' 赤黒の処理回数
            Koutei.KAKUCC3 = ""                     ' 確定コード
            Koutei.SUMDAYC3 = CalcSumcoTime(Koutei.KDAYC3) 'SUMCO時間
            Koutei.PAYCLASSC3 = ""                  '　転送先工場フラグ
'            Koutei.SUMITSNDC3 = ""                  ' SUMIT送信日付
            
'            Koutei.SSENDNOC3 = ""
            
            iRtn = CreateXSDC3(Koutei, wErrMsg)     '工程実績に在庫減情報登録
            If iRtn = FUNCTION_RETURN_FAILURE Then  '工程実績追加エラー
                MsgBox wErrMsg
                Exit Function
            End If
        End If
    Next i

    XSDC3Proc5 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.MAIber
    XSDC3Proc5 = FUNCTION_RETURN_FAILURE
'    gErr.HandleError
'    sErrMsg = GetMsgStr("EXXXX", sDBName)
    Resume proc_exit

End Function

Public Function HairetuOpe(HinInf() As typ_trans_info, HinNum As Integer, HINFLG As Integer)
    Dim recCnt As Integer
    Dim i, j As Integer
    Dim sflg As Integer
    
    sflg = 0
    recCnt = UBound(HinInf())
    
    If (HINFLG = 1) Then    'HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
        For i = HinNum + 1 To recCnt '既存の配列に空き場所を探す
            If (HinInf(i).LEN <= 0) Then    'i番目に空きがあったのでデータをずらす
                For j = i To HinNum + 1 Step -1
                    HinInf(j).hinban = HinInf(j - 1).hinban
                    HinInf(j).LEN = HinInf(j - 1).LEN
                    HinInf(j).WAT = HinInf(j - 1).WAT
                    HinInf(j).MAI = HinInf(j - 1).MAI
                    HinInf(j).KCKNT = HinInf(j - 1).KCKNT
                Next j
                sflg = 1
                Exit For
            End If
        Next i
        If (sflg = 0) Then  '空き見つからず
            ReDim Preserve HinInf(recCnt + 1)
            For i = recCnt + 1 To HinNum + 1 Step -1
                HinInf(i).hinban = HinInf(i - 1).hinban
                HinInf(i).LEN = HinInf(i - 1).LEN
                HinInf(i).WAT = HinInf(i - 1).WAT
                HinInf(i).MAI = HinInf(i - 1).MAI
                HinInf(i).KCKNT = HinInf(i - 1).KCKNT
            Next i
        End If
        'HinNum+1番目を空きにする
        HinInf(HinNum + 1).hinban = ""
        HinInf(HinNum + 1).LEN = 0
        HinInf(HinNum + 1).MAI = 0
        HinInf(HinNum + 1).WAT = 0
        HinInf(HinNum + 1).KCKNT = 0

    Else    'HinNum番目の配列を削除する(配列データを前につめる)
        i = HinNum
        HinInf(HinNum).hinban = ""
        HinInf(HinNum).LEN = 0
        HinInf(HinNum).MAI = 0
        HinInf(HinNum).WAT = 0
        HinInf(HinNum).KCKNT = 0
        For j = HinNum + 1 To recCnt
            If (HinInf(j).LEN > 0) Then 'HinNum以降でデータが存在していた時
                HinInf(i).hinban = HinInf(j).hinban
                HinInf(i).LEN = HinInf(j).LEN
                HinInf(i).MAI = HinInf(j).MAI
                HinInf(i).WAT = HinInf(j).WAT
                HinInf(i).KCKNT = HinInf(j).KCKNT
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
                i = i + 1
            Else
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
             End If
        Next j
End If

End Function

Public Function HairetuOpe_Mai(HinInf() As typ_trans_info, HinNum As Integer, HINFLG As Integer)
    Dim recCnt As Integer
    Dim i, j As Integer
    Dim sflg As Integer
    
    sflg = 0
    recCnt = UBound(HinInf())
    
    If (HINFLG = 1) Then    'HinNum番目の配列を空きにする(配列データを後ろにずらして空ける)
        For i = HinNum + 1 To recCnt '既存の配列に空き場所を探す
            If (HinInf(i).MAI <= 0) Then    'i番目に空きがあったのでデータをずらす
                For j = i To HinNum + 1 Step -1
                    HinInf(j).hinban = HinInf(j - 1).hinban
                    HinInf(j).LEN = HinInf(j - 1).LEN
                    HinInf(j).WAT = HinInf(j - 1).WAT
                    HinInf(j).MAI = HinInf(j - 1).MAI
                    HinInf(j).KCKNT = HinInf(j - 1).KCKNT
                Next j
                sflg = 1
                Exit For
            End If
        Next i
        If (sflg = 0) Then  '空き見つからず
            ReDim Preserve HinInf(recCnt + 1)
            For i = recCnt + 1 To HinNum + 1 Step -1
                HinInf(i).hinban = HinInf(i - 1).hinban
                HinInf(i).LEN = HinInf(i - 1).LEN
                HinInf(i).WAT = HinInf(i - 1).WAT
                HinInf(i).MAI = HinInf(i - 1).MAI
                HinInf(i).KCKNT = HinInf(i - 1).KCKNT
            Next i
        End If
        'HinNum+1番目を空きにする
        HinInf(HinNum + 1).hinban = ""
        HinInf(HinNum + 1).LEN = 0
        HinInf(HinNum + 1).MAI = 0
        HinInf(HinNum + 1).WAT = 0
        HinInf(HinNum + 1).KCKNT = 0

    Else    'HinNum番目の配列を削除する(配列データを前につめる)
        i = HinNum
        HinInf(HinNum).hinban = ""
        HinInf(HinNum).LEN = 0
        HinInf(HinNum).MAI = 0
        HinInf(HinNum).WAT = 0
        HinInf(HinNum).KCKNT = 0
        For j = HinNum + 1 To recCnt
            If (HinInf(j).MAI > 0) Then 'HinNum以降でデータが存在していた時
                HinInf(i).hinban = HinInf(j).hinban
                HinInf(i).LEN = HinInf(j).LEN
                HinInf(i).MAI = HinInf(j).MAI
                HinInf(i).WAT = HinInf(j).WAT
                HinInf(i).KCKNT = HinInf(j).KCKNT
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
                i = i + 1
            Else
                HinInf(j).hinban = ""
                HinInf(j).LEN = 0
                HinInf(j).MAI = 0
                HinInf(j).WAT = 0
                HinInf(j).KCKNT = 0
             End If
        Next j
End If

End Function

