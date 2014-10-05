Attribute VB_Name = "mdlCommon"

'///////////////////////////////////////////////////
' @(S)
'       共通関数
'
' @(h)  mdlDWHCommon.bas ver 1.0 ( 1999.12.14 小川 宏幸 )
'
'///////////////////////////////////////////////////
Option Explicit

Public Const constInch As Integer = 25    ''1ｲﾝﾁの㎜サイズ

''背景色とLockedの定数列挙
Public Enum CtlKind
    NORMAL_CTL                      '' 0 : 入力可能・白
    CHECK_CTL                       '' 1 : 入力不可・グレー
    RED_CTL                         '' 2 : 入力可能・赤
End Enum

''ﾒｯｾｰｼﾞ／ログ　種別定数列挙
Public Enum MsgKind
    NORMAL_MSG                      '' 0 : 通常(画面表示)
    ERR_DISP                        '' 1 : 画面表示ｴﾗｰ
    ERR_LOG                         '' 2 : ログ出力ｴﾗｰ
    ERR_DISP_LOG                    '' 3 : 画面表示・ログ出力ｴﾗｰ
    DEBUG_DISP = 5                  '' 5 : 画面表示ﾃﾞﾊﾞｯｸﾞ
    DEBUG_LOG                       '' 6 : ログ出力ﾃﾞﾊﾞｯｸﾞ
    DEBUG_DISP_LOG                  '' 7 : 画面表示・ログ出力ﾃﾞﾊﾞｯｸﾞ
End Enum

''初期処理エラー種別定数列挙
Public Enum InitKind
    NORMAL_RET                      '' 0 : 正常
    EXITSUB_RET                     '' 1 : Exit Sub しなければならない
    MAINMENU_RET                    '' 2 : GotoMainMenu しなければならない
End Enum

''出力するメッセージ属性（マスク）
Private Const MsgKindMask = 7       '' 0 : 通常ﾒｯｾｰｼﾞが画面表示される
                                    '' 1 : 通常/画面表示ｴﾗｰが出力される
                                    '' 2 : 通常/ログ出力ｴﾗｰが出力される
                                    '' 3 : 通常/画面表示ｴﾗｰ/ログ出力ｴﾗｰが出力される
                                    '' 7 : 通常/画面表示ｴﾗｰ/ログ出力ｴﾗｰ/ﾃﾞﾊﾞｯｸﾞが出力される

''パス
Private Const LogDir = "..\LOG\" ''ログのパス

''グローバル変数
Public gobjOraSess      As Object   ''オラクルセッションオブジェクト
Public gobjOraDB        As Object   ''オラクルデータベースブジェクト

''ＥＸＥオプション
Public gsFactryCd       As String   ''工場コード
Public gsCallCd         As String   ''呼出区分
Public gsHinban         As String   ''品番
Public myFactryCd       As String   ''工場コード

Public gsCompName       As String   ''コンピュータ名
Public gsEXEName        As String   ''ＥＸＥファイル名

Public gbFTPFlg         As Boolean  ''FTP転送フラグ
Public mbMenuRet        As Boolean  ''メニュー遷移許可フラグ

'Public gsProcCode1      As String   ''  処理工程コード1
'Public gsProcCode2      As String   ''  処理工程コード2
'Public gsProcCode3      As String   ''  処理工程コード3
'Public gsProcCode4      As String   ''  処理工程コード4
'Public gsProcCode5      As String   ''  処理工程コード5
''モジュール変数
Private msLogFile       As String   ''ログファイル名
Private msMsgStr(100)   As String   ''メッセージ配列
'' ========== 変数説明 ===========
'' メッセージ配列は、ログ初期化処理内でメッセージ初期化し
'' ログ出力・画面表示にて使用する。

Private Const SOKUTEI_MAX = 8       '' 測定個所数最大値
Private Const NULL_CHECK = 999999   '' データ無しチェック
'' 2000/04/24 追加
Public Type RRG_CALC                '' RRG計算データ
    dTeikou As Double               '' 抵抗値
    sRRGFlg As String               '' 計算フラグ
End Type
Public Type TYPE_RRG                '' RRG算出抵抗値一覧
    iSampleNo As String             '' サンプルNo
    dTeikouDT(SOKUTEI_MAX) As RRG_CALC '' 抵抗値(A～I)
End Type

'Cs推定計算用ﾊﾟﾗﾒｰﾀ　06/04/20 ooba
Public Type CS_SUITEI_TYPE
    sSiWeight           As String   ''ﾁｬｰｼﾞ量(Kg)
    sTopWT              As String   ''ﾄｯﾌﾟ重量(Kg)
    sUpDm               As String   ''直径(mm)
    sCsHenseki          As String   ''ｶｰﾎﾞﾝ偏析係数(ｺｰﾄﾞﾏｽﾀに保持)
    sSamplePos          As String   ''ｻﾝﾌﾟﾙ位置
    sResCs              As String   ''ｻﾝﾌﾟﾙ測定値
    sInfPos             As String   ''推定位置
End Type
    
' 更新種別
Public Enum CHANGE_TYPE
    ST_NORMAL                ' 未更新
    ST_UPDATE                ' 更新
    ST_INSERT                ' 追加
    ST_DELETE                ' 削除
    ST_DELINS                ' 削除追加
End Enum
' Oi/Csテーブルデータ
Public Type ST_OICS
    sCrystalNo  As String     ' 結晶番号
    sCryBuiNo   As String      ' 結晶部位
    sMenPosIti  As String     ' 面内位置
    sSampleNo   As String      ' サンプルNo
    sBuiKubun   As String      ' 部位区分
    sCarbonAT   As String      ' カーボン(AT)
    sSansoAT    As String       ' 酸素(AT)
    sCarbonPpma As String    ' カーボン(ppma)
    sSansoPpma  As String     ' 酸素(ppma)
    sORGNo      As String         ' ORG
    sYMDData    As String       ' 処理日付
    sChgType    As CHANGE_TYPE  ' 変更種別
End Type
''===============================================================================

''コンピュータ名取得API
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpszName As String, lpcchBuffer As Long) As Long

'工程連番作成用
Public wMaxKcnt As Integer
'戻るフラグ(f_cmbc039_2からF_cmbc039_1で戻るときに設定)
Public intModoru As Integer


'*ADD*  TCS)K.Kunori 2004.11.29 START >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    '2004/9/27tcs  yamauchi 追加
    Public gsSysdate        As String   ''ｼｽﾃﾑ日付
    
    '2004/11/12tcs tagawa 追加 start-------------------------------------
    Public gsSystemCd       As String   ''システム区分(200、300)
    
    '*DEL* 重複宣言になっているので削除 TCS)K.Kunori 2004.11.29 START >>>
    '''Public Const SYSTEM_200 = "2"               ''200mmシステム時
    '''Public Const SYSTEM_300 = "3"               ''300mmシステム時
    '''
    '''''ｼｽﾃﾑ名
    '''Public Const SYSTEM_NAME_200 = "200mm結晶操業システム"
    '''Public Const SYSTEM_NAME_300 = "300mm結晶操業システム"
    '*DEL* 重複宣言になっているので削除 TCS)K.Kunori 2004.11.29 END <<<
    ''2004/11/12tcs tagawa 追加 end--------------------------------------
    
    '*ADD* TCS)K.Kunori 2004.11.17 START >>>
    Public gobjOraSess2     As Object   ''オラクルセッションオブジェクト(SQL文ログ作成用)
    Public gobjOraDB2       As Object   ''オラクルデータベースブジェクト(SQL文ログ作成用)
    '*ADD* TCS)K.Kunori 2004.11.17 END <<<
    
    '2004/9/27tcs Suenaga 追加 start-------------------------------------
    Private mtrlNo  As String          '原料No
    Private cryno   As String          '現品ロットNo
    Private PROCCD  As String          '工程コード
    Private staffCd As String          '担当者コード
    Private recW    As String          '受入重量
    Private sendW   As String          '払出重量
    Private lossW   As String          'ロス重量
    Private factCd  As String          '工場コード
    Private recCd   As String          '受入工場コード
    Private sendCd  As String          '払出工場コード
    Private disapp  As String          '消滅区分
    Private sikake  As String          '仕掛区分
    Private sysCd   As String          'システム区分コード
    Private conceK  As String          '濃度区分
    Private conceT  As String          '濃度値
    Private SENDFLG As String          '原料送信フラグ
    Private occuFlg As String          '発生フラグ
    Private conceM  As String          '元濃度
    Private planFac As String          '使用予定工場
    Private tanaKu  As String          '棚入れ区分
    '2004/9/27tcs Suenaga 追加 end-------------------------------------
    
    '*** UPDATE START T.TERAUCHI 2004/10/19 払出区分追加
    Private stowkkbb3 As String          '払出区分
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
    Private sChgNo  As String           'チャージNo(CC200/CC300登録用)　05/08/23 ooba
    
    '2004/9/17tcs Yamauchi 追加 start-------------------------------------
    
    ''背景色
    Public Const COLOR_GRAY_SPR = &HC0C0C0      ''灰色
    Public Const COLOR_PINK_SPR = &HFFC0FF      ''ﾋﾟﾝｸ
    Public Const COLOR_RED_SPR = &HFF&          ''赤
    
    Public Const SYSTEM_200 = "2"               ''200mmシステム時
    Public Const SYSTEM_300 = "3"               ''300mmシステム時
    
    ''ｼｽﾃﾑ名
    Public Const SYSTEM_NAME_200 = "200mm結晶操業システム"
    Public Const SYSTEM_NAME_300 = "300mm結晶操業システム"
    
    ''ﾒｯｾｰｼﾞ（通常）
    Public Const MSG0001 = "入力欄に入力後、抽出キーを押下"
    Public Const MSG0002 = "入力情報で廃棄しますが、宜しいですか"
    Public Const MSG0003 = "廃棄しました"
    Public Const MSG0004 = "担当者、横分割数を入力し、実行ボタン押下"
    Public Const MSG0005 = "長さ、重量、縦分割数を入力し、実行ボタン押下"
    Public Const MSG0006 = "重量、仕掛状態、使用予定工場を入力し、実行ボタン押下"
    Public Const MSG0007 = "入力情報で更新しますが、宜しいですか"
    Public Const MSG0008 = "対象のデータを選択し、実行キーを押下"
    Public Const MSG0009 = "更新しました"
    Public Const MSG0010 = "対象のデータを選択し、実行キーを押下"
    Public Const MSG0011 = "表示明細の払出を行います。宜しいでしょうか"
    Public Const MSG0012 = "追加しました"
    Public Const MSG0013 = "表示明細の洗浄受入を行います。宜しいでしょうか"
    Public Const MSG0014 = "入力情報で他工場に払出いたします。宜しいでしょうか"
    Public Const MSG0015 = "入力情報の棚入れ/置き場処理を行います。宜しいでしょうか"
    Public Const MSG0016 = "入力情報で払出しますが、宜しいですか"
    Public Const MSG0017 = "入力欄に入力後、払出キーを押下"
    
    ''ﾒｯｾｰｼﾞ（ｴﾗｰ用）
    Public Const ERR0001 = "長さの合計がブロック全体の長さを超えています"
    Public Const ERR0002 = "重量の合計がブロック全体の重量を超えています"
    Public Const ERR0003 = "重量の合計が横割り毎の重量を超えています"
    Public Const ERR0004 = "長さが入力されていません"
    Public Const ERR0005 = "重量が入力されていません"
    Public Const ERR0006 = "縦分割数が入力されていません"
    
    '*** update start T.TERAUCHI 2004/10/20
    'Public Const ERR0007 = "分割数が多すぎます"
    Public Const ERR0007 = "現品ロットNoを採番することができません"
    '*** UPDATE END   T.TERAUCHI 2004/10/20
    
    Public Const ERR0008 = "統合ロットの為、この機能は使用できません"
    Public Const ERR0009 = "濃度計算の情報が不足しています"
    Public Const ERR0010 = "濃度計算の情報が不正です"
    Public Const ERR0011 = "使用予定工場が選択されていません"
    Public Const ERR0012 = "払出工場が選択されていません"
    Public Const ERR0013 = "バスケットNoの入力に誤りがあります"
    Public Const ERR0014 = "洗浄状態の入力に誤りがあります"
    Public Const ERR0015 = "原料番号の連番はこれ以上採番できません"
    Public Const ERR0016 = "推定抵抗値取得の情報が不足しています"
    Public Const ERR0017 = "選択されたロットは、受入処理をすることはできません"
    Public Const ERR0018 = "入力された重量が不正です"
    Public Const ERR0019 = "入力されたブロック長さが不正です"
    Public Const ERR0020 = "Csアウト値計算の情報が不足しています"
    Public Const ERR0021 = "Csアウト値計算の情報が不正です"
    Public Const ERR0022 = "結晶番号は12桁入力してください"
    Public Const ERR0023 = "ブロック形状を選択してください"
    Public Const ERR0024 = "ライフタイム10Ω換算値計算の情報が不足しています"
    Public Const ERR0025 = "ライフタイム10Ω換算値計算の情報が不正です"
    Public Const ERR0026 = "電極材の為、切断処理することはできません"
    Public Const ERR0027 = "推定抵抗を計算する為の実測値がありません"
    Public Const ERR0028 = "推定抵抗を計算する為の結晶情報がありません"
    Public Const ERR0029 = "抵抗値が廃棄対象の為、受入処理することはできません"
    Public Const ERR0030 = "払出工場が選択されている為、実行処理することはできません"
    Public Const ERR0031 = "ブロック部位対象外の為、濃度計算はできません"
    
    ''工程ｺｰﾄﾞ
    Public Const PROCD_GENRYO_UKEIRE = "CB410"              ''精製原料受入
    Public Const PROCD_GENRYO_SETUDAN = "CB510"             ''精製原料切断
    Public Const PROCD_ROT_KOSEI = "CB610"                  ''ﾛｯﾄ構成
    Public Const PROCD_GENRYO_SENJYO_UKEIRE = "CB220"       ''精製原料洗浄受入
    Public Const PROCD_GENRYO_SENJYO_HARAIDASI = "CB225"    ''精製原料洗浄払出
    Public Const PROCD_GENRYO_TANAIRE = "CB230"             ''原料棚入
    Public Const PROCD_ZAIKO_SYUSEI = "RP10"                ''在庫修正
    
    ''円周率
    Public Const CIRCULAR_CONSTANT = 3.14159265358979
    
    ''シリコン比重
    Public Const SPECIFIC_GRAVITY = 0.00233
    '2004/9/17tcs Yamauchi 追加 end-------------------------------------

'*ADD*  TCS)K.Kunori 2004.11.29 END <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<




'///////////////////////////////////////////////////
' @(f)
' 機能    : コンピュータ名取得・変数セット
'
' 返り値  : コンピュータ名
'
' 引き数  :
'
' 機能説明: コンピュータ名を取得し、変数セット
'
'///////////////////////////////////////////////////
Public Function GetCompName() As String
''>>>>> PC名20文字対応 SETsw H.Iwamoto 2005/10/31
'    Dim sCompName As String * 9     ''ｺﾝﾋﾟｭｰﾀ名受取ﾊﾞｯﾌｧ
    Dim sCompName As String * 20    ''ｺﾝﾋﾟｭｰﾀ名受取ﾊﾞｯﾌｧ
''<<<<< PC名20文字対応 SETsw H.Iwamoto 2005/10/31
    Dim lCompNameLen As Long        ''ﾊﾞｯﾌｧｻｲｽﾞ渡し、ｺﾝﾋﾟｭｰﾀ名LEN受取
    Dim bResult As Boolean          ''取得結果受取
    
    ''サイズをセット
    lCompNameLen = LenB(sCompName) - 1
    ''取得
    bResult = GetComputerName(sCompName, lCompNameLen)
    If bResult Then                 ''取得成功
        gsCompName = left(sCompName, lCompNameLen)  ''ｸﾞﾛｰﾊﾞﾙ変数にセット
    Else                            ''取得失敗
        gsCompName = ""                             ''ｸﾞﾛｰﾊﾞﾙ変数にセット
    End If

''>>>>> 取敢えず今(2005/10/26現在)は8文字返す。 SETsw H.Iwamoto 2005/10/31
    gsCompName = IIf(Len(gsCompName) > 10, left(gsCompName, 8), gsCompName)
''<<<<< 取敢えず今(2005/10/26現在)は8文字返す。 SETsw H.Iwamoto 2005/10/31
    
    ''コンピュータ名を返す
    GetCompName = gsCompName
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 実行ファイル名取得・変数セット
'
' 返り値  : 実行ファイル名
'
' 引き数  :
'
' 機能説明: 実行ファイル名を取得し、変数セット
'
'///////////////////////////////////////////////////
Public Function GetEXEName() As String
    gsEXEName = App.EXENAME         ''VBのｱﾌﾟﾘｹｰｼｮﾝｵﾌﾞｼﾞｪｸﾄから実行ﾌｧｲﾙ名取得・変数ｾｯﾄ
    GetEXEName = Trim(gsEXEName)          ''実行ﾌｧｲﾙ名を戻す
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : コマンドライン引数取得・変数セット
'
' 返り値  : True:正　False:否
'
' 引き数  :
'
' 機能説明: コマンドライン引数を取得し・トークン切出し・変数セット
'
'///////////////////////////////////////////////////
Public Function GetCmdLine() As Boolean
    Dim sCmdLine As String
    
    ''コマンドライン取得
    sCmdLine = Command
    '' 0        1         2
    '' 1234567890123456789012
    ''"99_*******_***********"
    ''工場コード_呼出区分_品番
    
    ''固定でコマンドライン引数を切出す
    gsFactryCd = left(sCmdLine, 2)    ''工場コード(2桁)
    gsCallCd = Mid(sCmdLine, 4, 7)    ''呼出区分(7桁)
    gsHinban = Mid(sCmdLine, 12, 11)  ''品番(11桁)
    myFactryCd = Mid(sCmdLine, 24, 2)
    
    If Len(gsFactryCd) <> 2 Then Exit Function
    If Len(gsCallCd) <> 7 Then Exit Function
    If gsHinban = "00000000000" Then gsHinban = ""
    GetCmdLine = True
End Function


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : フォ－ムを中央に表示
'
' 返り値    : なし
'
' 引き数    : FrmName - フォーム名
'
' 機能説明  : フォ－ムを中央に表示
'
'///////////////////////////////////////////////////
Public Function FrmCenter(frmName As Form)
    With frmName
        If .WindowState <> 2 Then
            .left = (Screen.Width - .Width) / 2
            .top = (Screen.Height - .Height) / 2
            .Width = 12000
            .Height = 9000
        End If
    End With
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : キャンセル・修正ボタン有効／無効制御
' 返り値  : なし
' 引き数  : True:制御有効
'           False:制御無効
' 機能説明: アクティブになっているフォームの
'          キャンセル・修正ボタンを制御有効／無効にする
'///////////////////////////////////////////////////
Public Sub F3F4Enabled(bEnabled As Boolean)
    Screen.ActiveForm.cmdF(3).Enabled = bEnabled
    Screen.ActiveForm.cmdF(4).Enabled = bEnabled
End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能    : フォーム上のコントロールを使用不可にする
'
' 返り値  :
'
' 引き数  : フォーム
'
' 機能説明: 指定したフォームの「[Ｆ１]ﾒｲﾝﾒﾆｭｰ」ボタン以外
'           のコントロールを使用不可にする
'
'///////////////////////////////////////////////////
Public Sub CtrlCancel(frmForm As Form)
    Dim iIdx As Integer
    Dim ctlControl As Control
    
    ''フォーム上のコントロールを全て使用不可にする
    For Each ctlControl In frmForm.Controls
        If TypeOf ctlControl Is TextBox Then
' Mod 2005/11/18 M.Makino CHECK_CTL -> CTRL_DISABLE_GRAY
            Call CtrlEnabled(ctlControl, CTRL_DISABLE_GRAY)
        ElseIf TypeOf ctlControl Is ComboBox Then
' Mod 2005/11/18 M.Makino CHECK_CTL -> CTRL_DISABLE_GRAY
            Call CtrlEnabled(ctlControl, CTRL_DISABLE_GRAY)
        ElseIf TypeOf ctlControl Is CommandButton Then
            ctlControl.Enabled = False
        End If
    Next ctlControl
    
    ''「[Ｆ１]ﾒｲﾝﾒﾆｭｰ」ボタンを使用可能にする
    frmForm.cmdF(1).Enabled = True
End Sub

'///////////////////////////////////////////////////
' @(f)
'
' 機能      : メッセージ初期化
'
' 返り値    : なし
'
' 引き数    : なし
'
' 機能説明  : メッセージ配列設定
'
' 備考      :
'///////////////////////////////////////////////////
Private Sub MsgInit()
'   CSV項目説明
'   ﾒｯｾｰｼﾞｺｰﾄﾞ,メッセージ
    Const MsgData1 As String = _
    "01,入力欄に入力後、実行キーを押下," & _
    "02,表示された内容を確認し実行キーを押下," & _
    "03,入力欄修正後、実行キーを押下," & _
    "04,選択し実行キーを押下," & _
    "05,前画面ボタンで戻ります," & _
    "06,不正な文字が入力されています," & _
    "07,値が小数点数になっています," & _
    "08,指定した担当者コードは登録されていません," & _
    "09,指定した結晶番号は登録されていません," & _
    "10,指定したサンプルNo.は登録されていません," & _
    "11,桁数が足りません," & _
    "12,入力されてません," & _
    "13,数値を入力して下さい," & _
    "14,値が０になっています," & _
    "15,値がマイナスになっています," & _
    "16,指定した結晶番号の仕掛工程が違います," & _
    "17,指定した原料番号の仕掛工程が違います," & _
    "18,依頼重量が仕掛重量を超えています," & _
    "19,指定した原料種類は多結晶ではありません," & _
    "20,指定した原料種類は登録されていません,"
    Const MsgData2 As String = _
    "22,仕掛工程が違います," & _
    "23,ホールド処理中の結晶です," & _
    "25,払出完了済みの結晶です," & _
    "26,分割終了を行ってください," & _
    "27,発行済：続けて発行する場合は、入力して下さい," & _
    "30,表示された内容を確認しｷｬﾝｾﾙか修正を押下," & _
    "31,工程コードを選択して下さい," & _
    "32,集計区分を選択して下さい," & _
    "33,ＧＲ／ＢＳを選択して下さい," & _
    "40,運転日誌　印刷中…," & _
    "41,引上指示書　印刷中…," & _
    "42,加工検査票　印刷中…," & _
    "43,再各付け伺い書　印刷中…,"
    Const MsgData3 As String = _
    "50,入力したデータに誤りがあります," & _
    "51,入力した値に誤りがあります," & _
    "52,日付が正しく入力されていません," & _
    "53,日付の範囲指定が正しくありません," & _
    "54,既に登録済みです," & _
    "55,該当データがありませんでした," & _
    "56,データが重複しています," & _
    "57,印刷が失敗しました," & _
    "58,選択されていません," & _
    "60,コンピュータ名取得失敗," & _
    "61,ログ初期化失敗," & _
    "62,実行ファイル名取得失敗," & _
    "63,多重起動しました," & _
    "64,ｺﾏﾝﾄﾞﾗｲﾝ引数が不足しています," & _
    "65,実行ファイルを起動できません," & _
    "66,整数が有効桁数を超えています," & _
    "67,小数点数が有効桁数を超えています," & _
    "68,数値が最大値を超えています," & _
    "69,数値が最小値未満です,"
    Const MsgData4 As String = _
    "70,レコード検索失敗," & _
    "71,レコード挿入失敗," & _
    "72,レコード更新失敗," & _
    "73,レコード削除失敗," & _
    "100,オラクルエラー," & _
    "00,,"
    'データを統一化
    Const sMsgData As String = MsgData1 & MsgData2 & MsgData3 & MsgData4
    
    Dim iP1 As Integer
    Dim iP2 As Integer
    Dim iMsgCd As Integer
    
    iP2 = 1  '最初の位置を指定
    Do
        'ﾒｯｾｰｼﾞｺｰﾄﾞ抽出
        iP1 = InStr(iP2, sMsgData, ",")
        iMsgCd = val(Mid(sMsgData, iP2, iP1 - iP2))
        iP2 = iP1 + 1
        
        'メッセージ抽出
        iP1 = InStr(iP2, sMsgData, ",")
        msMsgStr(iMsgCd) = Mid(sMsgData, iP2, iP1 - iP2)
        iP2 = iP1 + 1
        
'Debug.Print iMsgCd; ":"; msMsgStr(iMsgCd)
    Loop While iMsgCd   'ｺｰﾄﾞが０になるまで
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : ログ出力関連の初期化
'
' 返り値    : True:正　False:否
'
' 引き数    : LoadModule - [i]ロードモジュール名(Kxxxxxx.EXE)
'             LogFile    - [i]ログ出力ファイル名(Kxxxxxx.LOG、ファイル名のみ)
'
' 機能説明  : ログ出力関連の初期化、プログラムの起動時にCALLする。
'///////////////////////////////////////////////////
Public Function LogInit() As Boolean
    Dim sLine As String '行データ
    Dim FN1, FN2 'ファイル番号
    Dim sTmp As String 'work file
    Dim m, n
    
    On Error Resume Next
    
    ''ログ処理失敗をセット
    LogInit = False
    
    ''メッセージ初期化============================
    Call MsgInit
    
    ''ログファイル名ｸﾞﾛｰﾊﾞﾙ変数セット
    msLogFile = LogDir & GetCompName() & ".Log" ''コンピュータ名をﾛｸﾞﾌｧｲﾙ名にする
    
    ''テンポラリファイル名作成
    m = Len(msLogFile)
    sTmp = left(msLogFile, m - 4) & ".tmp"      ''拡張子を取外し".tmp"を付加
    
    ''ディレクトリ存在チェック
    If Dir(LogDir, vbDirectory) = "" Then       ''ログディレクトリが無い
        MkDir (LogDir)                          ''ログディレクトリ作成
        If Dir(LogDir, vbDirectory) = "" Then   ''ログディレクトリが作成できなければ
            GoTo Er
        End If
    End If
    
    ''ログファイルのオープン
    FN1 = FreeFile                              ''未使用のファイル番号を取得します
    Err = 0
    Open msLogFile For Input As #FN1            ''ログファイルオープン
    If Err <> 0 Then                            ''ログファイルが無ければ
        LogInit = True                          ''正常
        Exit Function
    End If
    
    ''テンポラリファイル削除処理
    If Dir(sTmp) <> "" Then                     ''テンポラリファイルが既に存在していたら
        Kill sTmp                               ''テンポラリファイル削除
    End If
    
    ''テンポラリファイルオープン
    FN2 = FreeFile                              ''未使用のファイル番号を取得します
    Err = 0
    Open sTmp For Output As #FN2                ''テンポラリファイルオープン
    If Err <> 0 Then
        Debug.Print "ﾃﾝﾎﾟﾗﾘﾌｧｲﾙが開けない:" & sTmp
        Close #FN1
        GoTo Er
    End If
    
    ''一ヵ月以内のログのみテンポラリにコピー
    Do While Not EOF(FN1)                       ''ファイルの終端までループ
        Line Input #FN1, sLine                  ''ログファイル読込
        If IsDate(left(sLine, 19)) Then         ''行頭が日付なら
            If CDate(left(sLine, 19)) > CDate(DateAdd("m", -1, Now)) Then   '' 1ヶ月前以内なら
                Print #FN2, sLine               ''テンポラリに出力
            End If
        End If
    Loop
    
    Close #FN1
    Close #FN2
    Kill msLogFile                              ''ログファイル削除
    Name sTmp As msLogFile                      ''テンポラリをログファイルにする
    ''ログ処理正常をセット
    LogInit = True
    On Error GoTo 0
    Exit Function
Er:
    Close #FN1
    Close #FN2
    On Error GoTo 0
End Function


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : メッセージログ出力
'
' 返り値    : なし
'
' 引き数    : メッセージ
'
' 機能説明  : メッセージをログ出力する
'
' 備考      :
'///////////////////////////////////////////////////
Private Sub MsgLog(Msg As String)
    On Error Resume Next
    Dim fno                             ''ファイル番号
    
    fno = FreeFile                      '' 未使用のファイル番号を取得する
    Err = 0
    Open msLogFile For Append As #fno   '' オープンして
    If Err <> 0 Then
        Exit Sub
    End If
    Print #fno, Msg                     '' 出力して
    Close #fno                          '' 閉じる
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : メッセージ画面表示
'
' 返り値    : なし
'
' 引き数    : arg1:メッセージ
'
' 機能説明  : メッセージを画面表示する
'
' 備考      :
'
'           使用条件：ﾌｫｰﾑ上にlblMsgというｺﾝﾄﾛｰﾙを、
'                     貼り付けてあること
'
'///////////////////////////////////////////////////
Private Sub MsgDisp(Msg As String, Optional lForeColor As Long = 0)
    On Error Resume Next
'    Screen.ActiveForm.lblMsg.ForeColor = lForeColor
    Screen.ActiveForm.lblMsg = Msg
    Screen.ActiveForm.lblMsg.Refresh
    On Error GoTo 0
End Sub


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : ｵﾗｸﾙｴﾗｰﾒｯｾｰｼﾞよりｵﾗｸﾙｴﾗｰｺｰﾄﾞを切出す
'
' 返り値    : ｵﾗｸﾙｴﾗｰｺｰﾄﾞ
'             "ORA-????? "が見つからない場合 ""を返す
'
' 引き数    : (ｵﾗｸﾙｵﾌﾞｼﾞｪｸﾄ).LastServerErrText :"ORA-????? "を含む文字列
'
' 機能説明  : ｵﾗｸﾙｴﾗｰﾒｯｾｰｼﾞよりｵﾗｸﾙｴﾗｰｺｰﾄﾞを切出す
'
' 備考      :
'///////////////////////////////////////////////////
Private Function GetStrOraErrCd(LastServerErrText As String) As String
    Dim vPnt
    Dim vLen
    vPnt = InStr(LastServerErrText, "ORA-")             ''ｴﾗｰｺｰﾄﾞの先頭
    If vPnt < 1 Then
        GetStrOraErrCd = ""
        Exit Function
    End If
    vLen = InStr(vPnt, LastServerErrText, ":") - 1      ''ﾚﾝｸﾞｽﾁｪｯｸ
    If vLen < 1 Then
        GetStrOraErrCd = Mid(LastServerErrText, vPnt)
    Else
        GetStrOraErrCd = Mid(LastServerErrText, vPnt, vLen)
    End If
End Function


'///////////////////////////////////////////////////
' @(f)
'
' 機能      : メッセージ編集・画面表示・ログ出力
'
' 返り値    : なし
'
' 引き数    : arg1:ﾒｯｾｰｼﾞｺｰﾄﾞ 100:ｵﾗｸﾙｴﾗｰ 100以外:ｵﾗｸﾙｴﾗｰ以外
'             arg2:追加ﾒｯｾｰｼﾞ
'             arg3:ﾒｯｾｰｼﾞ属性　0:通常ﾒｯｾｰｼﾞ
'                              1:画面表示ｴﾗｰﾒｯｾｰｼﾞ（入力欄赤表示のｴﾗｰなど）
'                              2:ログ出力ｴﾗｰﾒｯｾｰｼﾞ
'                              3:画面表示・ログ出力ｴﾗｰﾒｯｾｰｼﾞ（ｵﾗｸﾙｴﾗｰなど）
'                              5:画面表示ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞ
'                              6:ログ出力ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞ
'             arg4:ｵﾗｸﾙｴﾗｰ時のﾛｸﾞ/画面表示ﾃｰﾌﾞﾙ名
'
' 機能説明  : ﾒｯｾｰｼﾞｺｰﾄﾞからﾒｯｾｰｼﾞを編集して画面出力し、
'             追加ﾒｯｾｰｼﾞを編集したﾒｯｾｰｼﾞに追加してログ出力し、
'             ﾒｯｾｰｼﾞｺｰﾄﾞがｴﾗｰ区分の場合、警告音を鳴らす。
' 備考      :
'       【ログ出力形式】
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg 改行
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg 改行
'           .
'           .
'       YYYY/MM/DD HH:NN:SS::LOADMODULE::MSGCD::Msg::AddMsg 改行
'       YYYY/MM/DD = 年月日
'       HH:NN:SS   = 時分秒
'       LOADMODULE = ロードモジュール名
'       MsgCd      = メッセージ番号
'       Msg        = 概要メッセージ(固定文字)
'       AddMsg     = 詳細メッセージ
'       【ログ出力例】
'       1998/04/01 10:10:00::Kxxxxxx.EXE::AA250100::アプリケーション起動::メッセージ詳細
'       1998/04/01 10:10:03::Kxxxxxx.EXE::AA250200::アプリケーション終了::メッセージ詳細
'
'///////////////////////////////////////////////////
Public Sub MsgOut(ByVal iMsgCd As Integer, Optional ByVal sAddMsgStr As String = "", _
           Optional ByVal eMsgKind As MsgKind = 0, Optional ByVal TABLENAME As String = "Unknown")
    Dim sMsg As String                              ''メッセージ
    Dim sOraErrCd As String                         ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞ
    
    'メッセージ初期化
    Call MsgInit

    'ﾒｯｾｰｼﾞ属性がﾒｯｾｰｼﾞ出力属性範囲外の場合出力しない（開発運用開始後、ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞを出力しないようにできる）
    If Not ((eMsgKind = NORMAL_MSG) Or _
            ((eMsgKind And MsgKindMask) <> 0)) Then
        Exit Sub                                    ''終了
    End If
    
    If iMsgCd < 100 Then                            ''ﾒｯｾｰｼﾞｺｰﾄﾞがｵﾗｸﾙ以外なら
        ''オラクル以外のメッセージ
        On Error Resume Next                        ''ｴﾗｰﾄﾗｯﾌﾟ
        sMsg = msMsgStr(iMsgCd)                     ''メッセージ取得
        On Error GoTo 0                             ''ｴﾗｰﾄﾗｯﾌﾟ解除
    Else                                            ''ﾒｯｾｰｼﾞｺｰﾄﾞがｵﾗｸﾙｴﾗｰなら
        ''オラクルのエラーメッセージ
        If gobjOraSess.LastServerErr Then           ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄのエラーならば
            sMsg = gobjOraSess.LastServerErrText    ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄｴﾗｰﾒｯｾｰｼﾞをセット
            gobjOraSess.LastServerErrReset          ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄｴﾗｰをリセット
        ElseIf gobjOraDB.LastServerErr Then         ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄのエラーならば
            ''ｵﾗｸﾙｴﾗｰﾒｯｾｰｼﾞよりｵﾗｸﾙｴﾗｰｺｰﾄﾞを切出す
            sOraErrCd = GetStrOraErrCd(gobjOraDB.LastServerErrText)
            If sOraErrCd <> "" Then                 ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞが入っていれば
                sMsg = "DBエラー（" & TABLENAME & ")" & sOraErrCd ''指定のフォーマットで編集
                sAddMsgStr = gobjOraDB.LastServerErrText & _
                             "::" & sAddMsgStr
            Else                                    ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞが入っていなければ
                sMsg = gobjOraDB.LastServerErrText  ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄｴﾗｰﾒｯｾｰｼﾞをセット
            End If
            gobjOraDB.LastServerErrReset            ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄｴﾗｰをリセット
        ElseIf Err.Number Then                      ''実はVBのｴﾗｰだったなら
            sMsg = Error(Err.Number)                ''VBのｴﾗｰﾒｯｾｰｼﾞをセット
        Else                                        ''実はｴﾗｰじゃないならば
            sMsg = "ｵﾗｸﾙ正常時にｴﾗｰ出力した"         ''警告
        End If
    End If
    
    If (eMsgKind = NORMAL_MSG) Or _
       (eMsgKind And ERR_DISP) Then                     ''通常ﾒｯｾｰｼﾞか画面表示ビットが立っていれば
        ''エラーなら赤表示
        If (eMsgKind = ERR_DISP) Or _
           (eMsgKind = ERR_DISP_LOG) Then
            If iMsgCd = 100 Then                        ''オラクルエラーの場合
                MsgDisp sMsg, vbRed                     ''メッセージを画面表示する
            Else
                MsgDisp sMsg & sAddMsgStr, vbRed        ''メッセージ & 追加メッセージを画面表示する
            End If
        ''それ以外は黒表示
        Else
            If iMsgCd = 100 Then                        ''オラクルエラーの場合
                MsgDisp sMsg                            ''メッセージを画面表示する
            Else
                MsgDisp sMsg & sAddMsgStr               ''メッセージ & 追加メッセージを画面表示する
            End If
        End If
    End If
    
    If eMsgKind And ERR_LOG Then                    ''ログ出力ビットが立っていれば
        MsgLog (Format(Now, "YYYY/MM/DD HH:NN:SS::") & App.EXENAME & "::" & _
            iMsgCd & "::" & sMsg & "::" & sAddMsgStr) ''メッセージをログ出力する
    End If
    
    If (eMsgKind = ERR_DISP) Or _
       (eMsgKind = ERR_LOG) Or _
       (eMsgKind = ERR_DISP_LOG) Then                       ''ﾒｯｾｰｼﾞ属性がエラーなら
        Beep
    End If
End Sub


'///////////////////////////////////////////////////
' @(f)
' 機能    :ＤＢにコネクトする
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : なし
'
' 機能説明: ＤＢにコネクトする
'           ｺﾈｸﾄ先は、ｺﾏﾝﾄﾞﾗｲﾝ引数の工場ｺｰﾄﾞにより換える
'
'///////////////////////////////////////////////////
Public Function OraConn() As Boolean
    Dim sDbName As String
    Dim sUID As String
    Dim sPWD As String
    
    Select Case gsFactryCd
    Case "10"               ''野田工場
        sDbName = "NODA"
        sUID = "oracle"
        sPWD = "oracle"
    Case "30"               ''生野工場
        sDbName = "IKNO"
        sUID = "oracle"
        sPWD = "oracle"
    Case "40"               ''米沢工場
        sDbName = "YONE"
        sUID = "oracle"
        sPWD = "oracle"
    Case "42"               '’３００ｍｍ
        sDbName = "cm1"
        sUID = "cm1"
        sPWD = "cm1"
    Case "43"               '’３００ｍｍ
        sDbName = "cmt"
        sUID = "cm1"
        sPWD = "cm1"
    Case "90"               ''テスト環境
        sDbName = "TEST0"
        sUID = "oracle"
        sPWD = "oracle"
    Case "91"               ''テスト環境(新) 2007/04/05追加 SETsw kubota
                            ''テスト環境(米沢) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "92"               ''テスト環境(生野) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA0X"
        sUID = "oracle"
        sPWD = "oracle"
    Case "93"               ''テスト環境(生野A1) 2010/04/14追加 SETsw kubota
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "94"               ''テスト環境(尼崎A1) 2009/11/16追加 SSS.Marushita
        sDbName = "CLA1"
        sUID = "oracle"
        sPWD = "oracle"
    Case "99"               ''仮
        sDbName = "BOIS"
        sUID = "BOIS"
        sPWD = "BOIS"
    Case "AM"               ''尼崎工場 2009/06/02追加 SSS.Marushita
        sDbName = "CLK0"
        sUID = "oracle"
        sPWD = "oracle"
    Case Else               ''外販
        sDbName = "oracle"
        sUID = "oracle"
        sPWD = "oracle"
    End Select
    
    On Error GoTo ConnError
    
    ''オラクル接続
    Set gobjOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set gobjOraDB = gobjOraSess.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
    
    OraConn = True
    Exit Function
    
ConnError:
    OraConn = False
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    :ＤＢにコネクトする
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : なし
'
' 機能説明: ＤＢにコネクトする
'           ｺﾈｸﾄ先は、ｺﾏﾝﾄﾞﾗｲﾝ引数の工場ｺｰﾄﾞにより換える
'
'///////////////////////////////////////////////////
Public Function OraConn2() As Boolean
    Dim sDbName2 As String
    Dim sUID2 As String
    Dim sPWD2 As String
    
        sDbName2 = "DWH"
        sUID2 = "dwhmgr"
        sPWD2 = "dwhmgr"
    
    On Error GoTo ConnError2
    
    ''オラクル接続
    Set gobjOraSess = CreateObject("OracleInProcServer.XOraSession")
    Set gobjOraDB = gobjOraSess.OpenDatabase(sDbName2, sUID2 & "/" & sPWD2, 0&)
    
    OraConn2 = True
    Exit Function
    
ConnError2:
    OraConn2 = False
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    :ＤＢ開放
'
' 返り値  : 正常 - true
'           異常 - false
'
' 機能説明: ＤＢ開放
'
'///////////////////////////////////////////////////
Public Function OraDisConn() As Boolean
    
    On Error GoTo ErrProc
    
    ''オラクル切断
    gobjOraDB.Close
    
    ''解放
    Set gobjOraDB = Nothing
    Set gobjOraSess = Nothing
    
    OraDisConn = True
    Exit Function
    
ErrProc:
    OraDisConn = False
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    :オラクルダイナセットの作成
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : ARG1 - ダイナセットセットオブジェクト
'           ARG2 - SQL文
'           ARG3 - ダイナセットオプション
'
' 機能説明: オラクルダイナセット作成
'
'///////////////////////////////////////////////////
Public Function DynSet(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''オラクルダイナセット作成
    Set objOraDynaset = gobjOraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSet = True
    Exit Function
    
DynErr:
    DynSet = False
End Function
'///////////////////////////////////////////////////
' @(f)
' 機能    :オラクルダイナセットの作成
'
' 返り値  : 正常 - true
'           異常 - false
'
' 引き数  : ARG1 - ダイナセットセットオブジェクト
'           ARG2 - SQL文
'           ARG3 - ダイナセットオプション
'
' 機能説明: オラクルダイナセット作成
'
'///////////////////////////////////////////////////
Public Function DynSet2(ByRef objOraDynaset As Object, sSqlStmt As String, Optional vOpt = &H4&) As Boolean
    On Error GoTo DynErr
    
    ''オラクルダイナセット作成
    Set objOraDynaset = OraDB.CreateDynaset(sSqlStmt, vOpt)
    DynSet2 = True
    Exit Function
    
DynErr:
    DynSet2 = False
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ＳＱＬ文実行
'
' 返り値  : 0以上:処理件数
'           　-1：異常
'
' 引き数  : ARG1 - SQL文
'
' 機能説明: ＳＱＬ文実行し、処理件数を返す
'
'///////////////////////////////////////////////////
Public Function SqlExec(sSqlStmt As String) As Long
    On Error GoTo ErrProc
    
    ''オラクルＳＱＬ実行
    SqlExec = gobjOraDB.DbExecuteSQL(sSqlStmt)
    
    Exit Function
    
ErrProc:
    SqlExec = -1
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ＳＱＬ文実行
'
' 返り値  : 0以上:処理件数
'           　-1：異常
'
' 引き数  : ARG1 - SQL文
'
' 機能説明: ＳＱＬ文実行し、処理件数を返す
'
'///////////////////////////////////////////////////
Public Function SqlExec2(sSqlStmt As String) As Long
    On Error GoTo ErrProc
    
    ''オラクルＳＱＬ実行
    SqlExec2 = OraDB.DbExecuteSQL(sSqlStmt)
    
    Exit Function
    
ErrProc:
    SqlExec2 = -1
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 担当者名取得
'
' 返り値  : True:正常
'           False:失敗
'
' 引き数  : 担当者ｺｰﾄﾞ
'
' 機能説明: 担当者ｺｰﾄﾞから担当者名を取得
'
'///////////////////////////////////////////////////
Public Function GetUserName(sUserCd As String, ByRef sUserName As String) As Boolean
    Dim sSqlStmt As String
    Dim objOraDyn As Object
    
    sUserName = vbNullString       ''担当者名クリアー
    ''ＳＱＬ文作成
    sSqlStmt = "SELECT NVL(nameja9, ' ')                    "
    sSqlStmt = sSqlStmt & "FROM koda9                       "
    sSqlStmt = sSqlStmt & "WHERE sysca9 = 'K'               "
    sSqlStmt = sSqlStmt & "  AND shuca9 = '55'              "
    sSqlStmt = sSqlStmt & "  AND codea9 = '" & sUserCd & "' "
    
    ''ダイナセット作成
    If DynSet(objOraDyn, sSqlStmt) = False Then
        ''ダイナセット作成失敗
        Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
        
        GetUserName = False
        Exit Function
    End If
    If objOraDyn.EOF Then
        ''該当する担当者ｺｰﾄﾞが無かった
        Call MsgOut(8, "", ERR_DISP)
        
        GetUserName = False
        Exit Function
    End If

    sUserName = objOraDyn(0)  ''担当者名取得
    
    GetUserName = True        ''処理成功を返す
    
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 品番編集
' 返り値  : True:正常
'           False:失敗
' 引き数  :
'
' 機能説明: キー項目編集
'　　　　　 品番を‘－’で分割編集または逆編集を行う。
'
'///////////////////////////////////////////////////
'Public Function GetHinbanHensyu(sbHinban As String, sflg As Integer, ByRef sahinban As String) As Boolean
'
'    '　"-" なしから　"-"有りに編集
'        If sflg = 1 Then
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
'        End If
'
'    '　"-" 有りから　"-"無しに編集
'        If sflg = 2 Then
'            sahinban = Replace(sbHinban, "-", "")
'            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
'            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
'            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
'        End If
'
'    GetHinbanHensyu = True        ''処理成功を返す
'
'End Function


Public Function GetHinbanHensyu(sbHinban As String, sFlg As Integer, ByRef sahinban As String) As Boolean
    If sbHinban = "G" Or sbHinban = "Z" Then
    '　"-" なしから　"-"有りに編集
        If sFlg = 1 Then
            sahinban = Format(sbHinban, "@")
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
        End If
        
    '　"-" 有りから　"-"無しに編集
        If sFlg = 2 Then
            sahinban = Replace(sbHinban, "-", "")
            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
        End If
    Else
    '　"-" なしから　"-"有りに編集
        If sFlg = 1 Then
            sahinban = Format(sbHinban, "@@@-@@@@-@")
'            sahinban = Format(sbHinban, "@@@-@@@@-@@@@")
        End If
        
    '　"-" 有りから　"-"無しに編集
        If sFlg = 2 Then
            sahinban = Replace(sbHinban, "-", "")
            'Mid(saHinban, 1, 6) = Mid(sbHinban, 1, 6)
            'Mid(saHinban, 7, 2) = Mid(sbHinban, 8, 2)
            'Mid(saHinban, 9, 3) = Mid(sbHinban, 11, 3)
        End If
    End If
    GetHinbanHensyu = True        ''処理成功を返す
    
End Function
                                                                                                                                                                                         
'///////////////////////////////////////////////////
' @(f)
' 機能    : 製番編集
' 返り値  : True:正常
'           False:失敗
' 引き数  :
'
' 機能説明: キー項目編集
'　　　　　 製番を‘－’で分割編集または逆編集を行う。
'
'///////////////////////////////////////////////////
Public Function GetSeibanHensyu(sbSeiban As String, sFlg As Integer, ByRef saSeiban As String) As Boolean
    
    '　"-" なしから　"-"有りに編集
        If sFlg = 1 Then
            saSeiban = Format(sbSeiban, "@@-@@@@@")
        End If
        
    '　"-" 有りから　"-"無しに編集
        If sFlg = 2 Then
            saSeiban = Replace(sbSeiban, "-", "")
            'Mid(saSeiban, 1, 2) = Mid(sbSeiban, 1, 2)
            'Mid(saSeiban, 3, 5) = Mid(sbSeiban, 4, 5)
        End If
    
    GetSeibanHensyu = True        ''処理成功を返す
    
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 結晶番号編集
' 返り値  : True:正常
'           False:失敗
' 引き数  :
'
' 機能説明: キー項目編集
'　　　　　 結晶番号を‘－’で分割編集または逆編集を行う。
'
'///////////////////////////////////////////////////
Public Function GetXtalHensyu(sbXtal As String, sFlg As Integer, ByRef saXtal As String) As Boolean
Dim wXTAL As String
    
    '　"-" なしから　"-"有りに編集
        If sFlg = 1 Then
            wXTAL = sbXtal
            If Len(sbXtal) > 0 Then wXTAL = wXTAL & "000000000000"
            saXtal = Format(Mid(wXTAL, 1, 12), "@@@@-@@@@@-@@@")
        End If
        
    '　"-" 有りから　"-"無しに編集
        If sFlg = 2 Then
            saXtal = Replace(sbXtal, "-", "")
            'Mid(saXtal, 1, 4) = Mid(sbXtal, 1, 4)
            'Mid(saXtal, 5, 3) = Mid(sbXtal, 6, 3)
            'Mid(saXtal, 8, 2) = Mid(sbXtal, 10, 2)
            'Mid(saXtal, 10, 3) = Mid(sbXtal, 13, 3)
        End If
        
        sFlg = 0
        
    GetXtalHensyu = True        ''処理成功を返す
    
End Function

' @(f)
'
' 機能      : マウスポインタ変更
'
' 返り値    : なし
'
' 引数      : ipID   -   0:標準
'                       1:砂時計
'
' 機能説明  : マウスポインタ変更
'
Public Sub SetMousePointer(ipID%)
    Select Case ipID
    Case 0
        Screen.MousePointer = vbDefault    ''マウスポインタ標準
    Case 1
        Screen.MousePointer = vbHourglass  ''マウスポインタ砂時計
    Case Else
        Screen.MousePointer = vbDefault    ''マウスポインタ標準
    End Select
End Sub


' @(f)
'
' 機能      : 日付チェック
'
' 返り値  : OK - TRUE
'           NG - FALSE
'
' 引数      : sDate - (String)日付
'             iKind - 0:yyyymmdd形式
'             iKind - 1:yymmdd形式
'             iKind - 2:そのまま変換可能な形式
'
' 機能説明  : 日付に変換可能かチェックする
'
Public Function DateCheck(sDate$, iKind%) As Boolean
    Dim sCheckDate As String
    
    Select Case iKind
    Case 0
        sCheckDate = Mid(sDate, 1, 4) & "/" & Mid(sDate, 5, 2) & "/" & Mid(sDate, 7)
    Case 1
        sCheckDate = Mid(sDate, 1, 2) & "/" & Mid(sDate, 3, 2) & "/" & Mid(sDate, 5)
    Case Else
        sCheckDate = sDate
    End Select
    
    If IsDate(sCheckDate) Then
        DateCheck = True
    Else
        DateCheck = False
    End If
End Function


' @(f)
' 機能    : 期間チェック
'
' 返り値  :  OK - TRUE
'            NG -FALSE
'
' 引き数  : ctlControlS : コントロール(開始日)
'           ctlControlE : コントロール(終了日)
'           sDateS      : 開始日
'           sDateE      : 終了日
'
' 機能説明: 期間チェックを行い8桁の年月日を返す
'Update - 2000/02/15
Public Function KikanCheck(ctlControlS As Control, ctlControlE As Control, _
        ByRef sDateS$, ByRef sDateE$) As Boolean
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlDWHCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    Dim sDtS    As String       ''集計期間開始日
    Dim sDtE    As String       ''集計期間終了日
    Dim sDtT    As String       ''システム日付
    Dim sDtL    As String       ''該当月の月末日(開始日)
    Dim sDtLE   As String       ''該当月の月末日(終了日)
    Dim sWk     As String
    
    KikanCheck = False
    
    ''システム日付取得([yymmdd])
    sDtT = Format(Date, "yymmdd")

    ''集計期間取得(6桁[yymmdd])
    sDtS = Trim(ctlControlS.text)
    sDtE = Trim(ctlControlE.text)
    
    
    ''開始日の桁チェック
    If Len(sDtS) = 0 Then
        If Len(sDtE) <> 0 Then
            '''終了日のみ入力はエラー
            Call MsgOut(0, "期間の開始日を入力して下さい", ERR_DISP)
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Exit Function
        End If
        '''未入力は当月初日～当日を設定
        sDtS = Mid(sDtT, 1, 4) & "00"
    
    ElseIf Mid(sDtS, 5) = "00" Then
        If Len(sDtE) <> 0 Then
            Call MsgOut(0, "期間の終了日の入力は必要ありません", ERR_DISP)
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Exit Function
        End If
    
    ElseIf Not DateCheck(sDtS, 1) Then
    
    
    
    
    
    ElseIf Mid(sDtS, 5) <> "00" Then
        If Len(sDtE) = 0 Then
            sDtE = sDtT
        End If
    
    End If
    
    
    ''開始日の日付チェック
    If DateCheck(Mid(sDtS, 1, 4) & "01", 1) Then
        If Mid(sDtS, 1, 4) = Mid(sDtT, 1, 4) Then
            '''該当月が当月の場合は月末日を当日に設定
            sDtL = sDtT
        Else
            '''該当月の月末日算出([yymmdd])
            sWk = DateAdd("m", 1, Mid(sDtS, 1, 2) & "/" & Mid(sDtS, 3, 2) & "/" & "01")
            sDtL = DateAdd("d", -1, sWk)
''            sDtLE = Format(sDtLE, "yy/mm/dd") '2003/10/24 tuku SUMCO殿改造内容追加
            sDtL = Format(sDtL, "yy/mm/dd") '2004/11/26 日付フォーマット不具合対応
            sDtL = Mid(sDtL, 1, 2) & Mid(sDtL, 4, 2) & Mid(sDtL, 7)
'            sDtL = Mid(sDtL, 3, 2) & Mid(sDtL, 6, 2) & Mid(sDtL, 9)
        End If
        Call CtrlEnabled(ctlControlS, NORMAL_CTL)
    Else
        '''日付に変換できない場合はエラー
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    
   
    If Len(sDtE) = 0 Then
        '''開始日の日項目が[00]の場合は該当月の初日から月末を設定
        '''(当月の場合は初日から当日)
        sDtE = sDtL
    End If
    
    
    ''終了日の日付チェック
    If DateCheck(Mid(sDtE, 1, 4) & "01", 1) Then
        If Mid(sDtE, 1, 4) = Mid(sDtT, 1, 4) Then
            '''該当月が当月の場合は月末日を当日に設定
            sDtLE = sDtT
        Else
            '''該当月の月末日算出([yymmdd])
            sWk = DateAdd("m", 1, Mid(sDtE, 1, 2) & "/" & Mid(sDtE, 3, 2) & "/" & "01")
            sDtLE = DateAdd("d", -1, sWk)
            sDtLE = Format(sDtLE, "yy/mm/dd") '2003/10/24 tuku SUMCO殿改造内容追加
'            sDtLE = Mid(sDtLE, 3, 2) & Mid(sDtLE, 6, 2) & Mid(sDtLE, 9)
            sDtLE = Mid(sDtLE, 1, 2) & Mid(sDtLE, 4, 2) & Mid(sDtLE, 7)
        End If
        Call CtrlEnabled(ctlControlE, NORMAL_CTL)
    Else
        '''日付に変換できない場合はエラー
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    
    ''集計期間の日付変換
    If Mid(sDtS, 5) = "00" Then
        '''開始日の日項目が[00]の場合は該当月の初日から月末を設定
        '''(当月の場合は初日から当日)
        sDtS = Mid(sDtS, 1, 4) & "01"
        sDtE = sDtL
    
    ElseIf Mid(sDtS, 5) > Mid(sDtL, 5) Then
        '''開始日が月末より大きい場合は開始日に月末日を設定し
        '''終了日が未入力の場合は当日を設定する
        sDtS = sDtL
        If Len(sDtE) = 0 Then
            sDtE = sDtT
        ElseIf Mid(sDtE, 5) > Mid(sDtLE, 5) Then
            sDtE = sDtLE
        End If
    
    Else
        '''それ以外の場合は集計期間の日付チェックを行う
        If Mid(sDtE, 5) > Mid(sDtLE, 5) Then
            '''終了日が月末より大きい場合は終了日に月末日を設定
            sDtE = sDtLE
        End If
'*********************
        If Mid(sDtE, 5) = "00" Then
        '''開始日の日項目が[00]の場合は該当月の初日から月末を設定
        '''(当月の場合は初日から当日)
            sDtE = Mid(sDtE, 1, 4) & "01"
        End If
'**********************
        If Not DateCheck(sDtS, 1) Then
            '''日付に変換できない場合はエラー(開始日)
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
        If Not DateCheck(sDtE, 1) Then
            '''日付に変換できない場合はエラー(終了日)
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    End If
    
    ''年月日の桁合わせ
    ''テスト用に1900年代も対応した(2000を足すだけだと1900年代に対応できない)
'    sDtS = "20" & sDtS
'    sDtE = "20" & sDtE
    sDtS = DateChange(sDtS)
    sDtE = DateChange(sDtE)
    
    ''日付の範囲チェック
    If val(sDtS) > val(sDtE) Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(53, "", ERR_DISP)
        Exit Function
    End If
    
    sDateS = sDtS
    sDateE = sDtE
    
    KikanCheck = True

End Function


' @(f)
' 機能    : 日付変換
'
' 返り値  : 変換後日付
'
' 引き数  : sDate - 日付
'
' 機能説明: 6桁[yymmdd]の日付を8桁[yyyymmdd]にする
'
Public Function DateChange$(sDate$)
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlDWHCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    If Mid(sDate, 1, 2) < 50 Then
        DateChange = "20" & sDate
    Else
        DateChange = "19" & sDate
    End If
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : チャージ量関数
'
' 返り値  : チャージ量
'
' 引き数  : 推定ﾁｬｰｼﾞ量
'           ﾄｯﾌﾟｶｯﾄ重量
'           肩ｶｯﾄ重量
'
' 機能説明: チャージ量の計算
'
'  チャージ量 = 推定ﾁｬｰｼﾞ量 - ﾄｯﾌﾟｶｯﾄ重量 - 肩ｶｯﾄ重量
'
'///////////////////////////////////////////////////
Public Function CHARGEWEIGHT(lSuiteiChargeWeight As Long, _
                      lTopCutWeight As Long, _
                      lShoulderCutWeight As Long _
                      ) As Long
    CHARGEWEIGHT = lSuiteiChargeWeight - lTopCutWeight - lShoulderCutWeight
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 偏析値関数
'
' 返り値  : 偏析値
'
' 引き数  : TOP実測抵抗
'           BOT実測抵抗
'           切断後重量
'           チャージ量
'
' 機能説明: 偏析値の計算
'
'           log(TOP実測抵抗 / BOT実測抵抗)
'  偏析値 = ──────────────── + 1
'           log(1 - 切断後重量 / チャージ量)
'
'           使用条件：予めチャージ量を計算しておく
'
'///////////////////////////////////////////////////
Public Function Henseki(dTopRes As Double, _
                 dBotRes As Double, _
                 lCutAfterWeight As Long, _
                 lChargeWeight As Long _
                 ) As Double
    Dim dVal As Double              ''途中計算用
    
    ''ゼロチェック
    If dBotRes = 0 Then             ''BOT実測抵抗ゼロチェック
        Call MsgOut(14, "BOT実測抵抗", ERR_DISP)
        Exit Function
    ElseIf dTopRes = 0 Then         ''TOP実測抵抗ゼロチェック
        Call MsgOut(14, "TOP実測抵抗", ERR_DISP)
        Exit Function
    ElseIf lChargeWeight = 0 Then   ''チャージ量ゼロチェック
        Call MsgOut(14, "チャージ量", ERR_DISP)
        Exit Function
    End If
        
    dVal = 1 - lCutAfterWeight / lChargeWeight ''1 - 切断後重量 / チャージ量
    If dVal = 0 Then                ''ゼロチェック
        Call MsgOut(14, "1-切断後重量/ﾁｬｰｼﾞ量", ERR_DISP)
        Exit Function
    ElseIf dVal < 0 Then
        Call MsgOut(14, "切断後重量 > ﾁｬｰｼﾞ量", ERR_DISP)
        Exit Function
    End If
    
    dVal = Log(dVal)                ''log(1 - 切断後重量 / チャージ量)
    If dVal = 0 Then                ''ゼロチェック
        Call MsgOut(14, "log(1-切断後重量/ﾁｬｰｼﾞ量)", ERR_DISP)
        Exit Function
    End If
    
    Henseki = Log(dTopRes / dBotRes) / dVal + 1
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 推定位置関数
'
' 返り値  : 推定位置（合格位置）
'
' 引き数  : 目標抵抗
'           TOP実測抵抗
'           偏析値
'           切断後重量
'           チャージ量
'           切断後長さ
'           ゼロ除算防止エラーログ出力フラグ
'
' 機能説明: 推定位置の計算
'                                           1
'                                       ─────
'                                       1 - 偏析値
'             1 - (目標抵抗 / TOP実測抵抗)
'  推定位置 = ────────────────────
'              切断後重量 / チャージ量 / 切断後長さ
'
'           使用条件：予めチャージ量と偏析値を計算しておく
'
'///////////////////////////////////////////////////
Public Function SuiteiIchi(dTargetRes As Double, _
                            dTopRes As Double, _
                            dHenseki As Double, _
                            lCutAfterWeight As Long, _
                            lChargeWeight As Long, _
                            iCutAfterSize As Integer, _
                            Optional bErrLogFlg As Boolean = True _
                            ) As Double
    
    If (1 - dHenseki) = 0 Then                 ''(1-偏析値)ゼロチェック
        If bErrLogFlg Then Call MsgOut(14, "1-偏析値", ERR_DISP)
        Exit Function
    ElseIf dTopRes = 0 Then                    ''TOP実測抵抗ゼロチェック
        If bErrLogFlg Then Call MsgOut(14, "TOP実測抵抗", ERR_DISP)
        Exit Function
    ElseIf lChargeWeight = 0 Then              ''チャージ量ゼロチェック
        If bErrLogFlg Then Call MsgOut(14, "チャージ量", ERR_DISP)
        Exit Function
    ElseIf iCutAfterSize = 0 Then              ''切断後長さゼロチェック
        If bErrLogFlg Then Call MsgOut(14, "切断後長さ", ERR_DISP)
        Exit Function
    End If
    
    SuiteiIchi = (CLng(1) - (dTargetRes / dTopRes) ^ (CLng(1) / (CLng(1) - dHenseki))) / _
                 (CLng(lCutAfterWeight) / CLng(lChargeWeight) / CLng(iCutAfterSize))
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 推定重量関数
'
' 返り値  : 推定重量（合格重量）
'
' 引き数  : 推定位置
'           チャージ量
'           切断後長さ
'
' 機能説明: 推定重量の計算
'
'  推定重量 = 推定位置 × チャージ量 / 切断後長さ
'
'           使用条件：予めチャージ量と推定位置を計算しておく
'
'///////////////////////////////////////////////////
Public Function SuiteiWeight(sSuiteiIchi As Double, _
                            lChargeWeight As Long, _
                            iCutAfterSize As Integer _
                            ) As Double
    
    If iCutAfterSize = 0 Then                   ''切断後長さゼロチェック
        Call MsgOut(14, "切断後長さ", ERR_DISP)
        Exit Function
    End If
    
    SuiteiWeight = sSuiteiIchi * lChargeWeight / iCutAfterSize
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 推定抵抗率関数
'
' 返り値  : 推定抵抗率
'
' 引き数  : 推定位置
'           切断後重量
'           チャージ量
'           切断後長さ
'           偏析値
'           TOP実測抵抗
'
' 機能説明: 推定抵抗率の計算
'                                                        (1 - 偏析値)
'  Ａ = (1 - 推定位置 × 切断後重量 / チャージ量 / 切断後長さ)
'
'  推定抵抗 = TOP実測抵抗 × Ａ
'
'           使用条件：予めチャージ量と偏析値と推定位置を計算しておく
'
'///////////////////////////////////////////////////
Public Function SuiteiRes(iSuiteiIchi As Integer, _
                        lCutAfterWeight As Long, _
                        lChargeWeight As Long, _
                        iCutAfterSize As Integer, _
                        dHenseki As Double, _
                        dTopRes As Double _
                        ) As Double
    Dim dA As Double
    Dim db As Double
    
    If lChargeWeight = 0 Then                   ''チャージ量ゼロチェック
        Call MsgOut(14, "チャージ量", ERR_DISP)
        Exit Function
    ElseIf iCutAfterSize = 0 Then               ''切断後長さゼロチェック
        Call MsgOut(14, "切断後長さ", ERR_DISP)
        Exit Function
    End If
    
    db = 1 - iSuiteiIchi * lCutAfterWeight / lChargeWeight / iCutAfterSize
    dA = CDbl(db) ^ (1 - dHenseki)
    SuiteiRes = dTopRes * dA
    SuiteiRes = val(Format(SuiteiRes, "######0.0######"))
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 結晶1㎜重量関数
'
' 返り値  : 結晶1㎜の重量
'
' 引き数  : 直径mm
'
' 機能説明: 結晶1㎜当りの重量の計算
'                            2
'                 （直径 / 2) × 3.14 × 2.33
'  結晶1㎜の重量 = ─────────────
'                            1000
'
'///////////////////////////////////////////////////
Public Function WeightPar1mm(sChokkei As Single) As Double
    WeightPar1mm = (((sChokkei / 2) ^ 2) * 3.14 * 2.33) / 1000
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 結晶重量関数
'
' 返り値  : 結晶の重量
'
' 引き数  : 直径mm
'           長さmm
'
' 機能説明: 直径と長さにより結晶重量を計算する
'
'  結晶重量 = 結晶1㎜重量関数 × 長さ
'
'///////////////////////////////////////////////////
Public Function WeightCompute(sChokkei As Single, sNagasa As Single) As Double
    WeightCompute = WeightPar1mm(sChokkei) * sNagasa
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : 上限ねらい抵抗関数
'
' 返り値  : 上限ねらい抵抗
'
' 引き数  : 規格抵抗上限
'           上限内側管理(パーセント)
'
'
' 機能説明: 上限ねらい抵抗の計算
'
' 上限ねらい抵抗 = (規格抵抗上限 - 規格抵抗上限 × 上限内側管理 × 0.01) × 0.97
'
'///////////////////////////////////////////////////
Public Function UperTergetRes(dUperRes As Double, iUperInPar As Integer) As Double
    UperTergetRes = (dUperRes - dUperRes * (iUperInPar * 0.01)) * 0.97
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ねらい抵抗に対応する不純物濃度関数
'
' 返り値  : 不純物濃度
'
' 引き数  : ねらい抵抗
'           係数Ａ
'           係数Ｂ
'
' 機能説明: ねらい抵抗に対応する不純物濃度計算
'
'                                     1       ((Log(ねらい抵抗)－係数Ｂ)÷係数Ａ)
' ねらい抵抗に対応する不純物濃度 = ─── ×10
'                                   2.33
'///////////////////////////////////////////////////
Public Function DopantPar1g(sngTergetRes As Single, _
                            sngA As Single, sngB As Single) As Single
    If sngA = 0 Then   ''係数Ａゼロチェック
        Call MsgOut(14, "係数Ａ", ERR_DISP)
        Exit Function
    End If
    DopantPar1g = (1 / 2.33) * 10 ^ ((Log(sngTergetRes) - sngB) / sngA)
End Function


'///////////////////////////////////////////////////
' @(f)
' 機能    : ドーパント量の計算
'
' 返り値  : ドーパント量
'
' 引き数  : チャージ量
'           ねらい抵抗に対応する不純物濃度
'           不純物濃度
'           ファクター
'           偏析値
'
' 機能説明: 不純物（ドーパント）量関数
'
'                チャージ量(g)×ねらい抵抗に対応する不純物濃度
' ドーパント量 = ─────────────────────── × ファクター × 偏析値
'                                   不純物濃度
'///////////////////////////////////////////////////
Public Function DopantWeight(sngCharge As Single, sngDopantPar1g As Single, _
                             sngDopant As Single, sngFactor As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngDopant = 0 Then   ''不純物濃度ゼロチェック
        Call MsgOut(14, "不純物濃度", ERR_DISP)
        Exit Function
    End If
    DopantWeight = ((sngCharge * sngDopantPar1g) / sngDopant) * sngFactor * sngHenseki
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : ドーパント量の計算２
'
' 返り値  : ドーパント量(Ｍｅｔｌ使用）
'
' 引き数  : チャージ量
'           ねらい抵抗に対応する不純物濃度
'           分子量
'           偏析値
'
' 機能説明: 不純物（ドーパント）量関数２
'
'                チャージ量(g)×ねらい抵抗に対応する不純物濃度 × 分子量         １
' ドーパント量 = ───────────────────────────── × ─────
'                                   0.6 × 10＾23                           偏析値
'
'///////////////////////////////////////////////////
Public Function DopantWeight2(sngCharge As Single, sngDopantPar1g As Single, _
                             sngBunsi As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngHenseki = 0 Then   ''偏析値ゼロチェック
        Call MsgOut(14, "偏析値", ERR_DISP)
        Exit Function
    End If
    DopantWeight2 = ((sngCharge * sngDopantPar1g * sngBunsi) / (0.6 * 10 ^ 23)) / sngHenseki
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : ドーパント量の計算３
'
' 返り値  : ドーパント量(リメルト使用）
'
' 引き数  : リメルト量
'           リメルト平均イオン濃度
'           不純物濃度
'           分子量
'           偏析値
'
' 機能説明: 不純物（ドーパント）量関数３
'
'                リメルト量(g)×リメルト平均イオン濃度
' ドーパント量 = ──────────────────── × ファクター × 偏析値
'                                   不純物濃度
'///////////////////////////////////////////////////
Public Function DopantWeight3(sngCharge As Single, sngDopantPar1g As Single, _
                             sngDopant As Single, sngFactor As Single, _
                             Optional sngHenseki As Single = 1) As Single
    If sngDopant = 0 Then   ''不純物濃度ゼロチェック
        Call MsgOut(14, "不純物濃度", ERR_DISP)
        Exit Function
    End If
    DopantWeight3 = ((sngCharge * sngDopantPar1g) / sngDopant) * sngFactor * sngHenseki
End Function


' @(f)
'
' 機能      : 桁詰処理
'
' 返り値    : なし
'
' 引き数    : ctlControl    -   コントロール
'             sChar         -   桁詰文字列
'
' 機能説明  : 指定した文字列でMaxLengthに足りない文字数分埋める
'
' 備考      :
'
Public Sub FillUpString(ctlControl As Control, sChar As String)
    Dim iLength As Integer
    If TypeOf ctlControl Is TextBox Then
        iLength = LenB(StrConv(Trim(ctlControl.text), vbFromUnicode))
        If iLength < ctlControl.MaxLength Then
            ctlControl.text = Trim(ctlControl.text) & String(ctlControl.MaxLength - iLength, sChar)
        End If
    End If
End Sub


' @(f)
'
' 機能      : 結晶RRG算出処理
'
' 返り値    : RRG値
'
' 引数      : CalcType  - 計算種別(保証方法)
'             TeikouDat - 抵抗値一覧
'
' 機能説明  :　計算種別(保証方法)によりRRGを算出する処理をする
'
' 備考      : RRG値が[999999]の場合、RRG計算NG又は該当保証方法無し
'
Public Function GetCalcRRG(iCalcType As Integer, tTeikouDat As TYPE_RRG) As Double
    Dim dMaxNum As Double
    Dim dMinNum As Double
    Dim dHeikin As Double
    Dim dChuou As Double
    Dim dRRG As Double
    
    '' 計算式フラグ判定                <--- 2000/04/24 追加
    If iCalcType < 1 Or iCalcType > 7 Then
        GetCalcRRG = NULL_CHECK
        Exit Function
    End If
    
    '' 最大値取得
    dMaxNum = GetMax(iCalcType, tTeikouDat)
    '' 抵抗値MAX判定
    If dMaxNum = NULL_CHECK Then
        GetCalcRRG = dMaxNum
        Exit Function
    End If
    
    '' 最小値取得
    dMinNum = GetMin(tTeikouDat)
    '' 抵抗値MIN判定
    If dMinNum = NULL_CHECK Then
        GetCalcRRG = dMinNum
        Exit Function
    End If
    
    '' 平均値取得
    dHeikin = GetHeikin(tTeikouDat)
    
    '' 中央値取得
    dChuou = GetChuou(tTeikouDat, dHeikin)
    dRRG = 0
    
    '' 計算不可判定処理(分子又は分母が０となる場合)
    If dMaxNum = 0 Or dMinNum = 0 Or (dMaxNum - dMinNum) = 0 Then
        GetCalcRRG = NULL_CHECK
        Exit Function
    End If
    
    '' 保証方法によりRRG算出を分岐する
    Select Case iCalcType
    Case 1
        dRRG = (dMaxNum - dMinNum) / dMaxNum * 100
    Case 2
        dRRG = (dMaxNum - dMinNum) / dMinNum * 100
    Case 3
        dRRG = (dMaxNum - dMinNum) / dChuou * 100
    Case 4
        dRRG = (dMaxNum - dMinNum) / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = (dMaxNum - dMinNum) / tTeikouDat.dTeikou(0) * 100
    Case 5
        dRRG = dMaxNum / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = dMaxNum / tTeikouDat.dTeikou(0) * 100
    Case 6
        dRRG = dMaxNum / tTeikouDat.dTeikouDT(0).dTeikou * 100
        '' dRRG = dMaxNum / tTeikouDat.dTeikou(0) * 100
    Case 7
        dRRG = Abs(tTeikouDat.dTeikouDT(0).dTeikou - dHeikin) / (tTeikouDat.dTeikouDT(0).dTeikou + dHeikin) / 2 * 100
        '' dRRG = Abs(tTeikouDat.dTeikou(0) - dHeikin) / (tTeikouDat.dTeikou(0) + dHeikin) / 2 * 100
    Case Else
        dRRG = NULL_CHECK
    End Select

    ' 桁合わせ処理(4桁に調整)
    GetCalcRRG = left(dRRG, 4)
End Function

' @(f)
'
' 機能      : 最大値取得
'
' 返り値    : 最大値
'
' 引数      : CalcType  - 計算種別(保証方法)
'             TeikouDat - 抵抗値一覧
'
' 機能説明  :　抵抗値の最大値を検索する処理をする
'
' 備考      :
'
Private Function GetMax(iCalcType As Integer, tTeikouDat As TYPE_RRG) As Double
    Dim dMaxData As Double
    Dim iCntI As Integer
    Dim iNextCnt As Integer
    
    iNextCnt = 0
    iCntI = 0
    
    With tTeikouDat
        '' 保証方法により分岐
        Select Case iCalcType
        Case 5
            '' フラグチェック
            If .dTeikouDT(0).sRRGFlg = "1" Then
                If .dTeikouDT(1).sRRGFlg = "1" And .dTeikouDT(1).dTeikou <> NULL_CHECK Then
                    dMaxData = .dTeikouDT(1).dTeikou - .dTeikouDT(0).dTeikou
                    iNextCnt = 2
                Else
                    For iNextCnt = 2 To SOKUTEI_MAX
                    If .dTeikouDT(iNextCnt).sRRGFlg = "1" And .dTeikouDT(iNextCnt).dTeikou <> NULL_CHECK Then
                        dMaxData = .dTeikouDT(iNextCnt).dTeikou - .dTeikouDT(0).dTeikou
                        Exit For
                    End If
                    Next iNextCnt
                End If
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = .dTeikou(1) - .dTeikou(0)
            For iCntI = iNextCnt To SOKUTEI_MAX
                '' フラグチェック                 <--- 2000/0424 変更
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' 最大値を比較する
                    If dMaxData < .dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' 大きい場合、最大値を更新する
                        dMaxData = .dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou
                    End If
                End If
                '' 計算フラグ追加により変更        <--- 2000/04/24 削除
                '' If dMaxData < .dTeikou(iCntI) - .dTeikou(0) And .dTeikou(iCntI) <> 999999 Then
                ''    '' 大きい場合、最大値を更新する
                ''    dMaxData = .dTeikou(iCntI) - .dTeikou(0)
                '' End If
            Next iCntI
        Case 6
            '' フラグチェック                     <--- 2000/04/24 変更
            If .dTeikouDT(0).sRRGFlg = "1" Then
                If .dTeikouDT(1).sRRGFlg = "1" And .dTeikouDT(1).dTeikou <> NULL_CHECK Then
                   dMaxData = Abs(.dTeikouDT(1).dTeikou - .dTeikouDT(0).dTeikou)
                Else
                    For iNextCnt = 2 To SOKUTEI_MAX
                    If .dTeikouDT(iNextCnt).sRRGFlg = "1" And .dTeikouDT(iNextCnt).dTeikou <> NULL_CHECK Then
                        dMaxData = Abs(.dTeikouDT(iNextCnt).dTeikou - .dTeikouDT(0).dTeikou)
                        Exit For
                    End If
                    Next iNextCnt
                End If
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = Abs(.dTeikou(1) - .dTeikou(0))
            For iCntI = 2 To SOKUTEI_MAX
                '' フラグチェック                  <--- 2000/04/24 変更
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' 最大値を比較する
                    If dMaxData < Abs(.dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou) And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' 大きい場合、最大値を更新する
                        dMaxData = Abs(.dTeikouDT(iCntI).dTeikou - .dTeikouDT(0).dTeikou)
                    End If
                End If
                '' 計算フラグ追加により変更        <--- 2000/04/24 削除
                '' If dMaxData < Abs(.dTeikou(iCntI) - .dTeikou(0)) And .dTeikou(iCntI) <> 999999 Then
                ''     '' 大きい場合、最大値を更新する
                ''     dMaxData = Abs(.dTeikou(iCntI) - .dTeikou(0))
                '' End If
            Next iCntI
        Case Else
            '' フラグチェック
            If .dTeikouDT(0).sRRGFlg = "1" Then
                dMaxData = .dTeikouDT(0).dTeikou
            Else
                dMaxData = NULL_CHECK
            End If
            '' dMaxData = .dTeikou(0)
            '' 測定個所数繰り返す(最大9回)
            For iCntI = 1 To SOKUTEI_MAX
                '' フラグチェック                   <--- 2000/04/24 変更
                If .dTeikouDT(iCntI).sRRGFlg = "1" Then
                    '' 最大値を比較する
                    If dMaxData < .dTeikouDT(iCntI).dTeikou And .dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                        '' 大きい場合、最大値を更新する
                        dMaxData = .dTeikouDT(iCntI).dTeikou
                    End If
                End If
                '' 計算フラグ追加により変更          <--- 2000/04/24 削除
                '' If dMaxData < .dTeikou(iCntI) And .dTeikou(iCntI) <> 999999 Then
                ''     '' 大きい場合、最大値を更新する
                ''     dMaxData = .dTeikou(iCntI)
                '' End If
            Next iCntI
        End Select
    End With
    GetMax = dMaxData
End Function

' @(f)
'
' 機能      : 最小値取得
'
' 返り値    : 最小値
'
' 引数      : TeikouDat - 抵抗値一覧
'
' 機能説明  :　抵抗値の最小値を検索する処理をする
'
' 備考      :
'
Private Function GetMin(tTeikouDat As TYPE_RRG) As Double
    Dim dMineData As Double
    Dim iCntI As Integer
    Dim bChkFlg As Boolean
    
    bChkFlg = False
    iCntI = 0
    dMineData = tTeikouDat.dTeikouDT(0).dTeikou
    '' dMineData = tTeikouDat.dTeikou(0)
    
    '' 測定個所数繰り返す(最大9回)
    For iCntI = 1 To SOKUTEI_MAX
        '' フラグチェック                   <--- 2000/04/24 変更
        If tTeikouDat.dTeikouDT(iCntI).sRRGFlg = "1" Then
            If tTeikouDat.dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                '' 抵抗値がある場合、抵抗値を比較する
                If dMineData >= tTeikouDat.dTeikouDT(iCntI).dTeikou Then
                    '' 小さい場合、最小値を更新する
                    dMineData = tTeikouDat.dTeikouDT(iCntI).dTeikou
                End If
                bChkFlg = True
            End If
        End If
        '' 計算フラグ追加により変更          <--- 2000/04/24 削除
        '' If tTeikouDat.dTeikou(iCntI) <> NULL_CHECK Then
        ''     '' 抵抗値がある場合、抵抗値を比較する
        ''     If dMineData > tTeikouDat.dTeikou(iCntI) Then
        ''         '' 小さい場合、最小値を更新する
        ''         dMineData = tTeikouDat.dTeikou(iCntI)
        ''     End If
        '' End If
    Next iCntI
    If bChkFlg = False Then
        GetMin = NULL_CHECK
    Else
        GetMin = dMineData
    End If
End Function

' @(f)
'
' 機能      : 中央値取得
'
' 返り値    : 中央値
'
' 引数      : TeikouDat - 抵抗値一覧
'             Heikin    - 平均値
'
' 機能説明  :　抵抗値の中央値を取得する処理をする
'
' 備考      :
'
Private Function GetChuou(tTeikouDat As TYPE_RRG, dHeikin As Double) As Double
    Dim dWorkIn(9) As Double
    Dim dJudgeDT As Double
    Dim iCntI As Integer
    Dim iJituCnt As Integer
    Dim iSortCnt As Integer
    Dim iChuouCnt As Integer
    iSortCnt = 0
    iJituCnt = 0
    iChuouCnt = 0
    
    ' 測定データ個数カウント(最大9回)
    For iCntI = 0 To SOKUTEI_MAX
        '' フラグチェック                         <--- 2000/04/24 変更
        If tTeikouDat.dTeikouDT(iCntI).sRRGFlg = "1" Then
            '' データ有無判定
            If tTeikouDat.dTeikouDT(iCntI).dTeikou <> NULL_CHECK Then
                '' 入力データがある場合、ワークエリアに代入
                dWorkIn(iJituCnt) = tTeikouDat.dTeikouDT(iCntI).dTeikou
                iJituCnt = iJituCnt + 1
            End If
        End If
        '' 計算フラグ追加により変更               <--- 2000/04/24 削除
        '' If tTeikouDat.dTeikou(iCntI) <> NULL_CHECK Then
        ''     '' 入力データがある場合、ワークエリアに代入
        ''     dWorkIn(iJituCnt) = tTeikouDat.dTeikou(iCntI)
        ''     iJituCnt = iJituCnt + 1
        '' End If
    Next iCntI

    '' データを並び替える
    Do
        iSortCnt = iSortCnt + 1
        ' 実個数分実施したか判定
        If iSortCnt >= iJituCnt Then
            Exit Do
        End If
        dJudgeDT = dWorkIn(iSortCnt - 1)
        '' データを比較する
        If dJudgeDT > dWorkIn(iSortCnt) Then
            '' データが小さかった場合、データを入れ替える
            dJudgeDT = dWorkIn(iSortCnt)
            dWorkIn(iSortCnt) = dWorkIn(iSortCnt - 1)
            dWorkIn(iSortCnt - 1) = dJudgeDT
            iSortCnt = 0
        End If
    Loop

    '' 実個数が偶数か判定
    If iJituCnt Mod 2 = 0 Then
        '' 偶数の場合、平均に近い方を設定する
        For iCntI = 0 To iJituCnt
        If dHeikin < dWorkIn(iCntI) Then
            iChuouCnt = iCntI
            Exit For
        End If
        Next iCntI
        '' 差分比較処理
        If (dWorkIn(iCntI) - dHeikin) >= (dHeikin - dWorkIn(iCntI - 1)) Then
            GetChuou = dWorkIn(iCntI - 1)
        Else
            GetChuou = dWorkIn(iCntI)
        End If
    Else
        '' 奇数の場合、真ん中の値を設定する
        iChuouCnt = iJituCnt \ 2
        GetChuou = dWorkIn(iChuouCnt)
    End If
End Function

' @(f)
'
' 機能      : 平均値取得
'
' 返り値    : 平均値
'
' 引数      : TeikouDat - 抵抗値一覧
'
' 機能説明  :　抵抗値の平均値を取得する処理をする
'
' 備考      :
'
Private Function GetHeikin(tTeikouDat As TYPE_RRG) As Double
    Dim dSumData As Double
    Dim iCnt As Integer
    Dim iKazu As Integer
    
    dSumData = 0
    iKazu = 0
    iCnt = 0
    
    '' 測定抵抗個所分繰り返す(A～I)
    For iCnt = 0 To SOKUTEI_MAX
        '' フラグチェック                        <--- 2000/04/24 変更
        If tTeikouDat.dTeikouDT(iCnt).sRRGFlg = "1" Then
            ' 抵抗値が[0]か判定
            If tTeikouDat.dTeikouDT(iCnt).dTeikou <> NULL_CHECK Then
                '' 抵抗値が有る場合、合計値を更新する
                dSumData = dSumData + tTeikouDat.dTeikouDT(iCnt).dTeikou
                iKazu = iKazu + 1
            End If
        End If
        '' 計算フラグ追加により変更              <--- 2000/04/24 削除
        '' If tTeikouDat.dTeikou(iCnt) <> NULL_CHECK Then
        ''     '' 抵抗値が有る場合、合計値を更新する
        ''     dSumData = dSumData + tTeikouDat.dTeikou(iCnt)
        ''     iKazu = iKazu + 1
        '' End If
    
    Next iCnt
    '' 平均値を算出する
    GetHeikin = dSumData / iKazu
End Function

' @(f)
'
' 機能      :   原料仕掛工程テーブル検索処理
'
' 返り値    :   成功・失敗
'
' 引数      :   sMateNum    -   原料番号
'               sKoutei     -   工程コード
'               sWeight     -   仕掛重量
'
' 機能説明  :　 指定された工程の仕掛重量を検索し呼び出し元に返す
'
' 備考      :
'
Public Function SelectSikakariWeightDat(sMateNum$, sKoutei$, sWeight$) As Boolean
    Dim sSql    As String
    Dim objDS   As Object
    SelectSikakariWeightDat = False
    sSql = "SELECT  NVL(TO_CHAR(siwb2,'FM999999999'),' ')   "
    sSql = sSql & "FROM     xodb2                           "
    sSql = sSql & "WHERE    polnob2 = '" & sMateNum & "'    "
    sSql = sSql & "  AND    wkktb2  = '" & sKoutei & "'     "
    ''  SQLダイナセット処理
    ''  エラー時はFALSEを返す
    If DynSet(objDS, sSql) = False Then
        Call MsgOut(100, sSql, ERR_DISP_LOG, "xodb2")
        GoTo Er
    End If
    If objDS.EOF = False Then
        sWeight = objDS(0)
    End If
    Set objDS = Nothing
    SelectSikakariWeightDat = True
    Exit Function
Er:
    On Error Resume Next
    Set objDS = Nothing
End Function

' @(f)
'
' 機能      :   起動時ボタン操作処理
'
' 返り値    :   なし
'
' 引数      :   sForm   -   処理対象フォーム
'
' 機能説明  :　 処理対象のフォームに対してボタンを有効・無効にする
'
' 備考      :
'
Public Function InitCtrlAction(frmForm As Form, Optional bFlg As Boolean = False, Optional bAllFlg As Boolean = False) As Boolean
    Dim iIdx As Integer
    Dim ctlControl As Control
    If bAllFlg = True Then
        ''　フォーム上のすべてのボタンを列挙する
        For Each ctlControl In frmForm.Controls
            If TypeOf ctlControl Is CommandButton Then
                ctlControl.Enabled = bFlg
            End If
        Next ctlControl
    Else
        frmForm.cmdF(1).Enabled = bFlg
        frmForm.cmdF(2).Enabled = bFlg
        frmForm.cmdF(5).Enabled = bFlg
        frmForm.cmdF(6).Enabled = bFlg
        frmForm.cmdF(12).Enabled = bFlg
    End If
End Function

' @(f)
'
' 機能      :   スプレッド出力処理
'
' 返り値    :   なし
'
' 引数      :   sForm   -   処理対象フォーム
'
' 機能説明  :　 スプレッドシートをＣＳＶファイルに出力する。
'
' 備考      :
'
Function SPRD_PRT(msOBJ As Variant, msPGCAP As String, msNO As String) As Integer

Dim f_path As String    'CSVﾌｧｲﾙのﾊﾟｽ
Dim f_name(100)   As String    'CSVﾌｧｲﾙの名前
Dim f_time(100)   As Variant   'CSVﾌｧｲﾙの時間
Dim old_f     As String    '最古のCSVﾌｧｲﾙの名前
Dim old_t     As Variant   '最古のCSVﾌｧｲﾙの時間
Dim f_cnt     As Integer   'ﾌｧｲﾙのｶｳﾝﾀ
Dim f_max     As Integer
Dim data_wk   As String
Dim i, j, rows_cnt, col_cnt As Long

On Error GoTo Err
       
    Screen.MousePointer = vbHourglass
    
    'シートのプロパティー（MAXrows,MAXcol)取得
    rows_cnt = msOBJ.MaxRows
    col_cnt = msOBJ.MaxCols
    Debug.Print rows_cnt, col_cnt
        
    '出力ＣＳＶファイルのＯＰＥＮ
    If Dir(App.Path & "\Copy", vbDirectory) = "" Then
        MkDir (App.Path & "\Copy")
    End If
    
    '2000.08.16 古山
    f_max = 1
    f_name(f_max) = Dir(App.Path & "\Copy\*.csv")
            
    'ＣＳＶファイル数検索
    Do While Not f_name(f_max) = ""
        f_max = f_max + 1
        f_name(f_max) = Dir
    Loop
    
    f_max = f_max - 1
    
    'ファイル数が５０の場合、一番古いファイルを削除する。
    'ファイル数の上限(１００）を変更したい場合は以下のＩＦ文のf_maxの値を変更してください。
    If f_max = 50 Then
                    
        old_f = ""
        old_t = "99999999999999"
        
        For f_cnt = 1 To f_max
            If Mid(f_name(f_cnt), 11, 14) < old_t Then
                old_f = f_name(f_cnt)
                old_t = Mid(f_name(f_cnt), 11, 14)
            End If
        Next f_cnt
        
        'ファイルを削除する。
        Kill App.Path & "\Copy\" & old_f

    End If
    '2000.08.16 古山 END
    
    f_path = App.Path & "\Copy\" & msPGCAP & "_" & msNO & "_" & Format(Now(), "YYYYMMDDhhnnss") & ".csv"
    
    Open f_path For Output As #1
    
    'シートのmaxrows分ループ
    For i = 0 To rows_cnt       '０はヘッダ
        data_wk = ""
        msOBJ.row = i
        For j = 0 To col_cnt
            msOBJ.col = j
            data_wk = data_wk & Chr(34) & msOBJ.text & Chr(34) & ","
            DoEvents
        Next j
        
        '出力ＣＳＶファイルに1行分書込み
        Print #1, data_wk
    Next i
    
    Screen.MousePointer = vbDefault
    
    '出力ＣＳＶファイルのＣＬＯＳＥ
    Close #1
    
On Error GoTo 0

    SPRD_PRT = 0
        
Exit Function
    
Err:
        
    Screen.MousePointer = vbDefault
    MsgLog ("SPREAD SHEETの出力に失敗しました。::" & f_path & Chr(13) & Chr(10))

End Function

' @(f)
'
' 機能      :   処理工程検索処理
'
' 返り値    :   TRUE：成功、FALSE：失敗
'
' 引数      :   sGamen  -   画面コード
'               sKCode1 -   作業工程１
'               sKCode2 -   作業工程２
'
' 機能説明  :   画面コードから処理工程コード（作業工程コード）を検索し作業工程１に格納する
'
' 備考      :   １画面で2回処理を行う場合（実績を２つ）または、結晶の画面だが原料の仕掛に対して
'               操作を行うような場合の工程コードは作業工程２に入っている。
'               しかし、結晶検索加工の場合は作業工程１にAGR、作業工程２にMGRの工程コードを登録しています
'
'Public Function GetMyProcessCode(sGamen As String) As Boolean
'    Dim sSql    As String
'    Dim objDS   As Object
'    Dim sCode   As String
'    GetMyProcessCode = False
'    sCode = Left$(sGamen, 6) & "0"
'    sSql = "SELECT  NVL(kcode01a9,' '), "
'    sSql = sSql & " NVL(kcode02a9,' '), "
'    sSql = sSql & " NVL(kcode03a9,' '), "
'    sSql = sSql & " NVL(kcode04a9,' '), "
'    sSql = sSql & " NVL(kcode05a9,' ')  "
'    sSql = sSql & " FROM koda9          "
'    sSql = sSql & " WHERE   codea9  =   '" & sCode & "' "
'    sSql = sSql & "   AND   shuca9  =   '95'            "
'    sSql = sSql & "   AND   sysca9  =   'K'             "
'    If DynSet(objDS, sSql) = False Then
'        Call MsgOut(100, sSql, ERR_DISP_LOG, "koda9")
'        Exit Function
'    End If
'    If objDS.EOF = False Then
'        Do Until objDS.EOF
'            gsProcCode1 = objDS(0)
'            gsProcCode2 = objDS(1)
'            gsProcCode3 = objDS(2)
'            gsProcCode4 = objDS(3)
'            gsProcCode5 = objDS(4)
'            objDS.MoveNext
'        Loop
'    Else
'        gsProcCode1 = ""
'        gsProcCode2 = ""
'        gsProcCode3 = ""
'        gsProcCode4 = ""
'        gsProcCode5 = ""
'    End If
'    GetMyProcessCode = True
'End Function


' @(f)
'
' 機能      :   有効桁数処理
'
' 返り値    :   なし
'
' 引数      :　 有効桁数、データ
'
' 備考      :
'
Function keta(ketasu As Integer, motodata As Variant) As Variant
Dim yukocnt As Integer
Dim ln As Integer
Dim lp As Integer
Dim ld As Integer
Dim lz As Integer
Dim work
Dim moji
Dim oflg As Integer

    '初期値セット
    ld = 0
    lz = 0
    oflg = 0
    yukocnt = 0
    work = motodata
    If InStr(work, ".") = 0 Then
        work = work & ".0"
        motodata = work
    End If
    ln = Len(work)
    

    If Format(work, "###0.0####") = "0.0" Then
        keta = " "
        Exit Function
    End If

    '元データ桁数有効桁数分ループしながら
    For lp = 1 To ln
        moji = Mid(work, lp, 1)
        Debug.Print moji
        If moji <> 0 And moji <> "." And moji <> "-" And moji <> "+" Then
            yukocnt = yukocnt + 1
        ElseIf yukocnt > 0 And moji <> "." And moji <> "-" And moji <> "+" Then
            yukocnt = yukocnt + 1
        End If
        
        If yukocnt >= ketasu Then
            oflg = 1
            Exit For
        End If
    Next lp
    
'' 小数点位置判定
    For ld = 1 To ln
        moji = Mid(work, lp, 1)
        If moji = "." Then Exit For
    Next ld

    keta = Mid(motodata, 1, lp)

'' 有効桁数不足分”０”埋め
    If oflg = 0 Then
        For lz = 1 To ketasu - yukocnt
            keta = keta & "0"
            lp = lp + 1
        Next lz
    End If
    
'' 整数部不足桁数分”０”埋め
    ld = ld - 1
    If ld > lp Then
        For lz = 1 To lp - ld
            keta = keta & "0"
        Next lz
    End If

End Function


'///////////////////////////////////////////////////
' @(f)
'
' 機能      :   ＯＦ位置／ノッチ位置／ＣＦ位置検査変換処理
'
' 返り値    :   正常：" 0 0 0"～"-1-1-1"
'               異常："ERR"
'
' 引数      :   ＯＦ位置／ノッチ位置／ＣＦ位置
'               入力欄項目名　"ＯＦ位置"など
'               未入力許可フラグ　True:未入力可  False:未入力不可
'
' 備考      :   ＯＦ位置／ノッチ位置／ＣＦ位置入力欄の検査し６桁化する
'
'///////////////////////////////////////////////////
Public Function PosChkCnv(ctlControl As Control, _
                          sMsg As String, _
                          Optional bUnInput As Boolean = False) As String
    Dim sPos As String      ''位置
    Dim sCnvPos As String   ''位置（変換）
    Dim sChr As String      ''1文字切り出し
    Dim iIdx As Integer     ''インデックス
    Dim iNumCnt As Integer  ''数値カウンタ
    Dim bHifen As Boolean   ''ハイフンフラグ
    
'    PosChkCnv = "ERR"
    
    ''コントロール判定
    If (TypeOf ctlControl Is TextBox) Or _
       (TypeOf ctlControl Is ComboBox) Then
        ''値取得
        sPos = ctlControl.text
    Else
        ''値取得
        sPos = ctlControl.Caption
    End If
    
    PosChkCnv = sPos
    ''桁数チェック
    If bUnInput And (sPos = "") Then    ''未入力許可で未入力なら
        PosChkCnv = sPos
        Exit Function
    ElseIf Len(sPos) = 6 Then           ''最初から６桁の場合
        ''ノーチェック
        Call CtrlEnabled(ctlControl, NORMAL_CTL)
        PosChkCnv = sPos
        Exit Function
    ElseIf Len(sPos) > 6 Then
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "が６桁以上です", ERR_DISP)
        Exit Function
    ElseIf Len(sPos) < 3 Then
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "が３桁未満です", ERR_DISP)
        Exit Function
    End If
    
    ''文字種チェック
    For iIdx = 1 To Len(sPos)
        Select Case Mid(sPos, iIdx, 1)
        Case " ", "-", "0" To "9"
        Case Else
            Call CtrlEnabled(ctlControl, RED_CTL)
            Call MsgOut(0, sMsg & "に不正な文字が含まれています", ERR_DISP)
            Exit Function
        End Select
    Next
    
    ''変換
    For iIdx = 1 To Len(sPos)
        sChr = Mid(sPos, iIdx, 1)               ''切り出し
        If bHifen Then                          ''前がハイフンの場合
            If IsNumeric(sChr) Then             ''数値なら
                sCnvPos = sCnvPos & sChr        ''その数値を格納
                bHifen = False                  ''今回数値
                iNumCnt = iNumCnt + 1           ''数値をカウント
            Else                                ''ハイフン２連続
                Call CtrlEnabled(ctlControl, RED_CTL)
                Call MsgOut(0, sMsg & "の文字の並びが不正です", ERR_DISP)
                Exit Function
            End If
        Else                                    ''前が数値の場合
            If IsNumeric(sChr) Then             ''数値なら
                sCnvPos = sCnvPos & " " & sChr  ''空白＆数値格納
                bHifen = False                  ''今回数値をセット
                iNumCnt = iNumCnt + 1           ''数値をカウント
            Else                                ''数値以外なら
                sCnvPos = sCnvPos & sChr        ''その文字を格納
                bHifen = True                   ''今回ハイフンをセット
            End If
        End If
    Next
    If bHifen Then  ''最後が数値以外の場合
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "の文字の並びが不正です", ERR_DISP)
        Exit Function
    ElseIf iNumCnt > 3 Then ''数値が３桁を超えたら
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "の数値が３桁以上含まれています", ERR_DISP)
        Exit Function
    ElseIf iNumCnt < 3 Then ''数値が３桁未満
        Call CtrlEnabled(ctlControl, RED_CTL)
        Call MsgOut(0, sMsg & "の数値が３桁未満です", ERR_DISP)
        Exit Function
    End If
    
    ''完了
    Call CtrlEnabled(ctlControl, NORMAL_CTL)
    PosChkCnv = sCnvPos
End Function


'///////////////////////////////////////////////////
' @(f)
'
' 機能      :   テキストボックス文字埋め処理
'
' 返り値    :   処理結果の文字列
'
' 引数      :   テキストボックスコントロール
'               文字埋めする桁数
'               文字埋めする文字
'
' 備考      :   文字埋めする桁数省略時はMaxLengthまで埋める
'               MaxLengthが設定されてない場合１２桁まで埋める
'               文字埋めする文字省略時は"0"で埋める
'
'///////////////////////////////////////////////////
Public Function TextBoxDap(ctlTextBox As Control, Optional ByVal iColumn As Integer, Optional ByVal sChar As String) As String
    Dim iCol As Integer     ''設定する桁数
    Dim iLackCnt As Integer ''不足桁数
    ''桁数が指定されていたら
    If iColumn Then
        iCol = iColumn
    ''MaxLengthが設定されていたら
    ElseIf ctlTextBox.MaxLength Then
        iCol = ctlTextBox.MaxLength
    ''桁数省略し、MaxLengthも設定されてない場合
    Else
        iCol = 12
    End If
    ''不足桁数算出
    iLackCnt = iCol - Len(ctlTextBox)
    ''不足していたら
    If iLackCnt > 0 Then
        ''文字埋めする文字が省略された場合"0"にする
        If sChar = "" Then sChar = "0"
        ''文字埋めする
        ctlTextBox = ctlTextBox & String(iLackCnt, sChar)
    End If
    TextBoxDap = ctlTextBox
End Function


' @(f)
'
' 機能      : ORGデータ設定
'
' 返り値    : ORG値
'
' 引数      : OiCsData  - 結晶Oi/Csデータ
'             BuiNumber - 結晶部位番号
'             sKeisanFlg -ＯＳＦ計算フラグ
'
' 機能説明  : ORGデータを計算し設定する
'
' 備考      : ORG計算
'            １＝｜（周辺－中心）｜÷中心×１００  ※周辺毎に計算し、計算結果の最大値を使う
'            ２＝｜（中心Ａ）－（周辺の平均）｜÷（中心Ａ）×１００
'            ３＝（ｍａｘ－ｍｉｎ）／ｍａｘ×１００
'            ４＝｜（中心Ａ）－（Ｂ、Ｃ、Ｆ、Ｇの平均）｜÷（中心Ａ）×１００
'            ５＝（｜中心値－周辺値｜）÷（中心値＋周辺値）×２００  ※周辺毎に計算し、計算結果の最大値を使う
'            ７＝（ｍａｘ－ｍｉｎ）／ｍｉｎ×１００
'            ８＝（中心Ａ）－（周辺のｍｉｎ）÷（中心Ａ）×１００
''20000908 小川 この関数全体を作り直した
Public Function SetORGData(sOiCsDataIti() As ST_OICS, sBuiNumber As String, Optional sKeisanFlg As String = "1") As String
    Dim iCntIdx  As Integer ''センター値のｲﾝﾃﾞｯｸｽ
    Dim iIdx     As Integer ''ｲﾝﾃﾞｯｸｽ
    Dim iPnt     As Integer ''測定点"A"～"I"は０～８
    Dim iCnt     As Integer ''入力された測定点数カウント
    Dim dOrg     As Double  ''最終ＯＲＧ計算結果
    Dim dOrgs(8) As Double  ''測定点毎のＯＲＧ計算結果
    Dim sOi(8)   As String  ''同部位の各測定点毎の酸素：ﾃﾞｰﾀ無しは長さ０の文字列
    Dim dAveOi   As Double  ''平均値
    Dim dMaxOi   As Double  ''最大値
    Dim dMinOi   As Double  ''最小値
    Dim bValGet  As Boolean ''値取得フラグ
    
    ''同一部位の各測定点の酸素値を取得する処理
    For iIdx = 0 To UBound(sOiCsDataIti)
        ''部位が同じだったら
        If val(sBuiNumber) = val(sOiCsDataIti(iIdx).sCryBuiNo) Then
            ''測定点"A"～"I"→０～８に変換
            Select Case sOiCsDataIti(iIdx).sMenPosIti
            Case "A": iPnt = 0: iCntIdx = iIdx ''センターのｲﾝﾃﾞｯｸｽ取得
            Case "B": iPnt = 1
            Case "C": iPnt = 2
            Case "D": iPnt = 3
            Case "E": iPnt = 4
            Case "F": iPnt = 5
            Case "G": iPnt = 6
            Case "H": iPnt = 7
            Case "I": iPnt = 8
            End Select
            ''測定点が"A"～"I"なら
            Select Case sOiCsDataIti(iIdx).sMenPosIti
            Case "A", "B", "C", "D", "E", "F", "G", "H", "I"
                ''測定点の酸素値取得
                sOi(iPnt) = sOiCsDataIti(iIdx).sSansoAT
            End Select
        End If
    Next
    
    ''OSF計算ﾌﾗｸﾞにより使用する計算式を切り替える
    Select Case sKeisanFlg
    Case "3", "7", "8"
        ''最大値・最小値を取得する処理
        dMaxOi = -999999  ''最大値に最小の値を設定
        dMinOi = 999999   ''最小値に最大の値を設定
        bValGet = False   ''値未取得を設定
            
        For iPnt = 0 To 8
            ''この測定点が入力されていたら
            If sOi(iPnt) <> "" Then
                ''この測定値が最大値より大きければ
                If val(sOi(iPnt)) > dMaxOi Then
                    dMaxOi = val(sOi(iPnt)) ''最大値取得
                    bValGet = True          ''値取得を設定
                End If
                ''この測定値が最小値より小さければ
                If val(sOi(iPnt)) < dMinOi Then
                    dMinOi = val(sOi(iPnt)) ''最小値取得
                    bValGet = True          ''値取得を設定
                End If
            End If
        Next
        If sKeisanFlg = "8" Then ''ＯＳＦ計算フラグが"8"なら
            bValGet = False   ''値未取得を設定
            For iPnt = 1 To 8
                ''この測定点が入力されていたら
                If sOi(iPnt) <> "" Then
                    ''この測定値が最小値より小さければ
                    If val(sOi(iPnt)) < dMinOi Then
                        dMinOi = val(sOi(iPnt)) ''最小値取得
                        bValGet = True          ''値取得を設定
                    End If
                End If
            Next
        End If
        ''最大値と最小値を取得できたら
        If bValGet Then
            ''ＯＳＦ計算フラグにより計算式を代える
            If sKeisanFlg = "8" Then ''ＯＳＦ計算フラグが"8"なら
                ''中心値は必ず必要
                If sOi(0) = "" Then    ''なければ
                    Exit Function      ''抜ける
                End If
                ''（中心－周辺最小値）／中心×１００
                dOrg = (val(sOi(0)) - dMinOi) / val(sOi(0)) * 100
            ElseIf sKeisanFlg = "3" Then ''ＯＳＦ計算フラグが"3"なら
                ''（測定最大値－測定最小値）／測定最大値×１００
                dOrg = (dMaxOi - dMinOi) / dMaxOi * 100
            Else                     ''ＯＳＦ計算フラグが"7"なら
                ''（測定最大値－測定最小値）／測定最小値×１００
                dOrg = (dMaxOi - dMinOi) / dMinOi * 100
            End If
            ''これを戻り値とする
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    Case "2", "4"
        ''中心値は必ず必要
        If sOi(0) = "" Then    ''なければ
            Exit Function      ''抜ける
        End If
        ''平均値を取得する処理
        dAveOi = 0        ''平均値クリア
        iCnt = 0          ''入力件数クリア
        For iPnt = 1 To 8 ''周辺値：測定点"B"～"I"
            If sKeisanFlg = "2" Then ''ＯＳＦ計算フラグが"2"なら
                ''この測定点が入力されており、測定点が"B"～"I"なら
                If (sOi(iPnt) <> "") Then
                    dAveOi = dAveOi + val(sOi(iPnt))  ''対象値を加算
                    iCnt = iCnt + 1                   ''対象の入力件数カウント
                End If
            End If
            If sKeisanFlg = "4" Then ''ＯＳＦ計算フラグが"4"なら
                ''この測定点が入力されており、測定点が"B"・"C"・"F"・"G"なら
                If (sOi(iPnt) <> "") And ( _
                    (iPnt = 1) Or _
                    (iPnt = 2) Or _
                    (iPnt = 5) Or _
                    (iPnt = 6)) Then
                    dAveOi = dAveOi + val(sOi(iPnt))  ''対象値を加算
                    iCnt = iCnt + 1                   ''対象の入力件数カウント
                End If
            End If
        Next
        ''周辺の値を取得できたら
        If iCnt > 0 Then
            ''平均値算出
            If dAveOi <> 0 Then
                dAveOi = dAveOi / iCnt
            End If
            ''絶対値（中心値）－（周辺の平均値）÷中心値×１００
            dOrg = Abs(val(sOi(0)) - dAveOi) / val(sOi(0)) * 100
            ''これを戻り値とする
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    Case Else
        ''中心値は必ず必要
        If sOi(0) = "" Then    ''なければ
            Exit Function      ''抜ける
        End If
        ''周辺の測定点毎のＯＲＧを計算する処理
        dOrg = -999999    ''ＯＲＧに最小値を設定
        bValGet = False   ''値未取得を設定
        For iPnt = 1 To 8 ''周辺値：測定点"B"～"I"
            ''この測定点が入力されていたら
            If sOi(iPnt) <> "" Then
                ''ＯＳＦ計算フラグにより計算式を代える
                If sKeisanFlg = "5" Then ''ＯＳＦ計算フラグが"5"なら
                    ''絶対値（中心値－周辺値）÷（中心値＋周辺値）×２００
                    dOrgs(iPnt) = Abs(val(sOi(0)) - val(sOi(iPnt))) / (val(sOi(0)) + val(sOi(iPnt))) * 200
                Else                     ''ＯＳＦ計算フラグが"1"か""なら
                    ''絶対値（周辺値－中心値）÷中心値×１００     8/16 Yam　100をかけてRoundをいれるように修正
                    dOrgs(iPnt) = Round((Abs(val(sOi(iPnt)) - val(sOi(0))) / val(sOi(0))) * 100, 1)
                End If
                ''一番大きいＯＲＧを取得
                If dOrg < dOrgs(iPnt) Then
                    dOrg = dOrgs(iPnt)
                    bValGet = True       ''値取得を設定
                End If
            End If
        Next
        ''最大値を取得できたら
        If bValGet Then
            ''これを戻り値とする
            SetORGData = left(CStr(dOrg), 6)
        End If
        
    End Select
    
    ''センター値が有れば
    If sOi(0) <> "" Then
        ''ＯＲＧを設定する
        sOiCsDataIti(iCntIdx).sORGNo = SetORGData
    End If
End Function

' @(f)
'
' 機能      :   マイナス補正関数
'
' 返り値    :   マイナス補正後の数値文字列
'
' 引き数    :   ARG1        - センター値
'               ARG2        - マイナス補正値
'
' 機能説明  :   マイナス補正を行い文字列にて値を返す
'
' 備考      :
'
Public Function ArgmentFormat(sDat As String) As String
    Dim iDat    As Integer
    Dim iDo     As Integer
    Dim iFun    As Integer
    If sDat <> "" Then
        iDat = CInt(val(sDat))
        iDo = iDat \ 60
        iFun = (iDat Mod 60)
        If iDo > -1 And iFun < 0 Then     'マイナスの場合の修正　1/9 Yam
            iFun = Abs(iDat Mod 60)
            ArgmentFormat = Format$(iDo, "-#0°") & Format$(iFun, "00′")
        Else
            iFun = Abs(iDat Mod 60)
            ArgmentFormat = Format$(iDo, "#0°") & Format$(iFun, "00′")
        End If
    Else
        ArgmentFormat = ""
    End If
End Function


' @(f)
'
' 機能      :   ＯＦ位置変換処理（新表示→旧表示）
'
' 返り値    :   変換後のＯＦ位置（旧表示）
'
' 引き数    :   ARG1        - 結晶軸
'               ARG2        - ＯＦ位置（新表示）
'
' 機能説明  :   ＯＦ位置の変換を行い文字列にて値を返す
'
' 備考      :
'
Public Function OfposChg(Xjiku As String, Ofposo As String) As String
    
    OfposChg = ""
    If Ofposo = "1" Then OfposChg = "110"
    If Ofposo = "2" Then OfposChg = "110"
    If Ofposo = "3" Then OfposChg = "110"
    If Ofposo = "4" Then OfposChg = "110"
    If Ofposo = "5" Then OfposChg = "100"
    If Ofposo = "6" Then OfposChg = "100"
    If Ofposo = "7" Then OfposChg = "100"
    If Ofposo = "8" Then OfposChg = "100"
    If Ofposo = "9" Then OfposChg = "110"
    If Ofposo = "10" Then OfposChg = "110"
    If Ofposo = "11" Then OfposChg = "2111"
    If Ofposo = "12" Then OfposChg = "2112"
    If Ofposo = "13" Then OfposChg = "111"
    If Ofposo = "14" Then OfposChg = "111"
    If Ofposo = "15" Then OfposChg = "111"
    If Ofposo = "16" Then OfposChg = "111"
    If Ofposo = "17" Then OfposChg = "211"
    If Ofposo = "18" Then OfposChg = "211"
    If Ofposo = "19" Then OfposChg = "211"
    If Ofposo = "20" Then OfposChg = "211"
    If Ofposo = "21" Then OfposChg = "111"
    If Ofposo = "22" Then OfposChg = "111"
    If Ofposo = "23" Then OfposChg = "110"
    If Ofposo = "24" Then OfposChg = "111"
    If Ofposo = "1" And Xjiku = "511" Then OfposChg = "1101"
    If Ofposo = "9" And Xjiku = "111" Then OfposChg = "1101"
    If Ofposo = "10" And Xjiku = "111" Then OfposChg = "1102"

End Function



' @(f)
'
' 機能      :   ＯＦ位置図番号判定処理(ＯＦパターン）
'
' 返り値    :   判定後のＯＦ位置図番号
'
' 引き数    :   Xjiku        - 結晶軸
'               Ofpos        - ＯＦ位置（新表示）
'               Cfkaku       - ＯＦ角度（新表示）
'               Cfpos        - ＯＦ位置（新表示）
'               Nopos        - ノッチ位置（新表示）
'               Ocnflg       - OF,CF,ﾉｯﾁ要･不要
'               Cfkijyn      - ＣＦ指定基準（新表示）
'
' 機能説明  :   ＯＦ位置図番号の判定を行い文字列にて値を返す
'
' 備考      :
'
Public Function OfposBangoChg(Xjiku As String, Ofpos As String, Cfkaku As String, Cfpos As String, _
                              Nopos As String, Ocnflg As String, Cfkijyn As String) As String
    
    OfposBangoChg = "99"
    If Ocnflg = "" Or Ocnflg = " " Or Ocnflg = "0" Then OfposBangoChg = "00"
    If Ocnflg = "1" Then
        If Xjiku = "111" And Ofpos = "9" Then OfposBangoChg = "01" Else
        If Xjiku = "111" And Ofpos = "10" Then OfposBangoChg = "02" Else
        If Xjiku = "111" And Ofpos = "12" Then OfposBangoChg = "03" Else
        If Xjiku = "111" And Ofpos = "11" Then OfposBangoChg = "04" Else
        If Xjiku = "511" And Ofpos = "1" Then OfposBangoChg = "05" Else
        If Xjiku = "100" And Ofpos = "1" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "2" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "3" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "4" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Ofpos = "5" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "6" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "7" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Ofpos = "8" Then OfposBangoChg = "07" Else
        If Xjiku = "110" And Ofpos = "9" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Ofpos = "10" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Ofpos = "5" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Ofpos = "6" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Ofpos = "13" Then OfposBangoChg = "10" Else
        If Xjiku = "110" And Ofpos = "15" Then OfposBangoChg = "10" Else
        If Xjiku = "   " Or Xjiku = "" Then OfposBangoChg = "98"
    End If

    If Ocnflg = "3" Then
        If Xjiku = "111" And Nopos = "9" Then OfposBangoChg = "01" Else
        If Xjiku = "111" And Nopos = "10" Then OfposBangoChg = "02" Else
        If Xjiku = "111" And Nopos = "12" Then OfposBangoChg = "03" Else
        If Xjiku = "111" And Nopos = "11" Then OfposBangoChg = "04" Else
        If Xjiku = "511" And Nopos = "1" Then OfposBangoChg = "05" Else
        If Xjiku = "100" And Nopos = "1" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "2" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "3" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "4" Then OfposBangoChg = "06" Else
        If Xjiku = "100" And Nopos = "5" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "6" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "7" Then OfposBangoChg = "07" Else
        If Xjiku = "100" And Nopos = "8" Then OfposBangoChg = "07" Else
        If Xjiku = "110" And Nopos = "9" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Nopos = "10" Then OfposBangoChg = "08" Else
        If Xjiku = "110" And Nopos = "5" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Nopos = "6" Then OfposBangoChg = "09" Else
        If Xjiku = "110" And Nopos = "13" Then OfposBangoChg = "10" Else
        If Xjiku = "110" And Nopos = "15" Then OfposBangoChg = "10" Else
        If Xjiku = "   " Or Xjiku = "" Then OfposBangoChg = "98"
    End If

    If Ocnflg = "2" And Cfkijyn = "1" Then
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 10800 Then OfposBangoChg = "11" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 5400 Then OfposBangoChg = "12" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 2700 Then OfposBangoChg = "13" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 18900 Then OfposBangoChg = "14" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 13500 Then OfposBangoChg = "15" Else
        If Xjiku = "111" And Ofpos = "9" And Cfkaku = 8100 Then OfposBangoChg = "16" Else
        If Xjiku = "111" And Ofpos = "12" And Cfkaku = 10800 Then OfposBangoChg = "17" Else
        If Xjiku = "111" And Ofpos = "12" And Cfkaku = 16200 Then OfposBangoChg = "18" Else
        If Xjiku = "111" And Ofpos = "11" And Cfkaku = 13500 Then OfposBangoChg = "19" Else
        If Xjiku = "111" And Ofpos = "11" And Cfkaku = 8100 Then OfposBangoChg = "20" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 10800 Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 5400 Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 13500 Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 8100 Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "1" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "2" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "3" And Cfkaku = 2700 Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "4" And Cfkaku = 2700 Then OfposBangoChg = "25"
    End If

    If Ocnflg = "2" And Cfkijyn = "2" Then
        If Xjiku = "111" And Ofpos = "9" And Cfpos = "10" Then OfposBangoChg = "11" Else
        If Xjiku = "111" And Ofpos = "9" And Cfpos = "12" Then OfposBangoChg = "12" Else
        If Xjiku = "111" And Ofpos = "12" And Cfpos = "11" Then OfposBangoChg = "17" Else
        If Xjiku = "111" And Ofpos = "12" And Cfpos = "9" Then OfposBangoChg = "18" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "1" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "2" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "4" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "3" Then OfposBangoChg = "21" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "3" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "4" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "1" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "2" Then OfposBangoChg = "22" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "6" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "5" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "8" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "7" Then OfposBangoChg = "23" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "7" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "8" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "6" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "5" Then OfposBangoChg = "24" Else
        If Xjiku = "100" And Ofpos = "2" And Cfpos = "5" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "1" And Cfpos = "6" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "3" And Cfpos = "7" Then OfposBangoChg = "25" Else
        If Xjiku = "100" And Ofpos = "4" And Cfpos = "8" Then OfposBangoChg = "25"
    End If

End Function

'///////////////////////////////////////////////////
'
' @(f)
' 機能         : 印刷ボタン処理
' 返り値       : なし
' 引き数       : フォーム名
' 機能説明     : 印刷ボタン処理
'                <2001.02.23 kuro>
'
' ﾌﾟﾛﾊﾟﾃｨ説明 ：　Copies   :　印刷部数
'                 FromPage :　印刷開始ページ
'                 ToPage   :　印刷終了ページ
'                 hDC      :　選択されたプリンタのデヴァイスコンテキスト
'
'///////////////////////////////////////////////////
'
Public Function mdlHard_Copy(Getform As Form)
    Dim BeginPage, EndPage, NumCopies, i
    
    With Getform
        .CommonDialog1.CancelError = True
    
        On Error GoTo ErrHandler
    
        'ダイアログボックスの［ページ指定］オプションボタンを無効にする
        .CommonDialog1.Flags = &H8&
    
        'ダイアログボックスの［選択した部分］オプションボタンを無効にする
        .CommonDialog1.Flags = &H4&
    
        '［印刷］ダイアログボックスを表示
        .CommonDialog1.ShowPrinter
    
        'ユーザーの選択した値をダイアログボックスから取得
        'BeginPage = .CommonDialog1.FromPage   ''スタートページ
        'EndPage = .CommonDialog1.ToPage       ''エンドページ
        NumCopies = .CommonDialog1.Copies      ''印刷部数
        For i = 1 To NumCopies
            '印刷フォームをプリンタに送信
            .PrintForm
            Printer.EndDoc
        Next i
        Exit Function
    End With

ErrHandler:
    'ユーザーが［キャンセル］をクリックしました
    Exit Function
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : バージョンチェック関数
'
' 返り値  : True :ﾊﾞｰｼﾞｮﾝ一致
'           False:ﾊﾞｰｼﾞｮﾝ取得失敗／ﾊﾞｰｼﾞｮﾝが古い
' 引き数  :
'
' 機能説明: ＤＢに登録されているバージョンと
'           モジュールのバージョンが一致しているか比較し、
'           モジュールのバージョンが古い場合
'           フォームの「ﾒｲﾝﾒﾆｭｰ」ボタン以外のｺﾝﾄﾛｰﾙを
'           使用不可にする
'
'///////////////////////////////////////////////////
Public Function VerChk(frmCurrent As Form) As Boolean
    Dim objOraDyn As Object         ''ダイナセットオブジェクト
    Dim sSql As String              ''ＳＱＬ文
    Dim iMajorx As String           ''ＤＢ取得メジャーバージョン
    Dim iMajor As Integer           ''ＤＢ取得メジャーバージョン
    Dim iMinor As Integer           ''ＤＢ取得マイナーバージョン
    Dim iRevision As Integer        ''ＤＢ取得リビジョン
    
    gbFTPFlg = False    ''FTP起動フラグクリア
        
    ''DB登録ﾊﾞｰｼﾞｮﾝ取得
    ''ＳＱＬ文作成
    sSql = "SELECT   NVL(ctr01a9,0),    "
    sSql = sSql & "  NVL(ctr02a9,0),    "
    sSql = sSql & "  NVL(ctr03a9,0)     "
    sSql = sSql & "FROM  koda9         "
    sSql = sSql & "WHERE sysca9 = 'K'  "
    sSql = sSql & "AND   shuca9 = '01' "
    sSql = sSql & "AND   codea9 = '" & UCase(gsEXEName) & "'"
    
    If DynSet2(objOraDyn, sSql) = False Then
        ''取得失敗
        Call MsgOut(100, sSql, ERR_DISP_LOG, "koda9")
        VerChk = False
        GoTo Er
    End If
    If objOraDyn.EOF Then
        ''見つからなかった
        Call MsgOut(55, "ﾊﾞｰｼﾞｮﾝ情報", ERR_DISP)
        VerChk = False
        GoTo Er
    End If
    'GetMyProcessCode frmCurrent.Caption
    
    iMajor = objOraDyn(0)
    iMinor = objOraDyn(1)
    iRevision = objOraDyn(2)
    
    If iMajor = 0 And iMinor = 0 And iRevision = 0 Then
       VerChk = True
       Exit Function
    End If
    
    ''バージョン不一致をセット
    VerChk = False
    ''メジャーバージョンチェック
    If iMajor <> val(App.Major) Then
        Call MsgOut(0, "ﾒｼﾞｬｰﾊﾞｰｼﾞｮﾝ不一致ﾒｲﾝﾒﾆｭｰﾎﾞﾀﾝ押下", ERR_DISP)
        GoTo Er
    End If
    ''マイナーバージョンチェック
    If iMinor <> val(App.Minor) Then
        Call MsgOut(0, "ﾏｲﾅｰﾊﾞｰｼﾞｮﾝ不一致ﾒｲﾝﾒﾆｭｰﾎﾞﾀﾝ押下", ERR_DISP)
        GoTo Er
    End If
    ''リビジョンチェック
    If iRevision <> val(App.Revision) Then
        Call MsgOut(0, "ﾘﾋﾞｼﾞｮﾝﾊﾞｰｼﾞｮﾝ不一致ﾒｲﾝﾒﾆｭｰﾎﾞﾀﾝ押下", ERR_DISP)
        GoTo Er
    End If
    ''バージョン一致をセット
    VerChk = True
     ''  処理工程コード取得
    Exit Function
Er:
    ''バージョン不一致
    Call CtrlCancel(frmCurrent)     ''ﾒｲﾝﾒﾆｭｰ以外のｺﾝﾄﾛｰﾙを使えなくする
    gbFTPFlg = True    ''FTP起動フラグセット
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : Ｓｕｍｃｏ時間取得
'
' 返り値  : Ｓｕｍｃｏ時間（Date型：YYYY/MM/DD）
'
' 引き数  : 変換したい時刻（Date型：YYYY/MM/DD）
'
' 機能説明: パラメータの時刻からＳｕｍｃｏ時間に変換する
'
'  Ｓｕｍｃｏ時間 = パラメータの日付 －　調整時間
'
'           エラーがあった場合はパラメータの日付を
'           そのまま戻す。
'
'///////////////////////////////////////////////////
Public Function CalcSumcoTime(tParmDate As Date) As Date

    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
    Dim vChoseiTime     As Variant          '調整時間

    'エラーハンドラの設定
    On Error GoTo PROC_ERR

    'デフォルト戻り値設定
    CalcSumcoTime = tParmDate

    'SUMCO時間作成の為、調整時間取得
    sql = "SELECT KCODE01A9"
    sql = sql & " FROM koda9 "
    sql = sql & " WHERE SYSCA9 = 'X'"
    sql = sql & "   AND SHUCA9 = '80'"
    sql = sql & "   AND CODEA9 = '1'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    '存在しない時、処理終了
    If rs Is Nothing Then
        Exit Function
    End If
    If Not rs.EOF Then
        If IsNull(rs.Fields("KCODE01A9")) = True Then
            Exit Function
        Else
    '調整時間取得
            vChoseiTime = CDate(rs.Fields("KCODE01A9"))
        End If
    End If
    rs.Close

    'SUMCO時間=パラメータの日付-KODA9.調整時間
    CalcSumcoTime = Format(tParmDate - CDate(vChoseiTime), "yyyy/mm/dd")

    Exit Function

proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function

PROC_ERR:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.Number
    Resume proc_exit
End Function

'///////////////////////////////////////////////////
' @(f)
' 機能    : サーバー時間取得
'
' 返り値  : サーバー時間（Date型：YYYY/MM/DD）
'
' 機能説明: ORACLEより現在時間を取得する。
'
'
'///////////////////////////////////////////////////
Public Function getSvrTime() As Date
                                
    Dim sql             As String           'ＳＱＬ
    Dim rs              As OraDynaset       'レコードセット
                                
    'エラーハンドラの設定
    On Error GoTo PROC_ERR
                                
    'デフォルト戻り値設定(端末時刻）
    getSvrTime = Format(Now, "yyyy/mm/dd hh:mm:ss")
                                
    'サーバー時間取得
    sql = "SELECT SYSDATE"
    sql = sql & " FROM DUAL "
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
                                
    '存在しない時、処理終了
    If rs Is Nothing Then
        Exit Function
    End If
    If Not rs.EOF Then
        If IsNull(rs.Fields("SYSDATE")) = True Then
            Exit Function
        Else
    '時間取得
            getSvrTime = CDate(rs.Fields("SYSDATE"))
        End If
    End If
    rs.Close
                                
'    'SUMCO時間=パラメータの日付-KODA9.調整時間
'    CalcSumcoTime = Format(tParmDate - CDate(vChoseiTime), "yyyy/mm/dd")
'
'    Exit Function
                                
proc_exit:
    '' 終了
'    gErr.Pop
    Exit Function
                                
PROC_ERR:
    '' エラーハンドラ
    Debug.Print Err.Description & "：" & Err.Number
    Resume proc_exit
End Function
                                
                                
'*ADD* ﾓｼﾞｭｰﾙ統一 TCS)K.Kunori 2004.11.29 START >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    '2004/9/27tcs Yamauchi 追加 start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    ' 機能    : 担当者名取得(権限ﾁｪｯｸを含む場合）
    '           300mm限定
    '
    ' 返り値  : True:正常
    '           False:失敗
    '
    ' 引き数  : 担当者ｺｰﾄﾞ
    '
    ' 機能説明: 担当者ｺｰﾄﾞから担当者名を取得
    '
    '///////////////////////////////////////////////////
    Public Function GetAuthorityUser_300(ByVal STAFFID As String, ByVal ProcID As String, _
                                                                ByRef sUserName As String) As Boolean
        Dim sSqlStmt As String
        Dim objOraDyn As Object
        
        sUserName = vbNullString       ''担当者名ｸﾘｱ
        
        '' 引数で指定された社員IDと工程ｺｰﾄﾞで検索を行う
        sSqlStmt = ""
        sSqlStmt = sSqlStmt & "select   t1.JFMLNAME,t1.JFSTNAME         " & vbLf
        sSqlStmt = sSqlStmt & "from     TBCMB001 t1,TBCMB004 t4         " & vbLf
        sSqlStmt = sSqlStmt & "where    t1.EXECODE = t4.AUTHCODE        " & vbLf
        sSqlStmt = sSqlStmt & "and      t1.STAFFID = '" & STAFFID & "'  " & vbLf
        sSqlStmt = sSqlStmt & "and      t4.TRANID = '" & ProcID & "'    " & vbLf
        
        ''ダイナセット作成
        If DynSet2(objOraDyn, sSqlStmt) = False Then
            ''ダイナセット作成失敗
            Call MsgOut(100, sSqlStmt, ERR_DISP_LOG)
            
            GetAuthorityUser_300 = False
            Exit Function
        End If
        If objOraDyn.EOF Then
            ''該当する担当者ｺｰﾄﾞが無かった
            Call MsgOut(0, "認定担当者ではありません", ERR_DISP)
            
            GetAuthorityUser_300 = False
            Exit Function
        End If
    
        sUserName = NulltoStr(objOraDyn(0)) & Space(1) & NulltoStr(objOraDyn(1))    ''担当者名取得
        
        GetAuthorityUser_300 = True         ''処理成功を返す
        
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能      :   ｼｽﾃﾑ日付取得
    '
    ' 返り値    :　 成否
    '
    ' 引数      :　 なし
    '
    ' 機能説明  :   ｼｽﾃﾑ日付を取得する
    '
    ' 備考      :   取得した値をﾊﾟﾌﾞﾘｯｸ変数に格納
    '
    '///////////////////////////////////////////////////
    Public Function GetSysdate() As Boolean
    
        Dim sSql        As String
        Dim objOraDyn   As Object
        
    On Error GoTo ErrHand
    
        GetSysdate = False
        
        sSql = ""
        sSql = sSql & "select to_char(SYSDATE,'YYYY/MM/DD HH24:MI:SS')  " & vbLf
        sSql = sSql & "from dual                                        " & vbLf
        
        ''ﾀﾞｲﾅｾｯｯﾄ作成
        If DynSet2(objOraDyn, sSql) = False Then
            ''ﾀﾞｲﾅｾｯｯﾄ作成失敗
            Call MsgOut(100, sSql, ERR_DISP_LOG)
            GetSysdate = False
            Exit Function
        End If
        
        ''ﾊﾟﾌﾞﾘｯｸ変数に格納
        gsSysdate = objOraDyn.Fields(0).Value
        
        '開放
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
        
        GetSysdate = True
        Exit Function
        
ErrHand:
    
        '開放
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
    
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "")
    
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f)
    ' 機能    : コマンドライン引数取得・変数セット
    '
    ' 返り値  : True:正　False:否
    '
    ' 引き数  :
    '
    ' 機能説明: コマンドライン引数を取得し・トークン切出し・変数セット
    '
    '///////////////////////////////////////////////////
    Public Function GetCmdLine_Re() As Boolean
        Dim sCmdLine As String
        
        ''コマンドライン取得
        sCmdLine = Command
        '' 0        1         2
        '' 1234567890123456789012
        ''"99_*******_***********"
        ''工場コード_呼出区分_品番
        
        ''固定でコマンドライン引数を切出す
        gsFactryCd = left(sCmdLine, 2)    ''工場コード(2桁)
        gsCallCd = Mid(sCmdLine, 4, 7)    ''呼出区分(7桁)
        gsHinban = Mid(sCmdLine, 12, 11)  ''品番(11桁)
        myFactryCd = Mid(sCmdLine, 24, 2)
        
        If Len(gsFactryCd) <> 2 Then Exit Function
        If Len(gsCallCd) <> 7 Then Exit Function
        If gsHinban = "00000000000" Then gsHinban = ""
        
    '2004/11/12 TCS TAGAWA 追加 Start--------------------------------------------------
        Select Case gsFactryCd
            Case "10"               ''野田工場
                gsSystemCd = SYSTEM_200
            Case "30"               ''生野工場
                gsSystemCd = SYSTEM_200
            Case "40"               ''米沢工場
                gsSystemCd = SYSTEM_200
            Case "42"               ''３００ｍｍ
                gsSystemCd = SYSTEM_300
            Case "43"               ''３００ｍｍ
                gsSystemCd = SYSTEM_300
            Case "90"               ''テスト環境
                gsSystemCd = SYSTEM_200
            Case "91"               ''テスト環境(新) 2007/04/05追加 SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "92"               ''テスト環境(生野) 2009/11/20追加 SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "AM"               ''尼崎工場 2009/06/02追加 SSS.Marushita
                gsSystemCd = SYSTEM_200
            Case "93"               ''生野A1テスト 2010/04/14追加 SETsw kubota
                gsSystemCd = SYSTEM_200
            Case "94"               ''尼崎テスト 2009/06/03追加 SSS.Marushita
                gsSystemCd = SYSTEM_200
            Case Else               ''外販
                gsSystemCd = SYSTEM_200
        End Select
    '2004/11/12 tcs tagawa 追加  end---------------------------------------------------
    
        GetCmdLine_Re = True
    End Function
    
    '///////////////////////////////////////////////////
    ' @(f) 2004/11/12 TCS TAGAWA
    '
    ' 機能      : ﾌｫｰﾑのｷｬﾌﾟｼｮﾝ設定
    '
    ' 返り値    :　なし
    '
    ' 引数      : frmForm       - ﾌｫｰﾑ
    '             sProgramId    - ﾌﾟﾛｸﾞﾗﾑID
    
    '
    ' 機能説明  :　ﾌｫｰﾑのｷｬﾌﾟｼｮﾝ設定
    '
    ' 備考      :
    '
    '///////////////////////////////////////////////////
    Public Sub SetFormCaption(frmForm As Form, ByVal sProgramId As String)
    
        ''200mmの場合
        If gsSystemCd = SYSTEM_200 Then
            frmForm.Caption = sProgramId & " - " & SYSTEM_NAME_200
        ''300mmの場合
        Else
            frmForm.Caption = sProgramId & " - " & SYSTEM_NAME_300
        End If
    
    End Sub
    
    '概要      :プログラム起動時の初期化処理
    '説明      :
    Public Function InitExe_Re() As FUNCTION_RETURN
        
        '' プログラム起動時の初期化処理
        DoEvents
        
        '' パラメータ初期化
        InitExe_Re = FUNCTION_RETURN_SUCCESS
       
        '' エラー出力オブジェクト作成
        Init_ErrHandler_Re
        
        ''コマンドライン引数取得
        If GetCmdLine_Re() = False Then
            ''コマンドライン引数無し
            Call MsgOut(64, "", ERR_DISP_LOG)
            Exit Function
        End If
        
           ''実行ファイル名の取得
        If GetEXEName = "" Then
            ''コマンドライン引数無し
            Call MsgOut(62, "", ERR_DISP_LOG)
            Exit Function
        End If
     
        '' 多重起動チェック
        If App.PrevInstance = True Then
            '' 多重起動している場合
            '' エラーメッセージ＆ログ出力
            MsgBox "すでにプログラムが起動されています。", vbOKOnly + vbInformation
            InitExe_Re = FUNCTION_RETURN_FAILURE
            End
        End If
        
        '' データベース接続
        OraDBOpen
        
        '' 処理終了
    
    End Function
    '概要      :プログラム起動時の初期化処理
    '説明      :多重起動を許可 2008/04/30 追加 Info.Kameda
    Public Function InitExe_Re_Ref() As FUNCTION_RETURN
        
        '' プログラム起動時の初期化処理
        DoEvents
        
        '' パラメータ初期化
        InitExe_Re_Ref = FUNCTION_RETURN_SUCCESS
       
        '' エラー出力オブジェクト作成
        Init_ErrHandler_Re
        
        ''コマンドライン引数取得
        If GetCmdLine_Re() = False Then
            ''コマンドライン引数無し
            Call MsgOut(64, "", ERR_DISP_LOG)
            Exit Function
        End If
        
           ''実行ファイル名の取得
        If GetEXEName = "" Then
            ''コマンドライン引数無し
            Call MsgOut(62, "", ERR_DISP_LOG)
            Exit Function
        End If
     
        ''' 多重起動チェック
        'If App.PrevInstance = True Then
        '    '' 多重起動している場合
        '    '' エラーメッセージ＆ログ出力
        '    MsgBox "すでにプログラムが起動されています。", vbOKOnly + vbInformation
        '    InitExe_Re = FUNCTION_RETURN_FAILURE
        '    End
        'End If
        
        '' データベース接続
        OraDBOpen
        
        '' 処理終了
    
    End Function
    Private Sub Init_ErrHandler_Re()
        Set gErr = New CErrHandler
        With gErr
            .AppTitle = App.Title
            .Destination = App.Path & "\Err.log"
            .DisplayMsgOnError = True
            .MaxProcStackItems = 20
            .IncludeExpandedInfo = False
        End With
    End Sub
    
    '*** UPDATE START T.TERAUCHI 2004/11/17 追加
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能    : 実行したSQL文をログに残す
    '
    ' 返り値  : 成否
    '
    ' 引き数  :　sHostName  -マシン名（起動元）
    '           sAppName    -プログラム名（起動元）
    '           sFncName    -関数名（起動元）
    '           sSQL        -実行クエリー
    '           sMemo       -メモ
    '
    ' 機能説明: 実行したSQL文をTBCMC003テーブルに書き込む
    '
    ' 備考    :
    '
    '///////////////////////////////////////////////////
    Public Function WriteDBLog_Re(ByVal sHostName As String, ByVal sAppName As String, _
                              ByVal sFncName As String, ByVal sSqlLog As String, _
                              ByVal sMemo As String) As Boolean
    
        Dim sSql        As String
        Dim iRet        As Long
        Dim sDbName     As String
        Dim sUID        As String
        Dim sPWD        As String
        Dim bErrFlag    As Boolean
        Dim sTableName  As String
        
    On Error GoTo ErrHand
    
        WriteDBLog_Re = False
        
        Select Case gsFactryCd
        Case "10"               ''野田工場
            sDbName = "NODA"
            sUID = "oracle"
            sPWD = "oracle"
        Case "30"               ''生野工場
            sDbName = "IKNO"
            sUID = "oracle"
            sPWD = "oracle"
        Case "40"               ''米沢工場
            sDbName = "YONE"
            sUID = "oracle"
            sPWD = "oracle"
        Case "42"               '’３００ｍｍ
            sDbName = "cm1"
            sUID = "cm1"
            sPWD = "cm1"
        Case "43"               '’３００ｍｍ
            sDbName = "cmt"
            sUID = "cm1"
            sPWD = "cm1"
    '2001/02/24 FFC start
        'Case "43"               ''SUMCO
        '    sDbName = "CMT"
        '    sUID = "cm1"
        '    sPWD = "cm1"
        '    gsFactryCd = "40"
    '2001/02/24 FFC end
        Case "90"               ''テスト環境
            sDbName = "TEST0"
            sUID = "oracle"
            sPWD = "oracle"
        Case "91"               ''テスト環境(新) 2007/04/05追加 SETsw kubota
            sDbName = "CLA0X"
            sUID = "oracle"
            sPWD = "oracle"
        Case "99"               ''仮
            sDbName = "BOIS"
            sUID = "BOIS"
            sPWD = "BOIS"
        Case Else               ''外販
            sDbName = "oracle"
            sUID = "oracle"
            sPWD = "oracle"
        End Select
        
        ''オラクル接続
        Set gobjOraSess2 = CreateObject("OracleInProcServer.XOraSession")
        Set gobjOraDB2 = gobjOraSess2.OpenDatabase(sDbName, sUID & "/" & sPWD, 0&)
        
        If sHostName = "" Then
            sHostName = " "
        End If
        
        If sAppName = "" Then
            sAppName = " "
        End If
        
        If sFncName = "" Then
            sFncName = " "
        End If
        
        If sSqlLog = "" Then
            sSqlLog = " "
        End If
        
        If sMemo = "" Then
            sMemo = " "
        End If
        
        sSqlLog = Replace(sSqlLog, "'", "''")
        
        ''テーブル名取得 2005/07/13 tuku
        If gsSystemCd = SYSTEM_200 Then
            sTableName = "KODZL"
        Else
            sTableName = "TBCMC003"
        End If
        
        sSql = ""
        sSql = sSql & "insert into " & sTableName & " (           " & vbLf
        sSql = sSql & "                     L_DATE                  " & vbLf    ''ログをとった日時
        sSql = sSql & "                     ,SEQ                    " & vbLf    ''ログのシーケンス
        sSql = sSql & "                     ,HOSTNAME               " & vbLf    ''起動元マシン名
        sSql = sSql & "                     ,APPNAME                " & vbLf    ''起動元プログラム名
        sSql = sSql & "                     ,FNCNAME                " & vbLf    ''起動元関数名
        sSql = sSql & "                     ,SQL                    " & vbLf    ''SQL文ログ
        sSql = sSql & "                     ,MEMO               )   " & vbLf    ''メモ
        sSql = sSql & "values(              sysdate                 " & vbLf    ''ログをとった日時
        sSql = sSql & "                     ,log_seq.nextval        " & vbLf    ''ログのシーケンス
        sSql = sSql & "                     ,'" & sHostName & "'    " & vbLf    ''起動元マシン名
        sSql = sSql & "                     ,'" & sAppName & "'     " & vbLf    ''起動元プログラム名
        sSql = sSql & "                     ,'" & sFncName & "'     " & vbLf    ''起動元関数名
        sSql = sSql & "                     ,'" & sSqlLog & "'      " & vbLf    ''SQL文ログ
        sSql = sSql & "                     ,'" & sMemo & "'    )   " & vbLf    ''メモ
            
        Set gobjOraSess2 = CreateObject("OracleInProcServer.XOraSession")
            
            
        ''ﾄﾗﾝｻﾞｸｼｮﾝ開始
        gobjOraSess2.BeginTrans
        
        ''オラクルＳＱＬ実行
        iRet = gobjOraDB2.DbExecuteSQL(sSql)
            
        'SQL実行
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC003")
            ''ｴﾗｰ処理
            bErrFlag = True
            GoTo ErrHand
        ElseIf iRet = 0 Then
            Call MsgOut(71, "TBCMC003", ERR_DISP_LOG)
            ''ｴﾗｰ処理
            bErrFlag = True
            GoTo ErrHand
            Exit Function
        End If
        
        ''コミット処理
        gobjOraSess2.CommitTrans
        
        ''オラクル切断
        gobjOraDB2.Close
        
        ''解放
        Set gobjOraDB2 = Nothing
        Set gobjOraSess2 = Nothing
        
        WriteDBLog_Re = True
        
        Exit Function
    
ErrHand:
    
        ''ﾛｰﾙﾊﾞｯｸ処理
        gobjOraSess2.Rollback
        
        ''オラクル切断
        gobjOraDB2.Close
        
        ''解放
        Set gobjOraDB2 = Nothing
        Set gobjOraSess2 = Nothing
    
        ''VBエラー
        If Not bErrFlag Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "")
        End If
        
    End Function
    '2004/9/27tcs Yamauchi 追加 end-------------------------------------
    
    '2004/9/17tcs Suenaga 追加 start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    ' 機能    : 精製原料管理の更新を行う
    '
    ' 返り値  : True:成功　False:失敗
    '
    ' 引き数  :  genryoNo  - 原料No
    '           genpinNo  - 現品ロットNo
    '           kouteiCd  - 工程コード
    '           tantoCd   - 担当者コード
    '           ukeireW   - 受入重量
    '           haraiW    - 払出重量
    '           losW      - ロス重量
    '           koujyoCd  - 工場コード
    '           ukeireCd  - 受入工場コード
    '           haraiCd   - 払出工場コード
    '           syoumetu  - 消滅区分(1:現品消滅 2現品ロットNo消滅)
    '           shikake   - 仕掛区分(1:仕掛あり)
    '           sisutemu  - システム区分コード
    '           noudoK    - 濃度区分
    '           noudoT    - 濃度値
    '           sousinFlg - 原料送信フラグ(1:送信 7:送信なし)
    '           hasseiFlg - 発生フラグ(0:未送信 7:送信なし)
    '           motoConce - 元濃度
    '           yoteiFac  - 使用予定工場
    '           tanaire   - 棚入れ区分(0:棚入れ無し 1:棚入れ)
    '           sChg      - チャージNo(CC200/CC300登録用)　05/08/23 ooba
    '           motoSyscd - 原料取得元(200mm/300mm)　2008/06/16 SET/miyatake
    '
    ' 機能説明: 原料番号(XODB1)、原料仕掛工程(XODB2)、原料工程実績(XODB#)
    '          の更新を行う
    '
    '///////////////////////////////////////////////////
    Public Function Upd_XODB123(genryoNo, _
                                genpinNo, _
                                kouteiCd, _
                                tantoCd, _
                                ukeireW, _
                                HARAIW, _
                                losW, _
                                koujyoCd, _
                                ukeireCd, _
                                haraiCd, _
                                syoumetu, _
                                shikake, _
                                sisutemu, _
                                noudoK, _
                                noudoT, _
                                sousinFlg, _
                                hasseiFlg, _
                                motoConce, _
                                yoteiFac, _
                                tanaire, _
                                Optional sChg As String = "", _
                                Optional motoSyscd As String = "") As Boolean
        
        Dim objOraDyn As Object     'ダイナセットオブジェクト
        Dim sSql      As String     'SQL文
        Dim sUserName As String     '担当者名
        Dim nowdate   As String     'システム日付
        Dim cyoku     As String     '直区分
        Dim year      As String     '実績日　年
        Dim month     As String     '　　　　月
        Dim day       As String     '　　　　日
        Dim hour      As String     '　　　　時
        Dim Min       As String     '　　　　分
        Dim before    As String     '前工程
        Dim after     As String     '次工程
        Dim sikakeK   As String     '仕掛工程
        
        '戻り値設定
        Upd_XODB123 = False
        
    'エラーハンドラ
    On Error GoTo ErrHand
        
        '原料Noがない時
        If genryoNo = "" Then
            'メッセージ表示
            Call MsgOut(0, "原料Noがありません", ERR_DISP)
            '処理を抜ける
            Exit Function
        '原料Noがあるとき
        Else
            'プライベート変数に格納
            mtrlNo = genryoNo
        End If
        
    '---- UPD [空文字変換前にTrimを追加] 2004/10/18 TCS)R.Kawaguchi START ----
        ''パラメータがNULLの時、空文字に変換
        '現品ロットNo
        cryno = NulltoStr(Trim(genpinNo))
        '工程コード
        PROCCD = NulltoStr(Trim(kouteiCd))
        '工程コードが5文字(CB220等)の場合、右4文字のみを取得
        If Len(Trim(PROCCD)) = 5 Then
            PROCCD = Right(PROCCD, 4)
        End If
        '担当者コード
        staffCd = NulltoStr(Trim(tantoCd))
        '受入重量
        recW = NulltoStr(Trim(ukeireW))
        '払出重量
        sendW = NulltoStr(Trim(HARAIW))
        'ロス重量
        lossW = NulltoStr(Trim(losW))
        '工場コード
        factCd = NulltoStr(Trim(koujyoCd))
        '受入工場コード
        recCd = NulltoStr(Trim(ukeireCd))
        '払出工場コード
        sendCd = NulltoStr(Trim(haraiCd))
        '消滅区分
        disapp = NulltoStr(Trim(syoumetu))
        '仕掛区分
        sikake = NulltoStr(Trim(shikake))
        'システム区分コード
        sysCd = NulltoStr(Trim(sisutemu))
        '濃度区分
        conceK = NulltoStr(Trim(noudoK))
        '濃度値
        conceT = NulltoStr(Trim(noudoT))
        '原料送信フラグ
        SENDFLG = NulltoStr(Trim(sousinFlg))
        '発生フラグ
        occuFlg = NulltoStr(Trim(hasseiFlg))
        '元濃度
        conceM = NulltoStr(Trim(motoConce))
        '使用予定工場
        planFac = NulltoStr(Trim(yoteiFac))
        '棚入れ区分
        tanaKu = NulltoStr(Trim(tanaire))
    '---- UPD [空文字変換前にTrimを追加] 2004/10/18 TCS)R.Kawaguchi END ----
    
    '*** UPDATE START T.TERAUCHI 2004/10/19 払出区分追加
        '払出区分
        stowkkbb3 = " "
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
        'チャージNo(CC200/CC300登録用)　05/08/23 ooba
        sChgNo = sChg
    
        '日付を取得
    '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
    
        'プライベート変数に格納
        nowdate = gsSysdate
    
        ''300mmの入力担当者の権限チェック
        If Not GetAuthorityUser_300(staffCd, "C" & PROCCD, sUserName) Then
            '処理を抜ける
            Exit Function
        End If
        
        '原料Noが取得できなかった時
        If SqlCheck = False Then
            ''メッセージ表示
            Call MsgOut(0, "原料Noの入力に誤りがあります", ERR_DISP)
            '処理を抜ける
            Exit Function
        End If
        
    '---- UPD 2004/10/18 TCS)R.Kawaguchi START ----
    ''    'サーバーシステム日付を実働日に変更
    ''    nowdate = GetJITUDATE(nowdate)
    ''
    ''    '実働日より直区分を判定
    ''    cyoku = GetCYOKU(nowdate)
    ''
    ''    '実績日から切り取り
    ''    year = Mid(nowdate, 1, 4)     '年
    ''    month = Mid(nowdate, 6, 2)    '月
    ''    day = Mid(nowdate, 9, 2)      '日
    ''    hour = Mid(nowdate, 12, 2)    '時
    ''    min = Mid(nowdate, 15, 2)     '分
    
       'サーバーシステム日付を実働日に変更
        nowdate = GetJITUDATE(Format(nowdate, "yyyymmddhhmmss"))
        
        '実働日より直区分を判定
        cyoku = GetCYOKU(gsSysdate)
        
        '実績日から切り取り
        year = Mid(nowdate, 1, 4)     '年
        month = Mid(nowdate, 5, 2)    '月
        day = Mid(nowdate, 7, 2)      '日
        hour = Mid(nowdate, 9, 2)    '時
        Min = Mid(nowdate, 11, 2)     '分
    '---- UPD 2004/10/18 TCS)R.Kawaguchi END ----
        
        '実績日チェック
        If CheckDateFormat_Re(year, month, day, hour, Min) = False Then
            '実績日が妥当でない時、メッセージ表示
            Call MsgOut(0, "実績の時間フォーマットが不正です", ERR_DISP_LOG)
            '処理を抜ける
            Exit Function
        End If
        
        '工程コードから前工程、次工程、仕掛工程設定
        ' upd 原料在庫統合による修正  2008/06/16 SET/miyatake ===================> START
        'Call Settei(before, after, sikakeK)
        Call Settei(before, after, sikakeK, motoSyscd)
        ' upd 原料在庫統合による修正  2008/06/16 SET/miyatake ===================> END
        
        '原料番号(XODB1)更新
        If Not Upd_XODB1() Then GoTo ErrHand
        
        '原料仕掛工程(XODB2)更新
        If Not Upd_XODB2(sikakeK) Then GoTo ErrHand
        
        '原料仕掛工程(XODB2)追加
        If Not Ins_XODB2(after, cyoku) Then GoTo ErrHand
        
        '原料工程実績(XODB3)追加
        If Not Ins_XODB3(before, after, cyoku, staffCd, sUserName) Then GoTo ErrHand
        
        '戻り値設定
         Upd_XODB123 = True
        Exit Function
    
    'エラー時
ErrHand:
        Call MsgOut(72, "", ERR_DISP_LOG)
    End Function
    
    ' @(f)
    ' 機能      : 原料Noチェック
    '
    ' 返り値    : True:原料No存在 False:原料No非存在
    '
    ' 引き数    : なし
    '
    ' 機能説明  : XODB1に原料Noが存在するかチェックする
    '
    Private Function SqlCheck() As Boolean
        Dim sSql As String      'SQL文格納
        Dim objOraDyn As Object
    
        '戻り値設定
        SqlCheck = False
        
        ''SQL文作成
        sSql = ""
        sSql = sSql & " SELECT polnob1                    " & vbLf
        sSql = sSql & " FROM   xodb1                      " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "' "
        
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = False Then
            '原料Noが取得できなかった時、処理を抜ける
            Exit Function
        End If
        
        '戻り値設定
        SqlCheck = True
    End Function
    
    
    ' @(f)
    '
    ' 機能    : 前工程、次工程、仕掛工程取得
    '
    ' 返り値  : なし
    '
    ' 引き数  : before :前工程
    '           after  :次工程
    '           sikakeK:仕掛工程
    '           motoSyscd:原料取得元(200mm/300mm)
    '
    ' 機能説明: 工程コードから前工程、次工程、仕掛工程取得
    '
    ' 備考    :
    '
    Private Function Settei(ByRef before As String, ByRef after As String, ByRef sikakeK As String, Optional ByVal motoSyscd As String = "")
        
    '*** UPDATE START T.TERAUCHI 2004/10/18 使用しないので削除
    '    '精製原料受入
    '        If procCd = "B410" Then
    '            If disapp = "1" Then          '廃棄
    '                before = "B410"
    '                after = "ZZZZ"
    '            '受入仕掛
    '            Else                          '切断仕掛
    '                before = "C450"
    '                after = "B510"
    '            End If
    '        End If
    '
    '        '精製原料切断
    '        If procCd = "B510" Then
    '            before = "B410"
    '            If disapp = "1" Then         '原料No消滅
    '                after = "ZZZZ"
    '            ElseIf disapp = "2" Then     '現品ロット消滅
    '                after = "ZZZZ"
    '            ElseIf sikake = "1" Then     '切断仕掛
    '                after = "B510"
    '            Else                         '洗浄受入仕掛
    '                after = "B220"
    '            End If
    '        End If
    '
    '        'ロット構成
    '        If procCd = "B610" Then
    '            before = "B510"
    '            sikakeK = "B220"
    '            If disapp = "1" Then         '原料ロット統合前
    '                after = "ZZZZ"
    '            Else                         '原料ロット統合後
    '                after = "B220"
    '            End If
    '        End If
    '*** UPDATE END T.TERAUCHI 2004/10/18
       
            '洗浄受入
            If PROCCD = "B220" Then          '洗浄払出仕掛
                before = "B510"
                after = "B225"
                sikakeK = "B220"
            End If
            
            '洗浄受入仕掛
            If PROCCD = "B620" Then
                If sendCd <> "" Then         '他工場払出
                    before = "B510"
                    after = "RP00"
                    sikakeK = "B220"
                ElseIf disapp = "1" Then     '廃棄
                    before = "B510"
                    after = "ZZZZ"
                    sikakeK = "B220"
                                
                '*** UPDATE START T.TERAUCHI 2004/10/18 洗浄受入一覧の廃棄時、実績工程はB220とする
                    PROCCD = "B220"
                '*** UPDATE END   T.TERAUCHI 2004/10/18
    
                End If
            End If
            
            '洗浄払出
            If PROCCD = "B225" Then
                before = "B220"
                sikakeK = "B225"
                
            '*** UPDATE START T.TERAUCHI 2004/10/8 引上投入仕掛の場合
                after = "C200"
            '*** UPDATE END T.TERAUCHI 2004/10/8
            
                If sendCd <> "" Then         '他工場払出
                    after = "RP00"
                ElseIf disapp = "1" Then     '廃棄
                    after = "ZZZZ"
                ElseIf sikake = "1" Then     '再洗浄
                    after = "B225"
                End If
            End If
            
            '原料棚入れ(他工場)
            If PROCCD = "B230" Then
                If recCd <> "" Then
                    before = "RP00"
                    sikakeK = "B230"
                    If sendCd <> "" Then     '他工場受入-引上投入仕掛
                        after = "RP00"
                    Else
                        after = "C200"       '他工場受入-他工場払出
                    End If
                Else
                    If sendCd <> "" Then     '他工場払出
                        before = "B225"
                        after = "RP00"
                        sikakeK = "C200"
                    End If
                End If
            End If
            
        '*** UPDATE START T.TERAUCHI 2004/10/18 原料受入工程
            If PROCCD = "B240" Then
                before = "B240"
                
                If disapp = "1" Then     '廃棄
                
                    after = "ZZZZ"
                    sikakeK = "C200"
                    
                '*** UPDATE START T.TERAUCHI 2004/10/19 払出区分＝1(ロス)を追加
                    stowkkbb3 = "1"
                '*** UPDATE END   T.TERAUCHI 2004/10/19
                
                Else
                    after = "C200"
                    
            '*** UPDATE START T.TERAUCHI 2004/10/20 処理区分対応
                    If sysCd <> "" Then
                        stowkkbb3 = sysCd
                    End If
            '*** UPDATE END   T.TERAUCHI 2004/10/20
            
            '*** UPDATE START T.TERAUCHI 2004/10/20
                    '新原料修正の場合
                    If sikake = "1" Then
                        sikakeK = "C200"
                    '新原料追加の場合
                    Else
                        sikakeK = ""
                    End If
            '*** UPDATE END   T.TERAUCHI 2004/10/20
                
                End If
                
            End If
        
        '*** UPDATE END   T.TERAUCHI 2004/10/18
    
    
    
            '引上投入
        '*** UPDATE START T.TERAUCHI 2004/10/18 引上投入の工程変更
    '        If procCd = "C200" Then           '引上終了仕掛
    '            before = "B225"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
            '引上投入
        '*** UPDATE START T.TERAUCHI 2004/10/18 引上投入の工程変更
    '        If procCd = "C200" Then           '引上終了仕掛
    '            before = "B225"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
    
    ''Start 2004/10/22 Upd M.Yamauchi---------------------------------------
    '        If procCd = "C250" Then           '引上終了仕掛
    '            before = "C200"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
    ' upd 原料在庫統合による修正  2008/06/16 SET/miyatake ===================> START
    '        If PROCCD = "C200" Then           '引上終了仕掛
    '            before = "C100"
    '            after = "C300"
    '            sikakeK = "C200"
    '        End If
            If PROCCD = "C200" Then           '引上終了仕掛
                Dim work As String
                If motoSyscd = SYSTEM_200 Then
'>>>>> 工程コード変更 2008/07/25 SET.Marushita
                    before = "B690"
                    'before = "B990"
'<<<<< 工程コード変更 2008/07/25 SET.Marushita
                    after = "C300"
                    sikakeK = "C200"
                Else
                    before = "C100"
                    after = "C300"
                    sikakeK = "C200"
                End If
            End If
    ' upd 原料在庫統合による修正  2008/06/16 SET/miyatake ===================> END
    ''End 2004/10/22---------------------------------------------------------

        '*** UPDATE END   T.TERAUCHI 2004/10/18
            
            '引上終了
            If PROCCD = "C300" Then
'                If disapp = "1" Then          '引上終了完了         'ｺﾒﾝﾄ化　05/08/23 ooba
                    
            '*** UPDATE START T.TERAUCHI 2004/11/02
    '            '*** UPDATE START T.TERAUCHI 2004/10/18
    '            '    before = "C200"
    '                before = "C250"
    '            '*** UPDATE END   T.TERAUCHI 2004/10/18
                    
                    before = "C200"
            
            '*** UPDAT END    T.TERAUCHI 2004/11/02
                    
                    after = "ZZZZ"
                    sikakeK = "C300"
'                End If
            End If
        
    End Function
    
    
    
    
    
    ' @(f)
    '
    ' 機能    : 原料番号更新処理
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : なし
    '
    ' 機能説明: 原料番号(XODB1)更新処理
    '
    ' 備考    :
    '
    Private Function Upd_XODB1() As Boolean
        Dim sSql    As String        'SQL文格納
        Dim iRet    As Integer       'データ更新数
        Dim renban  As String        '工程連番
        Dim objOraDyn As Object
    
        '工程コードがB410の時は処理を抜ける
        If PROCCD = "B410" Then
            '戻り値設定
            Upd_XODB1 = True
            Exit Function
        End If
        
        '工程コードがB450で消滅区分が2の時、処理を抜ける
        If PROCCD = "B450" And disapp = "2" Then
            '戻り値設定
            Upd_XODB1 = True
            Exit Function
        End If
    
    'エラーハンドラ
    On Error GoTo ErrHand
        
        '戻り値設定
        Upd_XODB1 = False
        
        '工程連番取得
        sSql = ""
        sSql = sSql & " SELECT kcntb1                               " & vbLf
        sSql = sSql & " FROM   xodb1                                " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'           " & vbLf
        
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = True Then
            '取得したデータを格納
            renban = NulltoStr(objOraDyn.Fields("kcntb1").Value)
        End If
        
        'SQL文作成
        sSql = ""
        sSql = sSql & "UPDATE xodb1                                 " & vbLf
        
        '工程連番がNULLの時
        If renban = "" Then
            sSql = sSql & "SET kcntb1 = 1                           " & vbLf  '工程連番
        Else
            sSql = sSql & "SET kcntb1 = kcntb1 + 1                  " & vbLf
        End If
        
        'システム日付取得
    '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
        
        '*** UPDATE START T.TERAUCHI 2004/10/14 払出工場ｺｰﾄﾞも変更
            If sendCd <> "" Then
                sSql = sSql & "    ,toworkb1 = '" & sendCd & "'"
            End If
        '*** UPDATE END   T.TERAUCHI 2004/10/14
        
        '*** UPDATE START T.TERAUCHI 2004/10/18 修正日付も更新
        sSql = sSql & "        ,rdayb1 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/18
        
        '*** UPDATE START T.TERAUCHI 2004/11/2 送信日付は設定しない
        'sSql = sSql & "       ,sdayb1 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf '送信日付
        '*** UPDATE END   T.TERAUCHI 2004/11/2
        
        sSql = sSql & "       ,sndkb1 = ' '                        " & vbLf   '送信区分
        sSql = sSql & "WHERE  polnob1 = '" & mtrlNo & "'           " & vbLf
    
        '実行
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB1")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB1", ERR_DISP_LOG)
            Exit Function
        End If
    
        '戻り値設定
        Upd_XODB1 = True
        Exit Function
    
    'エラー時
ErrHand:
        ''メッセージ表示
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB1")
    End Function
    
    
    
    ' @(f)
    '
    ' 機能    : 原料仕掛工程更新処理
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : sikakeK:仕掛工程コード
    
    '
    ' 機能説明: 原料仕掛工程(XODB2)更新処理
    '
    ' 備考    :
    '
    Private Function Upd_XODB2(sikakeK As String) As Boolean
        Dim sSql      As String        'SQL文格納
        Dim KUBUN     As String        '原料コード
        Dim syurui    As String        '原料種類コード
        Dim objOraDyn As Object
        Dim iRet      As Integer       'データ更新数
         
    'エラーハンドラ
    On Error GoTo ErrHand
        
        '仕掛工程コードがない時は処理を抜ける
        If sikakeK = "" Then
            '戻り値設定
            Upd_XODB2 = True
            Exit Function
        End If
        
        '棚入れ区分が1の時は処理を抜ける
        If tanaKu = "1" Then
            '戻り値設定
            Upd_XODB2 = True
            Exit Function
        End If
        
        '原料No、仕掛工程コードの存在チェック
        'SQL文作成
        sSql = ""
        sSql = sSql & " SELECT polnob2,                   " & vbLf
        sSql = sSql & "        wkktb2                     " & vbLf
        sSql = sSql & " FROM   xodb2                      " & vbLf
        sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "' " & vbLf
        sSql = sSql & " AND    wkktb2 = '" & sikakeK & "' " & vbLf
        
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = False Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        Else
            If objOraDyn.EOF = True Then
                Upd_XODB2 = True
                '原料No、仕掛工程コードが存在しない時は処理を抜ける
                Exit Function
            End If
        End If
        
        '戻り値設定
        Upd_XODB2 = False
    
        'XODB1からXODB2更新のためのパラメータ取得
        'SQL文作成
        sSql = ""
        sSql = sSql & " SELECT pokubb1,                            " & vbLf  '原料区分
        sSql = sSql & "        pokidcb1                            " & vbLf  '原料種類コード
        sSql = sSql & " FROM   xodb1                               " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'          " & vbLf
        
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = True Then
            '取得したデータを格納
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
        
        '消滅区分が1の時
        If disapp = "1" Then
            'SQL文作成
            sSql = ""
            sSql = sSql & " UPDATE xodb2                           " & vbLf
            sSql = sSql & " SET    siwb2 = 0                       " & vbLf '仕掛重量
            sSql = sSql & "        ,sikosub2 = 0                   " & vbLf '仕掛個数
            
            'システム日付取得
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
            
            '*** UPDATE START T.TERAUCHI 2004/10/18 修正日付も更新
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
            
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '原料区分
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '原料種類コード
            
        '*** UPDATE START T.TERAUCHI 2004/10/25
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '送信区分
            sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    wkktb2 = '" & sikakeK & "'      " & vbLf
        
        '消滅区分が1以外のとき
        Else
            'SQL文作成
            sSql = ""
            sSql = sSql & " UPDATE xodb2                           " & vbLf
            sSql = sSql & " SET    siwb2 = siwb2 - " & val(recW) & vbLf     '仕掛重量
            sSql = sSql & "        ,sikosub2 = sikosub2 - 1        " & vbLf '仕掛個数
            
            'システム日付取得
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
            '*** UPDATE START T.TERAUCHI 2004/10/18 修正日付も更新
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
        
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '原料区分
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '原料種類コード
            
        '*** UPDATE START T.TERAUCHI 2004/10/25 送信日付の変更はなし
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss') " & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
                
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '送信区分
            sSql = sSql & " WHERE  polnob2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    wkktb2 = '" & sikakeK & "'      " & vbLf
        End If
        
        '実行
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB2", ERR_DISP_LOG)
            Exit Function
        End If
    
        '戻り値設定
        Upd_XODB2 = True
        Exit Function
    
    'エラー時
ErrHand:
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB2")
    End Function
    
    
    
    ' @(f)
    '
    ' 機能    : 原料仕掛工程追加処理
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : after:次工程コード
    '           cyoku:直区分
    '
    ' 機能説明: 原料仕掛工程(XODB2)追加処理
    '
    ' 備考    :
    '
    Private Function Ins_XODB2(after As String, cyoku As String) As Boolean
        Dim sSql      As String       'SQL文格納
        Dim KUBUN     As String       '原料区分
        Dim syurui    As String       '原料種類コード
        Dim objOraDyn As Object
        Dim iRet    As Integer        'データ追加数
        
        '次工程コードがRP00、ZZZZ、B510の時
        If after = "RP00" Or after = "ZZZZ" Or after = "B510" Then
            Ins_XODB2 = True
            '処理を抜ける
            Exit Function
        End If
        
        'エラーハンドラ
        On Error GoTo ErrHand
    
        '戻り値設定
        Ins_XODB2 = False
        
        'XODB1からXODB2更新のためのパラメータ取得
        'SQL文作成
        sSql = ""
        sSql = sSql & " SELECT pokubb1,                        " & vbLf  '原料区分
        sSql = sSql & "        pokidcb1                        " & vbLf  '原料種類コード
        sSql = sSql & " FROM   xodb1                           " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'      " & vbLf
        
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = True Then
            '取得したデータを格納
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
        
    
        '原料No、次工程コードの存在チェック
        sSql = ""
        sSql = sSql & " SELECT polnob2,                            " & vbLf
        sSql = sSql & "        wkktb2                              " & vbLf
        sSql = sSql & " FROM   xodb2                               " & vbLf
        sSql = sSql & " WHERE  POLNOB2 = '" & mtrlNo & "'          " & vbLf
        sSql = sSql & " AND    WKKTB2 = '" & after & "'            " & vbLf
    
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = False Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        End If
        
        '該当データがある場合
        If objOraDyn.EOF = False Then
            
            'SQL文作成
            sSql = ""
            sSql = sSql & " UPDATE xodb2                       " & vbLf
     '*** UPDATE START T.TERAUCHI 2004/10/19 受入重量ではなく、払出重量を設定
        '    sSql = sSql & " SET    siwb2 = siwb2 + " & val(recW) & vbLf '仕掛重量
             sSql = sSql & " SET    siwb2 = siwb2 + " & val(sendW) & vbLf '仕掛重量
     '*** UPDATE END   T.TERAUCHI 2004/10/19
            sSql = sSql & "        ,sikosub2 = sikosub2 + 1    " & vbLf '仕掛個数
            
            'システム日付取得
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
        
            '*** UPDATE START T.TERAUCHI 2004/10/18 修正日付も更新
            sSql = sSql & "        ,rdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            '*** UPDATE END   T.TERAUCHI 2004/10/18
        
            sSql = sSql & "        ,gndayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
            sSql = sSql & "        ,pokubb2 = '" & KUBUN & "'      " & vbLf '原料区分
            sSql = sSql & "        ,pokidcb2 = '" & syurui & "'    " & vbLf '原料種類コード
        
        '*** UPDATE START T.TERAUCHI 2004/10/25 修正日付の更新はなし
        '    sSql = sSql & "        ,sdayb2 = to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')" & vbLf
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "        ,sndkb2 = ' '                   " & vbLf '送信区分
            sSql = sSql & " WHERE  POLNOB2 = '" & mtrlNo & "'      " & vbLf
            sSql = sSql & " AND    WKKTB2 = '" & after & "'        " & vbLf
    
            
        '該当データがなかった場合
        Else
            sSql = ""
            sSql = sSql & "INSERT INTO XODB2                       " & vbLf
            sSql = sSql & "            (polnob2,                   " & vbLf
            sSql = sSql & "            wkktb2,                     " & vbLf
            sSql = sSql & "            PLACB2,                     " & vbLf
    
    '*** UPDATE START T.TERAUCHI 2004/10/18 登録日付も追加
            sSql = sSql & "            TDAYB2,                     " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
            sSql = sSql & "            RDAYB2,                     " & vbLf
            sSql = sSql & "            SDAYB2,                     " & vbLf
            sSql = sSql & "            SNDKB2,                     " & vbLf
            sSql = sSql & "            SAKJB2,                     " & vbLf
            sSql = sSql & "            POKUBB2,                    " & vbLf
            sSql = sSql & "            POKIDCB2,                   " & vbLf
            sSql = sSql & "            SIWB2,                      " & vbLf
            sSql = sSql & "            SIKOSUB2,                   " & vbLf
            sSql = sSql & "            GNDAYB2,                    " & vbLf
            sSql = sSql & "            GNCYOKB2)                   " & vbLf
            sSql = sSql & " VALUES     ('" & mtrlNo & "',          " & vbLf '原料番号
            
            If after = "" Then
                sSql = sSql & "        ' ',                        " & vbLf '工程コード
            Else
                sSql = sSql & "        '" & after & "',            " & vbLf '工程コード
            End If
            
            sSql = sSql & "            ' ',                        " & vbLf 'ラインコード
    '*** UPDATE START T.TERAUCHI 2004/10/18 登録日付、修正日付に値設定
        '    sSql = sSql & "            null,                       " & vbLf '修正日付
            sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '登録日付"
            
        '*** UPDATE START T.TERAUCHI 2004/10/25 修正日付の登録はなし
        '    sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '修正日付"
            sSql = sSql & "            null,                       " & vbLf '修正日付
        '*** UPDATE END   T.TERAUCHI 2004/10/25
    
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
            'システム日付取得
            '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''        If Not GetSysdate() Then Exit Function
        
        '*** UPDATE START T.TERAUCHI 2004/10/25 送信日付の登録はなし
        '    sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '送信日付"
            sSql = sSql & "            null,                       " & vbLf '送信日付
        '*** UPDATE END   T.TERAUCHI 2004/10/25
            
            sSql = sSql & "            ' ',                        " & vbLf          '送信区分
            sSql = sSql & "            '0',                        " & vbLf          '削除区分
            sSql = sSql & "            '" & KUBUN & "',            " & vbLf          '原料区分
            sSql = sSql & "            '" & syurui & "',           " & vbLf          '原料種類コード
            sSql = sSql & "            " & val(sendW) & ",         " & vbLf          '仕掛重量
            sSql = sSql & "            1,                          " & vbLf          '仕掛個数
            sSql = sSql & "            to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss')," & vbLf '最新仕掛日付
            sSql = sSql & "            '" & cyoku & "')            " & vbLf          '直区分
        End If
        
        '実行
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB2")
            Exit Function
        End If
    
        '戻り値設定
        Ins_XODB2 = True
        Exit Function
    
    'エラー時
ErrHand:
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB2")
    End Function
    
    ' @(f)
    '
    ' 機能    : 原料工程実績追加処理
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : before  :前工程コード
    '          after    :次工程コード
    '          cyoku    :直区分
    '          staffCd  :担当者コード
    '          sUserName:担当者名
    '
    ' 機能説明: 原料工程(XODB3)追加処理
    '
    ' 備考    :
    '
    Private Function Ins_XODB3(before As String, _
                                after As String, _
                                cyoku As String, _
                                staffCd As String, _
                                sUserName As String) As Boolean
        Dim sSql       As String         'SQL文格納
        Dim iRet       As Integer        'データ追加数
        Dim renban     As String         'XODB1の工程連番
        Dim KUBUN      As String         '       原料区分
        Dim syurui     As String         '       原料種類コード
        Dim objOraDyn  As Object
        Dim year       As String         'システム日付(年)
        Dim month      As String         '           (月)
        Dim day        As String         '           (日)
        Dim hour       As String         '           (時)
        Dim Min        As String         '           (分)
        Dim sNowdate   As String         'SUMCO時間対応　05/08/23 ooba
    
    'エラーハンドラ
    On Error GoTo ErrHand
    
        '戻り値設定
        Ins_XODB3 = False
    
        'SUMCO時間対応　05/08/23 ooba START =======================================>
        sNowdate = gsSysdate
        'サーバーシステム日付を実績日に変更
        sNowdate = GetJITUDATE(Format(sNowdate, "yyyymmddhhmmss"))
        '実績日から切り取り
        year = Mid(sNowdate, 1, 4)     '年
        month = Mid(sNowdate, 5, 2)    '月
        day = Mid(sNowdate, 7, 2)      '日
        hour = Mid(sNowdate, 9, 2)     '時
        Min = Mid(sNowdate, 11, 2)     '分
        'SUMCO時間対応　05/08/23 ooba END =========================================>
        
        'XODB1からデータ取得
        'SQL文作成
        sSql = ""
        sSql = sSql & " SELECT kcntb1,                       " & vbLf   '工程連番
        sSql = sSql & "        pokubb1,                      " & vbLf   '原料区分
        sSql = sSql & "        pokidcb1                      " & vbLf   '原料種類コード
        sSql = sSql & " FROM   xodb1                         " & vbLf
        sSql = sSql & " WHERE  polnob1 = '" & mtrlNo & "'    " & vbLf
    
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = True Then
            '取得したデータを格納
            renban = NulltoStr(objOraDyn.Fields("kcntb1").Value)
            KUBUN = NulltoStr(objOraDyn.Fields("pokubb1").Value)
            syurui = NulltoStr(objOraDyn.Fields("pokidcb1").Value)
        End If
    
        '原料工程実績(XODB3)更新
        sSql = ""
        sSql = sSql & "insert into XODB3                        " & vbLf
        sSql = sSql & "            (POLNOB3,                    " & vbLf   '原料番号
        sSql = sSql & "            KCNTB3,                      " & vbLf   '工程連番
        sSql = sSql & "            CRSEQB3,                     " & vbLf   '処理連番
        sSql = sSql & "            TDAYB3,                      " & vbLf   '登録日付
        sSql = sSql & "            RDAYB3,                      " & vbLf   '修正日付
        sSql = sSql & "            SDAYB3,                      " & vbLf   '送信日付
        sSql = sSql & "            SNDKB3,                      " & vbLf   '送信区分
        sSql = sSql & "            SAKJB3,                      " & vbLf   '削除区分
        sSql = sSql & "            POKUBB3,                     " & vbLf   '原料区分
        sSql = sSql & "            POKIDCB3,                    " & vbLf   '原料種類コード
        sSql = sSql & "            POLTNB3,                     " & vbLf   '原料ロットNo
        sSql = sSql & "            MODKBB3,                     " & vbLf   '赤黒区分
        sSql = sSql & "            SUMKBB3,                     " & vbLf   '集計区分
        sSql = sSql & "            WKKTB3,                      " & vbLf   '工程コード
        sSql = sSql & "            PLACB3,                      " & vbLf   'ラインコード
        sSql = sSql & "            FRWB3,                       " & vbLf   '受入重量
        sSql = sSql & "            TOWB3,                       " & vbLf   '払出重量
        sSql = sSql & "            LOSWB3,                      " & vbLf   'ロス重量
        sSql = sSql & "            FRWKKTB3,                    " & vbLf   '受入重量コード
        sSql = sSql & "            TOWKKTB3,                    " & vbLf   '払出工程コード
        sSql = sSql & "            TOWKKBB3,                    " & vbLf   '払出工程区分
        sSql = sSql & "            TOWORKB3,                    " & vbLf   '払出工場コード
        sSql = sSql & "            TOPLACB3,                    " & vbLf   '払出ラインコード
        sSql = sSql & "            CHGNB3,                      " & vbLf   'チャージNo
        sSql = sSql & "            EYYB3,                       " & vbLf   '実績日付(年)
        sSql = sSql & "            EMMB3,                       " & vbLf   '実績日付(月)
        sSql = sSql & "            EDDB3,                       " & vbLf   '実績日付(日)
        sSql = sSql & "            ECYOKB3,                     " & vbLf   '直区分
        sSql = sSql & "            EHHB3,                       " & vbLf   '実績時間(時)
        sSql = sSql & "            EMIB3,                       " & vbLf   '実績時間(分)
        sSql = sSql & "            MANB3,                       " & vbLf   '担当者
        sSql = sSql & "            MANJB3,                      " & vbLf   '担当者名
        sSql = sSql & "            DENKB3,                      " & vbLf   '濃度区分
        sSql = sSql & "            DENSITYB3,                   " & vbLf   '濃度値
        sSql = sSql & "            GSNDFLGB3,                   " & vbLf   '原料送信フラグ
        sSql = sSql & "            HFLGB3,                      " & vbLf   '発生フラグ
        sSql = sSql & "            mdensityb3,                  " & vbLf   '元濃度
        sSql = sSql & "            plworkb3,                    " & vbLf   '使用予定工場
        sSql = sSql & "            htkbnb3)                     " & vbLf   '廃棄/連合区分
        sSql = sSql & "VALUES      ('" & mtrlNo & "',           " & vbLf   '原料番号
        
        '工程連番が空の時
        If renban = "" Then
            sSql = sSql & "        0,                           " & vbLf
        Else
            sSql = sSql & "        " & val(renban) & ",         " & vbLf   '工程連番
        End If
        
        sSql = sSql & "            1,                           " & vbLf   '処理連番
        
        'システム日付取得
        '--- DEL 2004/10/18 TCS)R.Kawaguchi
    ''    If Not GetSysdate() Then Exit Function
        
        sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '登録日付
    '*** UPDATE START T.TERAUCHI 2004/10/18
    '    sSql = sSql & "            null,                        " & vbLf   '修正日付
    
        '*** UPDATE START T.TERAUCHI 2004/10/25 修正日付の登録はなし
    '    sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '修正日付
        sSql = sSql & "            null,                        " & vbLf   '修正日付
        '*** UPDATE END   T.TERAUCHI 2004/10/25
    
    '*** UPDATE END   T.TERAUCHI 2004/10/18
    
        '*** UPDATE START T.TERAUCHI 2004/10/25 送信日付の登録はなし
    '    sSql = sSql & " to_date('" & gsSysdate & "','yyyy/mm/dd hh24:mi:ss'), " & vbLf   '送信日付
        sSql = sSql & "            null,                        " & vbLf   '修正日付
        '*** UPDATE END   T.TERAUCHI 2004/10/25
        
        sSql = sSql & "            ' ',                         " & vbLf   '送信区分
        sSql = sSql & "            '0',                         " & vbLf   '削除区分
        sSql = sSql & "            '" & KUBUN & "',             " & vbLf   '原料区分
        sSql = sSql & "            '" & syurui & "',            " & vbLf   '原料種類コード
        sSql = sSql & "            ' ',                         " & vbLf   '原料ロット番号
        sSql = sSql & "            ' ',                         " & vbLf   '赤黒区分
        sSql = sSql & "            ' ',                         " & vbLf   '集計区分
        sSql = sSql & "            '" & PROCCD & "',            " & vbLf   '工程コード
        sSql = sSql & "            '　',                        " & vbLf   'ラインコード
        
    ''Upd start t.terauchi 2005/04/15   B240の時、原料受入と、在庫修正で処理を分ける
        ''B240以外はこれまで通り
        If PROCCD <> "B240" Then
            sSql = sSql & "            " & val(recW) & ",           " & vbLf   '受入重量
        ''B240の時
        Else
            ''在庫修正(廃却)の時
            If disapp = "1" Then
                sSql = sSql & "            0,          " & vbLf   '受入重量
            ''在庫修正の時
            ElseIf lossW <> "0" Then
                sSql = sSql & "            0,           " & vbLf   '受入重量
            ''原料受入の時(受入設定)
            Else
                sSql = sSql & "            " & val(recW) & ",           " & vbLf   '受入重量
            End If
        End If
    '*** UPDATE START T.TERAUCHI 2004/10/21 廃棄の際はロス重量に加算する
    '    sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '払出重量
    '    sSql = sSql & "            " & val(lossW) & ",          " & vbLf   'ロス重量
        
        '廃棄のとき
'        2004/12/17 TCS NAKAJIMA update-start
         If disapp = "1" And PROCCD <> "C300" Then
'        If disapp = "1" Then
'        2004/12/17 TCS NAKAJIMA update-end
            sSql = sSql & "            0,                           " & vbLf   '払出重量
            sSql = sSql & "            " & val(sendW) & ",          " & vbLf   'ロス重量
        Else
            
        ''Upd start t.terauchi 2005/04/15
            ''B240以外はこれまで通り
            If PROCCD <> "B240" Then
                sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '払出重量
            ''B240の時
            Else
                ''在庫修正の時
                If lossW <> "0" Then
                    sSql = sSql & "            0,          " & vbLf   '払出重量
                ''原料受入の時(受入設定)
                Else
                    sSql = sSql & "            " & val(sendW) & ",          " & vbLf   '払出重量
                End If
            End If
            
            sSql = sSql & "            " & val(lossW) & ",          " & vbLf   'ロス重量
        End If
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
        ' upd 原料在庫統合による修正  2008/06/17 SET/miyatake ===================> START
'        sSQL = sSQL & "            '" & before & "',            " & vbLf   '受入工程コード
'        sSQL = sSQL & "            '" & after & "',             " & vbLf   '払出工程コード
        If PROCCD = "C200" And recW < 0 Then
            'マイナス実績(返却時)は受入と払出を逆に設定する
            sSql = sSql & "            '" & after & "',             " & vbLf   '受入工程コード
            sSql = sSql & "            '" & before & "',            " & vbLf   '払出工程コード
        Else
            sSql = sSql & "            '" & before & "',            " & vbLf   '受入工程コード
            sSql = sSql & "            '" & after & "',             " & vbLf   '払出工程コード
        End If
        ' upd 原料在庫統合による修正  2008/06/17 SET/miyatake ===================> END
        
    '*** UPDATE START T.TERAUCHI 2004/10/19
    '    sSql = sSql & "            ' ',                         " & vbLf   '払出区分
        sSql = sSql & "            '" & stowkkbb3 & "',          " & vbLf   '払出区分
    '*** UPDATE END   T.TERAUCHI 2004/10/19
    
    '*** UPDATE START T.TERAUCHI 2004/10/18 払出工場ｺｰﾄﾞが設定されていない場合は自工場を設定する
        If sendCd <> "" Then
            sSql = sSql & "            '" & sendCd & "',            " & vbLf   '払出工場コード
        Else
            sSql = sSql & "            '" & factCd & "',            " & vbLf   '払出工場コード
        End If
    '*** UPDATE END   T.TERAUCHI 2004/10/18
        
        sSql = sSql & "            ' ',                         " & vbLf   '払出ラインコード
        
'        sSql = sSql & "            ' ',                         " & vbLf   'チャージNo
        'チャージNoセット　05/08/23 ooba START ==============================================>
        If PROCCD <> "C300" And PROCCD <> "C200" Then
            sSql = sSql & "            ' ',                         " & vbLf   'チャージNo
        Else
            sSql = sSql & "            '" & sChgNo & "',            " & vbLf   'チャージNo
        End If
        'チャージNoセット　05/08/23 ooba END ================================================>

        'システム日付から年を切り取り
'        year = Left(gsSysdate, 4)       'ｺﾒﾝﾄ化　05/08/23 ooba
        sSql = sSql & "            '" & year & "',              " & vbLf   '実績日付(年)
        'システム日付から月を切り取り
'        month = Mid(gsSysdate, 6, 2)    'ｺﾒﾝﾄ化　05/08/23 ooba
        sSql = sSql & "            '" & month & "',             " & vbLf   '実績日付(月)
        'システム日付から日を切り取り
'        day = Mid(gsSysdate, 9, 2)      'ｺﾒﾝﾄ化　05/08/23 ooba
        sSql = sSql & "            '" & day & "',               " & vbLf   '実績日付(日)
        sSql = sSql & "            '" & cyoku & "',             " & vbLf   '直区分
        'システム日付から時を切り取り
'        hour = Mid(gsSysdate, 12, 2)    'ｺﾒﾝﾄ化　05/08/23 ooba
        sSql = sSql & "            '" & hour & "',              " & vbLf   '実績時間(時)
        'システム日付から分を切り取り
'        min = Mid(gsSysdate, 15, 2)     'ｺﾒﾝﾄ化　05/08/23 ooba
        sSql = sSql & "            '" & Min & "',               " & vbLf   '実績時間(分)
        sSql = sSql & "            '" & Right(staffCd, 7) & "',  " & vbLf  '担当者
        sSql = sSql & "            '" & sUserName & "',         " & vbLf   '担当者名
        sSql = sSql & "            '" & conceK & "',            " & vbLf   '濃度区分
        sSql = sSql & "            " & val(conceT) & ",         " & vbLf   '濃度値
        sSql = sSql & "            '" & SENDFLG & "',           " & vbLf   '原料送信フラグ
        sSql = sSql & "            '" & occuFlg & "',           " & vbLf   '発生フラグ
        sSql = sSql & "            " & val(conceM) & ",         " & vbLf   '元濃度
        sSql = sSql & "            '" & planFac & "',           " & vbLf   '使用予定工場
        
        '消滅区分1の時
        If disapp = "1" Then
            sSql = sSql & "        '2')                         " & vbLf   '廃棄/連合区分
        '消滅区分2の時
        ElseIf disapp = "2" Then
            sSql = sSql & "        '9')                         " & vbLf
        '上記以外の時
        Else
            sSql = sSql & "        '1')                         " & vbLf
        End If
    
        '実行
        iRet = SqlExec2(sSql)
        If iRet < 0 Then
            Call MsgOut(100, sSql, ERR_DISP_LOG, "XODB3")
            Exit Function
        ElseIf iRet = 0 Then
            Call MsgOut(71, "XODB3", ERR_DISP_LOG)
            Exit Function
        End If
    
        '戻り値設定
        Ins_XODB3 = True
        Exit Function
    
    'エラー時
ErrHand:
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "XODB3")
    End Function
    
    ' @(f)
    '
    ' 機能    : 実績日チェック
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : before  :前工程コード
    '          after    :次工程コード
    '          cyoku    :直区分
    '          staffCd  :担当者コード
    '          sUserName:担当者名
    '
    ' 機能説明: 実績日が日付として正しいかチェックする
    '
    ' 備考    :
    '
    Public Function CheckDateFormat_Re(sYear As String, _
                                       sMonth As String, _
                                       sDay As String, _
                                       sHour As String, _
                                       sMinute As String) As Boolean
                                        
        ''年チェック2000年よりも小さければエラー
        If sYear < "2000" Or Len(sYear) <> 4 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''月チェック　12よりも大きければエラー
        If sMonth > "12" Or Len(sMonth) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''日チェック　31よりも大きければエラー
        If sDay > "31" Or Len(sDay) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''時チェック　23よりも大きければエラー
        If sHour > "23" Or Len(sHour) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
        
        ''分チェック　59よりも大きければエラー
        If sMinute > "59" Or Len(sMinute) <> 2 Then
            CheckDateFormat_Re = False
            Exit Function
        End If
             
        '戻り値設定
        CheckDateFormat_Re = True
    End Function
    
    ' @(f)
    ' 機能      : 実績日付作成
    '
    ' 返り値    : 実績日付(YYYYMMDDhhmmss型)
    '
    ' 引き数    : nowdate  -  日付データ(YYYYMMDDhhmmss型)
    '
    ' 機能説明  : 実績日付を作成し呼び出し側に返す
    '
    '---- UPD [mdlSP_Re(200mm用)のソース反映] 2004/10/18 TCS)R.Kawaguchi START ----
    ''Private Function GetJITUDATE(nowdate As String) As String
    ''    Dim jitudate As String     '変換用の日付を格納する変数(YYYY/MM/DD型)
    ''    Dim jitutime As String     '判定用の時刻を格納する変数
    ''
    ''    '判定用に日付、時刻を切り出す
    ''    jitudate = Left(nowdate, 10)
    ''    jitutime = Mid(nowdate, 11, 6)
    ''
    ''    '0:00から7:59の間の場合は日付を前日に設定
    ''    If jitutime >= "000000" And jitutime < "080000" Then
    ''        jitudate = Format(DateAdd("d", -1, jitudate), "YYYYMMDD")
    ''        jitudate = Replace(jitudate, "/", "")
    ''        GetJITUDATE = jitudate & jitutime
    ''    Else
    ''        GetJITUDATE = nowdate
    ''    End If
    ''End Function
    Public Function GetJITUDATE(ByVal systemdate As String) As String
    
        Dim jitudate As String     '変換用の日付を格納する変数(YYYY/MM/DD型)
        Dim jitutime As String     '判定用の時刻を格納する変数
    
        '判定用に日付、時刻を切り出す
        jitudate = left(systemdate, 4) & "/" & Mid(systemdate, 5, 2) & "/" & Mid(systemdate, 7, 2)
        jitutime = Mid(systemdate, 9, 6)
    ''    jitudate = Format(systemdate, "YYYYMMDD")
    ''    jitutime = Format(systemdate, "hhmmss")
        '0:00から7:59の間の場合は日付を前日に設定
        If jitutime >= "000000" And jitutime < "080000" Then
    ''        jitudate = jitudate - 1 'DateAdd("d", -1, jitudate)
    '        jitudate = DateAdd("d", -1, jitudate)
            jitudate = Format(DateAdd("d", -1, jitudate), "YYYYMMDD")   ''修正(2004/01/14)
            jitudate = Replace(jitudate, "/", "")
            GetJITUDATE = jitudate & jitutime
        Else
            GetJITUDATE = systemdate
        End If
    
    End Function
    '---- UPD [mdlSP_Re(200mm用)のソース反映] 2004/10/18 TCS)R.Kawaguchi END ----
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能      :   NULL文字変換
    '
    ' 返り値    :　 変換後のﾃﾞｰﾀ
    '
    ' 引数      :　 val - 変換するﾃﾞｰﾀ
    '
    ' 機能説明  :   NULL文字変換
    '
    ' 備考      :
    '
    '///////////////////////////////////////////////////
    Public Function NulltoStr(val As Variant) As String
        
        If IsNull(val) Then
            NulltoStr = ""
        Else
            NulltoStr = val
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
    
    ' @(f)
    ' 機能      : SQL数値変換関数
    '
    ' 返り値    : <入力数値> or NULL
    '
    ' 引き数    : 変換対象数値
    '
    ' 機能説明  : 渡された数値がNULLであれば"NULL"をそうでなければそのまま出力する
    Private Function Cnv2Number(vinput) As String
        If IsNull(vinput) Or vinput = "NULL" Then
            vinput = ""
        End If
        
        If vinput = "" Then
            Cnv2Number = "NULL"
        Else
            Cnv2Number = vinput
        End If
    End Function
    '2004/9/17tcs Suenaga 追加 end-------------------------------------
    
    '2004/10/15tcs Yamauchi 追加 start-------------------------------------
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能    : 四捨五入、切上、切捨を行う
    '
    ' 返り値  : 四捨五入した値
    '
    ' 引き数  : vValue          - 捨入するﾃﾞｰﾀ
    '           lExp            - 捨入する桁位置を１０の乗数で渡す
    '                             例：０．００１桁目を四捨五入するなら「－３」
    '           iIs             - 捨入する数
    '                             切上の場合は０、切捨の場合は１０
    '
    ' 機能説明: 四捨五入を行う
    '
    ' 備考    :
    '
    '///////////////////////////////////////////////////
    Public Function Round_Re(vValue As Variant, ByVal lExp As Long, Optional ByVal iIs As Integer = 5) As Double
    
        Dim lPeriod     As Long
        Dim vValueTemp  As Variant
      
        'データ初期化
        Round_Re = 0
        vValue = CDec(vValue)
        
        ' 数字の判定
        If Not IsNumeric(vValue) Then Exit Function
            
        vValueTemp = vValue
        
        ' ピリオドの位置を取得
        lPeriod = InStr(vValueTemp, Chr(46))
        If vValue < 0 Then
            lPeriod = lPeriod - 1
            vValueTemp = Mid(vValueTemp, 2)
        End If
    
        '小数点がなく、少数以下四捨五入の時は無視
        If lPeriod <= 0 And lExp < 0 Then
            Round_Re = CDbl(vValueTemp)
            Exit Function
        End If
        
        '小数点があり、少数長がlExpより短いとき無視
        If lPeriod > 0 And lExp * -1 > Len(vValueTemp) - lPeriod Then
            Round_Re = CDbl(vValueTemp)
            Exit Function
        End If
        
        'lExpが小数点以上の桁数より小さい場合
        If lExp + 2 < lPeriod Then
            If lExp < 0 Then         '捨入位置が小数点以下の場合
                '捨入位置以上を取得
                Round_Re = CDbl(left(vValueTemp, lPeriod - (lExp + 1)))
                '捨入
                If CInt(Mid(vValueTemp, lPeriod - lExp, 1)) >= iIs Then Round_Re = Round_Re + 10 ^ (lExp + 1)
            Else                         '小数点以上の場合
                '捨入位置以上を取得
                Round_Re = CDbl(left(vValueTemp, lPeriod - (lExp + 2))) * (10 ^ (lExp + 1))
                '捨入
                If CInt(Mid(vValueTemp, lPeriod - (lExp + 1), 1)) >= iIs Then Round_Re = Round_Re + 10 ^ (lExp + 1)
            End If
        'lExpが最上位桁のとき
        ElseIf lExp + 2 = lPeriod Then
            '捨入
            If CInt(left(vValueTemp, 1)) >= iIs Then
                Round_Re = 10 ^ (lExp + 1)
            Else
                Round_Re = 0
            End If
        'それ以上の時
        Else
            Exit Function
        End If
        
        If vValue < 0 Then
            Round_Re = Round_Re * (-1)
        End If
    
    End Function
    
    '
    ' @(f)
    ' 機能    : 自工場取得
    ' 返り値  : なし
    ' 引き数  : sFactryCd       - 工場ｺｰﾄﾞ
    '           sSelfFactory    - 自工場
    ' 機能説明: 工場コードより、自工場を取得する
    '
    Public Sub GetSelfFactory(ByVal sFactryCd As String, sSelfFactory As String)
    
        Select Case sFactryCd
            Case "10"               ''野田工場
                sSelfFactory = "I"
            Case "30"               ''生野工場
                sSelfFactory = "I"
            Case "40"               ''米沢工場
                sSelfFactory = "Y"
            Case "42"               ''３００ｍｍ
                sSelfFactory = "Z"
            Case "43"               ''３００ｍｍ
                sSelfFactory = "Z"
            Case "90"               ''テスト環境
                sSelfFactory = "Y"
            Case "91"               ''テスト環境(米沢) 2007/04/05追加 SETsw kubota
                sSelfFactory = "Y"
            Case "92"               ''テスト環境(生野)
                sSelfFactory = "I"
            Case "93"               ''テスト環境(生野A1) 2010/04/14 SETsw kubota
                sSelfFactory = "I"
            Case Else               ''外販
                sSelfFactory = "I"
        End Select
    
    End Sub
    
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能    : 300用　XODCX登録時、基本情報取得SQL
    '
    ' 返り値  : なし
    '
    ' 引き数  :　ssql       SQL格納用変数
    '            sCrystalNo 結晶番号
    '
    ' 機能説明:
    '
    ' 備考    :
    '
    '///////////////////////////////////////////////////
    Public Sub GetAssistSQL_300(sSql As String, sCrystalNo As String)
    
        sSql = ""
        sSql = sSql & "select                                                   " & vbLf
        sSql = sSql & "         T1.ADDOPPC1                                     " & vbLf    ''追加ドープ投入位置
        sSql = sSql & "         ,T1.PUTCUTWC1                                   " & vbLf    ''トップWT
        sSql = sSql & "         ,T1.DTYPEC1　　　　                              " & vbLf    ''ドープタイプ
        sSql = sSql & "         ,(T1.DIA1C1 + T1.DIA2C1 + T1.DIA3C1)            " & vbLf
        sSql = sSql & "         / 3 AS UPDMCX                                   " & vbLf    ''引上げAV径
        sSql = sSql & "         ,T2.HINBAN || TRIM(TO_CHAR(T2.NMNOREVNO,'00'))  " & vbLf
        sSql = sSql & "         || T2.NFACTORY || T2.NOPECOND AS HINBCX         " & vbLf    ''品番
        
    '*** UPDATE START T.TERAUCHI 2004/11/27 結晶ドープ、結晶ドープ量変更対応
    '    sSql = sSql & "         ,T2.DPNTCLS                                     " & vbLf    ''結晶ドープ
    '    sSql = sSql & "         ,T2.DOPANT                                      " & vbLf    ''結晶ドープ量
        sSql = sSql & "         ,T6.CRYDOP DPNTCLS                               " & vbLf    ''結晶ドープ
        sSql = sSql & "         ,T6.CRYDOPVL DOPANT                              " & vbLf    ''結晶ドープ量
    '*** UPDATE END   T.TERAUCHI 2004/11/27
        
        sSql = sSql & "         ,T2.PGID                                        " & vbLf    ''PGID
        sSql = sSql & "         ,T3.HSXTYPE                                     " & vbLf    ''タイプ
        sSql = sSql & "         ,(T3.HSXD1MIN + T3.HSXD1MAX) / 2 AS PRODMCX     " & vbLf    ''製品径
        sSql = sSql & "         ,T1.SUICHARGE                                  " & vbLf     ''推定チャージ量
        sSql = sSql & "         ,T4.HSXLTHWS                                    " & vbLf    ''ライフタイム仕様有無
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "         ,T5.CTR01A9 * 1000 AS CTR01A9" & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
    
    '*** UPDATE START T.TERAUCHI 2004/12/08
        sSql = sSql & "         ,T1.LENTKC1 AS LENTKC1" & vbLf                              ''直胴長さ
    '*** UPDATE END   T.TERAUCHI 2004/12/08
    
    '*** UPDATE START Marushita 2011/03/23 TSMC品識別対応
        sSql = sSql & "         ,T7.MTRLCHKFLG AS MTRLCHKFLG" & vbLf                        ''精製原料チェックフラグ(規格)
    '*** UPDATE END   Marushita 2011/03/23
    
        sSql = sSql & "from     XSDC1 T1                                        " & vbLf
        sSql = sSql & "         ,TBCMH001 T2                                    " & vbLf
        sSql = sSql & "         ,TBCME018 T3                                    " & vbLf
        sSql = sSql & "         ,TBCME019 T4                                    " & vbLf
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "             ,KODA9 T5 " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
    '*** UPDATE START T.TERAUCHI 2004/11/27 結晶ドープ、結晶ドープ量変更対応
        sSql = sSql & "             ,TBCMH002 T6 "
    '*** UPDATE END   T.TERAUCHI 2004/11/27
    '*** UPDATE START Marushita 2011/03/23 TSMC品識別対応
        sSql = sSql & "         ,TBCME036 T7 " & vbLf
    '*** UPDATE END   Marushita 2011/03/23
    
    '*** UPDATE START T.TERAUCHI 2004/10/21
    '    sSql = sSql & "where    SUBSTRB(T1.XTALC1,1,9)                          " & vbLf
    '    sSql = sSql & "         = SUBSTRB('" & sCrystalNo & "',1,9)             " & vbLf
        sSql = sSql & "where    T1.XTALC1                          " & vbLf
        sSql = sSql & "         = '" & sCrystalNo & "'             " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
        
        sSql = sSql & "and      SUBSTRB(T1.XTALC1,1,7) || '0' || SUBSTRB(T1.XTALC1,9,1)     " & vbLf
        sSql = sSql & "         = SUBSTRB(T2.UPINDNO,1,7) || '0' || SUBSTRB(T2.UPINDNO,9,1) " & vbLf
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T3.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T3.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T3.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T3.OPECOND                                    " & vbLf
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T4.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T4.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T4.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T4.OPECOND                                    " & vbLf
        
    '*** UPDATE START T.TERAUCHI 2004/10/21
        sSql = sSql & "AND          T5.SYSCA9 = 'K' " & vbLf
        sSql = sSql & "AND          T5.SHUCA9 = 'A7' " & vbLf
        sSql = sSql & "AND          T5.CODEA9 = '300'" & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/10/21
    
    '*** UPDATE START T.TERAUCHI 2004/11/27
        sSql = sSql & "and      SUBSTRB(T1.XTALC1,1,7) || '0' || SUBSTRB(T1.XTALC1,9,1)     " & vbLf
        sSql = sSql & "         = SUBSTRB(T6.UPINDNO,1,7) || '0' || SUBSTRB(T6.UPINDNO,9,1) " & vbLf
    '*** UPDATE END   T.TERAUCHI 2004/11/27
    '*** UPDATE START Marushita 2011/03/23 TSMC品識別対応
        sSql = sSql & "and      T2.HINBAN                                       " & vbLf
        sSql = sSql & "         = T7.HINBAN                                     " & vbLf
        sSql = sSql & "and      T2.NMNOREVNO                                    " & vbLf
        sSql = sSql & "         = T7.MNOREVNO                                   " & vbLf
        sSql = sSql & "and      T2.NFACTORY                                     " & vbLf
        sSql = sSql & "         = T7.FACTORY                                    " & vbLf
        sSql = sSql & "and      T2.NOPECOND                                     " & vbLf
        sSql = sSql & "         = T7.OPECOND                                    " & vbLf
    '*** UPDATE END   Marushita 2011/03/23 TSMC品識別対応
    
    End Sub
    '2004/10/15tcs Yamauchi 追加 end-------------------------------------
    
    '---- ADD [精製原料管理、原料工程実績作成対応] 2004/10/29 TCS)R.Kawaguchi START ----
    '///////////////////////////////////////////////////
    ' @(f)
    '
    ' 機能    : 肩重量取得
    '
    ' 返り値  : 成否
    '
    ' 引き数  : sDiameterKbn    - 直径区分
    '           dKataWeight     - 肩重量
    '
    ' 機能説明: 肩重量取得
    '
    ' 備考    :
    '
    '///////////////////////////////////////////////////
    Public Function GetKataWeight(ByVal sDiameterKbn As String, dKataWeight As Double) As Boolean
    
        Dim sSql        As String
        Dim objOraDyn   As Object
    
    On Error GoTo ErrHand
    
        GetKataWeight = False
    
        sSql = ""
        sSql = sSql & "select   CTR01A9" & vbLf
        sSql = sSql & "from     KODA9" & vbLf
        sSql = sSql & "where    SYSCA9 = 'K'" & vbLf
        sSql = sSql & "and      SHUCA9 = 'A7'" & vbLf
        sSql = sSql & "and      CODEA9 = '" & sDiameterKbn & "'" & vbLf
    
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = False Then
            ''取得失敗
            Call MsgOut(100, sSql, ERR_DISP_LOG, "KODA9")
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        ''ﾃﾞｰﾀなし
        If objOraDyn.EOF Then
            Call MsgOut(55, "管理ｺｰﾄﾞﾃｰﾌﾞﾙ", ERR_DISP)
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        dKataWeight = objOraDyn.Fields("CTR01A9").Value
        
        '開放
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
        
        GetKataWeight = True
        Exit Function
    
ErrHand:
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
        '開放
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
    
    End Function
    '---- ADD [精製原料管理、原料工程実績作成対応] 2004/10/29 TCS)R.Kawaguchi END ----
    
'>>>>>>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -----------------START
'>>>>>>>>>> Ins_TBCMC001_New関数をs_cmzclabel.basに移動のためコメント化 -----
'
'    '*** UPDATE START T.TERAUCHI 2004/12/06 精製原料候補発生時、精製原料ラベル発行用ﾓｼﾞｭｰﾙ追加
'    ' @(f)
'    '
'    ' 機能    : ラベル発行用ﾓｼﾞｭｰﾙ
'    '
'    ' 返り値  : True:成功 False:失敗
'    '
'    ' 引き数  : after:次工程コード
'    '           cyoku:直区分
'    '
'    ' 機能説明:　精製原料候補発生時、精製原料ラベルを発行する
'    '
'    ' 備考    :
'    '       引数：
'    '           sProcCode   工程ｺｰﾄﾞ
'    '           sEtcPrKind  その他ﾗﾍﾞﾙ種類
'    '           sStaffID    要求担当者
'    '           sPrKey01    帳票ｷｰﾃﾞｰﾀ1
'    '           sSysdate    ｷｭｰ日付
'    '           sRegDate    登録日付　※登録日付はPKの為、一回の処理で複数件登録する場合、
'    '                                   呼び出し元で1秒ずらす等の制御が必要
'    '       使用ﾌﾟﾛｸﾞﾗﾑ：
'    '           cmbc008     ｸﾘｽﾀﾙｶﾀﾛｸﾞ検索格上げ
'    '           cmbc030     結晶総合判定
'    '           cmbc018     切断・ｻﾝﾌﾟﾙ指示照会
'    '
'    Public Function Ins_TBCMC001_New(sProcCode As String, sEtcPrKind As String, sStaffID As String, sPrKey01 As String, sSysDate As String) As Boolean
'        Dim sSql      As String       'SQL文格納
'        Dim iRet    As Integer        'データ追加数
'
'    'エラーハンドラ
'    On Error GoTo ErrHand
'
'        '戻り値設定
'        Ins_TBCMC001_New = False
'
'        'ｺﾝﾋﾟｭｰﾀ名設定
'        gsCompName = GetCompName
'
'        '登録用ｸｴﾘｰ設定
'        sSql = ""
'        sSql = sSql & "insert into tbcmc001(" & vbLf    ''
'        sSql = sSql & "                 quedate                     " & vbLf    ''キュー日付
'        sSql = sSql & "                 ,reqkind                    " & vbLf    ''印刷要求区分
'        sSql = sSql & "                 ,printkind                  " & vbLf    ''印刷種類
'        sSql = sSql & "                 ,endflg                     " & vbLf    ''完了区分
'        sSql = sSql & "                 ,status                     " & vbLf    ''終了ステータス
'        sSql = sSql & "                 ,blockidumu                 " & vbLf    ''ブロックID有無区分
'        sSql = sSql & "                 ,proccode                   " & vbLf    ''工程コード
'        sSql = sSql & "                 ,etcprkind                  " & vbLf    ''その他ラベル種類
'        sSql = sSql & "                 ,crynum                     " & vbLf    ''結晶番号
'        sSql = sSql & "                 ,ingotpos                   " & vbLf    ''結晶内位置
'        sSql = sSql & "                 ,smplno                     " & vbLf    ''サンプルNo
'        sSql = sSql & "                 ,mtrlnum                    " & vbLf    ''原料番号
'        sSql = sSql & "                 ,smtrlnum                   " & vbLf    ''精製原料番号
'        sSql = sSql & "                 ,blockid                    " & vbLf    ''ブロックID
'        sSql = sSql & "                 ,hinban                     " & vbLf    ''品番
'        sSql = sSql & "                 ,revnum                     " & vbLf    ''製品番号改定番号
'        sSql = sSql & "                 ,factory                    " & vbLf    ''工場
'        sSql = sSql & "                 ,opecond                    " & vbLf    ''操業条件
'        sSql = sSql & "                 ,cryindrs                   " & vbLf    ''結晶検査指示(Rs)
'        sSql = sSql & "                 ,cryindoi                   " & vbLf    ''結晶検査指示(Oi)
'        sSql = sSql & "                 ,cryindb1                   " & vbLf    ''結晶検査指示(B1)
'        sSql = sSql & "                 ,cryindb2                   " & vbLf    ''結晶検査指示(B2)
'        sSql = sSql & "                 ,cryindb3                   " & vbLf    ''結晶検査指示(B3)
'        sSql = sSql & "                 ,cryindl1                   " & vbLf    ''結晶検査指示(L1)
'        sSql = sSql & "                 ,cryindl2                   " & vbLf    ''結晶検査指示(L2)
'        sSql = sSql & "                 ,cryindl3                   " & vbLf    ''結晶検査指示(L3)
'        sSql = sSql & "                 ,cryindl4                   " & vbLf    ''結晶検査指示(L4)
'        sSql = sSql & "                 ,cryindcs                   " & vbLf    ''結晶検査指示(Cs)
'        sSql = sSql & "                 ,cryindgd                   " & vbLf    ''結晶検査指示(Gd)
'        sSql = sSql & "                 ,cryindt                    " & vbLf    ''結晶検査指示(T)
'        sSql = sSql & "                 ,cryindep                   " & vbLf    ''結晶検査指示(EPD)
'        sSql = sSql & "                 ,staffid                    " & vbLf    ''要求担当者
'        sSql = sSql & "                 ,machine                    " & vbLf    ''要求マシン名
'        sSql = sSql & "                 ,regdate                    " & vbLf    ''登録日付
'        sSql = sSql & "                 ,upddate                    " & vbLf    ''更新日付
'        sSql = sSql & "                 ,prkey01                     " & vbLf    ''帳票キーデータ１
'        sSql = sSql & "     )                                       " & vbLf
'        sSql = sSql & "values(                                      " & vbLf
'        sSql = sSql & "                 to_date('" & sSysDate & "','yyyy/mm/dd hh24:mi:ss')       " & vbLf    ''キュー日付
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''印刷要求区分
'        sSql = sSql & "                 ,'1'                        " & vbLf    ''印刷種類
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''完了区分
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''終了ステータス
'        sSql = sSql & "                 ,'0'                        " & vbLf    ''ブロックID有無区分
'        sSql = sSql & "                 ,'" & sProcCode & "'        " & vbLf    ''工程コード
'        sSql = sSql & "                 ,'" & sEtcPrKind & "'       " & vbLf    ''その他ラベル種類
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶番号
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶内位置
'        sSql = sSql & "                 ,null                       " & vbLf    ''サンプルNo
'        sSql = sSql & "                 ,null                       " & vbLf    ''原料番号
'        sSql = sSql & "                 ,null                       " & vbLf    ''精製原料番号
'        sSql = sSql & "                 ,null                       " & vbLf    ''ブロックID
'        sSql = sSql & "                 ,null                       " & vbLf    ''品番
'        sSql = sSql & "                 ,null                       " & vbLf    ''製品番号改定番号
'        sSql = sSql & "                 ,null                       " & vbLf    ''工場
'        sSql = sSql & "                 ,null                       " & vbLf    ''操業条件
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Rs)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Oi)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B1)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B2)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B3)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L1)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L2)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L3)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L4)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Cs)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Gd)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(T)
'        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Epd)
'        sSql = sSql & "                 ,'" & sStaffID & "'         " & vbLf    ''要求担当者名
'        sSql = sSql & "                 ,'" & gsCompName & "'       " & vbLf    ''
'        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''登録日付
'        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''更新日付
'        sSql = sSql & "                 ,'" & sPrKey01 & "'         " & vbLf    ''帳票キーデータ１
'        sSql = sSql & "                             )               " & vbLf
'
'        '実行
'        iRet = SqlExec2(sSql)
'
'        If iRet < 0 Then
'            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
'            Exit Function
'        End If
'
'        '戻り値設定
'        Ins_TBCMC001_New = True
'
'    'エラー時
'ErrHand:
'        ''ｴﾗｰ
'        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMC001")
'    End Function
'    '*** UPDATE END T.TERAUCHI 2004/12/06 精製原料候補発生時、精製原料ラベル発行用ﾓｼﾞｭｰﾙ追加
''*ADD* ﾓｼﾞｭｰﾙ統一 TCS)K.Kunori 2004.11.29 END <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -----------------END

'概要      :Cs推定計算
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:tCsSuitei    ,I  ,CS_SUITEI_TYPE   ,Cs推定計算用ﾊﾟﾗﾒｰﾀ
'      　　:dCsSuitei    ,O  ,Double           ,Cs推定値
'      　　:戻り値       ,O   ,Boolean         ,計算の成否
'説明      :Cs推定値を計算する
'履歴      :06/04/20 ooba
Public Function GetCsSuiteiMain(tCsSuitei As CS_SUITEI_TYPE, dCsSuitei As Double) As Boolean
    Dim vGS         As Variant
    Dim vCSSC0      As Variant
    Dim vCS0        As Variant
    Dim vGT         As Variant
    Dim vCSTC0      As Variant
    
    Dim dSiWeight   As String   ''ﾁｬｰｼﾞ量(Kg)
    Dim dTopWT      As String   ''ﾄｯﾌﾟ重量(Kg)
    Dim dUpDm       As String   ''直径(mm)
    Dim dCsHenseki  As String   ''ｶｰﾎﾞﾝ偏析係数(ｺｰﾄﾞﾏｽﾀに保持)
    Dim dSamplePos  As String   ''ｻﾝﾌﾟﾙ位置
    Dim dResCs      As String   ''ｻﾝﾌﾟﾙ測定値
    Dim dInfPos     As String   ''推定位置
    
On Error GoTo ErrHand

    GetCsSuiteiMain = False

    ''変数格納
    dSiWeight = CDbl(tCsSuitei.sSiWeight) / 1000    ''ﾁｬｰｼﾞ量(Kg)
    dTopWT = CDbl(tCsSuitei.sTopWT) / 1000          ''ﾄｯﾌﾟ重量(Kg)
    dUpDm = CDbl(tCsSuitei.sUpDm)                   ''直径(mm)
    dCsHenseki = CDbl(tCsSuitei.sCsHenseki)         ''ｶｰﾎﾞﾝ偏析係数(ｺｰﾄﾞﾏｽﾀに保持)
    dSamplePos = CDbl(tCsSuitei.sSamplePos)         ''ｻﾝﾌﾟﾙ位置
    dResCs = CDbl(tCsSuitei.sResCs)                 ''ｻﾝﾌﾟﾙ測定値
    dInfPos = CDbl(tCsSuitei.sInfPos)               ''推定位置

    ''GS = (直径 / 20) ^ 2 * 3.14 * 2.33 * ｻﾝﾌﾟﾙ位置 / (ﾁｬｰｼﾞ量 - TOP重量) / 1000
    vGS = (dUpDm / 20) ^ 2 * 3.14 * 2.33 * dSamplePos / (dSiWeight - dTopWT) / 10000

    ''CSSC0 = ｶｰﾎﾞﾝ偏析係数 * (1 - GS) ^ (ｶｰﾎﾞﾝ偏析係数 - 1)
    vCSSC0 = dCsHenseki * (1 - vGS) ^ (dCsHenseki - 1)

    ''CS0 = ｻﾝﾌﾟﾙ測定値 / CSSC0
    vCS0 = dResCs / vCSSC0
    
    ''GT = (直径 / 20) ^ 2 * 3.14 * 2.33 * 推定位置 / (ﾁｬｰｼﾞ量 - TOP重量) / 1000
    vGT = (dUpDm / 20) ^ 2 * 3.14 * 2.33 * dInfPos / (dSiWeight - dTopWT) / 10000
    
    ''CSTC0 = ｶｰﾎﾞﾝ偏析係数 * (1 - GT) ^ (ｶｰﾎﾞﾝ偏析係数 - 1)
    vCSTC0 = dCsHenseki * (1 - vGT) ^ (dCsHenseki - 1)
    
    ''推定値 = CS0 * CSTC0
    dCsSuitei = vCS0 * vCSTC0
    
    GetCsSuiteiMain = True

    Exit Function
    
ErrHand:

    ''ｴﾗｰ
'    Call MsgOut(100, "", ERR_DISP_LOG, "")

End Function

'関数名    :権限コード獲得
'概要      :社員ごとに登録されている権限（１０ユニット）から
'           画面ごとに設定されたている権限ユニットコード(1～10迄の値)に
'           該当する権限コードを抜き出して戻り値に設定する。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :staffID       ,I  ,String    ,社員ID
'          :picname       ,I  ,String    ,画面名
'          :戻り値        ,O  ,String    ,権限コード（ 1 : 参照○ 更新× 承認×
'                                                      2 : 参照○ 更新○ 承認×
'                                                      3 : 参照○ 更新× 承認○
'                                                      4 : 参照○ 更新○ 承認○ ）
'説明      :見つからなかった場合は、戻り値に０を返す
'
'
'変更履歴   2009/09 SUMCO Akizuki 権限設定の改修

Public Function Getstaffauthority(STAFFID$, picname) As Integer
Dim dbIsMine As Boolean         'ＤＢオープンフラグ
Dim rs1 As OraDynaset           'KODA9(画面マスタ)用ダイナセット
Dim rs2 As OraDynaset           'KODA9(社員マスタ)用ダイナセット
Dim sql As String               'ＳＱＬ文格納領域
Dim picauthority As Integer     '権限ユニットコード

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function Getstaffauthority"
    
    Getstaffauthority = 0

    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    'KODA9(画面マスタ)より、当該画面の権限ユニットコード獲得
    sql = ""
    sql = "select KCODE01A9 from KODA9 where SYSCA9='K' and SHUCA9='01' and CODEA9='" & picname & "'"
    Set rs1 = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs1("KCODE01A9")) Then
        '権限ユニットコードがＮＵＬＬまたはスペースの場合、
        '(=つまり、権限チェックの設定を掛けていない場合)　チェックを行わず全操作可能を返す
        Getstaffauthority = 4
    ElseIf Trim(rs1("KCODE01A9")) = "" Then
        Getstaffauthority = 4
    Else
        '獲得した権限ユニットコード
        picauthority = CInt(rs1("KCODE01A9"))
        
        'KODA9(社員マスタ)より、当該社員の権限コード獲得
        sql = ""
        sql = "select KCODE03A9 from KODA9 where SYSCA9='K' and SHUCA9='55' and CODEA9='" & STAFFID & "'"
        Set rs2 = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
        If rs2.RecordCount = 0 Then
            ''該当する社員の情報が見つからなかった場合、０を返す
            Getstaffauthority = 0
            
        '>>>2009/09/07 SUMCO Akizuki
        ElseIf rs2.RecordCount = 1 Then
            ''権限コードがＮＵＬＬの場合、権限0[操作不可]に
            If IsNull(rs2("KCODE03A9")) Then
                'Getstaffauthority = 4
                Getstaffauthority = 0
            
            ''権限コードがスペースの場合も権限0[操作不可]に
            ElseIf Trim(rs2("KCODE03A9")) = "" Then
                'Getstaffauthority = 4
                Getstaffauthority = 0
            
            ''権限コードがユニットコードより小さい = ｢設定なし｣の場合も、権限0[操作不可]に
            ElseIf (picauthority <= 0) Or (picauthority > Len(Trim(rs2("KCODE03A9")))) Then
                Getstaffauthority = 0
        '<<<
        
            Else
            '権限コードが０（初期値）の場合、権限チェックを行わず権限0[操作不可]に
            '2007/08/09 kaga
                'If CInt(Left(rs2("KCODE03A9"), picauthority)) = 0 Then
                If CInt(Mid(Trim(rs2("KCODE03A9")), picauthority, 1)) = 0 Then
                    Getstaffauthority = 0
                '権限コードが０（初期値）の以外場合、当該画面に対する権限コードを返す
                Else
                    '2007/08/09 kaga
                    'Getstaffauthority = CInt(Left(rs2("KCODE03A9"), picauthority))
                    Getstaffauthority = CInt(Mid(Trim(rs2("KCODE03A9")), picauthority, 1))
                End If
            End If
        End If
    End If
    
    If dbIsMine Then
        OraDBClose
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

PROC_ERR:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

    
'--------------- 2008/08/25 INSERT START  By Systech ---------------
'概要      :テーブル「TBCME036」から条件にあったレコードのDK温度を抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :bFlg          ,I  ,Boolean         ,TRUE=仕様, FALSE=実績
'          :xsdcs         ,I  ,typ_XSDCS       ,新サンプル管理
'          :戻り値        ,O  ,String          ,DK温度
'説明      :
'履歴      :
Public Function GetDKTmpCode(bFlg As Boolean, xsdcs As typ_XSDCS) As String
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    
    GetDKTmpCode = ""

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function GetDKTmpName"
    
    ''SQLを組み立てる
    If bFlg Or xsdcs.CRYINDRSCS = "1" Then
        ' DK温度(仕様)又は状態FLG(Rs)が通常の場合、当該品番のDK温度を取得
        sql = "SELECT NVL(HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036"
        sql = sql & " WHERE HINBAN = '" & xsdcs.HINBCS & "'"
        sql = sql & " AND MNOREVNO = " & xsdcs.REVNUMCS
        sql = sql & " AND FACTORY = '" & xsdcs.FACTORYCS & "'"
        sql = sql & " AND OPECOND = '" & xsdcs.OPECS & "'"
    Else
        ' DK温度(実績)の場合、反映元品番のDK温度を取得
        sql = "SELECT NVL(A.HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036 A, XSDCS B"
        sql = sql & " WHERE B.XTALCS = '" & xsdcs.XTALCS & "'"
        sql = sql & " AND B.CRYSMPLIDRSCS = '" & xsdcs.CRYSMPLIDRSCS & "'"
        sql = sql & " AND B.CRYINDRSCS = '1'"
        sql = sql & " AND A.HINBAN = B.HINBCS"
        sql = sql & " AND A.MNOREVNO = B.REVNUMCS"
        sql = sql & " AND A.FACTORY = B.FACTORYCS"
        sql = sql & " AND A.OPECOND = B.OPECS"
    End If
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        ' DK温度を返す
        GetDKTmpCode = rs("HSXDKTMP")
    End If
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

'概要      :テーブル「TBCME036」から条件にあったレコードのDK温度を抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :bFlg          ,I  ,Boolean         ,TRUE=仕様, FALSE=実績
'          :xsdcw         ,I  ,typ_XSDCW       ,新サンプル管理
'          :戻り値        ,O  ,String          ,DK温度
'説明      :
'履歴      :
Public Function GetWfDKTmpCode(bFlg As Boolean, xsdcw As typ_XSDCW) As String
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    
    GetWfDKTmpCode = ""

    'エラーハンドラの設定
    On Error GoTo PROC_ERR
    gErr.Push "mdlCommon.bas -- Function GetWfDKTmpCode"
    
    ''SQLを組み立てる
    If bFlg Or xsdcw.WFINDRSCW = "1" Then
        ' DK温度(仕様)又は状態FLG(Rs)が通常の場合、当該品番のDK温度を取得
        sql = "SELECT NVL(HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036"
        sql = sql & " WHERE HINBAN = '" & xsdcw.HINBCW & "'"
        sql = sql & " AND MNOREVNO = " & xsdcw.REVNUMCW
        sql = sql & " AND FACTORY = '" & xsdcw.FACTORYCW & "'"
        sql = sql & " AND OPECOND = '" & xsdcw.OPECW & "'"
    Else
        ' DK温度(実績)の場合、反映元品番のDK温度を取得
        sql = "SELECT NVL(A.HSXDKTMP, ' ') AS HSXDKTMP"
        sql = sql & " FROM TBCME036 A, XSDCW B"
        sql = sql & " WHERE B.XTALCW = '" & xsdcw.XTALCW & "'"
        sql = sql & " AND B.WFSMPLIDRSCW = '" & xsdcw.WFSMPLIDRSCW & "'"
        sql = sql & " AND B.WFINDRSCW = '1'"
        sql = sql & " AND A.HINBAN = B.HINBCW"
        sql = sql & " AND A.MNOREVNO = B.REVNUMCW"
        sql = sql & " AND A.FACTORY = B.FACTORYCW"
        sql = sql & " AND A.OPECOND = B.OPECW"
    End If
    
    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        ' DK温度を返す
        GetWfDKTmpCode = rs("HSXDKTMP")
    End If
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

'概要      :DK温度名称より、先頭の数値("℃"以前)を返す
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :szBuf         ,I  ,String       , DK温度名称(汎用コードマスタのコード内容)
'          :戻り値        ,O  ,String       , DK温度名称先頭の数値
'説明      :
'履歴      :
Public Function GetDKTmpDispName(szBuf As String) As String
    Dim i       As Integer
    
    i = InStr(1, szBuf, "℃") - 1
    If i > 0 And Len(szBuf) >= i Then
        GetDKTmpDispName = left(szBuf, i)
    Else
        GetDKTmpDispName = szBuf
    End If

End Function
'--------------- 2008/08/25 INSERT  END   By Systech ---------------

'概要      :溝・Notch位置方位を表示用に変換
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :sDPDIR        ,I  ,String       ,溝・Notch位置方位
'          :戻り値        ,O  ,String       ,表示文字列
'説明      :溝・Notch位置方位を表示用に変換
'2009/05/22 SETsw kubota
Public Function CnvMizoNotchDisp(ByVal sDPDIR As String) As String

    CnvMizoNotchDisp = ""
    
    Select Case sDPDIR
    Case "B1", "B2", "B3", "B4", "D3", "D4"
        CnvMizoNotchDisp = "0"
    Case "B5", "B6", "B7", "B8", "D1", "D2"
        CnvMizoNotchDisp = "45"
    End Select

End Function

'///////////////////////////////////////////////////
' @(f)
'
' 機能      : メッセージ編集・画面表示・ログ出力
'
' 返り値    : なし
'
' 引き数    : arg1:ﾒｯｾｰｼﾞｺｰﾄﾞ 100:ｵﾗｸﾙｴﾗｰ 100以外:ｵﾗｸﾙｴﾗｰ以外
'             arg2:セッション
'             arg3:DB
'             arg4:追加ﾒｯｾｰｼﾞ
'             arg5:ﾒｯｾｰｼﾞ属性　0:通常ﾒｯｾｰｼﾞ
'                              1:画面表示ｴﾗｰﾒｯｾｰｼﾞ（入力欄赤表示のｴﾗｰなど）
'                              2:ログ出力ｴﾗｰﾒｯｾｰｼﾞ
'                              3:画面表示・ログ出力ｴﾗｰﾒｯｾｰｼﾞ（ｵﾗｸﾙｴﾗｰなど）
'                              5:画面表示ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞ
'                              6:ログ出力ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞ
'             arg6:ｵﾗｸﾙｴﾗｰ時のﾛｸﾞ/画面表示ﾃｰﾌﾞﾙ名
'
' 機能説明  : MsgOutのDB接続Object指定版
' 備考      :
'2009/05/28追加 SETsw kubota
'///////////////////////////////////////////////////
Public Sub MsgOut_DB(ByVal iMsgCd As Integer _
                   , ByRef objSess As Object _
                   , ByRef objDB As Object _
                   , Optional ByVal sAddMsgStr As String = "" _
                   , Optional ByVal eMsgKind As MsgKind = 0 _
                   , Optional ByVal TABLENAME As String = "Unknown" _
                   )
    Dim sMsg As String                              ''メッセージ
    Dim sOraErrCd As String                         ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞ
    
    'メッセージ初期化
    Call MsgInit

    'ﾒｯｾｰｼﾞ属性がﾒｯｾｰｼﾞ出力属性範囲外の場合出力しない（開発運用開始後、ﾃﾞﾊﾞｯｸﾞﾒｯｾｰｼﾞを出力しないようにできる）
    If Not ((eMsgKind = NORMAL_MSG) Or _
            ((eMsgKind And MsgKindMask) <> 0)) Then
        Exit Sub                                    ''終了
    End If
    
    If iMsgCd < 100 Then                            ''ﾒｯｾｰｼﾞｺｰﾄﾞがｵﾗｸﾙ以外なら
        ''オラクル以外のメッセージ
        On Error Resume Next                        ''ｴﾗｰﾄﾗｯﾌﾟ
        sMsg = msMsgStr(iMsgCd)                     ''メッセージ取得
        On Error GoTo 0                             ''ｴﾗｰﾄﾗｯﾌﾟ解除
    Else                                            ''ﾒｯｾｰｼﾞｺｰﾄﾞがｵﾗｸﾙｴﾗｰなら
        ''オラクルのエラーメッセージ
        If objSess.LastServerErr Then           ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄのエラーならば
            sMsg = objSess.LastServerErrText    ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄｴﾗｰﾒｯｾｰｼﾞをセット
            objSess.LastServerErrReset          ''ｵﾗｸﾙｾｯｼｮﾝｵﾌﾞｼﾞｪｸﾄｴﾗｰをリセット
        ElseIf objDB.LastServerErr Then         ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄのエラーならば
            ''ｵﾗｸﾙｴﾗｰﾒｯｾｰｼﾞよりｵﾗｸﾙｴﾗｰｺｰﾄﾞを切出す
            sOraErrCd = GetStrOraErrCd(objDB.LastServerErrText)
            If sOraErrCd <> "" Then                 ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞが入っていれば
                sMsg = "DBエラー（" & TABLENAME & ")" & sOraErrCd ''指定のフォーマットで編集
                sAddMsgStr = objDB.LastServerErrText & _
                             "::" & sAddMsgStr
            Else                                    ''ｵﾗｸﾙｴﾗｰｺｰﾄﾞが入っていなければ
                sMsg = objDB.LastServerErrText  ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄｴﾗｰﾒｯｾｰｼﾞをセット
            End If
            objDB.LastServerErrReset            ''ｵﾗｸﾙﾃﾞｰﾀﾍﾞｰｽｵﾌﾞｼﾞｪｸﾄｴﾗｰをリセット
        ElseIf Err.Number Then                      ''実はVBのｴﾗｰだったなら
            sMsg = Error(Err.Number)                ''VBのｴﾗｰﾒｯｾｰｼﾞをセット
        Else                                        ''実はｴﾗｰじゃないならば
            sMsg = "ｵﾗｸﾙ正常時にｴﾗｰ出力した"         ''警告
        End If
    End If
    
    If (eMsgKind = NORMAL_MSG) Or _
       (eMsgKind And ERR_DISP) Then                     ''通常ﾒｯｾｰｼﾞか画面表示ビットが立っていれば
        ''エラーなら赤表示
        If (eMsgKind = ERR_DISP) Or _
           (eMsgKind = ERR_DISP_LOG) Then
            If iMsgCd = 100 Then                        ''オラクルエラーの場合
                MsgDisp sMsg, vbRed                     ''メッセージを画面表示する
            Else
                MsgDisp sMsg & sAddMsgStr, vbRed        ''メッセージ & 追加メッセージを画面表示する
            End If
        ''それ以外は黒表示
        Else
            If iMsgCd = 100 Then                        ''オラクルエラーの場合
                MsgDisp sMsg                            ''メッセージを画面表示する
            Else
                MsgDisp sMsg & sAddMsgStr               ''メッセージ & 追加メッセージを画面表示する
            End If
        End If
    End If
    
    If eMsgKind And ERR_LOG Then                    ''ログ出力ビットが立っていれば
        MsgLog (Format(Now, "YYYY/MM/DD HH:NN:SS::") & App.EXENAME & "::" & _
            iMsgCd & "::" & sMsg & "::" & sAddMsgStr) ''メッセージをログ出力する
    End If
    
    If (eMsgKind = ERR_DISP) Or _
       (eMsgKind = ERR_LOG) Or _
       (eMsgKind = ERR_DISP_LOG) Then                       ''ﾒｯｾｰｼﾞ属性がエラーなら
        Beep
    End If
End Sub

'概要      :大画面用ﾒﾆｭｰ遷移ボタン処理
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :iIndex        ,I  ,Integer      ,1:ﾒｲﾝﾒﾆｭｰ遷移 2:ｻﾌﾞﾒﾆｭｰ遷移
'説明      :
'2009/08/12 SETsw kubota
Public Sub execSubClick(ByVal iIndex As Integer)
    
    Select Case iIndex
    Case 1      'ﾒｲﾝﾒﾆｭｰ
        GotoMainMenu
    Case 2      'ｻﾌﾞﾒﾆｭｰ
        GotoSubMenu
    End Select

End Sub

