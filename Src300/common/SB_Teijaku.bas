Attribute VB_Name = "SB_Teijaku"
'-----------------------------------------------------------------------------------------------------------------
'       定尺カット対応共通関数モジュール
'
'
'                                                   作成日      20003.09.10
'                                                   変更日      20003.XX.XX
'
'
'
'   作成者         システムブレイン
'
'
'
'-----------------------------------------------------------------------------------------------------------------

Option Explicit
Option Base 1

Public Enum EnumCutFlag
    DoCut = 1
    NoCut = 0
End Enum


Public Type typ_CutBlkHinban
    INPOS As Integer
    Cut As EnumCutFlag
    hinban As tFullHinban
    LENGTH As Integer
End Type


Public Type typ_ChangeHinb
    INPOS As Integer
    Cut As EnumCutFlag
    hinban As tFullHinban
    LENGTH As Integer
End Type


'====================================================================================================================
' ・指定された結晶の情報から(切断状況)、定尺カットを行なった際のブロック切断情報を呼び出し元に返す。                        *
' ・品番の仕様制約等の理由により、定尺カットが不可能な場合、定尺カット不可を返す。                                         *
' ・製品の切断領域区分により、「単独部位として切断」の場合、定尺カット可能な連続領域のみ対象とし、                          *
' ・単独部位は変更なしとする。                                                                                         *
' ・最下位部の長さが100mm未満の場合、100mm以上になるように調整を行なう。                                                 *
'====================================================================================================================
'  参照テーブル                         TBCME036                                                                    *
'  項目名　　　ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ        BLOCKHFLAG  0: 指定された、狙い品番の仕様値を取得する。                         *
'                                                   1: 指定されている全品番の仕様値を取得する。（配列作成）             *
'                                                                                                                   *
'   戻り値                                                                                                          *
'    正常終了               0       （なし）                                                                        *
'    正常終了                                                                                                      *
'   (定尺不可)              TJ001 ﾌﾞﾛｯｸｶｯﾄ品番がある為、定尺ｶｯﾄできません。                                         *
'   　                      TJ002 最下位部長さを取得できない為､定尺カットできません｡                                *
'                           TJ003 引上げ長 < 最下位部長さの為､定尺カットできません｡                                 *
'                           TJ004 Z / G品番がTOP / BOT以外の為､定尺カットできません｡                                *
'                           TJ005 ※最下位部チェックを行う。最下部以下のカット長がある場合はエラー                      *
'                                                                                                                   *
'-------------------------------------------------------------------------------------------------------------------*
'                           0                   正常終了                                                            *
'                           1                   正常終了 (定尺不可)                                                 *
'                           -1                  異常終了                                                            *
'===================================================================================================================*

Public Function funGetFixLengCut(ByVal sProccd As String, ByVal sCryNo As String, sTgetHinban As tFullHinban, _
                                 ByVal iFixCutLeng As Integer, ByVal iAllLeng As Integer, ByVal iSprFlg As Integer, _
                                 ByRef tCutBlkHinban() As typ_CutBlkHinban, ByVal iErr_Code As Integer, sErr_Msg As String) As Integer
' 1   sProccd                     String              ○      I       工程番号
' 2   sCryNo                      String              ○      I       引上指示№、又は、結晶番号
' 3   sTgetHinban                 String              ○      I       狙い品番
' 4   iFixCutLeng                 Integer             ○      I       定尺幅
' 5   iAllLeng                    Integer             ○      I       引上げ長
' 6   iSprFlg                     Integer             ○      I       狙い品番で定尺(0：狙い品番定尺,1；配列変数で定尺)
' 7   tCutBlkHinban()             typ_CutBlkHinban    ○      I/O     切断ﾌﾞﾛｯｸ品番構造体(配列)
' 8   iErr_Code                   Integer             ○      O       ｴﾗｰｺｰﾄﾞ(正常時は0)
' 9   sErr_Msg                    String              ○      O       ｴﾗｰﾒｯｾｰｼﾞｺｰﾄﾞ

    On Error GoTo ErrorHandler
    
    '配列カウント
    Dim hinban_ichi_flg             As Integer
    Dim hinban_ichi_flg1            As Integer
    
    '定尺可能長
    Dim w_FixCutLeng  As Integer
    Dim teijyaku_ok_length          As Integer              '定尺カット可能長さ保持
    Dim top_teijyaku, bot_teijyaku  As Integer
    
    '使用値の最下位部長
    Dim under_length                As String               '最下部長保持変数 関数呼び出し用
    Dim iunder_leng                 As Integer              '実際の値
    
    '添字
    Dim w_i                         As Integer
    Dim w_x                         As Integer
    Dim w_y                         As Integer
    Dim indx                        As Integer
    
    '定尺配列格納
    Dim eCutBlkHinban()             As typ_ChangeHinb       '定尺調整後にＧＺ品番がある場合に付加する配列
    Dim wCutBlkHinban()             As typ_ChangeHinb       'ＧＺ品番を除いた配列
    Dim wCutBlkHinban1()            As typ_ChangeHinb       '最下位部によるチェック後に調整が必要な場合に必要
    Dim ChangeHin()                 As typ_ChangeHinb       '画面の定尺用に分解する。また最下位部も考慮し分解する。
    Dim intHin()                    As typ_ChangeHinb       '画面の定尺に沿った配列を同一品番かつ切断にて集約する。
    
    
    '最下位部調整時に必要変数
    Dim flg                         As Boolean
    Dim w_LENG1                     As Integer
    Dim w_LENG2                     As Integer
    Dim w_sa                        As Integer
    Dim w_pos1                      As Integer
    Dim w_pos2                      As Integer
    Dim cnt                         As Integer
    
    '配列へ格納時に対象配列添字
    Dim c_pos                       As Integer
    
    '狙い品番での配列作成時に幾つ配列が必要か求める
    Dim W_HIN1                      As Double
    
    'ＧＺ判断変数
    Dim w_gztop                     As Boolean
    Dim w_gztail                    As Boolean
    
    Dim lp                          As Integer
    Dim cpCutBlkHinban()            As typ_CutBlkHinban
    
    
    
    funGetFixLengCut = 0

    teijyaku_ok_length = 0
    hinban_ichi_flg = UBound(tCutBlkHinban)
    
    
    '---------------------------------------------------------------------------------------------------------------
    If iSprFlg = 1 Then
    ' --処理追加-- 03/10/22  引上げ長以上の長さは処理できない
            ' 下側から引上げ長以上の位置がないかチェック
            For lp = hinban_ichi_flg To 1 Step -1
                
                If tCutBlkHinban(lp).INPOS > iAllLeng Then
                    sErr_Msg = "TJ006"
                    iErr_Code = 1
                    funGetFixLengCut = 1
                    Exit Function
                End If
            Next
    ' --処理追加-- 03/10/22
    
    ' --処理追加-- 03/10/22　昇順に並べ替える
            ReDim Preserve cpCutBlkHinban(hinban_ichi_flg + 1)      ' 一時コピー用の配列
    
            For lp = 1 To hinban_ichi_flg
                Dim pos     As Integer
                Dim lp2     As Integer
                                
                pos = lp
                For lp2 = 1 To lp - 1
                    If tCutBlkHinban(lp).INPOS < cpCutBlkHinban(lp2).INPOS Then
                        pos = lp2
                        Exit For
                    End If
                Next
                
                If pos <> lp Then
                    For lp2 = pos To lp - 1
                        cpCutBlkHinban(lp2 + 1) = cpCutBlkHinban(lp2)
                    Next
                End If
                
                cpCutBlkHinban(pos) = tCutBlkHinban(lp)
            Next
    ' --処理追加-- 03/10/22
    
    ' --処理追加-- 03/10/17  最終位置が必ず配列に設定されているとは限らない
        ' 切断ﾌﾞﾛｯｸ品番構造体の最終データが引上げ長に等しいか？
        If cpCutBlkHinban(hinban_ichi_flg).INPOS <> iAllLeng Then
            cpCutBlkHinban(hinban_ichi_flg).LENGTH = iAllLeng       ' 長さが設定されていないので設定
            hinban_ichi_flg = hinban_ichi_flg + 1                   ' 最終位置を追加
            cpCutBlkHinban(hinban_ichi_flg).INPOS = iAllLeng        '
            cpCutBlkHinban(hinban_ichi_flg).Cut = 1                 ' カット指定する事
            cpCutBlkHinban(hinban_ichi_flg).hinban.hinban = ""      '
        End If
    
    
        Erase tCutBlkHinban                                         ' イレースしないと配列を再定義できない
        ReDim tCutBlkHinban(hinban_ichi_flg)                        ' ※引数配列の為か？
        
        For lp = 1 To hinban_ichi_flg                               ' コピーを書き戻す
            tCutBlkHinban(lp) = cpCutBlkHinban(lp)
            
            ' 正しい順番にデータが設定されていなかった場合に長さ情報も不正になっているので再設定
            If lp <> hinban_ichi_flg Then
                tCutBlkHinban(lp).LENGTH = cpCutBlkHinban(lp + 1).INPOS
            End If
        Next
    End If
    ' --処理追加-- 03/10/17
    '---------------------------------------------------------------------------------------------------------------
    
    
    '-------------------------------------------------------------------------------------------
    '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    'パターン　Ａ　－－－－－＞品番位置指定なし
    'If hinban_ichi_flg = 0 Then
    If iSprFlg = 0 Then
        '狙い品番の仕様値を取得
        '---------------------------------------------------------------------------------------
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '３、ブロック単位保証フラグをターゲット品番でチェックする
        If Check_TBCME36_DB(sTgetHinban) = False Then
            'エラー : カット不可能です
            sErr_Msg = "TJ001"
            iErr_Code = 1
            funGetFixLengCut = 1
            Exit Function
        End If
        '|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        teijyaku_ok_length = iAllLeng
    Else
    'パターン　Ｂ　－－－－－＞品番位置指定あり
     
        '----------------------------------------------------------------------------------------
        '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
        '３、確認ブロック単位保証フラグ
        '一つでもフラグが立っていなければ定尺カットは出来ない
        For indx = 1 To hinban_ichi_flg '最後の２個前までがチェック対象範囲
''''' --03/10/17--  途中Ｚ，Ｇ品番がブロック単位保障フラグチェックでエラーになる恐れがある
'''''            '先頭または後尾がＺ，Ｇ品番のときはＤＢチェックは行わない
'''''            If indx = 1 Or indx = hinban_ichi_flg Then
                If StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "Z", vbTextCompare) <> 0 And StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "G", vbTextCompare) <> 0 Then
                    '品番がなかったら（空白だったら）フラグチェックをしない
                    If Trim$(tCutBlkHinban(indx).hinban.hinban) <> "" Then
                        'ＤＢブロック単位保障フラグチェック
                        If Check_TBCME36_DB(tCutBlkHinban(indx).hinban) = False Then
                            'エラー : カット不可能です
                            sErr_Msg = "TJ001"
                            iErr_Code = 1
                            funGetFixLengCut = 1
                            Exit Function
                        End If
                    End If
                End If
'''''            Else
'''''                '品番がなかったら（空白だったら）フラグチェックをしない
'''''                If Trim$(tCutBlkHinban(indx).hinban.hinban) <> "" Then
'''''                    'ＤＢブロック単位保障フラグチェック
'''''                    If Check_TBCME36_DB(tCutBlkHinban(indx).hinban) = False Then
'''''                        'エラー : カット不可能です
'''''                        sErr_Msg = "TJ001"
'''''                        iErr_Code = 1
'''''                        funGetFixLengCut = 1
'''''                        Exit Function
'''''                    End If
'''''               End If
'''''            End If
''''' --03/10/17--  途中Ｚ，Ｇ品番がブロック単位保障フラグチェックでエラーになる恐れがある
            '３ END ---------------------------------------------------------------------------------
            
            '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
            '４ 指定品番チェック　＆＆　定尺カット可能長さ算出処理
            
            If indx > 1 And indx < hinban_ichi_flg - 1 Then
                ' 品番/位置指定　＝１のときのみ処理する
                '①　Ｚ、Ｇ品番中途存在確認
                If StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(indx).hinban.hinban), "G", vbTextCompare) = 0 Then
                    sErr_Msg = "TJ004"
                    iErr_Code = 1
                    funGetFixLengCut = 1
                    Exit Function
                End If
            End If
        Next indx
        
        '４－②  定尺カット可能な長さを算出-----------------------------------------------------
        ' 途中の 定尺カット可能長さ足し込み
        
        '先頭の長さを取得
        ' 先頭がＧ，Ｚ品番のとき
        If StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "G", vbTextCompare) = 0 Then
          top_teijyaku = tCutBlkHinban(2).INPOS  '２番目の長さを取得
        Else
          top_teijyaku = 0 '無条件に先頭は０とする
        End If
        
        '後尾の長さを取得-----------------------------------------
        '最後尾が　Z,Ｇ品番だったら 一個前（配列では２個前）の長さを取得
        If StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "G", vbTextCompare) = 0 Then
           bot_teijyaku = tCutBlkHinban(hinban_ichi_flg - 1).INPOS  '１つ前の長さを取得 == (自分の位置)
        Else
        'Ｚ，Ｇ品番ではない
           bot_teijyaku = tCutBlkHinban(hinban_ichi_flg).INPOS  '最後尾の長さを取得
        End If
        
        '定尺可能長さの解
        teijyaku_ok_length = bot_teijyaku - top_teijyaku
''        teijyaku_ok_length = bot_teijyaku
        '４－② END-----------------------------------------------------------------------------
        
    '４ END ---------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    End If
    
    '----------------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '５　最下位部の長さの取得
    If funCodeDBGet("SB", "TJ", "LEN", 0, " ", under_length) <> 0 Or Val(under_length) < 0 Then
        'データ取得時エラー
        sErr_Msg = "TJ002"
        iErr_Code = 1
        funGetFixLengCut = 1
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------
    'エラーチェック
    '定尺可能長　＜　最下部長　の時エラー
    iunder_leng = Val(under_length)
     
    If teijyaku_ok_length < Val(under_length) Then
       sErr_Msg = "TJ003"
       iErr_Code = 1
       funGetFixLengCut = 1
        Exit Function
    End If
    
    '５ END ---------------------------------------------------------------------------------
    '||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    '---------------------------------------------------------------------------------------

    
    
    '６定尺カットデータ作成処理

    '---------------------------------------------------------------------------------------
    '６－①     位置／品番指定なし　＝０の時の処理
    
    If hinban_ichi_flg = 0 Then
    
        Erase tCutBlkHinban
        c_pos = 1
        w_x = 0
        w_y = 0
        '画面の定尺に沿って配列を作成する。
        w_FixCutLeng = teijyaku_ok_length
        W_HIN1 = w_FixCutLeng / iFixCutLeng
        w_i = K_fncRoundUp(W_HIN1, 0)
        For w_x = 0 To w_i
            ReDim Preserve tCutBlkHinban(c_pos)
            If w_x = w_i Then
                tCutBlkHinban(c_pos).hinban.hinban = Space(8)
                tCutBlkHinban(c_pos).INPOS = w_FixCutLeng
                tCutBlkHinban(c_pos).LENGTH = 0
            Else
                tCutBlkHinban(c_pos).hinban.hinban = sTgetHinban.hinban
                tCutBlkHinban(c_pos).INPOS = w_y
                If w_y + iFixCutLeng > w_FixCutLeng Then
                    tCutBlkHinban(c_pos).LENGTH = w_FixCutLeng
                Else
                    tCutBlkHinban(c_pos).LENGTH = w_y + iFixCutLeng
                End If
            End If
            w_y = w_y + iFixCutLeng
            c_pos = c_pos + 1
        Next
        hinban_ichi_flg = UBound(tCutBlkHinban)
    End If
    
    '６－①　処理の終わり　-----------------------------------------------------------------
    
    Erase wCutBlkHinban     'ＧＺ品番を除いた配列
    Erase eCutBlkHinban     '定尺調整後にＧＺ品番がある場合に付加する配列
    Erase ChangeHin         '画面の定尺用に分解する。また最下位部も考慮し分解する。
    Erase intHin
    c_pos = 1
    
    'Z,Ｇ品番を別配列に格納し、処理完了後に配列に設定する。
    ' 先頭がＧ，Ｚ品番のとき
    If StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(1).hinban.hinban), "G", vbTextCompare) = 0 Then
        w_gztop = True
    Else
        w_gztop = False
    End If
    '最後尾が　Z,Ｇ品番のとき
    If StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(hinban_ichi_flg - 1).hinban.hinban), "G", vbTextCompare) = 0 Then
        w_gztail = True
    Else
        w_gztail = False
    End If
    For w_i = 1 To hinban_ichi_flg
        'Z,Ｇ品番を以外を別配列に格納し、処理完了後にＧＺ品番と組み合わせる
        If StrComp(Trim(tCutBlkHinban(w_i).hinban.hinban), "Z", vbTextCompare) = 0 Or StrComp(Trim(tCutBlkHinban(w_i).hinban.hinban), "G", vbTextCompare) = 0 Then
                If w_i = hinban_ichi_flg - 1 Then
                    ReDim Preserve wCutBlkHinban(c_pos)
                    wCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg).hinban
                    wCutBlkHinban(c_pos).INPOS = tCutBlkHinban(w_i).INPOS
                    wCutBlkHinban(c_pos).LENGTH = tCutBlkHinban(hinban_ichi_flg).LENGTH
                    Exit For
                End If
        Else
                ReDim Preserve wCutBlkHinban(c_pos)
                wCutBlkHinban(c_pos).hinban = tCutBlkHinban(w_i).hinban
                wCutBlkHinban(c_pos).INPOS = tCutBlkHinban(w_i).INPOS
                wCutBlkHinban(c_pos).LENGTH = tCutBlkHinban(w_i).LENGTH
                c_pos = c_pos + 1
        End If
    Next
    
    '最下位部調整時に使用する為待避
    wCutBlkHinban1() = wCutBlkHinban()
    
    c_pos = 1
    w_pos1 = 0
    w_pos2 = 0
    w_sa = 0
    w_LENG1 = 0
    w_LENG2 = 0
    cnt = 0
    hinban_ichi_flg1 = UBound(wCutBlkHinban)
    '画面の定尺に沿って配列を分解する。
    w_FixCutLeng = iFixCutLeng
    flg = True
    For w_i = 1 To hinban_ichi_flg1
        ReDim Preserve ChangeHin(c_pos)
        '画面の定尺と変数が等しいとき（カットのはじまり）を判断
        If w_FixCutLeng = iFixCutLeng Then
            ChangeHin(c_pos).Cut = EnumCutFlag.DoCut
        Else
            ChangeHin(c_pos).Cut = EnumCutFlag.NoCut
        End If
        '最下位部調整が必要なカットがあるかチェックする。ある場合はカット長を調整
        If Not flg Then
            If w_pos1 = wCutBlkHinban(w_i).INPOS Then
                w_FixCutLeng = w_LENG1
                iFixCutLeng = w_LENG1
                cnt = cnt + 1
            Else
                If cnt = 1 Then
                    If ChangeHin(c_pos).Cut = EnumCutFlag.DoCut Then
                        w_FixCutLeng = w_LENG2
                        iFixCutLeng = w_LENG2
                        flg = True
                    End If
                End If
            End If
        End If
        '画面の定尺に合わせて配列を分解
        If (wCutBlkHinban(w_i).LENGTH - wCutBlkHinban(w_i).INPOS) <= w_FixCutLeng Then
            ChangeHin(c_pos).hinban = wCutBlkHinban(w_i).hinban
            ChangeHin(c_pos).INPOS = wCutBlkHinban(w_i).INPOS
            ChangeHin(c_pos).LENGTH = wCutBlkHinban(w_i).LENGTH - wCutBlkHinban(w_i).INPOS
            w_FixCutLeng = w_FixCutLeng - ChangeHin(c_pos).LENGTH
            '次のカットを切断
            If w_FixCutLeng = 0 Then
                w_FixCutLeng = iFixCutLeng
            End If
        Else
            ChangeHin(c_pos).hinban = wCutBlkHinban(w_i).hinban
            ChangeHin(c_pos).INPOS = wCutBlkHinban(w_i).INPOS
            ChangeHin(c_pos).LENGTH = w_FixCutLeng
            wCutBlkHinban(w_i).INPOS = wCutBlkHinban(w_i).INPOS + w_FixCutLeng
            w_FixCutLeng = iFixCutLeng
            w_i = w_i - 1
        End If
        '最終カットは強制的に切断
        If w_i = hinban_ichi_flg1 Then
            ChangeHin(c_pos).Cut = EnumCutFlag.DoCut
           '最終カットが最下位部より小さい場合再配列調整が必要かチェック
            flg = True
            For w_x = c_pos - 1 To 1 Step -1
                If flg Then
                    If ChangeHin(w_x).Cut = EnumCutFlag.DoCut Then
                        If ChangeHin(c_pos).INPOS - ChangeHin(w_x).INPOS < Val(under_length) Then
                            w_sa = Val(under_length) - (ChangeHin(c_pos).INPOS - ChangeHin(w_x).INPOS)
                            w_LENG2 = Val(under_length)
                            w_pos2 = ChangeHin(w_x).INPOS
                            flg = False
                        Else
                            Exit For
                        End If
                    End If
                Else
                    If ChangeHin(w_x).Cut = EnumCutFlag.DoCut Then
                            w_pos1 = ChangeHin(w_x).INPOS
                            w_LENG1 = (w_pos2 - ChangeHin(w_x).INPOS) - w_sa
                            Exit For
                    End If
                End If
            Next
            '最下位部調整が必要な為、再度配列作成（各変数初期設定に戻す）
            If Not flg Then
                Erase ChangeHin
                w_FixCutLeng = iFixCutLeng
                wCutBlkHinban() = wCutBlkHinban1()
                w_i = 0
                c_pos = 0
            End If
        End If
        c_pos = c_pos + 1
    Next
    
    c_pos = 0
    '上記の画面の定尺に沿った配列を同一品番かつ切断にて集約する。
    For w_i = 1 To UBound(ChangeHin)
        If ChangeHin(w_i).Cut = EnumCutFlag.DoCut Then
            c_pos = c_pos + 1
             ReDim Preserve intHin(c_pos)
             intHin(c_pos).hinban = ChangeHin(w_i).hinban
            intHin(c_pos).INPOS = ChangeHin(w_i).INPOS
             intHin(c_pos).Cut = ChangeHin(w_i).Cut
        Else
            If Trim$(ChangeHin(w_i - 1).hinban.hinban) <> Trim$(ChangeHin(w_i).hinban.hinban) Then
                 c_pos = c_pos + 1
                 ReDim Preserve intHin(c_pos)
                 intHin(c_pos).hinban = ChangeHin(w_i).hinban
                 intHin(c_pos).INPOS = ChangeHin(w_i).INPOS
                 intHin(c_pos).Cut = ChangeHin(w_i).Cut
            End If
        End If
    Next
    
    '最下位部チェックを行う。最下部以下のカット長がある場合はエラー
    c_pos = 0
    For w_i = 1 To UBound(intHin)
        '切断カットか
        If intHin(w_i).Cut = EnumCutFlag.DoCut Then
            For w_x = w_i + 1 To UBound(intHin)
                '切断カットか
                If intHin(w_x).Cut = EnumCutFlag.DoCut Then
                    If intHin(w_x).INPOS - intHin(w_i).INPOS < Val(under_length) Then
                        sErr_Msg = "TJ005"
                        iErr_Code = 1
                        funGetFixLengCut = 1
                        Exit Function
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    
    'Z,G品番を合成
    c_pos = 0
    If w_gztop = True Then
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(1).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(1).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(1).Cut
    End If
    For w_i = 1 To UBound(intHin)
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = intHin(w_i).hinban
        eCutBlkHinban(c_pos).INPOS = intHin(w_i).INPOS
        eCutBlkHinban(c_pos).Cut = intHin(w_i).Cut
    Next
    If w_gztail = True Then
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg - 1).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(hinban_ichi_flg - 1).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(hinban_ichi_flg - 1).Cut
        c_pos = c_pos + 1
        ReDim Preserve eCutBlkHinban(c_pos)
        eCutBlkHinban(c_pos).hinban = tCutBlkHinban(hinban_ichi_flg).hinban
        eCutBlkHinban(c_pos).INPOS = tCutBlkHinban(hinban_ichi_flg).INPOS
        eCutBlkHinban(c_pos).Cut = tCutBlkHinban(hinban_ichi_flg).Cut
    End If
    
    'ＬＥＮＧＴＨを調整
    c_pos = 0
    For w_i = 1 To UBound(eCutBlkHinban) - 1
        eCutBlkHinban(w_i).LENGTH = eCutBlkHinban(w_i + 1).INPOS - eCutBlkHinban(w_i).INPOS
    Next
    
    '７　結果セットを呼出元へ返す。
    tCutBlkHinban() = eCutBlkHinban()

    
    
Error:
    Exit Function
    
ErrorHandler:
    funGetFixLengCut = -1
    GoTo Error

End Function
' funGetFixLengCut  END OF FUNCTION          ---------------------------------------------------------------------------------------------------

'====================================================================================================================
Rem                                                                                                                 *
Rem ・ブロック単位保証 テーブルを参照しブロック単位保証フラグをチェックするファンクション　                         *
'====================================================================================================================
'  参照テーブル                         TBCME036                                                                    *
'  項目名　　　ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ        BLOCKHFLAG  0: 指定された、狙い品番の仕様値を取得する。                     *
'                                                   1: 指定されている全品番の仕様値を取得する。（配列作成）         *
'-------------------------------------------------------------------------------------------------------------------*
'                                                                                                                   *
' 引数群                                                                                                            *
'　第１引数　　品番   :フラグを参照する品番                                                                         *
'                                                                                                                   *
'-------------------------------------------------------------------------------------------------------------------*
'   戻り値                                                                                                          *
'   正常終了               :TRUE       ブロックフラグ＝0（定尺可能）                                              *
'                           :FALSE      ブロックフラグ＝1 （定尺不可能）                                           *
'-------------------------------------------------------------------------------------------------------------------*
'                           0                   正常終了                                                            *
'                           1                   正常終了 (定尺不可)                                                 *
'                           -1                  異常終了                                                            *
'===================================================================================================================*

Function Check_TBCME36_DB(Fhinban As tFullHinban) As Boolean
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim recCnt      As Long         'レコード数
    Dim i           As Long
    Dim dbIsMine    As Boolean

    On Error GoTo ErrorHandler

    ''SQLを組み立てる
    sql = "Select BLOCKHFLAG From TBCME036"
    sql = sql & " where HINBAN   = '" & Fhinban.hinban & "' "
    sql = sql & "   and MNOREVNO =  " & Fhinban.mnorevno
    sql = sql & "   and FACTORY  = '" & Fhinban.factory & "' "
    sql = sql & "   and OPECOND  = '" & Fhinban.opecond & "' "

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        Check_TBCME36_DB = False
        Exit Function
    End If

    If rs.RecordCount <> 0 Then
        If rs.Fields(0) = "0" Then
                Check_TBCME36_DB = True
         Else
                Check_TBCME36_DB = False
        End If
    Else
         Check_TBCME36_DB = False
    End If
        
Error:
    Exit Function
    
ErrorHandler:
    Check_TBCME36_DB = False
    GoTo Error

End Function
'============================================================
'
'機能            :切り上げ
'
'引き数          :①切り上げする数値
'　　　           ②小数第？位の？桁
'
'戻り値          :①切り上げ後の数値
'
'機能説明        :数値の切り上げ時に使用する。
'
'備考            :
'
'============================================================
Public Function K_fncRoundUp(Mdbl_Su As Double, Mint_keta As Integer) As Currency
    
    Dim Mcur_work   As Currency
    Dim Mdbl_ret    As Double
    
    On Error GoTo ErrorHandler
    
    K_fncRoundUp = True
    
    Mcur_work = Mdbl_Su
    
    Mdbl_ret = 10 ^ Abs(Mint_keta)
    
    If Mint_keta > 0 Then
        K_fncRoundUp = Int(Mcur_work * Mdbl_ret + 0.9999) / Mdbl_ret
    Else
        K_fncRoundUp = Int(Mcur_work / Mdbl_ret + 0.9999) * Mdbl_ret
    End If
    
Error:
    Exit Function
    
ErrorHandler:
    K_fncRoundUp = False
    GoTo Error

End Function
