Attribute VB_Name = "mdlVbx5xx2"
' @(h) mdlVbx5XX2.BAS              ver 1.00 ( '00.01.06 和泉沢)
' @(s)
Option Explicit
''抽出条件設定・未設定フラグ
''設定済み=True 未設定=False
Public gbFlgVbx5xx2 As Boolean
''メインSQLと結合する項目を指定する。
''分割結晶番号：xtalc2,3等 引上結晶番号：owxtalc2 品番：hinbc2
Public gsKeyVbx5xx2 As String
'' 結合WHERE条件文グローバル変数
'' テーブル名：JKB=分割結晶 JKC=引上結晶 JKD=製造指示 JKE=仕様
Public gsSqlWhereVbx5xx2 As String

'テーブルアクセスフラグ
'項目により結合するテーブルを選択する。
'使用=True,未使用=False （CheckF4Vbx5XX2で設定）
Private bXODC1 As Boolean   '引上結晶
Private bXODC2 As Boolean   '9/13 Yam
Private bXODE2 As Boolean   '製造指示
Private bSIYO1 As Boolean   '仕様１
Private bSIYO2 As Boolean   '仕様２
Private bSIYO3 As Boolean   '仕様３

''VBX5041納期設定・未設定フラグ
''設定済み=True 未設定=False
Public gbFlgVbx5040Nouki As Boolean
'変換前の機種コードの保存(11/17 Yam追加）
Public kisyuNm As String

Type CDNAMEDAT  ''コード・名前
    Cd As String
    Nm As String
End Type

' @(f)
' 機能      : キー制御処理
' 返り値    : なし
' 引数      : KeyCode   -   キーコード
' 機能説明  : キーコードによって処理を振り分ける処理を行う
' 備考      : 画面状態フラグ     0:初期状態
'                               1:確認実行
'                               2:登録実行
'
Public Sub KeyActionVbx5XX2(KeyCode As Integer)

gbFlgVbx5040Nouki = False   ''2000/06/07修正

    With frmSub
        ''コマンドボタン機能振分け
        Select Case KeyCode
        Case vbKeyF3
            '''キャンセル
            If .cmdF(3).Enabled = False Then Exit Sub
            ''画面初期化
            Call InitVbx5XX2(True)
            ''抽出条件設定OFF
            gbFlgVbx5xx2 = False
            ''メイン画面復帰
            frmMain.Show
            frmSub.Hide
        Case vbKeyF4
            '''修正
            If .cmdF(4).Enabled = False Then Exit Sub
            ''抽出画面専用問合せ文作成
            If vbKeyActionF4Vbx5XX2() Then
                ''画面初期化
                Call InitVbx5XX2(False)
                ''抽出条件設定ON
                gbFlgVbx5xx2 = True
                ''メイン画面復帰
                frmMain.Show
                frmSub.Hide
            End If
        End Select
    End With
End Sub

' @(f)
' 機能      : 抽出条件WHERE文作成（MAIN）
' 返り値    : なし
' 引数      : なし
' 機能説明  : 抽出条件画面で設定した項目の条件文を作成する。
' 備考      :

Private Function vbKeyActionF4Vbx5XX2()
    Dim i           As Integer      ''ループカウンタ
    Dim sWk         As String       ''作業領域
    Dim iWild       As Integer      ''ワイルドカード使用フラグ
    vbKeyActionF4Vbx5XX2 = False
    gbFlgVbx5040Nouki = False
    ''入力項目のチェック
    If CheckF4Vbx5XX2() = False Then
        Exit Function
    End If
    
    ''抽出条件画面 問合せ文の作成
    ''ex)作成イメージ
    ''      ,分割結晶 JKB, 引上結晶 JKC, 製造指示 JKD, 仕様 JKE
    ''      WHERE （メイン画面SQL）JKA.結合キー = JKB.対応するキー
    ''      AND JKB.引上結晶 = JKC.引上結晶
    ''      AND JKB.品番 = JKD.品番
    ''      AND JKB.品番 = JKE.品番
    ''      AND 以下選択項目により可変
    ''■使用するテーブルの記述
    ''  分割結晶テーブルは必ず結合する。
    ''  分割結晶テーブルから引上結晶・製造指示、仕様テーブルを結合する。
    If bXODC2 Then   '9/13 Yam
        gsSqlWhereVbx5xx2 = ", xodc2 JKB "          ''分割結晶結合(別名：JKB)
    Else
        gsSqlWhereVbx5xx2 = " "                     ''分割結晶結合(別名：JKB)
    End If
    If bXODC1 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", xodc1 JKC"     ''引上結晶(別名：JKC)
    End If
    'If bXODE2 Then
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", xode2 JKD"     ''製造指示結合(別名：JKD)
    'End If
    If bSIYO1 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods1 JKE"     ''仕様１結合(別名：JKE)
    End If
    If bSIYO2 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods2_es JKF"  ''仕様１結合(別名：JKE)
    End If
    If bSIYO3 Then
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & ", sods2_pr JKG"  ''仕様１結合(別名：JKE)
    End If
        
    ''■キー結合部に記述
    ''  メイン画面SQLと使用テーブルの結合条件を作成する。
    ''  分割結晶(JKB)とは必ず結合する。
    ''メインSQL＝分割結晶(JKB).（メインSQLが保持するキーに依存）
    ''分割結晶で結合する場合
    If bXODC2 Then  '9/13 Yam
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " where JKA.ck1 = JKB." & gsKeyVbx5xx2
    Else
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " where JKA.ck1 = JKA.ck1 "
    End If
    If bXODC1 Then
        ''分割結晶(JKB).引上結晶番号=引上結晶.引上結晶番号
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 = JKC.xtalc1"
    End If
    'If bXODE2 Then
        ''分割結晶(JKB).引上結晶番号=製造指示.品番
        ''  2000/06/19  修正開始    分割結晶テーブルの品番と結合しないで工程実績の
        ''  品番と結合するため修正（リビジョンも結合するよ) ｂｙ  牟田
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 || JKB.hinbrc2 = JKD.hinbe2 || JKD.hinbre2"
        ''02/14/2000 製造指示(XODE2)の一位化のために追加
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   =   JKD.hinbe2"
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   =   JKD.hinbre2"
    '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.knnoc2 = JKD.knnoe2"
        ''  2000/06/19  修正ここまで    分割結晶テーブルの品番と結合しないで工程実績の
        ''  品番と結合するため修正（リビジョン、製番も結合するよ) ｂｙ  牟田
    'End If
    If bSIYO1 Then
        ''分割結晶(JKB).引上結晶番号=仕様.品番
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 || JKB.hinbrc2 = JKE.hinbc3 || JKE.hinbrc3"
        ''  2000/06/19  修正開始    分割結晶テーブルの品番と結合しないで工程実績の
        ''  品番と結合するため修正（リビジョンも結合するよ) ｂｙ  牟田
'        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.hinbc2 = JKE.hinbc3 and JKB.hinbrc2 = JKE.hinbrc3"
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKE.specnos1 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKE.specnors1"
        ''  2000/06/19  修正ここまで    分割結晶テーブルの品番と結合しないで工程実績の
        ''  品番と結合するため修正（リビジョンも結合するよ) ｂｙ  牟田
    End If
    If bSIYO2 Then
        ''分割結晶(JKB).引上結晶番号=仕様.品番
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKF.specnos2 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKF.specnors2"
    End If
    If bSIYO3 Then
        ''分割結晶(JKB).引上結晶番号=仕様.品番
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK2   = JKG.specnos2 "
        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.cK3   = JKG.specnors2"
    End If
    
    ''■抽出画面項目の条件文を記述
    With frmVBX5XX2
        ''品番(分割結晶)
        If .optHinban(0).Value Then
            ''一致
            If Trim(.txtHinban(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 >= '" & Trim(UCase(.txtHinban(0).Text)) & "'"
            End If
            If Trim(.txtHinban(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 <= '" & Trim(UCase(.txtHinban(1).Text)) & "'"
            End If
        Else
            ''不一致
            If Trim(.txtHinban(0).Text) <> "" And Trim(.txtHinban(1).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKA.ck2 || JKA.ck3 < '" & Trim(UCase(.txtHinban(0).Text)) & "'"
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " or JKA.ck2 || JKA.ck3 > '" & Trim(UCase(.txtHinban(1).Text)) & "')"
            Else
                If Trim(.txtHinban(0).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 <= '" & Trim(UCase(.txtHinban(0).Text)) & "'"
                End If
                If Trim(.txtHinban(1).Text) <> "" Then
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKA.ck2 || JKA.ck3 >= '" & Trim(UCase(.txtHinban(1).Text)) & "'"
                End If
            End If
        End If
        ''機種(分割結晶)
        If Trim(.txtKisy.Text) <> "" Then
            If .optKisy(0).Value Then
                ''一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and SUBSTR(JKB.kisyuc2,1,2) = '" & Trim(.txtKisy.Text) & "'"
            Else
                ''不一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and SUBSTR(JKB.kisyuc2,1,2) != '" & Trim(.txtKisy.Text) & "'"
            End If
        End If
        
        .txtKisy.Text = kisyuNm   '11/17 追加(Yam)
        
        ''引上方法(分割結晶)
        sWk = ""
        For i = 0 To 2
            If Trim(.txtHikiageX(i).Text) <> "" Then
                sWk = sWk & "'" & Trim(.txtHikiageX(i).Text) & "',"
            End If
        Next
        If sWk <> "" Then
            sWk = Mid(sWk, 1, Len(sWk) - 1) '最後のカンマをとる。
            If .optHikiageX(0).Value Then
                ''一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.pumethc2 in(" & sWk & ")"
            Else
                ''不一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.pumethc2 not in(" & sWk & ")"
            End If
        End If
        ''PG-ID(引上結晶)
        .txtPgid.Text = UCase(.txtPgid.Text)   'Yam追加
        If Trim(.txtPgid.Text) <> "" Then
            sWk = Trim(.txtPgid.Text)
            ''ワイルドカード文字[?]を[_]に変換する(1文字ワイルドカード)
            Do
                iWild = InStr(sWk, "?")
                If iWild = 0 Then
                    Exit Do
                Else
                    sWk = Mid(sWk, 1, iWild - 1) & "_" & Mid(sWk, iWild + 1)
                End If
            Loop
            ''8桁に満たない場合は最後に[%]を付ける
            If Len(sWk) < 8 Then
                If Right(sWk, 1) <> "%" Then
                    sWk = sWk & "%"
                End If
            End If
            If .optPgid(0).Value Then
                ''一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKC.pgidc1 like '" & sWk & "'"
            Else
                ''不一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKC.pgidc1 not like '" & sWk & "'"
            End If
        End If
        ''製品区分(仕様)  1/22 Yam 修正
        If Trim(.txtSeizoKbn.Text) <> "" Then
                '4:その他の場合
            If Trim(.txtSeizoKbn.Text) = "9" Then
                If .optSeizoKbn(0).Value Then
                    ''一致
                    For i = 1 To 3
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 != '" & Right("00" & Trim(i), 2) & "'"
                    Next
                Else
                    ''不一致
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and ((JKE.comprdgrs1 = '01') or (JKE.comprdgrs1 = '02') or (JKE.comprdgrs1 = '03'))"
                End If
                '1,2,3:その他以外の場合
            Else
                If .optSeizoKbn(0).Value Then
                    ''一致
                        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 = '" & Right("00" & Trim(.txtSeizoKbn.Text), 2) & "'"
                Else
                    ''不一致
                    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comprdgrs1 != '" & Right("00" & Trim(.txtSeizoKbn.Text), 2) & "'"
                End If
            End If
        End If
        ''使用目的(仕様)
        sWk = ""
        For i = 0 To 5
            If Trim(.txtMokuteki(i).Text) <> "" Then
                sWk = sWk & "'" & Trim(.txtMokuteki(i).Text) & "'" & ","
            End If
        Next
        If sWk <> "" Then
            sWk = Mid(sWk, 1, Len(sWk) - 1) '最後のカンマをとる。
            If .optMokuteki(0).Value Then
                ''一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comcususes1 in(" & sWk & ")"
            Else
                ''不一致
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.comcususes1 not in(" & sWk & ")"
            End If
        End If
        ''向先(製造指示) ＜kuro分割結晶DBより製造指示DBに変更＞
        'sWk = ""
        'For i = 0 To 1
        '    If Trim(.txtMukaisaki(i).Text) <> "" Then
        '        sWk = sWk & "'" & Trim(.txtMukaisaki(i).Text) & "',"
        '    End If
        'Next
        'If sWk <> "" Then
        '    sWk = Mid(sWk, 1, Len(sWk) - 1) '最後のカンマをとる。
        '    If .optMukaisaki(0).Value Then
        '        ''一致
        '        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKD.swplace2 in(" & sWk & ")"
        '    Else
        '        ''不一致
        '        gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKD.swplace2 not in(" & sWk & ")"
        '    End If
        'End If
        ''納期(製造指示)
        'If Trim(.txtNoki(0).Text) <> "" Then
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and TO_CHAR(JKD.snyye2,'FM0000') || TO_CHAR(JKD.snmme2,'FM00') || TO_CHAR(JKD.sndde2,'FM00') >= '"
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & DateChange(Trim(.txtNoki(0).Text)) & "'"
        'End If
        'If Trim(.txtNoki(1).Text) <> "" Then
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and TO_CHAR(JKD.snyye2,'FM0000') || TO_CHAR(JKD.snmme2,'FM00') || TO_CHAR(JKD.sndde2,'FM00') <= '"
        '    gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & DateChange(Trim(.txtNoki(1).Text)) & "'"
        'End If
        ''号機(分割結晶)
        If Trim(.txtGoki(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 >= '" & Trim(.txtGoki(0).Text) & "'"
        End If
        If Trim(.txtGoki(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.owxtalc2 <= '" & Trim(.txtGoki(1).Text) & "999999999'"
        End If
        ''格上区分(分割結晶)
        If Trim(.txtKakuage.Text) <> "" Then
        ' 1/23 Yam修正   gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.laupc2 != '" & Trim(.txtKakuage.Text) & "ZZZZZZZZZ'"
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKB.laupc2 != '" & Trim(.txtKakuage.Text) & "'"
        End If
        ''直径区分(仕様)
        If Trim(.txtChokkei(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.commxdiadvs1 >= '" & Trim(.txtChokkei(0).Text) & "'"
        End If
        If Trim(.txtChokkei(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKE.commxdiadvs1 <= '" & Trim(.txtChokkei(1).Text) & "'"
        End If
        ''伝導型(仕様)
        '////Chihi 11/14 追加///
        .txtDendo.Text = UCase(.txtDendo.Text)
        If Trim(.txtDendo.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.mxtyps2 = '" & Trim(.txtDendo.Text) & "'"
        End If
        ''ドーパント(仕様)
        '/// Chihi 11/13追加
        .txtDoba.Text = UCase(.txtDoba.Text)
        If Trim(.txtDoba.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.mxdops2 = '" & Trim(.txtDoba.Text) & "'"
        End If
        ''方位(仕様)
        If Trim(.txtHoui.Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKG.axaxiss2 = '" & Trim(.txtHoui.Text) & "'"
        End If
        ''抵抗率(仕様）
        If Trim(.txtTeikoKbn) <> "" Then
            If Trim(.txtTeikouritsu(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalls2 >= " & Trim(.txtTeikouritsu(0).Text)
            End If
            If Trim(.txtTeikouritsu(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalls2 <= " & Trim(.txtTeikouritsu(1).Text)
            End If
        Else
            If Trim(.txtTeikouritsu(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalus2 >= " & Trim(.txtTeikouritsu(0).Text)
            End If
            If Trim(.txtTeikouritsu(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.rsxtalus2 <= " & Trim(.txtTeikouritsu(1).Text)
            End If
        End If
        ''抵抗(レンジ)(仕様)
        If Trim(.txtTeikou(0).Text) <> "" Then
            ''▼スラグ抵抗min(JKE.teikou_s)がnullの時のレコードは抽出できません。
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / JKF.teikou_s) >= " & Trim(.txtTeikou(0).Text)
            ''▼スラグ抵抗min(JKE.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0を1に置き換える場合↓)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / DECODE(JKF.teikou_s,0,1,JKF.teikou_s)) >= " & Trim(.txtTeikou(0).Text)
            ''▼スラグ抵抗min(JKE.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0をnullに置き換える場合↓)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.rsxtalus2 / DECODE(JKF.rsxtalls2,0,null,JKF.rsxtalls2)) >= " & Trim(.txtTeikou(0).Text)

        End If
        If Trim(.txtTeikou(1).Text) <> "" Then
            ''▼スラグ抵抗min(JKE.teikou_s)がnullの時のレコードは抽出できません。
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / JKF.teikou_s) <= " & Trim(.txtTeikou(1).Text)
            ''▼スラグ抵抗min(JKE.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0を1に置き換える場合↓)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.teikou_u / DECODE(JKF.teikou_s,0,1,JKF.teikou_s)) <= " & Trim(.txtTeikou(1).Text)
            ''▼スラグ抵抗min(JKE.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0をnullに置き換える場合↓)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.rsxtalus2 / DECODE(JKF.rsxtalls2,0,null,JKF.rsxtalls2)) <= " & Trim(.txtTeikou(1).Text)
        End If
        ''酸素濃度(仕様）
        If Trim(.txtSansoKbn) <> "" Then
            If Trim(.txtSanso(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgls2 >= " & Trim(.txtSanso(0).Text)
            End If
            If Trim(.txtSanso(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgls2 <= " & Trim(.txtSanso(1).Text)
            End If
        Else
            If Trim(.txtSanso(0).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgus2 >= " & Trim(.txtSanso(0).Text)
            End If
            If Trim(.txtSanso(1).Text) <> "" Then
                gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oislgus2 <= " & Trim(.txtSanso(1).Text)
            End If
        End If
        ''Oi(レンジ)(仕様)
        If Trim(.txtOi(0).Text) <> "" Then
            ''▼スラグ[Oi]min(JKF.teikou_s)がnullの時のレコードは抽出できません。
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / JKF.sanso_s) >= " & Trim(.txtOi(0).Text)
            ''▼スラグ[Oi]min(JKF.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0を1に置き換える場合↓)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / DECODE(JKF.sanso_s,0,1,JKF.sanso_s)) >= " & Trim(.txtOi(0).Text)
            ''▼スラグ[Oi]min(JKF.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0をnullに置き換える場合↓)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.oislgus2 / DECODE(JKF.oislgls2,0,null,JKF.oislgls2)) >= " & Trim(.txtOi(0).Text)
        End If
        If Trim(.txtOi(1).Text) <> "" Then
            ''▼スラグ[Oi]min(JKF.teikou_s)がnullの時のレコードは抽出できません。
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / JKF.sanso_s) <= " & Trim(.txtOi(1).Text)
            ''▼スラグ[Oi]min(JKF.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0を1に置き換える場合↓)
'            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.sanso_u / DECODE(JKF.sanso_s,0,1,JKF.sanso_s)) <= " & Trim(.txtOi(1).Text)
            ''▼スラグ[Oi]min(JKF.teikou_s)が0の時には除数が0になり、エラーが返ってきます。(除数0をnullに置き換える場合↓)
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and (JKF.oislgus2 / DECODE(JKF.oislgls2,0,null,JKF.oislgls2)) <= " & Trim(.txtOi(1).Text)
        End If
        ''ORG(仕様)
        If Trim(.txtOrg(0).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oiorgs2 >= " & Trim(.txtOrg(0).Text)
        End If
        If Trim(.txtOrg(1).Text) <> "" Then
            gsSqlWhereVbx5xx2 = gsSqlWhereVbx5xx2 & " and JKF.oiorgs2 <= " & Trim(.txtOrg(1).Text)
        End If
    End With
    
    vbKeyActionF4Vbx5XX2 = True
End Function

' @(f)
' 機能      : キーチェック処理
' 返り値    : 正常=True,異常=False
' 引数      : なし
' 機能説明  : 入力項目チェック
' 備考      :
'
Private Function CheckF4Vbx5XX2() As Boolean
    Dim i           As Integer      ''ループカウンタ
    Dim akisy As String

    CheckF4Vbx5XX2 = False
    bXODC1 = False
    bXODC2 = False '9/13 Yam
    bXODE2 = False
    bSIYO1 = False
    bSIYO2 = False
    bSIYO3 = False
    With frmSub
        ''品番チェック
        If Trim(.txtHinban(0).Text) <> "" And Trim(.txtHinban(1).Text) <> "" Then
            Call FillUpString(.txtHinban(0), "0")
            Call FillUpString(.txtHinban(1), "9")
            If Val(.txtHinban(0).Text) > Val(.txtHinban(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtHinban(0), RED_CTL)
                Call CtrlEnabled(.txtHinban(1), RED_CTL)
                Exit Function
            End If
            '分割結晶テーブル項目のためテーブルフラグは立てない。
        End If
        
        ''機種チェック            '11/17　Yam追加
        .txtKisy.Text = UCase(Trim(.txtKisy.Text))
        kisyuNm = .txtKisy.Text
        If Trim(.txtKisy.Text) <> "" Then
            If GetkisyNo(Trim(.txtKisy.Text), akisy) = False Then
                Exit Function
            End If
            .txtKisy.Text = akisy
            bXODC2 = True           '9/13 Yam
        End If
        ''引上方法チェック
        For i = 0 To 2
            If Trim(.txtHikiageX(i).Text) <> "" Then
            bXODC2 = True           '9/13 Yam
            End If
        Next
        ''PG-IDチェック
        If Trim(.txtPgid.Text) <> "" Then
            bXODC1 = True           '引上結晶ON
            bXODC2 = True           '9/13 Yam
        End If
        ''製品区分チェック
        If Trim(.txtSeizoKbn.Text) <> "" Then
            If (Val(.txtSeizoKbn.Text) < 1) Or (Val(.txtSeizoKbn.Text) > 4) _
            And (Val(.txtSeizoKbn.Text) <> 9) Or (IsNumeric(.txtSeizoKbn.Text) = False) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtSeizoKbn, RED_CTL)
                Exit Function
            End If
            bSIYO1 = True            '仕様ON
        End If
        ''使用目的チェック
        For i = 0 To 5
            If Trim(.txtMokuteki(i).Text) <> "" Then
                '順番に入力されていること
                If i > 0 Then
                    If Trim(.txtMokuteki(i - 1).Text) = "" Then
                        Call MsgOut(50, "", ERR_DISP)
                        Call CtrlEnabled(.txtMokuteki(i - 1), RED_CTL)
                        Exit Function
                    End If
                End If
                bSIYO1 = True        '仕様ON
            End If
        Next
        ''向先チェック
        'For i = 0 To 1
        '    If Trim(.txtMukaisaki(i).Text) <> "" Then
        '        '順番に入力されていること
        '        If i > 0 Then
        '            If Trim(.txtMukaisaki(i - 1).Text) = "" Then
        '                Call MsgOut(50, "", ERR_DISP)
        '                Call CtrlEnabled(.txtMukaisaki(i - 1), RED_CTL)
        '                Exit Function
        '            End If
        '        End If
        '        '製造指示ON kuro追加
        '        If Len(Trim(.txtMukaisaki(i).Text)) <> 0 Then
        '            bXODE2 = True
        '        End If
        '    End If
        'Next
        ''納期チェック
        'If Trim(.txtNoki(0).Text) <> "" Or Trim(.txtNoki(1).Text) <> "" Then
        '    '桁が入力されていた場合桁を埋める
        '    If Len(Trim(.txtNoki(0).Text)) <> 0 Then
        '        FillUpString .txtNoki(0), "0"
        '    ElseIf Len(Trim(.txtNoki(1).Text)) <> 0 Then
        '        FillUpString .txtNoki(1), "9"
        '    End If
            '期間チェック抽出条件用
        '    If KikanCheckVbx5XX2(.txtNoki(0), .txtNoki(1)) = False Then
        '        Exit Function
        '    End If
        '    bXODE2 = True            '製造指示ON
        '    gbFlgVbx5040Nouki = True '納期設定済み
        'End If
        ''号機チェック
        If Trim(.txtGoki(0).Text) <> "" Or Trim(.txtGoki(1).Text) <> "" Then
            If Trim(.txtGoki(0).Text) > Trim(.txtGoki(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtGoki(0), RED_CTL)
                Call CtrlEnabled(.txtGoki(1), RED_CTL)
                Exit Function
            End If
            bXODC2 = True           '9/13 Yam
        End If
        ''格上区分チェック
        If Trim(.txtKakuage.Text) <> "" Then
            If Trim(.txtKakuage.Text) <> "1" Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtKakuage, RED_CTL)
                Exit Function
            End If
            bXODC2 = True           '9/13 Yam
        End If
        ''直径区分チェック
        If Trim(.txtChokkei(0).Text) <> "" Or Trim(.txtChokkei(1).Text) <> "" Then
            If Trim(.txtChokkei(0).Text) > Trim(.txtChokkei(1).Text) Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtChokkei(0), RED_CTL)
                Call CtrlEnabled(.txtChokkei(1), RED_CTL)
                Exit Function
            End If
            bSIYO1 = True            '仕様ON
        End If
        ''伝導型チェック
        .txtDendo.Text = UCase(.txtDendo.Text)
        If Trim(.txtDendo.Text) <> "" Then
            If Trim(.txtDendo.Text) <> "P" And Trim(.txtDendo.Text) <> "N" Then
                Call MsgOut(50, "", ERR_DISP)
                Call CtrlEnabled(.txtDendo, RED_CTL)
                Exit Function
            End If
            bSIYO3 = True            '仕様ON
        End If
        ''ドーパントチェック
        If Trim(.txtDoba.Text) <> "" Then
            bSIYO3 = True            '仕様ON
        End If
        ''方位チェック
        If Trim(.txtHoui.Text) <> "" Then
            bSIYO3 = True            '仕様ON
        End If
        ''抵抗率チェック
        If Trim(.txtTeikouritsu(0).Text) <> "" Then
            If Not IsNumeric(.txtTeikouritsu(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtTeikouritsu(1).Text) <> "" Then
            If Not IsNumeric(.txtTeikouritsu(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtTeikouritsu(0).Text) <> "" And Trim(.txtTeikouritsu(1).Text) <> "" Then
            If Val(Trim(.txtTeikouritsu(0).Text)) > Val(Trim(.txtTeikouritsu(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikouritsu(0), RED_CTL)
                Call CtrlEnabled(.txtTeikouritsu(1), RED_CTL)
                Exit Function
            End If
        End If
        ''抵抗率参照チェック
        If Trim(.txtTeikoKbn.Text) <> "" And Trim(.txtTeikoKbn.Text) <> "1" And _
            (Trim(.txtTeikouritsu(0).Text) <> "" Or Trim(.txtTeikouritsu(1).Text) <> "") Then
            Call MsgOut(50)
            Call CtrlEnabled(.txtTeikoKbn, RED_CTL)
            Exit Function
        End If
        ''抵抗(レンジ)チェック
        If Trim(.txtTeikou(0).Text) <> "" Then
            If Not IsNumeric(.txtTeikou(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtTeikou(1).Text) <> "" Then
            If Not IsNumeric(.txtTeikou(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtTeikou(0).Text) <> "" And Trim(.txtTeikou(1).Text) <> "" Then
            If Val(Trim(.txtTeikou(0).Text)) > Val(Trim(.txtTeikou(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtTeikou(0), RED_CTL)
                Call CtrlEnabled(.txtTeikou(1), RED_CTL)
                Exit Function
            End If
        End If
        ''酸素濃度チェック
        If Trim(.txtSanso(0).Text) <> "" Then
            If Not IsNumeric(.txtSanso(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtSanso(1).Text) <> "" Then
            If Not IsNumeric(.txtSanso(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtSanso(0).Text) <> "" And Trim(.txtSanso(1).Text) <> "" Then
            If Val(Trim(.txtSanso(0).Text)) > Val(Trim(.txtSanso(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtSanso(0), RED_CTL)
                Call CtrlEnabled(.txtSanso(1), RED_CTL)
                Exit Function
            End If
        End If
        ''酸素濃度参照チェック
        If Trim(.txtSansoKbn.Text) <> "" And Trim(.txtSansoKbn.Text) <> "1" And _
            (Trim(.txtSanso(0).Text) <> "" Or Trim(.txtSanso(1).Text) <> "") Then
            Call MsgOut(50)
            Call CtrlEnabled(.txtSansoKbn, RED_CTL)
            Exit Function
        End If
        ''oi(レンジ)チェック
        If Trim(.txtOi(0).Text) <> "" Then
            If Not IsNumeric(.txtOi(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtOi(1).Text) <> "" Then
            If Not IsNumeric(.txtOi(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtOi(0).Text) <> "" And Trim(.txtOi(1).Text) <> "" Then
            If Val(Trim(.txtOi(0).Text)) > Val(Trim(.txtOi(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOi(0), RED_CTL)
                Call CtrlEnabled(.txtOi(1), RED_CTL)
                Exit Function
            End If
        End If
        ''ORGチェック
        If Trim(.txtOrg(0).Text) <> "" Then
            If Not IsNumeric(.txtOrg(0).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(0), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtOrg(1).Text) <> "" Then
            If Not IsNumeric(.txtOrg(1).Text) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(1), RED_CTL)
                Exit Function
            End If
            bSIYO2 = True            '仕様ON
        End If
        If Trim(.txtOrg(0).Text) <> "" And Trim(.txtOrg(1).Text) <> "" Then
            If Val(Trim(.txtOrg(0).Text)) > Val(Trim(.txtOrg(1).Text)) Then
                Call MsgOut(51)
                Call CtrlEnabled(.txtOrg(0), RED_CTL)
                Call CtrlEnabled(.txtOrg(1), RED_CTL)
                Exit Function
            End If
        End If
    End With
    CheckF4Vbx5XX2 = True
End Function

' @(f)
' 機能    : 期間チェック（VBX5XX2専用）
' 返り値  :  OK - TRUE
'            NG -FALSE
' 引き数  : ctlControlS : コントロール(開始日)
'           ctlControlE : コントロール(終了日)
' 機能説明:

Private Function KikanCheckVbx5XX2(ctlControlS As Control, ctlControlE As Control) As Boolean
    'xxxxxxxxxxxxxxxxxxxxxxx
    '   mdlCommon.bas?
    'xxxxxxxxxxxxxxxxxxxxxxx
    Dim sDtS    As String       ''集計期間開始日
    Dim sDtE    As String       ''集計期間終了日
    Dim sDtT    As String       ''システム日付
    Dim sDtL    As String       ''該当月の月末日(開始日)
    Dim sDtLE   As String       ''該当月の月末日(終了日)
    Dim sWk     As String
    
    KikanCheckVbx5XX2 = False
    
    ''集計期間取得変換(6桁[yymmdd]→8桁[yyyymmdd])
    sDtS = DateChange(Trim(ctlControlS.Text))
    sDtE = DateChange(Trim(ctlControlE.Text))
    
    ''開始日・終了日の数値チェック
    If Len(sDtS) = 8 Then
        If IsNumeric(Trim(sDtS)) = False Then
            Call CtrlEnabled(ctlControlS, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    ElseIf Len(sDtE) = 8 Then
        If IsNumeric(Trim(sDtE)) = False Then
            Call CtrlEnabled(ctlControlE, RED_CTL)
            Call MsgOut(52, "", ERR_DISP)
            Exit Function
        End If
    End If
    If Val(Mid(sDtS, 5, 2)) < 1 Or Val(Mid(sDtS, 5, 2)) > 12 Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    ElseIf Val(Mid(sDtE, 5, 2)) < 1 Or Val(Mid(sDtE, 5, 2)) > 12 Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call MsgOut(52, "", ERR_DISP)
        Exit Function
    End If
    ''整合性チェック（開始日＞終了日？）
    If Val(sDtS) > Val(sDtE) Then
        Call CtrlEnabled(ctlControlS, RED_CTL)
        Call CtrlEnabled(ctlControlE, RED_CTL)
        Call MsgOut(53, "", ERR_DISP)
        Exit Function
    End If
    
    KikanCheckVbx5XX2 = True
End Function

' @(f)
' 機能      : 画面初期化(Vbx5XX2)
' 返り値    : なし
' 引数      : bStatus：True=テキストクリア
'                    ：False=背景色のみ
' 機能説明  : 抽出画面の初期化

Public Sub InitVbx5XX2(bStatus As Boolean)
    Dim i As Integer    ''ループカウンタ

    With frmSub
        If bStatus Then
            ''ラジオボタン初期設定
            .optHinban(0).Value = True          ''品番 一致不一致
            .optKisy(0).Value = True            ''機種 一致不一致
            .optHikiageX(0).Value = True        ''引上方法 一致不一致
            .optPgid(0).Value = True            ''PGID 一致不一致
            .optSeizoKbn(0).Value = True        ''製品区分 一致不一致
            .optMokuteki(0).Value = True        ''使用目的 一致不一致
            '.optMukaisaki(0).Value = True       ''向先 一致不一致
        End If
        ''入力フィールドのクリア
        Call CtrlEnabled(.txtKisy, NORMAL_CTL, bStatus)            ''機種
        Call CtrlEnabled(.txtPgid, NORMAL_CTL, bStatus)            ''PGID
        Call CtrlEnabled(.txtSeizoKbn, NORMAL_CTL, bStatus)        ''製品区分
        Call CtrlEnabled(.txtKakuage, NORMAL_CTL, bStatus)         ''格上区分
        Call CtrlEnabled(.txtDendo, NORMAL_CTL, bStatus)           ''伝導型
        Call CtrlEnabled(.txtDoba, NORMAL_CTL, bStatus)            ''ドーパント
        Call CtrlEnabled(.txtHoui, NORMAL_CTL, bStatus)            ''方位
        Call CtrlEnabled(.txtTeikoKbn, NORMAL_CTL, bStatus)        ''抵抗率（下限値参照欄）
        Call CtrlEnabled(.txtSansoKbn, NORMAL_CTL, bStatus)        ''酸素濃度（下限値参照欄）
        For i = 0 To 1
            Call CtrlEnabled(.txtSansoKbn, NORMAL_CTL, bStatus)        ''酸素濃度
            Call CtrlEnabled(.txtHinban(i), NORMAL_CTL, bStatus)       ''品番
            'Call CtrlEnabled(.txtMukaisaki(i), NORMAL_CTL, bStatus)    ''向先
            'Call CtrlEnabled(.txtNoki(i), NORMAL_CTL, bStatus)         ''納期
            Call CtrlEnabled(.txtGoki(i), NORMAL_CTL, bStatus)         ''号機
            Call CtrlEnabled(.txtChokkei(i), NORMAL_CTL, bStatus)      ''直径区分
            Call CtrlEnabled(.txtTeikouritsu(i), NORMAL_CTL, bStatus)  ''抵抗率
            Call CtrlEnabled(.txtTeikou(i), NORMAL_CTL, bStatus)       ''レンジ
            Call CtrlEnabled(.txtSanso(i), NORMAL_CTL, bStatus)        ''酸素濃度
            Call CtrlEnabled(.txtOi(i), NORMAL_CTL, bStatus)           ''Oiレンジ
            Call CtrlEnabled(.txtOrg(i), NORMAL_CTL, bStatus)          ''ORG
        Next i
        For i = 0 To 2
            Call CtrlEnabled(.txtHikiageX(i), NORMAL_CTL, bStatus)     ''引上方法
        Next i
        For i = 0 To 5
            Call CtrlEnabled(.txtMokuteki(i), NORMAL_CTL, bStatus)     ''使用目的
        Next i
    End With
End Sub

' @(f)
' 機能      :   機種・引上方法を日本語文字列変換
' 返り値    :　 TRUE ：正常
'               FALSE：異常
' 引数      :   iKbn：処理区分  1:機種・引上方法
'                               2:機種・号機
'                               3:機種・号機・品番
'               sCds：(IN)機種コード＆？コード[＆品番]
'               sStr：(OUT)機種名＆？コード[＆品番]
'
' 機能説明  :　 機種＆引上方法／号機No[＆品番]から管理コードの日本語に文字列変換し、戻す。
'
' 備考      :   '2000/08/15 小川 この処理を追加した。

Public Function ChgKisyuStr(ByVal iKbn As Integer, ByVal sCds As String, ByRef sStr As String) As Boolean
    Static bReadad As Boolean
    Static Kisyu() As CDNAMEDAT
    Static Pumeth() As CDNAMEDAT
    
    Dim sKisyu As String
    Dim sCd1 As String
    Dim sCd2 As String
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''ダイナセット
    Dim iIdx As Integer
    Dim wk_Hinb As String
    
'    Debug.Print "コード：" & sCds
    ChgKisyuStr = False
    
    ''機種・引上方法をまだ取得してなければ
    If Not bReadad Then
        ''機種名取得処理SQL文作成
        sSQL = "SELECT NVL(  codea9,   ' '), "   ''個別コード
        sSQL = sSQL & "NVL(  namesja9, ' ')  "   ''コード名（日本語短縮）
        sSQL = sSQL & "FROM  koda9           "
        sSQL = sSQL & "WHERE shuca9 = '44'   "
        sSQL = sSQL & "  AND sysca9 = 'K'    "
        ''ダイナセット作成（検索実行）
        If DynSet(objOraDyn, sSQL) = False Then
            ''ダイナセット作成失敗
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''処理中止
            Exit Function
        End If
        ''再配置
        ReDim Kisyu(objOraDyn.RecordCount)
        ''取得確認
        iIdx = 0
        Do While Not objOraDyn.EOF
            With Kisyu(iIdx)
                .Cd = Trim(objOraDyn(0))    ''機種コード
                .Nm = Trim(objOraDyn(1))    ''機種名取得
            End With
            iIdx = iIdx + 1
            objOraDyn.MoveNext
        Loop
        
        ''引上方法取得処理SQL文作成
        sSQL = "SELECT NVL(  codea9,  ' '), "            ''コード名（日本語）
        sSQL = sSQL & "NVL(  nameja9, ' ')  "
        sSQL = sSQL & "FROM  koda9          "
        sSQL = sSQL & "WHERE shuca9 = '51'  "
        sSQL = sSQL & "  AND sysca9 = 'X'   "
        ''ダイナセット作成（検索実行）
        If DynSet(objOraDyn, sSQL) = False Then
            ''ダイナセット作成失敗
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''処理中止
            Exit Function
        End If
        ReDim Pumeth(objOraDyn.RecordCount)
        ''取得確認
        iIdx = 0
        Do While Not objOraDyn.EOF
            With Pumeth(iIdx)
                .Cd = Trim(objOraDyn(0))    ''機種コード
                .Nm = Trim(objOraDyn(1))    ''機種名取得
            End With
            iIdx = iIdx + 1
            objOraDyn.MoveNext
        Loop
        
        ''読込んだ
        bReadad = True
    End If
    
    ''コード切り出し
    sKisyu = Left(sCds, 2)      ''機種　切り出し
    Select Case iKbn
    Case 1                      ''機種別選択時は
        sCd1 = Mid(sCds, 3, 1)  ''引上方法 切り出し
        sCd2 = ""
    Case 2                      ''号機別選択時は
        sCd1 = Mid(sCds, 3)     ''号機 切り出し
        sCd2 = ""
    Case 3                      ''号機品番別選択時は
        sCd1 = Mid(sCds, 3, 3)  ''号機 切り出し
        sCd2 = Mid(sCds, 6)     ''品番 切り出し
    Case 5                      ''機種別選択時は
        sCd1 = Mid(sCds, 4, 1)  ''引上方法 切り出し
        sCd2 = ""
    Case 6                      ''号機別選択時は
        sCd1 = Mid(sCds, 4)     ''号機 切り出し
        sCd2 = ""
    End Select
    
    ''機種検索・名前取得
    For iIdx = 0 To UBound(Kisyu)
        If Kisyu(iIdx).Cd = sKisyu Then sKisyu = Kisyu(iIdx).Nm: Exit For
    Next
    If (iKbn And 3) = 1 Then    ''引上方法なら
        ''引上方法検索・名前取得
        For iIdx = 0 To UBound(Pumeth)
            If Pumeth(iIdx).Cd = sCd1 Then sCd1 = Pumeth(iIdx).Nm: Exit For
        Next
    End If
    
    ''戻し
    'sStr = sKisyu & " " & sCd1 & " " & Format(sCd2, "!@@@-@@@@-@@@@") & vbTab
    If GetHinbanHensyu(Trim(sCd2), 1, wk_Hinb) = True Then
        sStr = sKisyu & " " & sCd1 & " " & wk_Hinb & Chr(9) ' vbTab
    End If
    ChgKisyuStr = True
End Function

' @(f)
' 機能      :   製品コードから名称を取得
' 返り値    :　 TRUE ：正常
'               FALSE：異常
' 引数      :   製品コード
' 機能説明  :　 製品コードから名称を取得

Public Function GetGreadName(ByVal sCds As String, ByRef sStr As String) As Boolean
    Dim sName As String
    Dim sCd1 As String
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''ダイナセット
    
    GetGreadName = False
    
    ''製品コード名取得処理SQL文作成
    sSQL = "SELECT NVL(namjls9,' ') "   ''名称
    sSQL = sSQL & "FROM  sods9           "
    sSQL = sSQL & "WHERE nams9 = 'COMPRDGRS1'   "
    sSQL = sSQL & "  AND clss9 = '01'    "
    sSQL = sSQL & "  AND vals9 = '" & sCds & "'    "
    ''ダイナセット作成（検索実行）
    If DynSet(objOraDyn, sSQL) = False Then
        ''ダイナセット作成失敗
        Call MsgOut(100, "", ERR_DISP_LOG, "sods9")
        ''処理中止
        Exit Function
    End If
    If objOraDyn.EOF = True Then   '4/19 Yam 追加
        sStr = " " & Chr(9)
        GetGreadName = True
        Exit Function
    End If

    sName = objOraDyn(0)
    
    sStr = sCds & " " & sName & Chr(9)
    
    GetGreadName = True
End Function

' @(f)
' 機能      :   機種コード変換
' 返り値    :　 TRUE ：正常
'               FALSE：異常
' 機能説明  :　 機種管理コードから個別コードに変換し、戻す。
' 備考      :   '2000/11/17 Yam

Public Function GetkisyNo(ByVal bkisy As String, ByRef akisy As String) As Boolean
    Dim sSQL As String
    Dim objOraDyn As OraDynaset             ''ダイナセット
    
    GetkisyNo = False
        ''機種名取得処理SQL文作成
        sSQL = "SELECT NVL(  codea9,   ' ') "   ''個別コード
        sSQL = sSQL & "FROM  koda9           "
        sSQL = sSQL & "WHERE shuca9 = '44'   "
        sSQL = sSQL & "  AND sysca9 = 'K'    "
        sSQL = sSQL & "  AND kcodea9 =  '" & Trim(bkisy) & "'"
        Debug.Print sSQL
        ''ダイナセット作成（検索実行）
        If DynSet(objOraDyn, sSQL) = False Then
            ''ダイナセット作成失敗
            Call MsgOut(100, "", ERR_DISP_LOG, "KODA9")
            ''処理中止
            Exit Function
        End If
        If objOraDyn.EOF = True Then
            Exit Function
        End If
         
        akisy = objOraDyn(0)    ''個別コード
                
    GetkisyNo = True
End Function
